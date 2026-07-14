/**
 * れんらくアプリ自動保存
 *
 * れんらくアプリ（BusCatch）の通知メールを Gmail から探し、
 * パスワード認証つきのお知らせページにアクセスして、本文と添付ファイルを Google Drive に保存します。
 * 必要に応じて、下の CONFIG を変更してください。
 */

const CONFIG = {
  // Google Driveに作る保存先フォルダ名です。
  // 下の DRIVE_FOLDER_ID が空欄のときだけ使われます。
  // マイドライブ直下に同じ名前のフォルダがなければ、自動で作成します。
  // 例: 'renraku-app', '学校プリント', '幼稚園のお知らせ'
  DRIVE_FOLDER_NAME: 'renraku-app',

  // 既存フォルダに保存したい場合だけ、Google DriveのフォルダIDを入れてください。
  // よく分からなければ空欄のままでOKです。
  // 空欄の場合は、上の DRIVE_FOLDER_NAME のフォルダに保存します。
  // 例: '1abcDEFghijk...' のような文字列
  //
  // ここに直接書く代わりに、スクリプト プロパティに DRIVE_FOLDER_ID を設定しても上書きできます
  // （例: e-msg版と同じフォルダを指定したいが、フォルダIDをコードに残したくない場合）。
  DRIVE_FOLDER_ID: '',

  // Gmailの検索条件です。通常は変更不要です。
  // れんらくアプリから届くメールの送信元ドメインで検索します。
  GMAIL_QUERY: 'from:(@buscatch.net)',

  // 処理済み・失敗を管理するGmailラベルです。
  // 同じメールを何度も保存しないために使います。通常は変更不要です。
  PROCESSED_LABEL: 'renraku-app-to-GoogleDrive/saved',
  FAILED_LABEL: 'renraku-app-to-GoogleDrive/failed',

  // フォルダ名やログの日付に使うタイムゾーンです。
  // 日本で使う場合はこのままでOKです。
  TIMEZONE: 'Asia/Tokyo',

  // 1回の実行で処理する最大スレッド数です。
  // たくさん未処理メールがある場合でも、1回で処理しすぎないようにしています。
  // 通常は変更不要です。
  MAX_THREADS_PER_RUN: 50,

  // LINEへの転送を行うかどうかです。
  // false にすると、保存は行うがLINE送信だけスキップします。
  // スクリプト プロパティに LINE_TOKEN / LINE_GROUP_ID が未設定の場合も、自動でスキップされます。
  LINE_ENABLED: true,

  // LINEメッセージ本文の最大文字数です。通常は変更不要です。
  LINE_BODY_MAX_CHARS: 3500,
};

/**
 * まずはこの関数を選んで「実行」してください。
 * 新しい れんらくアプリ のメールがあれば、Google Drive に保存します。
 *
 * 実行する前に、スクリプト プロパティに LOGIN_PASSWORD を設定してください。
 * （詳しい手順は README.md を参照してください）
 */
function 今すぐ保存する() {
  return renrakuAppAutoSaver.saveNow();
}

/**
 * 動作確認ができたあとに、この関数を1回だけ実行してください。
 * 以後、1時間ごとに自動で「今すぐ保存する」が実行されます。
 */
function 初回に1回だけ実行する_自動保存を開始() {
  renrakuAppAutoSaver.startAutoSave();
}

/**
 * LINE連携の動作確認用です。テストメッセージを1件送信します。
 * スクリプト プロパティに LINE_TOKEN / LINE_GROUP_ID を設定してから実行してください。
 */
function LINE送信をテストする() {
  renrakuAppAutoSaver.testLine();
}

const renrakuAppAutoSaver = (() => {
  // メール本文に含まれる、お知らせ確認ページへのリンクを探すためのパターンです。
  const MAIL_LINK_PATTERN = /https:\/\/buscatch\.net\/mobile\/[A-Za-z0-9_-]+\/open_confirm_mail\/open\/\?[^\s"'<>]+/g;

  // お知らせページの中から、添付ファイル（PDF・画像など）のURLを探すためのパターンです。
  const ATTACHMENT_URL_PATTERN = /https:\/\/img\.[A-Za-z0-9.-]*buscatch\.net\/\d+\/mail\/img\/[^\s"'<>]+/g;

  function saveNow() {
    const config = getConfig();

    if (!config.loginPassword) {
      throw new Error(
        'スクリプト プロパティに LOGIN_PASSWORD が設定されていません。' +
          'プロジェクトの設定 > スクリプト プロパティ から設定してください。（README.md 参照）'
      );
    }

    const processedLabel = getOrCreateLabel(config.processedLabel);
    const failedLabel = getOrCreateLabel(config.failedLabel);
    const query = buildGmailQuery(config);
    const threads = GmailApp.search(query, 0, config.maxThreadsPerRun);

    // 古いメールから順番に保存します。
    threads.sort((a, b) => a.getLastMessageDate() - b.getLastMessageDate());
    Logger.log(`処理対象スレッド: ${threads.length}`);

    const parentFolder = getSaveFolder(config);
    let success = 0;
    let failed = 0;

    threads.forEach((thread) => {
      let threadFailed = false;

      thread.getMessages().forEach((message) => {
        try {
          processMessage(message, parentFolder, config);
          success++;
        } catch (error) {
          Logger.log(`ERROR: ${message.getSubject()} - ${error.stack || error}`);
          failed++;
          threadFailed = true;
        }
      });

      thread.addLabel(threadFailed ? failedLabel : processedLabel);
    });

    Logger.log(`完了: 成功 ${success} 件, 失敗 ${failed} 件`);
    return { success: success, failed: failed };
  }

  function startAutoSave() {
    // 二重登録を防ぐため、既存の自動実行設定を消してから作り直します。
    removeExistingTriggers();
    ScriptApp.newTrigger('今すぐ保存する').timeBased().everyHours(1).create();
    Logger.log('自動保存を開始しました: 1時間ごとに確認します');
  }

  function removeExistingTriggers() {
    ScriptApp.getProjectTriggers().forEach((trigger) => {
      const handlerName = trigger.getHandlerFunction();
      if (handlerName === '今すぐ保存する') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }

  function getConfig() {
    return {
      // スクリプト プロパティに DRIVE_FOLDER_ID があれば、CONFIG より優先します。
      driveFolderId: PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID') || CONFIG.DRIVE_FOLDER_ID,
      driveFolderName: CONFIG.DRIVE_FOLDER_NAME,
      gmailQuery: CONFIG.GMAIL_QUERY,
      processedLabel: CONFIG.PROCESSED_LABEL,
      failedLabel: CONFIG.FAILED_LABEL,
      timezone: CONFIG.TIMEZONE,
      maxThreadsPerRun: parsePositiveInteger(CONFIG.MAX_THREADS_PER_RUN, 50),
      // パスワードはコードに書かず、スクリプト プロパティから読み込みます。
      loginPassword: PropertiesService.getScriptProperties().getProperty('LOGIN_PASSWORD') || '',
      lineEnabled: CONFIG.LINE_ENABLED,
      lineBodyMaxChars: parsePositiveInteger(CONFIG.LINE_BODY_MAX_CHARS, 3500),
      // LINEのトークン・グループIDもコードに書かず、スクリプト プロパティから読み込みます。
      lineToken: PropertiesService.getScriptProperties().getProperty('LINE_TOKEN') || '',
      lineGroupId: PropertiesService.getScriptProperties().getProperty('LINE_GROUP_ID') || '',
    };
  }

  function parsePositiveInteger(value, fallback) {
    value = Number(value);
    return Number.isFinite(value) && value > 0 ? Math.floor(value) : fallback;
  }

  function buildGmailQuery(config) {
    // 処理済み/失敗ラベルが付いているメールは検索対象から外します。
    return [
      config.gmailQuery,
      `-label:"${config.processedLabel}"`,
      `-label:"${config.failedLabel}"`,
    ].join(' ');
  }

  function getSaveFolder(config) {
    // DRIVE_FOLDER_ID が指定されている場合は、そのフォルダへ保存します。
    if (config.driveFolderId) {
      return DriveApp.getFolderById(config.driveFolderId);
    }

    // 指定がない場合は、DRIVE_FOLDER_NAME のフォルダを探します。
    const folders = DriveApp.getFoldersByName(config.driveFolderName);
    if (folders.hasNext()) {
      return folders.next();
    }

    // フォルダがまだなければ、マイドライブ直下に作成します。
    Logger.log(`保存先フォルダを作成します: ${config.driveFolderName}`);
    return DriveApp.createFolder(config.driveFolderName);
  }

  function processMessage(message, parentFolder, config) {
    const date = message.getDate();
    const subject = (message.getSubject() || '無題').trim();
    const body = message.getPlainBody() || '';
    const mailLinks = uniqueMatches(body, MAIL_LINK_PATTERN);

    // お知らせ確認ページへのリンクがないメールは保存対象外です。
    if (mailLinks.length === 0) {
      Logger.log(`お知らせ確認ページのURLが見つかりません: ${subject}`);
      return;
    }

    const dateText = Utilities.formatDate(date, config.timezone, 'yyyyMMdd_HHmm');
    const folderName = sanitizeDriveName(`${dateText}_${subject}`);
    const subfolder = parentFolder.createFolder(folderName);

    // 先に添付ファイル（と、お知らせページから取れる本文プレビュー）を取得し、
    // その後にメール本文とあわせて保存します。
    const fileResult = saveLinkedFiles(subfolder, mailLinks, config);
    saveMessageBody(subfolder, message, subject, date, body, fileResult.bodyPreview, config);

    Logger.log(`保存完了: ${subject} / 添付ファイル ${fileResult.saved} 件`);

    // LINEに送る内容は、LINE_ENABLEDがfalseでも組み立ててログに出します。
    // 実行ログでプレビューしてから、送信を有効にするかどうか判断できるようにするためです。
    const lineMessage = buildLineMessage({
      subject: subject,
      date: date,
      folderUrl: subfolder.getUrl(),
      attachmentCount: fileResult.saved,
      bodyPreview: fileResult.bodyPreview,
    });
    Logger.log(`[renraku-app][LINEプレビュー] ${subject}\n${lineMessage}`);

    if (config.lineEnabled) {
      pushToLine(lineMessage, config);
    }
  }

  function saveMessageBody(subfolder, message, subject, date, body, bodyPreview, config) {
    const bodyHeader =
      `件名: ${subject}\n` +
      `送信日時: ${Utilities.formatDate(date, config.timezone, 'yyyy/MM/dd HH:mm')} (${config.timezone})\n` +
      `送信者: ${message.getFrom()}\n` +
      '------------------------------\n\n';

    // お知らせページから本文が取れた場合は、そちらを主に、メールの定型文は参考として末尾に添えます。
    const content = bodyPreview
      ? `${bodyHeader}${bodyPreview}\n\n------------------------------\n（メール本文）\n${body}`
      : bodyHeader + body;

    subfolder.createFile('本文.txt', content, MimeType.PLAIN_TEXT);
  }

  function saveLinkedFiles(subfolder, mailLinks, config) {
    let saved = 0;
    let bodyPreview = '';

    mailLinks.forEach((mailLink) => {
      const result = fetchAttachmentUrls(mailLink, config.loginPassword);

      if (!bodyPreview && result.bodyPreview) {
        bodyPreview = result.bodyPreview;
      }

      if (result.fileUrls.length === 0) {
        Logger.log(`添付ファイルが見つかりません: ${mailLink}`);
        return;
      }

      result.fileUrls.forEach((url) => {
        // 見つかった添付ファイルを1つずつ取得してDriveに保存します。
        const fileResponse = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const fileStatus = fileResponse.getResponseCode();

        if (fileStatus !== 200) {
          Logger.log(`添付ファイル取得失敗 ${url}: ${fileStatus}`);
          return;
        }

        const filename = sanitizeDriveName(getFilenameFromUrl(url) || `添付${saved + 1}${getExtensionFromUrl(url)}`);
        subfolder.createFile(fileResponse.getBlob().setName(filename));
        saved++;
      });
    });

    return { saved: saved, bodyPreview: bodyPreview };
  }

  /**
   * LINEグループにテキストメッセージをプッシュ送信します。
   * スクリプト プロパティに LINE_TOKEN と LINE_GROUP_ID が必要です。
   * 未設定の場合は、保存処理を止めずに警告ログだけ出してスキップします。
   */
  function pushToLine(text, config) {
    if (!config.lineToken || !config.lineGroupId) {
      Logger.log('LINE未設定: スクリプト プロパティに LINE_TOKEN / LINE_GROUP_ID を設定してください（今回はスキップ）');
      return;
    }

    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${config.lineToken}` },
      payload: JSON.stringify({
        to: config.lineGroupId,
        messages: [{ type: 'text', text: text }],
      }),
      muteHttpExceptions: true,
    });

    const status = response.getResponseCode();
    if (status >= 300) {
      Logger.log(`LINE送信失敗 (${status}): ${response.getContentText()}`);
    }
  }

  /**
   * LINEに送るテキストメッセージを組み立てます。
   */
  function buildLineMessage({ subject, date, folderUrl, attachmentCount, bodyPreview }) {
    const lines = [];
    lines.push(subject);
    lines.push(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm'));

    if (bodyPreview) {
      lines.push('');
      let text = removeSubjectDuplicateLine(bodyPreview, subject);
      if (text.length > CONFIG.LINE_BODY_MAX_CHARS) {
        text = `${text.substring(0, CONFIG.LINE_BODY_MAX_CHARS)}…(以下省略・Driveで全文確認)`;
      }
      lines.push(text);
    }

    lines.push('');
    lines.push(attachmentCount > 0 ? `📎 添付ファイル ${attachmentCount} 件` : '本文のみ（添付ファイルなし）');
    lines.push(`🔗 ${folderUrl}`);
    return lines.join('\n');
  }

  /**
   * 本文プレビューの先頭行が件名とほぼ同じ内容（「ただいま」等の前置きや句点の有無を除いて重複）の場合、
   * LINEの表示上はその1行を取り除きます（件名がすでに1行目に出ているため）。
   */
  function removeSubjectDuplicateLine(text, subject) {
    if (!text || !subject) {
      return text;
    }

    const lines = text.split('\n');
    const firstLine = (lines[0] || '').replace(/[。、\s]+$/, '').trim();
    const normalizedSubject = subject.trim();

    if (!firstLine || (!normalizedSubject.includes(firstLine) && !firstLine.includes(normalizedSubject))) {
      return text;
    }

    lines.shift();
    while (lines.length > 0 && lines[0].trim() === '') {
      lines.shift();
    }
    return lines.join('\n');
  }

  /**
   * お知らせ確認ページのURLにアクセスし、パスワード認証を行った上で、
   * ページ内から添付ファイルのURL一覧を取得します。
   *
   * UrlFetchAppの自動リダイレクト追跡（followRedirects既定値true）は、
   * リダイレクト先の最終レスポンスのヘッダーしか見えず、
   * リダイレクトの途中で発行されるセッションCookieを取りこぼすことがあります。
   * そのため、ここではリダイレクトを1段階ずつ自前で追いながらCookieを引き継ぎます。
   */
  function fetchAttachmentUrls(mailLink, loginPassword) {
    const linkParts = parseMailLink(mailLink);
    if (!linkParts) {
      throw new Error(`お知らせ確認ページのURLを解析できませんでした: ${mailLink}`);
    }

    const cookieJar = createCookieJar();

    // 1. お知らせ確認ページへアクセスし、セッションを開始します（パスワード入力画面が返ってきます）。
    const firstResponse = fetchFollowingRedirects(mailLink, { muteHttpExceptions: true }, cookieJar);
    Logger.log(
      `[renraku-app] 初回アクセス status=${firstResponse.getResponseCode()} cookie=${cookieJar.header() ? 'あり' : 'なし'}`
    );

    // 2. パスワードを送信して認証します。
    //    このレスポンス自体には本文は乗っておらず、ログイン後にもう一度同じリンクを開き直すと
    //    本文ページが表示される、という挙動が確認されているため、ここでは中身を使いません。
    const loginUrl = `https://buscatch.net/mobile/${linkParts.schoolCode}/certifications/confirm_password/`;
    const loginResponse = fetchFollowingRedirects(
      loginUrl,
      {
        method: 'post',
        payload: { password: loginPassword, u: linkParts.u, s: linkParts.s },
        muteHttpExceptions: true,
      },
      cookieJar
    );
    Logger.log(`[renraku-app] ログインPOST status=${loginResponse.getResponseCode()}`);

    // 3. 認証済みのセッションで、あらためて同じお知らせリンクへアクセスします。
    //    ここで返ってくるページに、本文と添付ファイルが含まれます。
    const contentResponse = fetchFollowingRedirects(mailLink, { muteHttpExceptions: true }, cookieJar);
    const html = contentResponse.getContentText('UTF-8');
    Logger.log(`[renraku-app] 本文ページ status=${contentResponse.getResponseCode()} 本文長=${html.length}`);

    if (isLoginFailed(html)) {
      throw new Error(
        'パスワード認証に失敗しました。スクリプト プロパティの LOGIN_PASSWORD が正しいか確認してください。'
      );
    }

    // 4. 本文ページから、添付ファイルのURLと本文プレビューを探します。
    const fileUrls = uniqueMatches(html, ATTACHMENT_URL_PATTERN);
    if (fileUrls.length === 0) {
      Logger.log(`[renraku-app] 添付ファイルURLが見つかりません。HTML先頭400文字: ${html.substring(0, 400)}`);
    }

    return { fileUrls: fileUrls, bodyPreview: extractBodyPreview(html) };
  }

  /**
   * お知らせページのHTMLから、本文プレビューを抜き出します。
   * 本文は <dl class="info"> 直下の最初の <dd> に入っており、
   * 添付ファイルがある場合はそれ以降の <dd> に添付リンクが続きます。
   */
  function extractBodyPreview(html) {
    const match = (html || '').match(/<dl class="info">\s*<dd>([\s\S]*?)<\/dd>/);
    if (!match) {
      return '';
    }

    let text = match[1];
    text = text.replace(/<br\s*\/?>/gi, '\n');
    text = text.replace(/<[^>]+>/g, '');
    text = decodeHtml(text);
    text = text.replace(/\n{3,}/g, '\n\n');
    return text.trim();
  }

  /**
   * リダイレクト(3xx)を1段階ずつ自前で追いかけながらfetchします。
   * 各レスポンスのSet-Cookieを毎回cookieJarに反映するため、
   * リダイレクト途中で発行されるセッションCookieも取りこぼしません。
   */
  function fetchFollowingRedirects(url, options, cookieJar, maxRedirects) {
    maxRedirects = maxRedirects || 5;
    let currentUrl = url;
    let currentOptions = options;
    let response;

    for (let i = 0; i <= maxRedirects; i++) {
      const requestOptions = Object.assign({}, currentOptions, {
        followRedirects: false,
        headers: Object.assign({}, currentOptions.headers, { Cookie: cookieJar.header() }),
      });

      response = UrlFetchApp.fetch(currentUrl, requestOptions);
      cookieJar.update(response);

      const status = response.getResponseCode();
      if (status < 300 || status >= 400) {
        break;
      }

      const location = getHeaderCaseInsensitive(response.getAllHeaders(), 'Location');
      if (!location) {
        break;
      }

      // リダイレクト先はGETで取得します（POSTの302/303リダイレクトはブラウザもGETに切り替えるのが一般的です）。
      currentUrl = resolveUrl(location);
      currentOptions = { muteHttpExceptions: true };
    }

    return response;
  }

  function getHeaderCaseInsensitive(headers, name) {
    const key = Object.keys(headers || {}).find((k) => k.toLowerCase() === name.toLowerCase());
    return key ? headers[key] : null;
  }

  function resolveUrl(location) {
    if (/^https?:\/\//i.test(location)) {
      return location;
    }
    if (location.startsWith('/')) {
      return `https://buscatch.net${location}`;
    }
    return location;
  }

  function parseMailLink(mailLink) {
    const match = mailLink.match(/https:\/\/buscatch\.net\/mobile\/([A-Za-z0-9_-]+)\/open_confirm_mail\/open\/\?(.+)/);
    if (!match) {
      return null;
    }

    const params = {};
    match[2].split('&').forEach((pair) => {
      const [key, value] = pair.split('=');
      if (key) {
        params[decodeURIComponent(key)] = decodeURIComponent(value || '');
      }
    });

    if (!params.u || !params.s) {
      return null;
    }

    return { schoolCode: match[1], u: params.u, s: params.s };
  }

  function isLoginFailed(html) {
    // 認証に失敗すると、パスワード入力フォームが再度表示されます。
    return /name="password"/.test(html || '');
  }

  function createCookieJar() {
    const store = {};

    return {
      update(response) {
        const headers = response.getAllHeaders();
        let setCookie = headers['Set-Cookie'];
        if (!setCookie) {
          return;
        }
        if (!Array.isArray(setCookie)) {
          setCookie = [setCookie];
        }
        setCookie.forEach((line) => {
          const pair = line.split(';')[0];
          const separatorIndex = pair.indexOf('=');
          if (separatorIndex > -1) {
            const name = pair.substring(0, separatorIndex).trim();
            const value = pair.substring(separatorIndex + 1).trim();
            store[name] = value;
          }
        });
      },
      header() {
        return Object.keys(store)
          .map((name) => `${name}=${store[name]}`)
          .join('; ');
      },
    };
  }

  function decodeHtml(text) {
    return String(text || '')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&nbsp;/g, ' ');
  }

  function sanitizeDriveName(name) {
    // Driveで見づらくなる文字を整えて、長すぎる名前を短くします。
    return String(name || '')
      .replace(/[\/\\]/g, '_')
      .replace(/\s+/g, ' ')
      .substring(0, 200)
      .trim();
  }

  function getFilenameFromUrl(url) {
    const match = String(url || '').match(/\/([^\/?#]+)(?:[?#].*)?$/);
    return match ? decodeURIComponent(match[1]) : '';
  }

  function getExtensionFromUrl(url) {
    const match = String(url || '').match(/\.([A-Za-z0-9]+)(?:[?#].*)?$/);
    return match ? `.${match[1]}` : '';
  }

  function uniqueMatches(text, regex) {
    const matches = [];
    const seen = {};
    let match;

    regex.lastIndex = 0;
    while ((match = regex.exec(text || '')) !== null) {
      const value = match[0];
      if (!seen[value]) {
        seen[value] = true;
        matches.push(value);
      }
    }

    return matches;
  }

  function getOrCreateLabel(name) {
    let label = GmailApp.getUserLabelByName(name);
    if (!label) {
      label = GmailApp.createLabel(name);
    }
    return label;
  }

  function testLine() {
    const config = getConfig();
    pushToLine('テスト送信です（renraku-app-to-GoogleDrive）', config);
  }

  return {
    saveNow: saveNow,
    startAutoSave: startAutoSave,
    testLine: testLine,
  };
})();

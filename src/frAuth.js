/******************************************************************
 * GAS　+　OAush2ライブラリ を用いて認証機能を実装し、freee　APIと連携する
******************************************************************/
// 参考サイト: https://qiita.com/TakeshiNickOsanai/items/62810b0e96bf37bd0eca

/*
参照ライブラリ
title        |OAuth2 (サーバー、アプリ、ユーザー間で安全なAPI利用ができるように認可の仕組みを標準化したライブラリ)
project_key  |1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
*/

// const Properties = PropertiesService.getScriptProperties();
const oAuth2Service = getFreeeOAuth2Service();


/** freee API　へのアクセス管理用の　OAuth2.0　サービスを取得する関数
 *  処理内容:
 *   - OAuth2サービスを作成し、 freeeAPIの認証エンドポイントとトークンエンドポイントを設定
 *   - ClientID & ClientSecret、コールバック関数を設定
 *   - プロパティストアをユーザープロパティに設定して認証情報を保存 
 *     (ユーザーごとに認証情報を管理でき、再認証の必要性を減らすことができる)
 */
function getFreeeOAuth2Service() {
  return OAuth2.createService('freeeAPI') 
  .setAuthorizationBaseUrl(
    'https://accounts.secure.freee.co.jp/public_api/authorize'
  )
  .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
  .setClientId(
    Properties.getProperty('FREEE_CLIENT_ID')
  )
  .setClientSecret(
    Properties.getProperty('FREEE_CLIENT_SECRET')
  )
  .setCallbackFunction('authCallback')
  .setPropertyStore(PropertiesService.getUserProperties());
}

/** 
 * モードレスダイアログを作成する関数 
 * モードレスダイアログ: 
 *   - showModelessDialog()を使用して表示させる
 *   - 画面上で他の要素と干渉せずに表示される
 *   - ダイアログが表示されていても、ユーザーはアプリの他の部分とやり取りが可能
 * 通常のダイアログ:
 *   - showModalDialog()、または、組み込みのブラウザダイアログ(alert(),confirm(), prompt()など)を使用して表示させる
 *   - 画面上でモーダルとして表示される
 *   - ダイアログが表示されている間は、ユーザーは他の操作を行えない
 */
function createModelessDialog(html, title) {
  // HTMLを作成してダイアログを表示
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(360)
    .setHeight(120);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}


/** freee API へのアクセスを許可するための認可処理を表示する関数 */
function showAuth() {
  // freee APIへのアクセス権を持っていない場合
  if (!oAuth2Service.hasAccess()) {
    // 認可用URLを取得
    const authorizationUrl = oAuth2Service.getAuthorizationUrl();
    // ダイアログのテンプレートを作成して表示
    const template = HtmlService.createTemplate(
      '<p>freee APIへのアクセスを許可するために認証を行ってください。</p>' +
      '<p>認証は<a href="<?= authorizationUrl ?>" target="_blank">こちらをクリック</a></p>'
    );
    template.authorizationUrl = authorizationUrl;
    const page = template.evaluate();
    const title = 'freeeアプリの認可処理';
    createModelessDialog(page, title);
  } else {
    showUser();
  } 
}


/** 認可コールバックを処理する関数 */
function authCallback( request ) {
  // 認可処理を行い、結果を返す
  const isAuthorized = oAuth2Service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証に成功しました。\nこのタブは閉じて大丈夫です。');
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。\nこのタブを閉じて、もう一度認証処理を実行してみてください。');
  }
}


/** ユーザー情報を表示する関数 */
function showUser() {
  const response = UrlFetchApp.fetch(
    'https://api.freee.co.jp/api/1/users/me',
    {
      headers: {
        Authorization: 'Bearer ' + oAuth2Service.getAccessToken(),
      },
    }
  );
  const myJson = JSON.parse(response);
  // const message = 'OAuth認証済みです。\\n認可済ユーザー名: ' + myJson.user.display_name;
  const message = 'OAuth認証済みです。';
  Browser.msgBox(message);
}

// /** ログアウト処理を行う関数 */
// function logout() {
//   oAuth2Service.reset();
//   const message = 'freeeアプリからログアウトしました。';
//   const logoutTitle = 'ログアウト終了';

//   createModelessDialog(message, logoutTitle);
// }

/** 定数定義 */
// LINEのイベントタイプ
const EVENT_TYPES = {
  MESSAGE: "message", //　メッセージイベント
  FOLLOW: "follow"    // 友達追加イベント
};

// LINE応答メッセージ発動ワード
const MESSAGES = {
  ORDER: "注文",
  NO: "いいえ",
  QUESTION: "お問い合わせ"
};

// テキストレスポンス
const TEXT_REPLY = {
  ORDER_CONFIRM: "注文を開始してよろしいですか？",
  CANCEL: "承知しました。\n何かございましたら、お気軽にお問い合わせください。",
  ANSWER: "通常のLINEメッセージ機能からお問い合わせ可能です。"
};



/** LINEからのリクエスト処理 (Webhook) */
function doPost(e) {
  try {
    // 受信データをパース
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const body = JSON.parse(raw);
    const events = Array.isArray(body.events) ? body.events : [];

    // イベントが無い場合も LINE に200を返す
    if (events.length === 0) {
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    // 複数イベントを順に処理
    events.forEach(event => {
      try {
        const eventType = event.type;
        const replyToken = event.replyToken;
        const lineUserId = event.source && event.source.userId;

        switch (eventType) {
          case EVENT_TYPES.MESSAGE:
            handleMessageEvent(event, lineUserId, replyToken);
            break;
          case EVENT_TYPES.FOLLOW:
            handleFollowEvent(lineUserId);
            break;
          default:
            // 未対応のイベントは無視
            console.log(`Unhandled event type: ${eventType}`);
            break;
        }
      } catch (innerErr) {
        // 個別イベントの処理失敗は全体に影響させない
        console.error(`Event handling error: ${innerErr.message}`);
      }
    });

    // 常に200を返す（LINEは本文は見ない）
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    console.error(`doPost error: ${error.message}`);
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  }
}



/** 
 * 内部ヘルパー関数
 */

/** メッセージイベントごとの処理 */
const handleMessageEvent = (event, lineUserId, replyToken) => {
  try {
    const userMessage = (event.message && event.message.text) ? event.message.text.trim() : "";

    // ユーザーからのメッセージが、配列内のワードのどれにも該当しないとき
    if (![MESSAGES.ORDER, MESSAGES.NO, MESSAGES.QUESTION].includes(userMessage)) { 
      return;
    }

    // 管理者登録ワード確認（存在するなら先に処理）
     try {
      const keyword = getAdminRegisterKeyword();
      if (keyword && userMessage === keyword) {
        registerAdminAccount(lineUserId);
        if (replyToken) {
          sendReplyMessage(replyToken, [{ type: "text", text: "管理者アカウントとして登録しました。" }]);
        }
        return;
      }
    } catch (e) {
      console.warn(`Admin keyword check skipped: ${e.message}`);
    }

  if (userMessage === MESSAGES.ORDER) {
      handleOrderEvent(lineUserId, replyToken);
    } else if (userMessage === MESSAGES.NO && replyToken) {
      sendReplyMessage(replyToken, [{ type: "text", text: TEXT_REPLY.CANCEL }]);
    } else if (userMessage === MESSAGES.QUESTION && replyToken) {
      sendReplyMessage(replyToken, [{ type: "text", text: TEXT_REPLY.ANSWER }]);
    }
  } catch (error) {
    throw new Error(`Error in handleMessageEvent: ${error.message}`);
  } 
};


/** 注文開始処理 */
const handleOrderEvent = (lineUserId, replyToken) => {
  try {
    if (!lineUserId || !replyToken) return;

    const now = new Date();
    const orderId = now.getTime().toString(16);  // UNIXタイムスタンプを16進数表記に変換
    // debug(`1.注文ID生成: ${orderId}`);
        
    // 注文IDをキーにして、LINEユーザーIDをキャッシュデータに保存
    cache = makeCache();  // キャッシュを初期化
    cache.put(orderId, lineUserId, 3600); // 有効期間は１時間
    // debug(`3.キャッシュ完了: ${cache.get(orderId)}`);

    // キャッシュデータをcacheシートにもバックアップ
    const lastRow = cacheSheet.getLastRow() + 1;  // キャッシュシートの最終行
    cacheSheet.getRange(lastRow, 1, 1, 3).setValues([[now, orderId, lineUserId]]);  // 日時、注文ID、LINEユーザーIDを格納

    // 注文IDをフォームの初期値にセットして、事前入力された公開用URLを作成する
    const prefilledUrl = generatePrefilledFormUrl(orderId);
    // debug(prefilledUrl);

    const replyMessage =  [
      {
        type: "template",
        altText: "注文受付",
        template: {
          type: "confirm",
          text: TEXT_REPLY.ORDER_CONFIRM,
          actions: [
            { type: "uri", label: "はい", uri: prefilledUrl },
            { type: "message", label: "いいえ", text: MESSAGES.NO }
          ]
        }
      }
    ];
    sendReplyMessage(replyToken, replyMessage);
  } catch (error) {
    throw new Error(`Error in handleOrderEvent: ${error.message}`);
  } 
};


/** フォローイベント処理 */
const handleFollowEvent = (lineUserId) => {
  try {
    if (!lineUserId) return;
    const displayName = getUserProfile(lineUserId);
    customerSheet.appendRow([lineUserId, displayName]);
    customerSheet.getDataRange().removeDuplicates([1]);  // 列（1列目:ID)を指定して重複判定

    // debug(`友だち追加イベント処理完了 : ${displayName}`);
  } catch (error) {
    throw new Error(`Error in handleFollowEvent: ${error.message}`);
  }
}


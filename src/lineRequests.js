/** ===========================
 * LINE Webhook エントリポイント
 * =========================== */

/** 入口: 即時 200 応答（重い処理は置かない） */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const body = JSON.parse(raw);
    const events = Array.isArray(body.events) ? body.events : [];

    // イベントが無い場合も LINE に200を返す
    if (events.length === 0) {
      return createOkResponse();
    }

    events.forEach(ev => {
      try {
        dispatchLineEvent(ev);
      } catch (err) {
        console.error('dispatchLineEvent error:', err && err.stack || err);
      }
    });
    return createOkResponse();
  } catch (err) {
    console.error('doPost global error:', err && err.stack || err);
    return createOkResponse(); // 失敗しても 200 を返す
  }
}


/** 200 OK のテキストレスポンスを生成する。 */
function createOkResponse() {
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

/** LINEイベント種別ごとでハンドラにディスパッチ */
function dispatchLineEvent(event) {
  const t = event && event.type;
  const uid = event && event.source && event.source.userId;
  const replyToken = event && event.replyToken;

  if (t === 'message') return handleMessageEvent(event, uid, replyToken);
  if (t === 'follow')  return handleFollowEvent(uid);
}


/** messageイベントハンドラ */
function handleMessageEvent(event, lineUserId, replyToken) {
  try {
    const text = (event && event.message && event.message.text || '').trim();
    if (!text) return;

    // 1) 管理者登録キーワード一致を優先
    try {
      const keyword = getAdminRegisterKeyword();
      if (keyword && text === keyword && lineUserId) {
        registerAdminAccount(lineUserId);
        if (replyToken) {
          LineMessagingService.reply(replyToken, [
            { type: 'text', text: AppConfig.line.reply.ADMIN_REGISTERED }]);
        }
        return;
      }
    } catch (e) {
      console.warn(`admin keyword unavailable: ${e && e.message}`);
    }

    // 2) コマンド群
    if (text === AppConfig.line.commands.ORDER) {
      return handleOrderEvent(lineUserId, replyToken);
    }
    if (text === AppConfig.line.commands.NO && replyToken) {
      return LineMessagingService.reply(replyToken, [
        { type: 'text', text: AppConfig.line.reply.CANCEL }
      ]);
    }
    if (text === AppConfig.line.commands.QUESTION && replyToken) {
      return LineMessagingService.reply(replyToken, [
        { type: 'text', text: AppConfig.line.reply.ANSWER }
      ]);
    }
  } catch (error) {
    console.error(`handleMessageEvent: ${error && error.message}`);
  }
}


/** 注文開始（事前入力フォームURLを返信） */
function handleOrderEvent(lineUserId, replyToken) {
  try {
    if (!lineUserId || !replyToken) return;

    const orderId = Utilities.getUuid(); // 一意の注文ID

    let prefilledUrl = '';

    // 1) 事前入力URLを生成（内部で「entry.xxxが無い」「注文IDがURLに無い」等ならthrow）
    try {
      prefilledUrl = generatePrefilledFormUrl(orderId);
    } catch (err) {
      // 生成失敗：配布せず、管理者とユーザーに知らせて終了
      notifyToAdmin(`⚠️ 注文フォームのURL生成に失敗: ${err && err.message}\norderId=${orderId}`);
      LineMessagingService.reply(replyToken, [{
        type: 'text',
        text: 'ただいま注文フォームの準備でエラーが発生しました。お手数ですが、1分ほどおいてからもう一度「注文」と送ってください。'
      }]);
      return;
    }

    // 2) フォームURLをユーザーへ提示（confirm テンプレ）
    LineMessagingService.reply(replyToken, [{
      type: 'template',
      altText: '注文受付',
      template: {
        type: 'confirm',
        text: AppConfig.line.reply.ORDER_CONFIRM,
        actions: [
          { type: 'uri', label: 'はい', uri: prefilledUrl },
          { type: 'message', label: AppConfig.line.commands.NO, text: AppConfig.line.commands.NO }
        ]
      }
    }]);

    // 3) URLの配布に成功した場合のみ、orderId ↔ userId を保存（返信紐付け用）
    try {
      const cache = makeCache();
      cache.put(orderId, lineUserId, 3600);
      const last = cacheSheet.getLastRow() + 1;
      cacheSheet.getRange(last, 1, 1, 3).setValues([[new Date(), orderId, lineUserId]]);
    } catch (persistErr) {
      console.warn(`注文キャッシュの永続化（保存）に失敗しました: ${persistErr && persistErr.message}`);
    }
  } catch (error) {
    console.error(`handleOrderEvent: ${error && error.message}`);
  }
}


/** followイベントハンドラ（LINE顧客IDシートへ登録） */
const handleFollowEvent = (lineUserId) => {
  try {
    if (!lineUserId) return;
    const displayName = getUserProfile(lineUserId);
    customerSheet.appendRow([lineUserId, displayName]);
    customerSheet.getDataRange().removeDuplicates([1]);  // 列（1列目:ID)を指定して重複判定

  } catch (error) {
    throw new Error(`Error in handleFollowEvent: ${error.message}`);
  }
}


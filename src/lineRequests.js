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

    // すべてのケースで replyToken は必須
    if (!replyToken) {
      console.warn('handleMessageEvent: replyToken missing; ignore. text=', text);
      return;
    }

    const { ORDER, NO, QUESTION } = AppConfig.line.commands;
    const adminKeyword = getAdminRegisterKeyword(); // null か 文字列

    if (adminKeyword && text === adminKeyword) {
      processAdminRegistration(lineUserId, replyToken);
      return;
    }
    // 受信コマンドで場合分け
    switch (text) { 
      case ORDER:
        return handleOrderEvent(lineUserId, replyToken);

      case NO:
        return LineMessagingService.reply(replyToken, [
          { type: 'text', text: AppConfig.line.reply.CANCEL }
        ]);

      case QUESTION:
        return LineMessagingService.reply(replyToken, [
          { type: 'text', text: AppConfig.line.reply.ANSWER }
        ]);

      default:
        // 未対応メッセージは無視（必要ならヘルプ返信を実装）
        return;
    }
  } catch (error) {
    console.error(`handleMessageEvent: ${error && error.message}`);
  }
}

/** 管理者キーワード受信時の返信処理 */
function processAdminRegistration(lineUserId, replyToken) {
  const result = registerAdminAccount(lineUserId);

  // エラー系
  if (!result || result.ok !== true) {
    LineMessagingService.reply(replyToken, [
      { type: 'text', text: '管理者キーワードが未設定か、または、LINEユーザーIDの取得に失敗しました。\nもう一度お試しください。' }
    ]);
    return;
  }

  // 既存登録
  if (result.already) {
    LineMessagingService.reply(replyToken, [
      { type: 'text', text: '既に登録済です' }
    ]);
    return;
  }

  // 新規登録
  LineMessagingService.reply(replyToken, [
    { type: 'text', text: AppConfig.line.reply.ADMIN_REGISTERED }
  ]);
}


/** 注文開始（事前入力フォームURLを返信） */
function handleOrderEvent(lineUserId, replyToken) {
  try {
    if (!lineUserId) return;

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
        text: 'ただいま注文フォームの送信でエラーが発生しました。お手数ですが、1分ほど待ってからもう一度お試しください。'
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

    // 3) URLの配布に成功した場合のみ、orderId <-> userId を保存（返信紐付け用）
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
    const lineName = getUserProfile(lineUserId);
    customerSheet.appendRow([lineUserId, lineName]);
    customerSheet.getDataRange().removeDuplicates([1]);  // 列（1列目:ID)を指定して重複判定

  } catch (error) {
    throw new Error(`Error in handleFollowEvent: ${error.message}`);
  }
}


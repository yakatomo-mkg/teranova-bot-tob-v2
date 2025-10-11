/** 
 * （友達追加時） LINEユーザーIDをもとに、プロフィール情報を取得してくる関数
 */
function getUserProfile(userId) {
  try {
    // ユーザ情報を取得するための、Messaging APIエンドポイント
    const LINE_GET_PROFILE_URL = `https://api.line.me/v2/bot/profile/${userId}`;
    const res = UrlFetchApp.fetch(LINE_GET_PROFILE_URL, {
      headers: {
        Authorization: `Bearer ${LINE_CHANNEL_TOKEN}`,
      },
    })
    return JSON.parse(res.getContent()).displayName;  // アカウント名を返す
  } catch (error) {
    throw new Error(`Error in getUserProfile: ${error.message}`);
  }

}


/**
 * 「注文ID」の質問に対してorderIdの値をセットし、事前入力フォームURLを返す関数
 * @params  {string} id - orderId
 * @returns {string} url - 事前入力フォームURL 
 */
function generatePrefilledFormUrl(id) {
  try {
    const item = form.getItems()[0];  // 最初の質問を取得
    const itemType = item.getType();  // アイテムのタイプを取得
    console.log(`First item type: ${itemType}`);  // 最初の質問のタイプをデバッグ出力
    console.log(`First item title: ${item.getTitle()}`);  // 最初の質問のタイトルをデバッグ出力
    if (!item && item.getTitle() !== "注文ID") {
      throw new Error("質問タイトル「注文ID」が存在しません。");
    }
    // フォームへの新しい回答を作成
    const formRes = form.createResponse();
    // 上記で取得した item をテキスト質問として扱い、orderIdを用いて回答を作成
    const itemRes = item.asTextItem().createResponse(id);
    // 作成した回答(itemRes)をフォーム回答に追加
    formRes.withItemResponse(itemRes);
    return formRes.toPrefilledUrl();  // 事前入力フォームURL 
  } catch (error) {
    throw new Error(`Error in generatePrefilledFormUrl: ${error.message}`);
  }  
}



/** 
 * スクリプトキャッシュ(データの一時的な保存サービス)を操作するためのヘルパー関数
 */
function makeCache() {
  const cache = CacheService.getScriptCache();  // スクリプトキャッシュのインスタンスを作成
  return {
    // getプロパティ : 指定されたキーに対応する値を取得する
    get: function(key) {
      return JSON.parse(cache.get(key));  // JSオブジェクトにパースして返す
    },

    // putプロパティ : 指定されたキーとvalueをキャッシュに保存する
    put: function(key, value, sec) {
      cache.put(key, JSON.stringify(value), (sec === undefined) ? 600 : sec);  // JSON文字列に変換して保存
      return value;
    },

    // removeプロパティ : 指定されたキーに対応するvalueをキャッシュから削除する   
    remove: function(key) {
      cache.remove(key);
      return true;  // キャッシュデータ削除成功時のの確認用返り値
    }
  };
}


/** 
 * LINEユーザーに応答メッセージを送信する関数
 */
function sendReplyMessage(replyToken, messages) {
  try {
    const LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply";
    const options = {
      method: "post",
      headers: {
        "Content-Type": "application/json; charset=UTF-8",
        Authorization: `Bearer ${LINE_CHANNEL_TOKEN}`,
      },
      payload: JSON.stringify({ replyToken, messages }),
    };
    return UrlFetchApp.fetch(LINE_REPLY_URL, options);

  } catch (error) {
    throw new Error(`Error in sendReplyMessage: ${error.message}`);
  }
};



/** 
 * 引数で指定したLINEユーザーに、pushメッセージを送信する関数
 * (任意のタイミングでメッセージを送信できる)
 */
const sendPushMessage = (lineUserId, message) => {
  try {
    const LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push";
    const postData = {
      to: lineUserId,
      messages: [
        {
          type: "text",
          text: message,
        },
      ],
    };
    const headers = {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: `Bearer ${LINE_CHANNEL_TOKEN}`,
    };
    const options = {
      method: "post",
      headers: headers,
      payload: JSON.stringify(postData),
    };
    const res = UrlFetchApp.fetch(LINE_PUSH_URL, options);
    return res;
  } catch (error) {
    throw new Error(`Error in sendPushMessage: ${error.message}`);
  }
  
}


/**
 * AdminsシートのB列（2列目）に登録されたLINE userIdを取得
 */
function getAdminUserIds() {
  const lastRow = adminSheet.getLastRow();
  if (lastRow < 2) return [];

  return adminSheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(v => v && v.startsWith('U'));
}


/**
 * 管理者へ一斉通知（最大10名想定）
 *  - まず /multicast でまとめて送信
 *  - エラー時は個別 /push でフォールバック
 */
function notifyToAdmin(message) {
  const adminIds = getAdminUserIds();
  if (adminIds.length === 0) {
    console.warn('No admin userIds found in Admins sheet (column B).');
    return;
  }

  try {
    lineMulticast(adminIds, message);
  } catch (e) {
    console.error(`multicast failed, fallback to individual push. Reason: ${e}`);
    adminIds.forEach(uid => {
      try { linePushMessage(uid, message); }
      catch (err) { console.error(`push failed for ${uid}: ${err}`); }
    });
  }
}

/** 個別Push */
function linePushMessage(userId, text) {
  const payload = { to: userId, messages: [{ type: 'text', text }] };
  const res = UrlFetchApp.fetch(LINE_PUSH_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${LINE_CHANNEL_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 400) {
    throw new Error(`LINE push failed: ${code} ${res.getContentText()}`);
  }
}

/** 一斉送信（/multicast） */
function lineMulticast(userIds, text) {
  const url = 'https://api.line.me/v2/bot/message/multicast';
  const payload = { to: userIds, messages: [{ type: 'text', text }] };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${LINE_CHANNEL_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 400) {
    throw new Error(`LINE multicast failed: ${code} ${res.getContentText()}`);
  }
}




/**
 * 管理者登録の合言葉を取得（「管理者アカウント」シートのE2）
 */
function getAdminRegisterKeyword() {
  const keyword = String(adminSheet.getRange('E2').getValue() || '').trim();
  if (!keyword) {
    throw new Error("管理者登録キーワードが未設定です（管理者アカウント!E2）");
  }
  return keyword;
}


/**
 * UID を「管理者アカウント」シートに記録
 */
function registerAdminAccount(userId) {
  const lastRow = adminSheet.getLastRow();
  const startRow = 2; // 1行目はヘッダー

  if (lastRow >= startRow) {
    const ids = adminSheet
      .getRange(startRow, 2, lastRow - startRow + 1, 1)
      .getValues()
      .flat()
      .map(v => String(v || '').trim());

    const idx = ids.findIndex(v => v === userId);
    if (idx !== -1) {
      // 既存の場合は A列の時刻を更新
      adminSheet.getRange(startRow + idx, 1).setValue(new Date());
      return;
    }
  }
  // 未登録なら末尾に追記
  adminSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[new Date(), userId]]);
}


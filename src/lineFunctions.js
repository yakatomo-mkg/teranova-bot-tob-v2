/** ===========================
 * LINE Messaging 送受信サービス
 * =========================== */

/** Messaging APIに渡す messages配列を “最低限送れる形” に整える */
function normalizeLineMessages(messages) {
  if (!Array.isArray(messages)) return [];
  return messages.map(m => {
    if (!m || typeof m !== 'object') return { type: 'text', text: String(m) };
    if (!m.type) return { type: 'text', text: JSON.stringify(m) };
    return m;
  });
}

/** 共通：LINE API 呼び出し（指数バックオフ・リトライ対応） */
function invokeLineApi(url, payload) {
  const { maxAttempts, baseDelayMs, retryStatuses } = AppConfig.line.backoff;
  let attempt = 0;

  while (attempt < maxAttempts) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${LINE_CHANNEL_TOKEN}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (retryStatuses.indexOf(code) === -1) {
      if (code >= 400) throw new Error(`LINE API error ${code}: ${body}`);
      return res;
    }
    attempt += 1;
    if (attempt >= maxAttempts) throw new Error(`LINE API retry exhausted with ${code}: ${body}`);
    Utilities.sleep(baseDelayMs * Math.pow(2, attempt - 1));
  }
}


/** サービス層（reply/push/multicast/管理者通知） */
const LineMessagingService = (() => {
  function reply(replyToken, messages) {
    if (!replyToken) throw new Error('replyToken is required');
    const payload = { replyToken, messages: normalizeLineMessages(messages) };
    invokeLineApi(AppConfig.line.endpoints.REPLY, payload);
  }
  function push(lineUserId, messages) {
    if (!lineUserId) throw new Error('lineUserId is required');
    const payload = { to: lineUserId, messages: normalizeLineMessages(messages) };
    invokeLineApi(AppConfig.line.endpoints.PUSH, payload);
  }
  function multicast(userIds, messages) {
    if (!Array.isArray(userIds) || userIds.length === 0) return;
    const payload = { to: userIds, messages: normalizeLineMessages(messages) };
    invokeLineApi(AppConfig.line.endpoints.MULTICAST, payload);
  }

  function notifyAdministrators(message) {
    const ids = getAdminUserIds();
    if (ids.length === 0) {
      console.warn('notifyAdministrators: no admin IDs');
      return;
    }
    try {
      multicast(ids, [{ type: 'text', text: message }]);
    } catch (e) {
      console.error(`multicast failed: ${e && e.message}`);
      ids.forEach(id => {
        try { push(id, [{ type: 'text', text: message }]); }
        catch (err) { console.error(`push failed for ${id}: ${err && err.message}`); }
      });
    }
  }
  return { reply, push, multicast, notifyAdministrators };
})();



/** 管理者 UID 一覧（管理者シート B列） */
function getAdminUserIds() {
  const lastRow = adminSheet.getLastRow();
  if (lastRow < 2) return [];
  const uidCol = AppConfig.adminSheet.uidCol;
  return adminSheet
    .getRange(2, uidCol, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(v => v && v.startsWith('U'));
}


/** 管理者への一斉通知（他ファイル互換ラッパー） */
function notifyToAdmin(message) {
  LineMessagingService.notifyAdministrators(message);
}

/** 管理者登録用キーワードを取得 */
function getAdminRegisterKeyword() {
  const cell = AppConfig.adminSheet.keywordCell;
  const keyword = String(adminSheet.getRange(cell).getValue() || '').trim();
  return keyword || null;
}

/** 管理者アカウントを登録 */
function registerAdminAccount(lineUserId) {
  try {
    if (!lineUserId) {
      return { ok: false, reason: 'NO_USER_ID' };
    }

    const { uidCol, startRow } = AppConfig.adminSheet;

    // adminSheetで 同一 UID の有無をチェック （存在すれば終了）
    const lastAdminRow = getLastDataRow(adminSheet, uidCol);
    if (lastAdminRow >= startRow) {
      const rowCount = lastAdminRow - startRow + 1;
      const adminUids = adminSheet
        .getRange(startRow, uidCol, rowCount, 1)
        .getValues()
        .flat()
        .map(v => String(v || '').trim());
      
      // 既に登録済みなら何もしない
      if (adminUids.includes(String(lineUserId))) {
        return { ok: true, already: true };  // 既に登録済み
      }
    }
    
    // displayName を取得（まずcustomerSheetをチェック -> 無ければ API ）
    let displayName = '';
    try {
      const customerIndex = buildCustomerIndexByUid(5);
      const rec = customerIndex[lineUserId];
      if (rec) {
        const lineNameColIdx = AppConfig.customerSheet.lineNameCol - 1;
        displayName = String(rec.values[lineNameColIdx] || '').trim();
      }
    } catch (e) {
      console.warn(`registerAdminAccount: buildCustomerIndexByUid failed: ${e && e.message}`);
    }
    if (!displayName) {
      try {
        // フォールバック: LINE API から取得
        displayName = String(getUserProfile(lineUserId) || '').trim();
      } catch (e) {
        console.warn(`registerAdminAccount: getUserProfile failed: ${e && e.message}`);
      }
    }

    // 一括書き込み（UID, displayName）
    const insertRow = (lastAdminRow >= startRow ? lastAdminRow : (startRow - 1)) + 1;
    adminSheet.getRange(insertRow, uidCol, 1, 2).setValues([[ lineUserId, displayName ]]);

    return { ok: true, created: true };  // 新規登録成功

  } catch (error) {
    logErrorToSheet('管理者登録に失敗', error);
    console.error(`registerAdminAccount: ${error && error.message}`);
    return { ok: false, reason: 'EXCEPTION' };
  }
  
}


/** LINE UID から 表示名 を取得 */
function getUserProfile(lineUserId) {
  try {
    const url = AppConfig.line.endpoints.profile(lineUserId);
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${LINE_CHANNEL_TOKEN}` },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText() || '';
    if (code >= 400) throw new Error(`LINE profile error ${code}: ${body}`);
    const json = JSON.parse(body || '{}');
    return String(json.displayName || '').trim();
  } catch (error) {
    throw new Error(`Error in getUserProfile: ${error.message}`);
  }
}

/** 事前入力フォームURLを生成（フォーム末尾の設問「注文ID」に orderId をセット） */
function generatePrefilledFormUrl(orderId) {
  try {
    // タイトル一致のテキスト設問を検索
    const matches = form.getItems(FormApp.ItemType.TEXT)
      .filter(item => item.getTitle() === AppConfig.form.titles.ORDER_ID);

    if (matches.length === 0) {
      throw new Error(`質問タイトル「${AppConfig.form.titles.ORDER_ID}」がフォーム内に見つかりません。`);
    }
    if (matches.length > 1) {
      throw new Error(`質問タイトル「${AppConfig.form.titles.ORDER_ID}」が複数存在します。1つのみになるようフォームを調整してください。`);
    }

    const item = matches[0];
    const formRes = form.createResponse();
    const itemRes = item.asTextItem().createResponse(orderId);
    formRes.withItemResponse(itemRes);
    const baseUrl = formRes.toPrefilledUrl();

    // 事前入力（entry.xxx=...）が入っていないURLは「注文ID未セット」と判定して即エラー
    // ex.「https://docs.../viewform だけ」 or 「entry.xxx が無い」
    if (baseUrl.indexOf('?') === -1 || !/[?&]entry\.\d+=/.test(baseUrl)) {
      throw new Error('Prefilled URL に 注文ID が含まれていません');
    }

    // セクション到達を保証（2ページ想定: 0,1）
    // 参照記事: https://qiita.com/Lucy_kgsmec/items/5d84e2dec15a80e14594
    const sep = baseUrl.indexOf('?') === -1 ? '?' : '&';
    const url = `${baseUrl}${sep}pageHistory=0,1`;

    // 注文IDが埋め込まれているか最終チェック（念のため）
    if (url.indexOf(encodeURIComponent(orderId)) === -1) {
      throw new Error('Prefilled URL に注文IDの値が含まれていません');
    }

    return url;
  } catch (error) {
    throw new Error(`Error in generatePrefilledFormUrl: ${error.message}`);
  }
}


/** 顧客シートを走査して UID -> 行データのインデックス を作る。 */
/** 生成されるデータ形式例: 
 * customerIndexByUid = {
    "U12345": {
      row: 7,
      values: [ "U12345", "LINEユーザー名", "飲食店名", "freeeユーザー名", "freeeID" ]
    },
    "U67890": {...}
  };
 */
function buildCustomerIndexByUid(columnCount = 5) {
  if (!Number.isInteger(columnCount) || columnCount <= 0) {
    throw new Error('buildCustomerIndexByUid: columnCount must be a positive integer.');
  }

  const { uidCol, startRow } = AppConfig.customerSheet;

  // 最終行を UID 列基準で取得
  const lastDataRow = getLastDataRow(customerSheet, uidCol);
  if (lastDataRow < startRow) return {};

  // 登録済の顧客データを一括取得
  const totalRows = lastDataRow - startRow + 1;  // 読み込む行数
  const totalCols = Math.max(columnCount, uidCol);  // 読み込む列数
  const customerData = customerSheet.getRange(startRow, 1, totalRows, totalCols).getValues();


  const customerIndexByUid = {};
  for (let i = 0; i < customerData.length; i++) {
    const rowCells = customerData[i]; // 1行データ
    const userId = String(rowCells[uidCol - 1] || '').trim();
    if (!userId) continue; // UID未入力行はスキップ

    customerIndexByUid[userId] = {
      row: startRow + i,  // 行番号に変換
      values: rowCells.slice(0, columnCount),  // 行データをキャッシュ
    };
  }
  return customerIndexByUid;
}



/** 既存互換の簡易キャッシュ（handleFormSubmit で参照） */
function makeCache() {
  const cache = CacheService.getScriptCache();
  return {
    get: key => {
      const v = cache.get(key);
      return v ? JSON.parse(v) : null;
    },
    put: (key, value, sec) => {
      cache.put(key, JSON.stringify(value), sec || 600);
      return value;
    },
    remove: key => (cache.remove(key), true)
  };
}


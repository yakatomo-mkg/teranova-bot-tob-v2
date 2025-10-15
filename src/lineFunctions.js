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

/** 管理者登録キーワード（管理者シート!E2） */
function getAdminRegisterKeyword() {
  const cell = AppConfig.adminSheet.keywordCell;
  const keyword = String(adminSheet.getRange(cell).getValue() || '').trim();
  if (!keyword) throw new Error(`管理者登録キーワードが未設定です（${AppConfig.sheetNames.ADMIN}!${cell}）`);
  return keyword;
}

/** UID を管理者シートに upsert（A:時刻, B:UID） */
function registerAdminAccount(userId) {
  const startRow = 2;
  const lastRow = adminSheet.getLastRow();
  if (lastRow >= startRow) {
    const ids = adminSheet.getRange(startRow, 2, lastRow - startRow + 1, 1)
      .getValues().flat().map(v => String(v || '').trim());
    const idx = ids.findIndex(v => v === userId);
    if (idx !== -1) {
      adminSheet.getRange(startRow + idx, 1).setValue(new Date());
      return;
    }
  }
  adminSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[new Date(), userId]]);
}


/** LINEプロフィール（ID & displayName）取得 */
function getUserProfile(userId) {
  try {
    const url = AppConfig.line.endpoints.profile(userId);
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${LINE_CHANNEL_TOKEN}` },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 400) throw new Error(`LINE profile error ${code}: ${res.getContentText()}`);
    const json = JSON.parse(res.getContentText() || '{}');
    return json.displayName || '';
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


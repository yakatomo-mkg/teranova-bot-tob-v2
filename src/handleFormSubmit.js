/** TODO: LINE ID や 注文ID が取得できずに、確認用メッセージの送信に失敗した場合（`Error in sendOrderNotification: Error: LINE user ID not found for order ID:`）
 * 現状 -> スプレッドシートに注文データが記録されてしまっている（参照：飲食店名「test_miss」）
 * 対応 -> 注文データが正しく受け取れなかったときはもう一度入力を依頼し、当該注文データは破棄（スプレッドシートにも記録を残さない）
*/


/** フォームから回答が送信されたときの処理 */
function handleFormSubmit(e) {
  const receivedOrder = getOrders(e);

  // 受信回答を Map 化（検索コスト削減）
  const answersByTitle = mapAnswersByTitle(receivedOrder); // title -> answer
  console.log('[submit] titles', receivedOrder.map(r => r.question));

  // タイトルで「注文ID」を取得
  const orderId = answersByTitle.get(AppConfig.form.titles.ORDER_ID) || "";

  console.log('[submit] ORDER_ID expected=', AppConfig.form.titles.ORDER_ID);
  console.log('[submit] orderId(from answersByTitle)=', orderId);

  // タイムスタンプを取得して、「受付日」を作成
  const timestamp = e.response.getTimestamp();
  const acceptedDate = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd");

  // menuシートから、フォーム登録中のメニューデータを取得
  const menuData = getRegistedMenuData(menuSheet, AppConfig.menuSheet);  
  
  const writtenRow = appendOrderToSheet(answersByTitle, receivedOrder, acceptedDate, menuData, orderSheet, AppConfig.orderSheet);

  // 注文データを管理者と注文ユーザーに送信
  sendOrderNotification(orderId, receivedOrder, menuData, cacheSheet);
}


/** フォームの回答から必要情報を取得 */
function getOrders(e) {
  try {
    const formResponses = e.response.getItemResponses();
    const orderData = [];   // 注文データを格納する配列
    for (const res of formResponses) {
      const qItem = res.getItem();
      const qItemId = qItem.getId();
      const qItemTitle = qItem.getTitle();
      let qItemAnswer = res.getResponse();
      qItemAnswer = String(qItemAnswer == null ? "" : qItemAnswer).trim();


      if (qItemTitle === AppConfig.form.titles.ORDER_ID) {
        console.log('[getOrders] FOUND ORDER_ID item',
          { id: qItemId, title: qItemTitle, answer: qItemAnswer });
      }
      orderData.push({
        id: qItemId,
        question: qItemTitle,
        answer: qItemAnswer
      });
    }
    return orderData;
  } catch (error) {
    throw new Error(`Error in getOrders: ${error}`);
  }
}


/** タイトル -> 回答 の Map を生成 */
function mapAnswersByTitle(receivedOrder) {
  const map = new Map();
  for (const it of receivedOrder) {
    map.set(it.question, it.answer);
  }
  return map;
}



// /** 質問ID -> （question, answer）の Map を生成 */
// function mapItemsById(receivedOrder) {
//   const map = new Map();
//   for (const it of receivedOrder) {
//     map.set(it.id, it);
//   }
//   return map;
// }


/** タイトル一致で回答を取得（見つからなければ空文字） */
function findAnswerByTitle(receivedOrder, title) {
  const item = receivedOrder.find(r => r.question === title);
  return item ? item.answer : "";
}

/** タイトル一致でアイテム全体を取得（見つからなければ null） */
function findItemByTitle(receivedOrder, title) {
  return receivedOrder.find(r => r.question === title) || null;
}


/** 受信したフォーム回答をorderシートに転記（バッチ書き込み） */
function appendOrderToSheet(answersByTitle, receivedOrder, acceptedDate, menuData, orderSheet, oS) {
  try {
    const targetRow = orderSheet.getLastRow() + 1;

    // 1) 先頭の固定カラム（1〜6列）は一括書き込み
    const headerRow = buildHeaderRow(answersByTitle, acceptedDate);
    orderSheet.getRange(targetRow, 1, 1, headerRow.length).setValues([headerRow]);

    // 2) メニュー回答行（最小範囲に圧縮して一発書き込み）
    const menuRowWrite = buildMenuWriteRow(receivedOrder, menuData, oS);
    if (menuRowWrite) {
      const { minCol, rowArray } = menuRowWrite;
      orderSheet.getRange(targetRow, minCol, 1, rowArray.length).setValues([rowArray]);
    }

    // 3) 1列目にチェックボックスをセット
    orderSheet.getRange(targetRow, 1).insertCheckboxes();

    return targetRow;
  } catch (error) {
    throw new Error(`Error at appendOrderToSheet: ${error}`);
  }
}


/** 先頭固定カラム行を作成（チェックボックス・受付日・注文ID・店名・お届け日・コメント） */
function buildHeaderRow(answersByTitle, acceptedDate) {
  const t = AppConfig.form.titles;
  const orderId      = answersByTitle.get(t.ORDER_ID)      || "";
  const shopName     = answersByTitle.get(t.SHOP_NAME)     || "";
  const deliveryDate = answersByTitle.get(t.DELIVERY_DATE) || "";
  const comment      = answersByTitle.get(t.COMMENT)       || "";


  console.log('[headerRow] orderId=', orderId, 'shopName=', shopName, 'deliveryDate=', deliveryDate, 'comment.len=', comment ? comment.length : 0);

  return [
    "",             // チェックボックス用の空白
    acceptedDate,   // 受付日
    orderId,        // 注文ID
    shopName,       // 飲食店名
    deliveryDate,   // お届け日
    comment         // コメント
  ];
}



/** メニュー回答の最小範囲行を作成（{minCol, rowArray} を返す） */
function buildMenuWriteRow(receivedOrder, menuData, oS) {
  // menuData の [0]=qId, [4]=orderシート列番号
  const menuIndex = new Map(menuData.map(r => [r[0], r]));
  const minCol = oS.menuStartCol;

  let maxColUsed = 0;
  const cells = {}; // {colNum: value}

  for (const { id, question, answer } of receivedOrder) {
    if (!answer || answer === "0") continue;
    if ( FIXED_QUESTIONS.has(question) ) continue; // ← 旧FORM_MENU_START_INDEXの代替

    const menuRow = menuIndex.get(id);
    if (!menuRow) continue;

    const colNum = menuRow[4]; // TODO: [グローバル変数に設定] orderシートにおける列番号
    if (colNum < minCol) continue;

    cells[colNum] = answer;
    if (colNum > maxColUsed) maxColUsed = colNum;
  }

  if (maxColUsed < minCol) return null;

  const width = maxColUsed - minCol + 1;
  const rowArray = new Array(width).fill('');
  for (const colStr in cells) {
    const col = Number(colStr);
    rowArray[col - minCol] = cells[col];
  }
  return { minCol, rowArray };
}


/** menuシートのフォーム登録メニュー管理エリアからデータを取得 */
function getRegistedMenuData(menuSheet, mS) {
  try {
    const lastRowFmArea = menuSheet.getRange(menuSheet.getMaxRows(), mS.fmStartCol)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const numRows = lastRowFmArea - mS.contStartRow + 1;
    const menuDataRange = menuSheet.getRange(mS.contStartRow, mS.fmStartCol, numRows, 5);
    return menuDataRange.getValues();
  } catch (error) {
    throw new Error(`Error in getRegistedMenuData: ${error}`);
  }
}




function sendOrderNotification(orderId, receivedOrder, menuData, cacheSheet) {
  try {
    const cache = makeCache();
    let lineUserId = cache.get(orderId);

    // キャッシュからユーザーIDが見つからなかった場合 */
    if (!lineUserId) {
      lineUserId = findLineUserId(orderId, cacheSheet);
      if (!lineUserId) {
        throw new Error(`LINE user ID not found for order ID: ${orderId}`);
      }
    }

    // 見つかった場合は注文ユーザーと管理者に通知
    sendOrderMessageToUser(lineUserId, receivedOrder, menuData);
    sendOrderMessageToAdmin(receivedOrder, menuData);
    
    // キャッシュデータを削除
    const isRemovedCache = removeCacheData(orderId, cacheSheet);
    if (isRemovedCache) {
      console.log(`Successfully removed cache data for order ID: ${orderId}`);
    } else {
      console.log(`Failed to remove cache data for order ID: ${orderId}`);
    }
  } catch (error) {
    throw new Error(`Error in sendOrderNotification: ${error}`);
  }
}

/** キャッシュシートから UID を検索 */
function findLineUserId(orderId, cacheSheet) {
  try {
    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 2) return null;
    const cacheData = cacheSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const cacheRow = cacheData.find(row => row[1] === orderId);
    return cacheRow ? cacheRow[2] : null;
  } catch (error) {
    throw new Error(`Error in findLineUserId: ${error}`);
  }
}


/** キャッシュ削除（依存注入） */
function removeCacheData(orderId, cacheSheet) {
  try {
    const cache = makeCache();
    cache.remove(orderId);

    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 2) return false;

    const cacheData = cacheSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const deleteIdx = cacheData.findIndex(row => row[1] === orderId);
    if (deleteIdx !== -1) {
      cacheSheet.deleteRow(deleteIdx + 2);
      return true;
    }
    return false;
  } catch (error) {
    throw new Error(`Error in removeCacheData: ${error}`);
  }
}


function sendOrderMessageToUser(lineUserId, receivedOrder, menuData) {
  try {
    let toUserMessage = `ご注文ありがとうございます。\n以下のご注文を承りました。`;
    toUserMessage += createOrderMessage(receivedOrder, menuData);
    LineMessagingService.push(lineUserId, [{ type: 'text', text: toUserMessage }]);
  } catch (error) {
    throw new Error(`Error in sendOrderMessageToUser: ${error}`);
  }
}


function sendOrderMessageToAdmin(receivedOrder, menuData) {
  try {
    let toAdminMessage = `\n以下の注文が届きました。`;
    toAdminMessage += createOrderMessage(receivedOrder, menuData);
    notifyToAdmin(toAdminMessage);
  } catch (error) {
    throw new Error(`Error in sendOrderMessageToAdmin: ${error}`);
  }
}


const createOrderMessage = (receivedOrder, menuData) => {
  try {
    const answers = mapAnswersByTitle(receivedOrder);
    const t = AppConfig.form.titles;
    const orderId      = answers.get(t.ORDER_ID)      || "";
    const name         = answers.get(t.SHOP_NAME)     || "";
    const deliveryDate = answers.get(t.DELIVERY_DATE) || "";
    const comment      = answers.get(t.COMMENT)       || "";

    let message = `
    
注文ID：${orderId}
お届け日：${deliveryDate}
お名前：${name}

【 ご注文内容 】`;

  // menuData の [0]=qId, [3]=unit
    for (const { id: qId, question, answer } of receivedOrder) {
      if (!answer || answer === "0") continue;
      if ( FIXED_QUESTIONS.has(question) ) continue;

      const menuRow = menuData.find(row => row[0] === qId);
      const unit = menuRow ? (menuRow[3] || "") : "";
      message += `\n${question}：${answer} ${unit}`;
    }

    if (comment && comment.trim() !== "") {
      message += `\n\nコメント：\n${comment}`;
    }
    return message;
  } catch (error) {
    throw new Error(`Error in createOrderMessage: ${error}`);
  }
};


/** フォームから回答が送信されたときの処理 */

function handleFormSubmit(e) {

  const receivedOrder = getOrders(e);
  console.log(receivedOrder[0].answer);  // 注文IDを出力
  const orderId = receivedOrder[0].answer;

  // タイムスタンプを取得して、「受付日」を作成
  const timestamp = e.response.getTimestamp();
  const acceptedDate = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd");

  const menuData = getRegistedMenuData();  // menuシートから、フォーム登録中のメニューデータを取得

  const lastRow = orderSheet.getLastRow() + 1; // orderシートにおいて転記する行を取得

  appendOrderToSheet(receivedOrder, acceptedDate, lastRow, menuData); // 転記
  orderSheet.getRange(lastRow, 1).insertCheckboxes();  // 1列目にチェックボックスをセット

  // 注文データを管理者と注文ユーザーに送信
  sendOrderNotification(orderId, receivedOrder, menuData);
}



/** フォームの回答データから必要情報を取得する関数 */
function getOrders(e) {
  try {
    console.log(`フォーム名: ${e.source.getTitle()}`);
    console.log(`フォームID: ${e.source.getId()}`);

    const formResponses = e.response.getItemResponses();

    const orderData = [];   // 注文データを格納する配列
    for (const res of formResponses) {
      const qItem = res.getItem();
      // const qItemIndex = qItem.getIndex(); // 質問インデックス
      const qItemId = qItem.getId();    // 質問ID
      const qItemTitle = qItem.getTitle();  // 質問タイトル
      let qItemAnswer = res.getResponse(); // 回答
      qItemAnswer = qItemAnswer.trim();  // 前後の空白文字がある場合は除去

      // debug(`${qItemId} - ${qItemTitle} - ${qItemAnswer}`);
      /** 
       * 1043428527 - 注文ID - 18e7cd6fc83
       * 368478372 - 飲食店名（お名前） - デンパーク
       * 740554813 - メールアドレス - denpark@gmail.com
       * 1011140994 - お届け日 - 2024年4月5日（金）
       * 1904804293 - 野菜セットM（約8品目） - 3
       * 1653605003 - にんじん（250g） - 
       * 935542138 - にんじん - 0.6
       * 1984324547 - だいこん（葉なし） - 10
       * 145849789 - コメント - にんじんは量り売りを注文
       */

      // 回答が空白のもの、または0のものはorderDataに追加しない
      // if (qItemAnswer !== "" && qItemAnswer !== "0") {
      orderData.push({
        id: qItemId,           // 質問ID
        question: qItemTitle,  // 質問タイトル
        answer: qItemAnswer.toString()  // 回答
      });
      // }
    }
    // debug(orderData);
    return orderData;
  } catch (error) {
    console.error(`Error in getOrders: ${error}`);
    throw new Error(`Error in getOrders: ${error}`);
  }
}


/** 受信したフォーム回答をorderシートに転記するためフォーマットを整える */
function appendOrderToSheet(receivedOrder, acceptedDate, lastRow, menuData) {
  try {
    // 基本情報に関する質問への回答を一括で転記
    const orderValues = [
      "",                       // チェックボックス用の空白
      acceptedDate,             // 受付日
      receivedOrder[0].answer,  // 注文ID
      receivedOrder[1].answer,  // 飲食店名
      receivedOrder[2].answer,  // お届け日
      receivedOrder[receivedOrder.length - 1].answer,  // コメント
    ];
    orderSheet.getRange(lastRow, 1, 1, orderValues.length).setValues([orderValues]);

    // 注文メニューの転記(１つずつ行う)
    for (let i = FORM_MENU_START_INDEX; i < receivedOrder.length; i++) {
      const order = receivedOrder[i];

      // 回答が空白または0の場合はスキップ
      if (order.answer === "" || order.answer === "0") {
        continue;
      }

      const menuIndex = new Map(menuData.map(r=>[r[0], r]));
      const menuRow = menuIndex.get(order.id);
      if (menuRow) {
        const orderStColNum = menuRow[4];  // menuDataの行データから、orderシートにおける列番号を取得
        orderSheet.getRange(lastRow, orderStColNum).setValue(order.answer);
      }
    }
  } catch (error) {
    console.log("Error at appendOrderToSheet: ", error);
    throw new Error(`Error at appendOrderToSheet: ${error}`);
  }
}


/** menuシートのフォーム登録メニュー管理エリアからデータを取得 */
function getRegistedMenuData() {
  try {
    const mS = MENU_SHEET_SETTINGS;
    const lastRowFmArea = menuSheet.getRange(menuSheet.getMaxRows(), mS.fmStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const numRows = lastRowFmArea - mS.contStartRow + 1;  // ヘッダーを除いたコンテンツが存在する行数を計算
    console.log(`フォーム登録メニュー数: ${numRows}`);
    const menuDataRange = menuSheet.getRange(mS.contStartRow, mS.fmStartCol, numRows, 5);  // 5列分
    console.log(menuDataRange.getValues());
    return menuDataRange.getValues();  // 二次元配列としてデータを取得
  } catch (error) {
    console.error(`Error in getRegistedMenuData: ${error}`);
    throw new Error(`Error in getRegistedMenuData: ${error}`);
  }
}



function sendOrderNotification(orderId, receivedOrder, menuData) {
  try {
    const cache = makeCache();
    let lineUserId = cache.get(orderId);

    /** キャッシュからLINEユーザーIDが見つからなかった場合 */
    if (!lineUserId) {
      lineUserId = findLineUserId(orderId);
      if (!lineUserId) {
        throw new Error(`LINE user ID not found for order ID: ${orderId}`);
      }
    }
    /** LINEユーザーIDが見つかった場合 */ 
    // 注文ユーザーと管理者に通知
    sendOrderMessageToUser(lineUserId, receivedOrder, menuData);
    sendOrderMessageToAdmin(receivedOrder, menuData);
    
    // キャッシュデータを削除
    const isRemovedCache = removeCacheData(orderId);
    if (isRemovedCache) {
      console.log(`Successfully removed cache data for order ID: ${orderId}`);
    } else {
      console.log(`Failed to remove cache data for order ID: ${orderId}`);
    }

  } catch (error) {
    throw new Error(`Error in sendOrderNotification: ${error}`);
  }
}


function findLineUserId(orderId) {
  try {
    const cacheData = cacheSheet.getRange(2, 1, cacheSheet.getLastRow(), 3).getValues();
    // debug(`cacheデータ: ${cacheData}`);
    const cacheRow = cacheData.find(row => row[1] === orderId);
    return cacheRow ? cacheRow[2] : null;
  } catch (error) {
    throw new Error(`Error in findLineUserId: ${error}`);
  }
}



function removeCacheData(orderId) {
  try {
    // キャッシュサービス内のキャッシュデータを削除
    const cache = makeCache();
    cache.remove(orderId);

    // cacheシートからキャッシュデータを削除
    const cacheData = cacheSheet.getRange(2, 1, cacheSheet.getLastRow(), 3).getValues();  // キャッシュシートの全キャッシュデータを取得
    const deleteIdx = cacheData.findIndex(row => row[1] === orderId);
    if (deleteIdx !== -1) {
      cacheSheet.deleteRow(deleteIdx + 2);  // deleteIndexは0から始まるため、+2して行数を取得する
      return true;  // キャッシュデータの削除が成功した場合はtrueを返す
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
    sendPushMessage(lineUserId, toUserMessage);
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
    const { answer: orderId } = receivedOrder[0];
    const { answer: name } = receivedOrder[1];
    const { answer: deliveryDate } = receivedOrder[2];

    let message = `
    
注文ID：${orderId}
お届け日：${deliveryDate}
お名前：${name}

【 ご注文内容 】`;

    for (let i = FORM_MENU_START_INDEX; i < receivedOrder.length - 1; i++) {
      const { id: qId, question, answer } = receivedOrder[i];
      // 回答が空白または0の場合はスキップ
      if (answer === "" || answer === "0") {
        continue;
      }

      let itemUnit = "";
      const menuRow = menuData.find(row => row[0] === qId);
      if (menuRow) {
        itemUnit = menuRow[3];  // menuシートから「単位」を取得
      }
      message += `
${question}：${answer} ${itemUnit}`;
    }
    const { answer: comment } = receivedOrder[receivedOrder.length - 1];
    if (comment && comment.trim() !== "") {
      message += `
      
コメント：
${comment}`;
    }
    return message;
  } catch (error) {
    throw new Error(`Error in createOrderMessage: ${error}`);
  }  
}


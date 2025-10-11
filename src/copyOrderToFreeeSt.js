/**
 * orderシートの注文データのうち、チェックをつけたデータを、freeeシートに転記する関数
 */
function copyOrderToFreeSheet() {
  try {
    const lrAcptDate = getLastDataRow(orderSheet, oS.acptDateCol);
    const lcOs = orderSheet.getLastColumn();
    console.log(`lrAcptDate = ${lrAcptDate}`);
    const orderValues = orderSheet.getRange(1, 1, lrAcptDate, lcOs).getValues();
    // console.log(orderValues);
    /** console.log(orderValues[1]); // 取得した二次元配列のうち、２番目の要素(=2行目の行データ)を出力 */
    /** [ '納品書作成', '注文受付日', '注文ID', '飲食店名', 'メールアドレス', 'お届け日', 'コメント', '野菜セットS', ... ,'ブロッコリー', 'かぶ' ] */

    // メニューヘッダー項目をキャッシュ
    const menuH = cacheMenuHeaders(orderValues);

    // コピー対象の注文データを配列に格納
    const targetOrders = getTargetOrders(orderValues, menuH);

    if (targetOrders.length === 0) {
      throw new Error("チェックされた項目が見つかりません。\nコピーしたい注文データにチェックをつけてください。")
    }

    const lrFs = getLastDataRow(freeeSheet, fS.deliDateCol);
    const copyData = buildCopyData(targetOrders);

    writeDataToSheet(freeeSheet, lrFs + 1, copyData);
    setPulldowns();

    const copiedRows = targetOrders.map(data => parseInt(data.rowIndex));
    console.log(copiedRows);
    rowColoredAndUnchecked(orderSheet, copiedRows, "#808080");

    showMessageDialog("成功", `コピーが完了した注文数 : ${targetOrders.length}`);   
  } catch (error) {
    showErrorDialog("copyOrderToFreeSheet", error);
  }
}


/**
 * 注文シートのメニューヘッダー情報をキャッシュする関数
 * @param {Array} values - orderシートから取得した二次元配列 (ヘッダー行を含む)
 * @param {object} oS - orderシートの設定オブジェクト
 * @return {Array} - メニューヘッダー情報の配列
 */
function cacheMenuHeaders(values) {
  const menuHeaders = [];
  
  for (let i = oS.menuStartCol - 1; i < values[1].length; i++) {
    const item = values[1][i];
    const unit = values[2][i];
    const amount = values[3][i];
    const price = values[4][i];

    menuHeaders.push({item, amount, unit, price});
  }
  return menuHeaders;
}

/**
 * 注文データ(行データ)を処理し、リクエスト(配列)を作成する関数
 * @param {Array} values - 注文データの二次元配列
 * @param {object} oS - orderシートの設定オブジェクト
 * @return {Array} - 注文データの配列
 */
function getTargetOrders(values, menuHeaders) {
  const orderData = [];
  for (let i = oS.contStartRow - 1; i < values.length; i++) {

    if (values[i][0]) {
      const row = values[i];
      const data = {
        rowIndex: i + 1,   // 行番号
        acptDate: row[oS.acptDateCol - 1],
        orderId: row[oS.orderIdCol - 1],
        deliDate: row[oS.deliDateCol - 1],
        shopName: row[oS.shopNameCol - 1],
        menuData: getMenuData(row, menuHeaders)
      };
      orderData.push(data);
    }
  }
  return orderData;
}

/**
 * メニューデータを抽出し、JSON文字列に変換する関数
 * @param {Array} row - 注文データの行
 * @param {object} oS - orderシートの設定オブジェクト
 * @param {Array} menuHeaders - メニューのヘッダー情報
 * @return {string} - JSON形式のメニューデータ
 */
function getMenuData(row, menuHeaders) {
  const menuArr = [];
  for (let i = oS.menuStartCol - 1; i < row.length; i++) {
    if (row[i] !== "") {
      const { item, amount, unit, price } = menuHeaders[i - (oS.menuStartCol - 1)];
      const quantity = row[i];
      menuArr.push({
        item,
        amount,
        unit,
        price,
        quantity
      });
    }
  }
  console.log(`${JSON.stringify(menuArr)}`);
  return JSON.stringify(menuArr);
}


/**
 * コピーするためのデータを構築する関数
 * @param {Array} orderData - 注文データの配列
 * @param {Array} menuHeaders - メニューヘッダー情報の配列
 * @return {Array} - コピー用のデータ配列
 */
function buildCopyData(orderData) {
  return orderData.map(data => ([
    "",             // チェックボックス
    data.acptDate,
    data.orderId,
    data.deliDate,
    data.shopName,
    "",
    "",
    data.menuData
  ]));
}

/**
 * 指定したシートにデータを書き込む関数
 * @param {Object} sheet - freeeシートのオブジェクト
 * @param {number} startRow - 書き込み開始行番号
 * @param {Array} data - 書き込むデータの二次元配列
 */
function writeDataToSheet(sheet, startRow, data) {
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(startRow, 1).offset(0, 0, data.length, 1).insertCheckboxes();
}


/**
 * 行データがセットされたタイミングで、「freee（取引先名）」 列にプルダウンをセットする関数
 */
function setPulldowns() {
  try {
    const startRowOfPs = pS.contStartRow;
    const ptnrDataLength = partnerSheet.getLastRow() - startRowOfPs + 1;

    const dataRange = partnerSheet.getRange(startRowOfPs, pS.displayCol, ptnrDataLength, 1);
    const nameList = dataRange.getValues().flat().filter(String);
    console.log(nameList);

    const pdRule = SpreadsheetApp.newDataValidation().requireValueInList(nameList, true).build();

    const lastRowOfFs = getLastDataRow(freeeSheet, fS.deliDateCol);
    const startRowOfFs = fS.contStartRow;
    const freeeDataLength = lastRowOfFs - startRowOfFs + 1;

    const pdRange = freeeSheet.getRange(startRowOfFs, fS.partnerNameCol, freeeDataLength, 1);
    pdRange.setDataValidation(pdRule);
  
  } catch (error) {
    throw new Error(`Error in setPulldowns: プルダウン設定時にエラーが発生しました。\n${error.message}`)
  }
}


/** freeeシートにおいて取引先名列に値がセットされた際に取引先IDを取得してID列にセットする関数 */
function detectSettedPartnerName(e) {
  const sheet = e.source.getActiveSheet();
  const editRange = e.range;
  const editCol = editRange.getColumn();
  console.log(editCol);
  if (sheet.getName() === "納品書作成" && editCol === fS.partnerNameCol && editRange.getValue() !== "") {
    console.log("納品書作成シートで取引先名の変更を検知しました!");
    const row = editRange.getRow();
    const newValue = e.value;
    console.log(newValue);

    const parts = newValue.split("]  ");
    const id = parts[0].slice(1);
    sheet.getRange(row, fS.partnerNameCol + 1).setValue(id);
  }
  console.log("終了");
}


/**
 * detectSettedPartnerName に対する onEdit トリガーを設定
 */
function createOnEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  const isTriggerSet = triggers.some(trigger => trigger.getHandlerFunction() === 'detectSettedPartnerName' && trigger.getEventType() === ScriptApp.EventType.ON_EDIT);
  
  if (!isTriggerSet) {
    ScriptApp.newTrigger('detectSettedPartnerName')
      .forSpreadsheet(SSID)
      .onEdit()
      .create();
  }
}



/** 
 * 【トリガー設定中】
 * スプレッドシートのツールバーに 「teranova Menu」 を作成 
 */

// let ui = SpreadsheetApp.getUi();

function setCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('てらのばメニュー')
    .addItem("注文フォームを作成", "updateOrderForm")
    .addSeparator()
    .addItem('転記（ order → freee ）', 'copyOrderToFreeSheet')
    .addSeparator()
    .addItem('グレー色の行データを削除', 'deleteColoredRows')
    .addSeparator()
    .addItem('［ freee ］ 認証', 'showAuth')
    .addItem('［ freee ］ 取引先リストを取得', 'fetchPartnersList')
    .addItem('［ freee ］ 納品書作成', 'createDeliverySlips')
    // .addItem('［freee］ ログアウト', 'logout')
    // .addItem('［freee］ 納品書取得', 'getchDeliverySlips')
    .addToUi();
}

function createOnOpenTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.some(trigger => trigger.getHandlerFunction() === 'setCustomMenu');


  if (!existingTrigger) {
    ScriptApp.newTrigger('setCustomMenu')
      .forSpreadsheet(SSID)
      .onOpen()
      .create();
  }
}


/**
 * スプレッドシートのツールバーに 「てらのばメニュー」 を作成
 */
// function setCustomMenu() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('てらのばメニュー')
//     .addItem("注文フォームを作成", "updateOrderForm")
//     .addSeparator()
//     .addItem('転記（ order → freee ）', 'copyOrderToFreeSheet')
//     .addSeparator()
//     .addItem('グレー色の行データを削除', 'deleteColoredRows')
//     .addSeparator()
//     .addItem('［ freee ］ 認証', 'showAuth')
//     .addItem('［ freee ］ 取引先リストを取得', 'fetchPartnersList')
//     .addItem('［ freee ］ 納品書作成', 'createDeliverySlips')
//     .addToUi();
// }

// function createOnOpenTrigger() {
//   const triggers = ScriptApp.getProjectTriggers();
//   const existingTrigger = triggers.some(tr => tr.getHandlerFunction() === 'setCustomMenu');
//   if (!existingTrigger) {
//     ScriptApp.newTrigger('setCustomMenu')
//       .forSpreadsheet(SPREADSHEET_ID) // <- TODO: [テスト環境用] 
//       .onOpen()
//       .create();
//   }
// }

/**
 * テスト用
 * 実行後は 削除 or コメントアウト すること
 */


/** 注文フォームの設問タイトルをログ出力 */
function logFormQuestionTitles() {
  try {
    const items = form.getItems();
    if (items.length === 0) {
      console.log("このフォームには設問がありません。");
      return;
    }

    console.log(`フォームタイトル: ${form.getTitle()}`);
    console.log("----- 設問一覧 -----");
    items.forEach((item, i) => {
      console.log(`${i + 1}. ${item.getTitle()} [${item.getType()}]`);
    });
    console.log("----- 出力完了 -----");

  } catch (error) {
    console.error("logFormQuestionTitles でエラーが発生:", error);
  }
}



// function doPost(e) {
//   try {
//     // 受信だけ記録（JSONパースもしない）
//     const size = (e && e.postData && e.postData.contents) ? e.postData.contents.length : 0;
//     console.log('[SMOKE] size=', size);
//   } catch (_) {}
//   return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
// }

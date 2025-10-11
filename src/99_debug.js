/** 
 * デバッグ関数
 * @param {String | Object} value - 出力する値またはオブジェクト 
 */
// function debug(value="デバッグテスト") {
//   const date = new Date();  // 現在の日時を取得
//   const targetRow = logSheet.getLastRow() + 1;  // ログを書き込む行を取得
//   let outPutValue = value;  // 出力する値を格納する変数を初期化

//   // 出力する値がオブジェクトの場合、JSON文字列に変換 (オブジェクトのままだと解釈できない表示になるため)
//   if (typeof value === "object") {
//     outPutValue = JSON.stringify(value);
//   }
//   logSheet.getRange("A" + targetRow).setValue(date);   // A列に出力日時をセット
//   logSheet.getRange("B" + targetRow).setValue(outPutValue);  // B列にログの出力をセット
// }
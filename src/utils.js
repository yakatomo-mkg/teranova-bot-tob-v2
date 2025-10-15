// const ui = SpreadsheetApp.getUi();

function showErrorDialog(functionName, error) {
  const ui = SpreadsheetApp.getUi();
  ui.alert("エラー", `Error in ${functionName}\n${error.message}`, ui.ButtonSet.OK );
}


function showMessageDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}


function getLastDataRow(sheet, colNumber) {
  const lastDataRow = sheet.getRange(sheet.getMaxRows(), colNumber)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  return lastDataRow;
}


function rowColoredAndUnchecked(sheet, rows, color, col=sheet.getLastColumn()) {
  rows.forEach(row => {
    sheet.getRange(row, 1, 1, col).setBackground(color);
    const checkedCell = sheet.getRange(row, 1);
    checkedCell.setValue(false);
  });
}



function deleteColoredRows() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = activeSheet.getName();

  if (sheetName === '注文データ' || sheetName === '納品書作成') {
    const lastRow = activeSheet.getLastRow();
    console.log(lastRow);

    for (let i = lastRow; i > 0; i--) {
      const row = activeSheet.getRange(i, 1, 1, activeSheet.getLastColumn());
      const rowBgColor = row.getBackground();

      if (rowBgColor === '#808080') {
        // console.log(i);
        activeSheet.deleteRow(i);
      }
    }
  } else {
    console.log("このシートでの行削除は許可されていません。");
  }
}


function logErrorToSheet(functionName, error) {
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1).setValue(new Date());
  logSheet.getRange(lastRow + 1, 2).setValue(`Error in ${functionName}`);
  logSheet.getRange(lastRow + 1, 3).setValue(error.message);
}


function updateOrderForm() {
  try {
    const formTitle = form.getTitle();
    const message = `\n${formTitle}\n\n上記のフォームを更新しても良いですか？`;
    if (!confirm(message)) return;
    const lastRowOfItem = getLastDataRow(menuSheet, mS.itemNameCol);
    console.log(`lastRowOfItem = ${lastRowOfItem}`);
    console.log(`mS.ckBoxCol = ${mS.ckBoxCol}`);

    if (lastRowOfItem - mS.contStartRow + 1 === 0) {
        throw new Error("メニューDBにおいて、メニューが見つかりません。");
    }
    
    const isCheckedData = menuSheet.getRange(mS.contStartRow, mS.ckBoxCol, lastRowOfItem - mS.contStartRow + 1, 1).getValues().flat().some(Boolean);
    if (!isCheckedData) {
        throw new Error("メニューシートにおいて、チェックされた項目が見つかりません。");
    }
    console.log('ここまでOK');
    
    /** フォーム更新処理を開始 */
    console.log(`mS.itemDetailCol = ${mS.itemDetailCol}`);
    const menuData = menuSheet.getRange(mS.contStartRow, 1, lastRowOfItem, mS.itemDetailCol).getValues();

    
    const items = form.getItems();

    items.forEach(item => form.deleteItem(item));

    clearMenuSheetValues();

    addBasicQuestions(form);

    addMenuItems(menuData, lastRowOfItem);

    form.addTextItem().setTitle("コメント");
    
    getQuestionId();

    getPublishedFormUrl();

    ss.toast("フォームの作成に成功しました。", "成功", 5);

  } catch (error) {
    showMessageDialog("updateOrderForm", error);
  }
}


function addBasicQuestions(form) {
  try {
    form.addTextItem().setTitle("注文ID").setHelpText("注文管理のためのIDですので、削除や変更を行わないでください。").setRequired(true);
    form.addTextItem().setTitle("飲食店名（お名前）").setRequired(true);

    form.addListItem().setTitle("お届け日").setRequired(true).setChoiceValues(generateDaysChoices(30)); 
  } catch (error) {
    throw new Error(`Error in addBasicQuestions: 基本情報アイテムの追加中にエラーが発生しました。\n${error.message}`);
  }  
}



function addMenuItems(menuData, lastRowOfItem) {
  try {
    for (let i = 0; i < lastRowOfItem - 1; i++) {
      const isChecked = menuData[i][mS.ckBoxCol - 1];
      if (isChecked) {
        const itemId = menuData[i][mS.itemIdCol - 1];
        const itemName = menuData[i][mS.itemNameCol - 1];
        const itemUnit = menuData[i][mS.itemUnitCol - 1];
        const itemAmount = menuData[i][mS.itemAmtCol - 1];
        const itemPrice = menuData[i][mS.itemPriceCol - 1];
        const type = menuData[i][mS.formTypeCol - 1];
        const itemDetail = menuData[i][mS.itemDetailCol - 1];
        let itemTitle = itemName;
        if (itemAmount) {
          itemTitle += "（" + itemAmount + "）";
        }

        let itemDescription = `¥${itemPrice} /${itemUnit}`;
        if (itemDetail) {
          itemDescription += `\n${itemDetail}`;
        }

        let qItem;
        switch (type) {
          case "記述式":
            if (!itemUnit) {
              throw new Error(`${itemName} の単位を入力してください。`);
            }
            qItem = form.addTextItem().setTitle(itemTitle).setHelpText(`${itemDescription}\n\n※「${itemUnit}」単位で、数値のみを入力してください。`);
            qItem.setValidation(FormApp.createTextValidation()
                .setHelpText("半角数値でご入力ください。")
                .requireNumber()
                .build());
            break;
          case "プルダウン":
            const orderUpperLimit = menuData[i][mS.upperLimitCol - 1]; 
            if (!orderUpperLimit || orderUpperLimit < 1) {
              throw new Error(`${itemName} の注文上限値を設定してください。`);
            }
            let choiceValues = [];
            for (let j = 1; j <= orderUpperLimit; j++) {
              choiceValues.push([j]);
            }
            qItem = form.addListItem().setTitle(itemTitle).setHelpText(itemDescription).setChoiceValues(choiceValues);
            break;
          default:
            throw new Error(`${itemName} の「フォーム形式」の列を設定してください`);
        }

        const osColNumber = oS.menuStartCol + i;
        writeToMenuSheet(qItem.getId(), itemId, itemName, itemUnit, osColNumber);
      } 
    }
  } catch (error) {
    throw new Error(`Error in addMenuItems: メニューアイテムの追加中にエラーが発生しました。\n${error.message}`);
  }
  
}


function writeToMenuSheet(qId, itemId, itemName, itemUnit, colNum) {

  const lastRowOfFmArea = getLastDataRow(menuSheet, mS.fmStartCol);
  console.log(`lastRowOfFmArea = ${lastRowOfFmArea}`);
  menuSheet.getRange(lastRowOfFmArea + 1, mS.fmStartCol, 1, 5).setValues([[qId, itemId, itemName, itemUnit, colNum]]);
}

function confirm(message) {
  const ui = SpreadsheetApp.getUi();
  return ui.alert(message, ui.ButtonSet.YES_NO) === ui.Button.YES;
}



function getItemTitles(items) {
  let itemTitles = [];
  for (let i = 0; i < items.length; i++) {
    const title = items[i].getTitle();
    itemTitles.push(title);
  }
  return itemTitles;
}



function clearMenuSheetValues() {
  try {
    const lastRowOfFmArea = getLastDataRow(menuSheet, mS.fmStartCol);

    // ログを追加
    console.log(`mS.contStartRow: ${mS.contStartRow}`);
    console.log(`mS.fmStartCol: ${mS.fmStartCol}`);
    console.log(`lastRowOfFmArea = ${lastRowOfFmArea}`);

    const numRows = lastRowOfFmArea - mS.contStartRow + 1;

    console.log(`numRows: ${numRows}`);
    
    if (numRows > 0) {
      menuSheet.getRange(mS.contStartRow, mS.fmStartCol, numRows, 5).clearContent();
    }
  } catch (error) {
    throw new Error(`Error in clearMenuSheetValues: フォームメニューの初期化に失敗しました。\n${error.massage}`);
  }
}


function getQuestionId() {
  try {
    const firstQuestion = form.getItems()[0];
    if (firstQuestion && firstQuestion.getTitle() === "注文ID") {
      const questionId = firstQuestion.getId();
      menuSheet.getRange(mS.questionIdCell).setValue(questionId);
    } else {
      throw new Error("質問タイトル「注文ID」がフォームの質問に存在しません。");
    }
  } catch (error) {
    throw new Error(`Error in getQuestionId: 「注文ID」の質問IDが取得できませんでした。\n${error.massage}`);
  }
}

function getPublishedFormUrl() {
  try {
    const formUrl = form.getPublishedUrl();
    menuSheet.getRange(mS.formPublishedUrlCell).setValue(formUrl);
  } catch (error) {
    throw new Error(`Error in getPublishedFormUrl: 更新後のフォームURLの取得に失敗しました。${error.massage}`);
  }
}

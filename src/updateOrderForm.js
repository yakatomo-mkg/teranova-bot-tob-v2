function updateOrderForm() {
  try {
    const mS = AppConfig.menuSheet;
    const oS = AppConfig.orderSheet;

    const formTitle = form.getTitle();
    const message = `\n注文フォーム『${formTitle}』を更新してもよろしいですか？`;
    if (!confirm(message)) return;

    const lastRowOfItem = getLastDataRow(menuSheet, mS.itemNameCol);
    if (lastRowOfItem - mS.contStartRow + 1 === 0) {
      throw new Error("メニュー管理シートに商品を登録してください。");
    }
    const isCheckedData = menuSheet
      .getRange(mS.contStartRow, mS.ckBoxCol, lastRowOfItem - mS.contStartRow + 1, 1)
      .getValues().flat().some(Boolean);
    if (!isCheckedData) {
      throw new Error("メニュー管理シートにおいて、チェックされた項目が見つかりません。");
    }
    
    /** フォーム更新処理 */
    const menuData = menuSheet.getRange(mS.contStartRow, 1, lastRowOfItem, mS.itemDetailCol).getValues();

    // 既存の全アイテムを削除
    form.getItems().forEach(item => form.deleteItem(item));

    // メニュー管理エリアを初期化
    clearMenuSheetValues(menuSheet, ms);

    // 1) 基本質問（※注文IDはここでは追加しない）
    addBasicQuestions(form);

    // 2) 商品メニュー群
    addMenuItems({ form, menuSheet, mS, oS, menuData, lastRowOfItem });

    // 3) コメント
    form.addTextItem().setTitle(AppConfig.form.titles.COMMENT);

    // 4) フォーム末尾に「注文ID」を追加
    addOrderIdQuestionAtEnd(form);

    // 5) 質問ID と 公開URL を メニューシートに反映（フォーム更新の管理のため）
    setOrderIdQuestionIdToSheet(form, menuSheet, mS);
    setPublishedFormUrl(form, menuSheet, mS);

    SpreadsheetApp.getActive().toast("フォームの作成に成功しました。", "成功", 5);


  } catch (error) {
    showMessageDialog("updateOrderForm", error);
  }
}


/** 基本質問 */
function addBasicQuestions(form) {
  try {
    const t = AppConfig.formConfig.titles;
    form.addTextItem().setTitle(t.SHOP_NAME).setRequired(true);
    form.addListItem().setTitle(t.DELIVERY_DATE)
      .setRequired(true).setChoiceValues(generateDaysChoices(30));
  } catch (error) {
    throw new Error(`Error in addBasicQuestions: ${error.message}`);
  }
}


/** 末尾に「注文ID」を追加 */
function addOrderIdQuestionAtEnd(form) {
  try {
    form.addTextItem()
      .setTitle(AppConfig.formConfig.titles.ORDER_ID)
      .setHelpText("※ 注文管理用（編集不可）")
      .setRequired(true);
  } catch (error) {
    throw new Error(`Error in addOrderIdQuestionAtEnd: ${error.message}`);
  }
}


function addMenuItems({ form, menuSheet, mS, oS, menuData, lastRowOfItem }) {
  try {
    for (let i = 0; i < lastRowOfItem - 1; i++) {
      const isChecked = menuData[i][mS.ckBoxCol - 1];
      if (!isChecked) continue;

      const itemId     = menuData[i][mS.itemIdCol    - 1];
      const itemName   = menuData[i][mS.itemNameCol  - 1];
      const itemUnit   = menuData[i][mS.itemUnitCol  - 1];
      const itemAmount = menuData[i][mS.itemAmtCol   - 1];
      const itemPrice  = menuData[i][mS.itemPriceCol - 1];
      const type       = menuData[i][mS.formTypeCol  - 1];
      const itemDetail = menuData[i][mS.itemDetailCol- 1];

      let itemTitle = itemName;
      if (itemAmount) itemTitle += `（${itemAmount}）`;

      let itemDescription = `¥${itemPrice} /${itemUnit}`;
      if (itemDetail) itemDescription += `\n${itemDetail}`;

      let qItem;
      switch (type) {
        case "記述式":
          if (!itemUnit) throw new Error(`${itemName} の単位を入力してください。`);
          qItem = form.addTextItem().setTitle(itemTitle)
                  .setHelpText(`${itemDescription}\n\n※「${itemUnit}」単位で、数値のみを入力してください。`);
          qItem.setValidation(
            FormApp.createTextValidation().setHelpText("半角数値でご入力ください。").requireNumber().build()
          );
          break;
        case "プルダウン":
          const orderUpperLimit = menuData[i][mS.upperLimitCol - 1];
          if (!orderUpperLimit || orderUpperLimit < 1) {
            throw new Error(`${itemName} の注文上限値を設定してください。`);
          }
          const choiceValues = Array.from({ length: orderUpperLimit }, (_, j) => String(j + 1));
          qItem = form.addListItem().setTitle(itemTitle).setHelpText(itemDescription).setChoiceValues(choiceValues);
          break;
        default:
          throw new Error(`${itemName} の「フォーム形式」の列を設定してください`);
      }

      const osColNumber = oS.menuStartCol + i;
      writeToMenuSheet(menuSheet, mS, qItem.getId(), itemId, itemName, itemUnit, osColNumber);
    }
  } catch (error) {
    throw new Error(`Error in addMenuItems: ${error.message}`);
  }
}


function writeToMenuSheet(menuSheet, mS, qId, itemId, itemName, itemUnit, colNum) {
  const lastRowOfFmArea = getLastDataRow(menuSheet, mS.fmStartCol);
  menuSheet.getRange(lastRowOfFmArea + 1, mS.fmStartCol, 1, 5)
    .setValues([[qId, itemId, itemName, itemUnit, colNum]]);
}



function confirm(message) {
  const ui = SpreadsheetApp.getUi();
  return ui.alert(message, ui.ButtonSet.YES_NO) === ui.Button.YES;
}



function clearMenuSheetValues(menuSheet, mS) {
  try {
    const lastRowOfFmArea = getLastDataRow(menuSheet, mS.fmStartCol);
    const numRows = lastRowOfFmArea - mS.contStartRow + 1;
    if (numRows > 0) {
      menuSheet.getRange(mS.contStartRow, mS.fmStartCol, numRows, 5).clearContent();
    }
  } catch (error) {
    throw new Error(`Error in clearMenuSheetValues: ${error.message}`);
  }
}



/**「注文ID」の質問IDを、タイトル検索で取得してメニューシートに書き込む */
function setQuestionId(form, menuSheet, mS) {
  try {
    const t = AppConfig.formConfig.titles;
    const matches = form.getItems(FormApp.ItemType.TEXT)
      .filter(item => item.getTitle() === t.ORDER_ID);

    if (matches.length === 0) {
      throw new Error(`質問タイトル「${t.ORDER_ID}」が注文フォームに存在しません。`);
    }
    if (matches.length > 1) {
      throw new Error(`質問タイトル「${t.ORDER_ID}」が複数存在します。1つのみになるよう注文フォームを調整してください。`);
    }
    const questionId = matches[0].getId();
    menuSheet.getRange(mS.questionIdCell).setValue(questionId);
  } catch (error) {
    throw new Error(`Error in setOrderIdQuestionId: ${error.message}`);
  }
}



function setPublishedFormUrl(form, menuSheet, mS) {
  try {
    const formUrl = form.getPublishedUrl();
    menuSheet.getRange(mS.formPublishedUrlCell).setValue(formUrl);
  } catch (error) {
    throw new Error(`Error in setPublishedFormUrl: ${error.message}`);
  }
}
function createDeliverySlips() {
  try {
    const values = getFreeeStValues();

    const accessToken = getFreeeOAuth2Service().getAccessToken();

    const { successOrders, requests } = processOrders(values, accessToken);

    if (requests.length === 0) {
      throw new Error('チェックされた項目が一つもありません。');
    }

    postOrderToApi(values, requests, successOrders);

    const successMsg = '成功した注文ID:\n' + successOrders.map(order => order.orderId).join('\n');
    showMessageDialog("納品書作成に成功", `${successMsg}`);

    const successRows = successOrders.map(order => order.rowNum);
    console.log(successRows);
    rowColoredAndUnchecked(freeeSheet, successRows, "#808080");

    
  } catch (error) {
    showErrorDialog("createDeliSlips", error);
  }
}


function getFreeeStValues() {
  try {
    const lRDeliDate = getLastDataRow(freeeSheet, fS.deliDateCol);
    console.log(`lRDeliDate = ${lRDeliDate}`);
    const values = freeeSheet.getRange(fS.contStartRow, 1, lRDeliDate - fS.contStartRow + 1, freeeSheet.getLastColumn()).getValues();
    // console.log(`取得データ : ${values}`);
    // console.log(`${values[0][fS.orderIdCol - 1]}`);
    // console.log(`取得データ数 : ${values[0].length}`);

    // // valuesの各行データを出力してみる
    // values.forEach((row, index) => {
    //   /** index + fS.contStartRow = 行番号の取得 */
    //   console.log(row[0])  // 各行の1列目のセルの値(=チェックボックスのチェックの有無)
    //   console.log(`Row ${index + fS.contStartRow}: ${JSON.stringify(row)}`);
    //   console.log(fS.contStartRow);
    // })
    return values;
  } catch (error) {
    throw new Error(`Error in getFreeeStValues: freeeシートからのデータ取得中にエラーが発生しました。\n${error.message}`)
  }
  
}


function processOrders (values, accessToken) {
  try {
    const successOrders = [];
    const originalVals = [...values];

    const filteredVals = values.filter(row => row[0]);
    const requests = filteredVals.map((row, filteredIdx) => {
        const orginalRowNum = originalVals.indexOf(row) + fS.contStartRow;
        console.log(`${filteredIdx} - ${orginalRowNum}: ${row}`);
        return createOrderRequest(row, orginalRowNum, filteredIdx, filteredVals, accessToken); 
      })
      .filter(req => req !== null);
    
    return { successOrders, requests };
  } catch (error) {
    throw new Error(`Error in processOrders: 注文データ処理中にエラーが発生しました。\n${error.message}`);
  }
}


function createOrderRequest(row, rowNum, idx, values, accessToken) {
  try {
    const ourCompanyId = partnerSheet.getRange(pS.ourCompanyIdCell).getValue();
    const deliveryDate = formatDate(row[fS.deliDateCol - 1]);
 
    const orderData = {
      company_id: ourCompanyId,
      delivery_slip_date: deliveryDate, 
      tax_entry_method: 'out',
      tax_fraction: 'round',
      line_amount_fraction: 'round',
      withholding_tax_entry_method: 'in',
      memo: values[idx][fS.orderIdCol - 1],
      partner_id: values[idx][fS.partnerIdCol - 1],
      partner_title: '御中', 
      lines: JSON.parse(row[fS.menuCol - 1]).map(itemObj => {
        let description = itemObj.item;
        if (itemObj.amount.trim() !== '') {
          description += `（${itemObj.amount}）`;
        }
        return {
          type: 'item',
          description: description,
          unit: itemObj.unit,
          quantity: itemObj.quantity,
          unit_price: String(itemObj.price),
          tax_rate: 8,
          reduced_tax_rate: true,
          withholding: false,
        };
      }),       
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken},
      payload: JSON.stringify(orderData),
      muteHttpExceptions: true,
      timeout: 60000
    };
    const requestUrl = `${BASE_URL}/iv/delivery_slips`;
    return {
      url: requestUrl,
      options: options,
      rowNum: rowNum,
    };
  } catch (error) {
    throw new Error(`Error in createOrderRequest: リクエストデータ作成時にエラーが発生しました。\n${error.message}`);
  }
}


function postOrderToApi(originValues, requests, successOrders) {
  try {
    requests.forEach(req => {
      const res = UrlFetchApp.fetch(req.url, req.options);
      if (res.getResponseCode() === 201) {
        const orderId = originValues[req.rowNum - fS.contStartRow][fS.orderIdCol - 1];
        console.log({ "rowNum": req.rowNum, "orderId": orderId });
        successOrders.push({ rowNum: req.rowNum, orderId: orderId });
      } else {
        throw new Error(`API送信エラー\n注文ID: ${req.orderId}\nステータスコード: ${res.getResponseCode()}\nレスポンス: ${res.getContentText()}`)
      }
    });
  } catch (error) {
    throw new Error(`Error in postOrderToApi: API送信時にエラーが発生しました。\n${error.message}`);
  }
}

function formatDate(dateString) {
  try {
    const months = {
      '1月': '01',
      '2月': '02',
      '3月': '03',
      '4月': '04',
      '5月': '05',
      '6月': '06',
      '7月': '07',
      '8月': '08',
      '9月': '09',
      '10月': '10',
      '11月': '11',
      '12月': '12'
    };
    const pattern = /(\d{4})年(\d{1,2})月(\d{1,2})日/; // 「YYYY年MM月DD日」にマッチさせるための正規表現
    const match = dateString.match(pattern);

    if (!match) {
      throw new Error('Invalid date format');
    }
    const year = match[1];
    const month = months[`${match[2]}月`];
    const day = ("0" + match[3]).slice(-2);

    const formattedDate = `${year}-${month}-${day}`;
    // console.log(formattedDate);
    return formattedDate;
  } catch (error) {
    throw new Error(`Error in formatDate: 日付フォーマット作成時にエラーが発生しました。\n${error.message}`);
  }
}


// /**
//  * freeeAPIにリクエストを送信して納品書を作成する関数
//  * 
//  * @param  {String} アクセストークン
//  * @param  {Object} 納品書データ
//  * @returns {}
//  */
// function postDeliverySlips(accessToken, postData) {
//   Utilities.sleep(1000);

//   const options = {
//     method: 'post',
//     contentType: 'application/json',
//     headers: {Authorization: 'Bearer ' + accessToken},
//     payload: JSON.stringify(postData),
//     muteHttpExceptions: true,
//     timeout: 60000 // 60秒のタイムアウト設定
//   };

//   const requestUrl = `${BASE_URL}/iv/delivery_slips`;
//   const response = UrlFetchApp.fetch(requestUrl, options);
//   Logger.log(response.getResponseCode());
//   return response;
// }


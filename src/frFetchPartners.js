function fetchPartnersList() {
  const ourCompanyId = partnerSheet.getRange(pS.ourCompanyIdCell).getValue();
  console.log(ourCompanyId);

  const accessToken = getFreeeOAuth2Service().getAccessToken();
  const requestUrl = `${BASE_URL}/api/1/partners?company_id=${ourCompanyId}&limit=3000`;

  try {
    const response = accessFreeeAPI(accessToken, requestUrl);
    const partners = response.partners;
    
    if (partners && partners.length > 0) {
      const partnarsData = partners.map(partner => [partner.name, partner.id]);
      partnerSheet.getRange(pS.contStartRow, pS.nameCol, partnarsData.length, partnarsData[0].length).setValues(partnarsData);
    }

    concatIdAndName();
  } catch (error) {
    showErrorDialog("fetchPartnersList", error); 
  }
}


function accessFreeeAPI(accessToken, url) {
  Utilities.sleep(1000);
  const options = {
    method: 'get',
    headers: {'Authorization': 'Bearer ' + accessToken},
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);

  if (res.getResponseCode() !== 200) {
    throw new Error(`API access error\n${res.getResponseCode()}: ${res.getContentText()}`);
  }

  const resData = JSON.parse(res.getContentText());
  return resData;
}


/** 取引先名と取引先IDを結合した表示名列を作成 */
function concatIdAndName() {
  const startRow = pS.contStartRow;
  const dataLen = partnerSheet.getLastRow() - startRow + 1;
  const dataRange = partnerSheet.getRange(startRow, 1, dataLen, 2);
  const partnersData = dataRange.getValues();
  console.log(partnersData.length);
  
  let concatValues = [];
  for (let i = 0; i < partnersData.length; i++) {
    const name = partnersData[i][pS.nameCol - 1];
    const id = partnersData[i][pS.idCol - 1];


    // const cntVal = name + '　-　' + id;
    const cntVal = "[" + id + "]  " + name;
    concatValues.push([cntVal]);
  }
  // console.log(concatValues);

  partnerSheet.getRange(startRow, pS.displayCol, concatValues.length, 1).setValues(concatValues);
}






/**
 * 【参考】 事業所名一覧を取得して、シートに転記する関数 
 * @returns {void} 何も返さない (シートに値がセットされるのみ)
 */
const getOurCompanyList = () => {
  const pS = PARTNERS_SHEET_SETTINGS;
  const accessToken = getFreeeOAuth2Service().getAccessToken();
  const requestUrl = `${BASE_URL}/api/1/companies`;

  try {
    const response = accessFreeeAPI(accessToken, requestUrl);
    const companies = response.companies;
    if (companies && companies.length > 0) {
      const companyData = companies.map(company => [company.name, company.id]);
      partnerSheet.getRange(pS.ourCompanyListRow, pS.ourCompanyListCol, companyData.length, 2).setValues(companyData);
    }
  } catch (error) {
    showErrorDialog("getCompanyList", error);
  }
}




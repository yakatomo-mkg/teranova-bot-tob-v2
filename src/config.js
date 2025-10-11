const Properties = PropertiesService.getScriptProperties();

/** ---------------------------
 * バインドしているフォームの設定
 * ---------------------------- */
const form = FormApp.openById('xxxxxx');
const FORM_MENU_START_INDEX = 3;  // フォームにおけるメニュー項目開始インデックス


/** ---------------------------
 * LINEアカウント設定
 * ---------------------------- */

const LINE_CHANNEL_TOKEN = Properties.getProperty('LINE_CHANNEL_TOKEN');
/** ---------------------------
 * SpreadSheetの設定
 * ---------------------------- */
const SSID = 'yyyyyy';
const ss = SpreadsheetApp.openById(SSID);

const orderSheet = ss.getSheetByName("注文データ");
const menuSheet = ss.getSheetByName("フォーム作成");
const freeeSheet = ss.getSheetByName("納品書作成");
const partnerSheet = ss.getSheetByName("取引先");
const customerSheet = ss.getSheetByName("customer");
const cacheSheet = ss.getSheetByName("cache");
const logSheet = ss.getSheetByName("logs");
const adminSheet = ss.getSheetByName("管理者アカウント");



/** 各シートにおけるヘッダー列番号の定義 */
const MENU_SHEET_SETTINGS = {
  contStartRow: 3,  // 操作対象のコンテンツが開始する行(1,2行目はヘッダー)
  ckBoxCol: 1,  // チェックボックス設定列
  itemIdCol: 2,
  itemNameCol: 3,
  itemUnitCol: 4,
  itemAmtCol: 5,
  itemPriceCol: 6,
  upperLimitCol: 7,
  formTypeCol: 8,
  itemDetailCol: 9,
  fmStartCol: 12,  // フォームメニュー管理エリアの起点列
  questionIdCell: "S3",  // 「注文ID」の質問IDをセットするセル位置
  formPublishedUrlCell: "S4",  // 更新後の公開用フォームURLをセットするセル位置
};

const ORDER_SHEET_SETTINGS = {
  ckBoxCol: 1,
  acptDateCol: 2,
  orderIdCol: 3,
  shopNameCol: 4,
  deliDateCol: 5,
  commentCol: 6,
  menuStartCol: 7,
  contStartRow: 6,
};


const FREEE_SHEET_SETTINGS = {
  // ヘッダー列の設定
  ckBoxCol: 1,
  acptDateCol: 2,
  orderIdCol: 3,
  deliDateCol: 4,
  shopNameCol: 5,
  partnerNameCol: 6,
  partnerIdCol: 7,
  menuCol: 8,
  // その他の設定
  contStartRow: 2,  // コンテンツ(order情報がスタートする行)
}

const PARTNERS_SHEET_SETTINGS = {
  // ヘッダー列の設定
  nameCol: 1,
  idCol: 2,
  displayCol: 3,
  // その他の設定
  contStartRow: 3,  // コンテンツ(order情報がスタートする行)
  ourCompanyNameCell: "E3",
  ourCompanyIdCell: "F3",
  ourCompanyListRow: 10, // 自社事業所一覧のスタート行
  ourCompanyListCol: 5, // 自社事業所一覧のスタート列
}

const mS = MENU_SHEET_SETTINGS;
const oS = ORDER_SHEET_SETTINGS;
const fS = FREEE_SHEET_SETTINGS;
const pS = PARTNERS_SHEET_SETTINGS;


// /** -------------------------
//  * フォーム設定
//  * -------------------------*/
// const FORM_ID = Properties.getProperty('FORM_ID');
// const form = FormApp.openById(FORM_ID);


/** -------------------------------
 * freee設定
 * -------------------------------　*/
const BASE_URL = 'https://api.freee.co.jp';  // freee API エンドポイントのベース部分





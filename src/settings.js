/**************************************************************
 * settings.gs
 * 目的:
 *   - 外部リソース（フォーム/スプレッドシート/各シート/トークン）の参照を作成
 *   - アプリ全体で使う“純粋な定数（マジックナンバー/文言/列番号など）”を一元管理
 *
 * 設計方針:
 *   - 【副作用あり】外部アクセス（Properties/Spreadsheet/Form）はトップレベルのグローバルへ
 *   - 【副作用なし】定数は AppConfig オブジェクトに集約（単なるデータ）
 *   - コードからは「グローバルの参照（form/ss/...）」と「AppConfigの定数」を組み合わせて使う
 **************************************************************/



/* ============================================================
 * 外部リソース参照：runtime handles (副作用あり)
 *    - Script Properties からID/トークンを取得
 *    - フォーム/スプレッドシート/各シートの参照を作成
 *    - ここで実行時リソースを“1回だけ”開くことで、重複アクセスを抑制
 * ========================================================== */

// Properties helper
const ScriptProps = PropertiesService.getScriptProperties();
const requiredProp = (k) => {
  const v = ScriptProps.getProperty(k);
  if (!v) throw new Error(`Missing Script Property: ${k}`);
  return v;
};

// プロジェクト設定
const FORM_ID         = requiredProp('FORM_ID');
const SPREADSHEET_ID  = requiredProp('SPREADSHEET_ID');
const LINE_CHANNEL_TOKEN = requiredProp('LINE_CHANNEL_TOKEN');
const FREEE_CLIENT_ID    = requiredProp('FREEE_CLIENT_ID');    // OAuthクライアントID
const FREEE_CLIENT_SECRET= requiredProp('FREEE_CLIENT_SECRET');

// 環境（prod/dev）モード別の可変設定
const ENV = ScriptProps.getProperty('ENV') || 'prod';

// 外部サービスのハンドル
const form = FormApp.openById(FORM_ID);
const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

// シート参照（必要シートを名前解決して一括初期化）
const orderSheet    = ss.getSheetByName('注文データ');
const menuSheet     = ss.getSheetByName('メニュー管理');
const freeeSheet    = ss.getSheetByName('納品書作成');
const partnerSheet  = ss.getSheetByName('freee取引先');
const customerSheet = ss.getSheetByName('LINE顧客ID');
const cacheSheet    = ss.getSheetByName('cache');
const logSheet      = ss.getSheetByName('ErrorLog');
const adminSheet    = ss.getSheetByName('管理者アカウント');



/* ============================================================
 * 純粋定数：pure config  (副作用なし)
 *    - マジックナンバー/文言/列番号/セル位置などを一元化
 *    - “データのみ”のためどこから参照しても安全
 * ========================================================== */
const AppConfig = {
  form: {
    titles: {
      ORDER_ID: '注文ID',
      SHOP_NAME: '飲食店名（お名前）',
      DELIVERY_DATE: 'お届け日',
      COMMENT: 'コメント',
    },
    menuStartIndex: 3, // メニュー設問項目の開始インデックス
    pageHistorySegments: ['0','1'], // セクション到達保証用
  },

  sheetNames: {
    ORDER: '注文データ',
    MENU:  'メニュー管理',
    FREEE: '納品書作成',
    PARTNER: 'freee取引先',
    CUSTOMER: 'LINE顧客ID',
    ADMIN: '管理者アカウント',
    CACHE: 'cache',
    LOGS:  'ErrorLog',
  },

  menuSheet: {
    startRow: 3,
    ckBoxCol: 1,
    itemIdCol: 2,
    itemNameCol: 3,
    itemUnitCol: 4,
    itemAmtCol: 5,
    itemPriceCol: 6,
    upperLimitCol: 7,
    formTypeCol: 8,
    itemDetailCol: 9,
    fmStartCol: 12,
    questionIdCell: 'S3',
    formPublishedUrlCell: 'S4',
  },

  orderSheet: {
    ckBoxCol: 1,
    acptDateCol: 2,
    orderIdCol: 3,
    shopNameCol: 4,
    deliDateCol: 5,
    commentCol: 6,
    menuStartCol: 7,
    startRow: 6,
  },

  freeeSheet: {
    ckBoxCol: 1,
    acptDateCol: 2,
    orderIdCol: 3,
    deliDateCol: 4,
    shopNameCol: 5,
    partnerNameCol: 6,
    partnerIdCol: 7,
    menuCol: 8,
    startRow: 2,
  },

  partnersSheet: {
    nameCol: 1,
    idCol: 2,
    displayCol: 3,
    startRow: 3,
    ourCompanyNameCell: 'E3',
    ourCompanyIdCell: 'F3',
    ourCompanyListRow: 10,
    ourCompanyListCol: 5,
  },

  adminSheet: {
    uidCol: 1,  // LINE UID
    keywordCell: 'E2',  // 管理者登録キーワードのセル
    startRow: 2,
  },

  customerSheet: {
    uidCol: 1,      // LINE UID
    lineNameCol: 2,
    shopName: 3,
    freeeName: 4,
    freeeId: 5,
    startRow: 2,
  },

  line: {
    endpoints: {
      REPLY: 'https://api.line.me/v2/bot/message/reply',
      PUSH: 'https://api.line.me/v2/bot/message/push',
      MULTICAST: 'https://api.line.me/v2/bot/message/multicast',
      profile: function(uid) { 
        return 'https://api.line.me/v2/bot/profile/' + uid;
      },
    },
    backoff: {
      maxAttempts: 3,
      baseDelayMs: 500,
      retryStatuses: [429, 500, 502, 503, 504],
    },
    commands: {
    ORDER: '注文',
    NO: 'いいえ',
    QUESTION: 'お問い合わせ',
   },
    reply: {
      ORDER_CONFIRM: '注文を開始してよろしいですか？',
      CANCEL: '承知しました。何かございましたら、お気軽にお問い合わせください。',
      ANSWER: '通常のLINEメッセージ機能からお問い合わせいただけます。',
      ADMIN_REGISTERED: '管理者アカウントとして登録しました。',
    },
  },

  freee: {
    baseUrl: 'https://api.freee.co.jp',

    endpoints: {
      usersMe:      '/api/1/users/me',
      partners:     '/api/1/partners',
      deliverySlips:'/iv/delivery_slips', // 納品書API
    },

    // // 既定パラメータや制限（変更されやすいものはここに集約）
    // defaults: {
    //   partners: { limit: 500, order: 'asc' }, // 必要に応じて
    //   pagination: { limitMax: 500 },
    // },

    // // リトライ/レート制限方針を明示
    // backoff: {
    //   maxAttempts: 5,
    //   baseDelayMs: 250,    // 0.25s, 0.5s, 1s...
    //   retryStatuses: [401, 403, 429, 500, 502, 503, 504],
    // },
  },
};


// 固定質問タイトル集合：フォーム回答のメニュー項目判定で使用
const FIXED_QUESTIONS = new Set(Object.values(AppConfig.form.titles));











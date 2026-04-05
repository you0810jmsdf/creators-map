// ============================================================
// 千葉ニュータウン クリエイターズマップ
// Code.gs - メインスクリプト
// ============================================================

// ---- 設定 ここを環境に合わせて変更 ----
const SHEET_NAME    = 'フォームの回答 2'; // フォーム回答シート名
const MAPS_API_KEY  = 'AIzaSyB5oRAuKsGBhK3GwGAPAoKPAdOM5C1IwGs'; // Google Maps APIキー

// フォーム回答の列インデックス（0始まり）
const COL = {
  TIMESTAMP   : 0,  // タイムスタンプ
  EMAIL_AUTO  : 1,  // メールアドレス（フォーム自動収集）
  NAME        : 2,  // 屋号 / 名前
  CATEGORY    : 3,  // 分野
  WORKS       : 4,  // 代表作・制作物
  DESCRIPTION : 5,  // 一言紹介
  ADDRESS     : 6,  // 住所
  EMAIL       : 7,  // 公開用メールアドレス
  URL         : 8,  // 公開用SNS/URL
  CONSENT     : 9,  // 掲載への同意
  LAT         : 10, // 緯度（スクリプトが自動入力）
  LNG         : 11, // 経度（スクリプトが自動入力）
  APPROVED    : 12, // 公開承認フラグ（TRUE/FALSE）
};
// ---- 設定ここまで ----


// ============================================================
// 全カテゴリ一括インポート（定期実行用）
// Apps Scriptのトリガーで monthlyImportAll を月1回自動実行推奨
// 設定方法: Apps Script → トリガー → 追加 → 月次タイマー → monthlyImportAll
// ============================================================
function monthlyImportAll() {
  Logger.log('=== 月次一括インポート開始 ===');
  importWebLeatherCrafters();   // 革・レザー
  importWebCeramicsWood();      // 陶・木・金属
  importWebTextile();           // 布・糸・手芸
  importWebFabLab();            // ファブ・デジタル工作
  importWebPhoto();             // 映像・写真
  importWebIllustPainting();    // 絵・デザイン
  importWebAccessories();       // 陶・木・金属（アクセサリー）
  importWebPhotography();       // 映像・写真
  importWebCeramics();          // 陶・木・金属（陶芸）
  importWebWoodcraft();         // 陶・木・金属（木工）
  importWebDesigners();         // IT・Web
  importWebFoodFarm();          // 食・農・発酵
  importWebFabLabExtra();       // ファブ・デジタル工作（追加分）
  Utilities.sleep(2000);
  geocodeAll();                 // 緯度経度を自動補完
  Logger.log('=== 月次一括インポート完了 ===');
}


// ============================================================
// 月次トリガーを自動設定（初回1回だけ実行）
// 実行後はAppsScript → トリガー画面で確認できる
// ============================================================
function setupMonthlyTrigger() {
  // 既存の monthlyImportAll トリガーを削除してから再作成（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'monthlyImportAll') {
      ScriptApp.deleteTrigger(t);
      Logger.log('既存トリガーを削除しました');
    }
  }
  // 毎月1日の午前2時に実行
  ScriptApp.newTrigger('monthlyImportAll')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  Logger.log('月次トリガーを設定しました（毎月1日 午前2時）');
}


// ============================================================
// 【革・レザー】WEB収集データ追記
// ============================================================
function importWebLeatherCrafters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL]
  const DATA = [
    ['LEATHER工房YANAI',
     'オーダーメイド革製品・修理・リメイク',
     '埼玉県所沢市上新井2-1-11',
     'leather.yanai@gmail.com',
     'https://leather-yanai.com/'],

    ['革工房-Way-',
     '手縫い革財布・キーホルダー・革小物',
     '千葉県千葉市中央区中央3丁目18-1 B1F',
     'info@way03.com',
     'https://way03.com/'],

    ['Atelier K.I.（アトリエケーアイ）',
     '財布・名刺入れ・バッグ・オーダーメイド革製品',
     '東京都台東区蔵前4-20-12 クラマエビル3B',
     'info@atelier-ki.com',
     'https://atelier-ki.com/'],

    ['FILLY（フィリー）',
     'オーダーメイド革製品・財布・バッグ',
     '東京都武蔵野市中町1-3-3',
     '',
     'https://filly-leathers.com/'],

    ['革工房ZE-NO',
     '革製品製造・OEM・リペア・修理',
     '東京都足立区梅田5-24-10',
     'zeno.bag@gmail.com',
     'https://www.zeno-bag.com/'],

    ['Xartifact（エクサルティファクト）',
     'ハンドメイド・オーダーメイド革製品・財布・小物',
     '',
     'contact@xartifact-leather.com',
     'https://www.xartifact-leather.com/'],

    ['革工房NAUTS',
     'オーダーメイド革製品・財布・バッグ・革小物',
     '千葉県大網白里市みどりが丘2-19-6',
     'info@nauts.jp',
     'https://nauts.jp/'],

    ['革工房Rim',
     '財布・鞄・革小物・フルオーダーメイド',
     '京都府京都市中京区鍛治屋町377-1',
     '',
     'https://www.rim-works.com/'],

    ['革工房すだく',
     '革製品オーダー・修理・メンテナンス',
     '京都府京都市上京区中町通丸太町下ル駒之町554-2',
     '',
     'https://www.sudaku.jp/'],

    ['ERI.S LEATHER STUDIO',
     '財布・革小物・ハンドメイド',
     '大阪府枚方市楠葉花園町15-1',
     'kouhonori.eris@gmail.com',
     'https://erisleatherstudio.shopinfo.jp/'],

    ['Japlish（ジャプリッシュ）',
     '革小物・レザーバッグ・ハンドメイド',
     '福岡県福岡市博多区山王1-12-30 Abundant 80 1F',
     'japlish-shop@japlish.jp',
     'https://japlish.jp/'],

    ['革ノ花宗（はなむね）',
     'オーダーメイド革製品・レザークラフト体験',
     '福岡県久留米市津福本町1649-4 みなとビル103',
     '',
     'https://hanamune.net/'],

    ['柏革工房',
     'フルオーダー・セミオーダー革製品・刺繍',
     '北海道旭川市東光4条3丁目3-9',
     '',
     'https://kashikawa-kobo.com/'],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '革・レザー';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = '';   // geocodeAll() で自動補完
    row[COL.LNG]         = '';
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebLeatherCrafters 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【陶・木・金属】WEB収集データ追記
// ============================================================
function importWebCeramicsWood() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL]
  const DATA = [
    ['益子焼窯元よこやま',
     '益子焼・オーダーメイド食器・世界にひとつの器づくり',
     '栃木県芳賀郡益子町',
     '',
     'https://tougei.net/'],

    ['木工房ひのかわ',
     'オーダーメイド木製家具・無垢材家具',
     '熊本県八代郡氷川町宮原671-1',
     'info@hinokawa.jp',
     'https://www.hinokawa.jp/'],

    ['アクロージュファニチャー',
     '無垢材オーダー家具・木製家具制作',
     '東京都新宿区築地町6 北星ビル2階',
     '',
     'https://www.acroge-furniture.com/'],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '陶・木・金属';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebCeramicsWood 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【布・糸・手芸】WEB収集データ追記
// ============================================================
function importWebTextile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  const DATA = [
    ['アトリエ草乃しずか',
     '日本刺繍・手刺繍作品・刺繍教室',
     '東京都杉並区浜田山3-15-17',
     '',
     'https://kusano-shizuka.com/'],

    ['アトリエ森繍',
     '日本刺繍・伝統工芸刺繍作品',
     '京都府京都市北区上賀茂中大路町11-4',
     '',
     'https://morinui.jp/'],

    ['アトリエKAZUE',
     '刺繍教室・刺繍作品・ハンドメイド',
     '千葉県市川市八幡5-6-29',
     '',
     'https://www.a-kazue.com/'],

    ['atelier Ruto（アトリエ・ルト）',
     '棒針編み・かぎ針編み・手編み教室',
     '東京都練馬区（自宅サロン・住所非公開）',
     '',
     'https://atelierruto.com/'],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '布・糸・手芸';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebTextile 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【食・農・発酵】WEB収集データ追記
// ============================================================
function importWebFoodFarm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['柴海農園',
     '有機野菜・農薬不使用・年間60品目・直販・加工食品',
     '千葉県印西市瀬戸459-1',
     'info@shibakai-nouen.com',
     'https://shibakai-nouen.com/',
     35.817, 140.133],

    ['えか自然農場',
     '有機JAS認証野菜・農薬不使用・有機栽培',
     '千葉県流山市西深井352-2',
     '',
     'https://www.ecafarm.jp/',
     35.877, 139.921],

    ['つるかめ農園',
     '自然栽培米・農薬不使用・肥料不使用・いすみ産コシヒカリ',
     '千葉県いすみ市',
     '',
     'https://www.tsurukamefarm.com/',
     35.258, 140.373],

    ['株式会社 寺田本家',
     '自然酒・発酵食品・天然醸造・自然酒五人娘',
     '千葉県香取郡神崎町神崎本宿1964',
     '',
     'https://www.teradahonke.co.jp/',
     35.876, 140.539],

    ['鎌倉味噌醸造',
     '天然醸造・古式製法の味噌醤油・白味噌・赤味噌',
     '神奈川県藤沢市大鋸3-2-27-1',
     'kamakuramisojyouzou@gmail.com',
     'https://kamakuramiso.com/',
     35.333, 139.479],

    ['ヤマキ醸造',
     '国産有機JAS認定味噌・醤油・豆腐・体験工房',
     '埼玉県児玉郡神川町下阿久原955',
     '',
     'https://yamaki-co.com/',
     36.186, 139.098],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '食・農・発酵';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebFoodFarm 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【ファブ・デジタル工作】追加分（関東エリア）
// ============================================================
function importWebFabLabExtra() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  const DATA = [
    ['ファブラボみなとみらい',
     '3Dプリンター・レーザー加工機・デジタルファブリケーション',
     '神奈川県横浜市西区みなとみらい4-5-3 神奈川大学みなとみらいキャンパス1F',
     'fablab-minatomirai@kanagawa-u.ac.jp',
     'https://www.kanagawa-u.ac.jp/cooperation/project/fablab/',
     35.458, 139.638],

    ['ファブラボ鎌倉',
     '3Dプリンター・レーザーカッター・デジタル工作・ものづくり',
     '神奈川県鎌倉市扇ヶ谷1-10-6 結の蔵 壱号室',
     '',
     'https://www.fablabkamakura.com/',
     35.322, 139.546],

    ['西千葉工作室',
     'レーザーカッター・3Dプリンター・電子工作・ものづくりコミュニティ',
     '千葉県千葉市稲毛区緑町2-16-3 萩原ビル1F',
     '',
     'https://www.facebook.com/nishichibaksks',
     35.646, 140.104],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = 'ファブ・デジタル工作';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebFabLabExtra 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【ファブ・デジタル工作】WEB収集データ追記
// ============================================================
function importWebFabLab() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['ファブラボ神田錦町',
     '3Dプリンター・レーザーカッター・UVプリンター・デジタルファブリケーション',
     '東京都千代田区神田錦町3-20 アイゼンビル1F',
     'info@fablabkn.tokyo',
     'https://fablabkn.tokyo/',
     35.694, 139.759],

    ['ファブラボ博多',
     '3Dプリンター・レーザーカッター・デジタル工作機器',
     '福岡県福岡市南区市崎1-2-8 高宮マンション1F 142号室',
     '',
     'https://fablabhakata.com/',
     33.570, 130.421],

    ['AbenoMakerSpace FabLabβ',
     '3Dプリンター・レーザー加工機・デジタル工作スペース',
     '大阪府大阪市阿倍野区松崎町2-9-36',
     '',
     'https://www.abenomakerspace.com/',
     34.644, 135.513],

    ['おおたfab（ファブラボ蒲田）',
     '3Dプリンター・レーザーカッター・CNC・モノづくりスペース',
     '東京都大田区西蒲田7-49-2 OTELOビル4F',
     '',
     'https://ot-fb.com/fablab/',
     35.561, 139.716],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = 'ファブ・デジタル工作';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebFabLab 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【映像・写真】WEB収集データ追記
// ============================================================
function importWebPhoto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  const DATA = [
    ['フォトアトリエ エノ',
     '家族写真・マタニティ・ニューボーン・自然光写真',
     '京都府京都市右京区山ノ内宮前町5-19',
     '',
     'https://photoatelier-eno.com/'],
  ];

  let addCount = 0;
  const today = new Date();
  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) { Logger.log('スキップ（重複）: ' + name); continue; }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '映像・写真';
    row[COL.WORKS]       = d[1];
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.APPROVED]    = true;
    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }
  Logger.log('=== importWebPhoto 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 全国レザークラフター名簿をシートに直接追記
// Apps Scriptエディタで importLeatherCrafters を選択して「実行」
// ============================================================
function importLeatherCrafters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  // 重複チェック用（屋号が既に存在する行はスキップ）
  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // ── 転記データ（全国レザークラフター名簿 14件） ──
  // [NAME, WORKS, ADDRESS, EMAIL, URL, LAT, LNG]
  const DATA = [
    ['革工房 Rigel（リゲル）',          'タンニン鞣し革財布・革鞄・革小物',                          '北海道札幌市西区宮の沢2条3丁目15-15-2',         'rigel-otaruleather@outlook.com',    'https://www.rigel-leather.net/',                              43.0718, 141.28  ],
    ['北海道レザークラフト工房',          '革製品全般',                                                '北海道赤平市茂尻旭町1丁目15番地1',               '',                                  '',                                                            43.35,   141.97  ],
    ['J\'s LEATHER（ジェーズレザー）',   '財布・バッグ・革小物',                                      '宮城県仙台市宮城野区田子字新入6-1',               'jsleather@earth.so-net.jp',         'https://www.instagram.com/jsleather_sendai/',                 38.27,   140.91  ],
    ['革工房KIT',                        '手縫い革財布・バッグ・革小物',                              '宮城県仙台市宮城野区安養寺1-31-5',               '',                                  '',                                                            38.272,  140.915 ],
    ['革工房 きくわん舎',                '革バッグ・財布・革小物（カワイイ小物から本格バッグまで）',  '東京都武蔵野市境南町3-7-8',                      'koubou@kikuwansha.co.jp',           'https://www.instagram.com/kikuwansha/',                       35.705,  139.56  ],
    ['革工房OHANA',                      '財布・バッグ・革小物（世界に一つのオーダーメイド）',        '埼玉県さいたま市大宮区桜木町3-181-2',            'oomiya@ohana-online.jp',            'https://www.instagram.com/kawakobo_ohana/',                   35.9069, 139.6239],
    ['革工房 Bon Craft（ボンクラフト）', 'バッグ・クロコダイル製品・ランドセルリメイク',              '千葉県習志野市本大久保2-10-35 1F',               '',                                  '',                                                            35.68,   140.02  ],
    ['革工房 Pimu Factory',              '比翼開閉式折財布・革財布・革小物',                          '愛知県名古屋市西区菊井2-21-18',                  'info@pimufactory.com',              'https://www.instagram.com/pimufactory.leatherworks/',         35.19,   136.88  ],
    ['犬山革工房 vinculum leather',      '革財布・革小物（セミオーダー品）',                          '愛知県犬山市東古券677',                           '',                                  'https://www.instagram.com/vinculum__official/',               35.38,   136.94  ],
    ['京都山本製革店',                   '財布・バッグ（栃木レザー・イタリアンレザー使用）',          '京都府京都市山科区日ノ岡石塚町33-3',             '',                                  'https://www.instagram.com/yamamotoseika/',                    34.98,   135.81  ],
    ['革工房むくり',                     '革綴じミニトートバッグ・車掌かばん・こはぜカードケース',    '京都府京都市左京区下鴨森本町3 フォーレフジタ1F', 'arata@kobo-mukuri.com',             'https://www.instagram.com/mukuri_leather_kyoto/',             35.04,   135.77  ],
    ['大阪レザーアート（OLA）',          'バッグ・財布（1点ものオーダーメイド）',                     '大阪府大阪市中央区南船場2-10-17 南愛ビル606',    '',                                  'https://www.instagram.com/osaka_leather_art/',                34.68,   135.50  ],
    ['レザークラフトショップAGO',        'レザークラフト材料・工具',                                  '大分県別府市弓ヶ浜町6番26号',                    'agomonogatari77@gmail.com',         '',                                                            33.28,   131.49  ],
    ['皮革工房凜',                       '財布・バッグ・コースター',                                  '鹿児島県姶良市（詳細住所要確認）',               '',                                  'https://www.instagram.com/kawakoubou_rin/',                   31.73,   130.63  ],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '革・レザー';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【イラスト・絵画】WEB収集データ追記
// ============================================================
function importWebIllustPainting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['アトリエColors',
     '絵画・イラスト・漫画・デッサン教室（子どもから大人まで）',
     '千葉県印西市草深942 2階',
     '',
     'https://ateliercolors.net/',
     35.82, 140.13],

    ['絵画教室アトリエMIWA',
     '水彩・油絵・デッサン絵画教室',
     '千葉県千葉市中央区汐見丘町15-3',
     '',
     'https://arts-am.com/',
     35.61, 140.10],

    ['ナカジマカツ絵画教室',
     '油絵・デッサン・水彩・アクリル画',
     '千葉県船橋市芝山7-33-19',
     'mail@art-katsu.com',
     'http://art-katsu.com/',
     35.69, 140.02],

    ['アトリエこうたき',
     '絵画・デッサン・美大受験指導',
     '千葉県千葉市中央区今井2-2-17',
     '',
     'https://atelier-kohtaki.jimdo.com/',
     35.60, 140.10],

    ['アトリエ・オクダ',
     '水彩画・パステル画・絵本・イラスト',
     '千葉県柏市柏3-7-21椎名ビル703',
     '',
     'https://atelier-okuda.com/',
     35.87, 139.97],

    ['絵画教室1KAN.',
     'デッサン・アクリル・水彩絵画教室',
     '千葉県千葉市中央区今井2-2-17 ダイアパレス京葉蘇我２',
     '',
     'https://1kan-art.studio.site/',
     35.60, 140.10],

    ['柏美術学院カルチャー教室',
     'イラスト・デッサン・水彩絵画・油絵指導',
     '千葉県柏市柏4-5-3 TJビル2F',
     '',
     'https://artsalon-kashibi.com/',
     35.87, 139.97],

    ['アトリエ色',
     '絵画・造形教室（油絵・アクリル・デッサン）',
     '千葉県印西市大森3581-2',
     '',
     'https://ateliersiki.net/',
     35.80, 140.10],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '絵・デザイン';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebIllustPainting 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【アクセサリー・ジュエリー】WEB収集データ追記
// ============================================================
function importWebAccessories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['ついぶ柏工房',
     '手作り指輪・ペアリング・バングル・結婚指輪',
     '千葉県柏市末広町1-1 柏高島屋ステーションモール9F',
     '',
     'https://tsuibukashiwa.com/',
     35.87, 139.97],

    ['winwin1006（ウィンウィンテンシックス）',
     'シルバーアクセサリー・手作り指輪体験',
     '千葉県流山市後平井105-3',
     '',
     'https://www.instagram.com/winwin1006_yubiwa_tukuru/',
     35.86, 139.91],

    ['A☆ローズ',
     'ビーズ・UVレジン・ハンドメイドアクセサリー教室',
     '千葉県',
     '',
     'https://a-rose28.jimdofree.com/',
     35.61, 140.12],

    ['わくわくシルバーアクセサリー',
     'シルバーアクセサリー手作り体験・銀粘土インストラクター',
     '千葉市緑区あすみが丘',
     '',
     'https://waku2silverk.com/',
     35.57, 140.19],

    ['maroi（マロイ）',
     '手作り結婚指輪・オーダーメイドジュエリー',
     '千葉県千葉市中央区登戸1-23-1 二藤ビル2F',
     'info@maroi.jp',
     'https://maroi.jp/',
     35.61, 140.11],

    ['創作ジュエリー工房SILVER ART KAORI',
     '彫金・シルバーオリジナルジュエリー制作',
     '岩手県盛岡市中ノ橋通1-1-21 ホテルブライトイン盛岡2F',
     'ck@silverartkaori.com',
     'http://silverartkaori.com/',
     39.70, 141.15],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '陶・木・金属';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebAccessories 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【カメラ・写真】WEB収集データ追記
// ============================================================
function importWebPhotography() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['タカオカメラ',
     '出張カメラマン・家族写真・記念撮影・ポートレート',
     '東京・千葉エリア',
     '',
     'https://www.takaocamera.com/takaophoto/syuccyou.html',
     35.70, 139.80],

    ['出張撮影キュートワン',
     'お宮参り・七五三・入学式・家族写真の出張撮影',
     '千葉県松戸市',
     '',
     'https://www.cuteone-jp.com/',
     35.78, 139.90],

    ['at FOME（アットホーム）',
     '家族写真・ライフスタイルフォト・出張撮影',
     '千葉県',
     '',
     'https://www.atfome.com/photo-chiba/',
     35.61, 140.12],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '映像・写真';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebPhotography 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【陶芸】WEB収集データ追記
// ============================================================
function importWebCeramics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['千葉陶芸工房',
     '陶芸体験・創作陶芸・器制作',
     '千葉県千葉市中央区松波3-10-10 1階',
     '',
     'http://www.c-tougei.com/',
     35.61, 140.11],

    ['苅田郷陶芸教室（かったごう）',
     '陶芸教室・陶芸体験・うつわ制作',
     '千葉県千葉市緑区刈田子町',
     '',
     'https://www.kattago-tougei.com/',
     35.56, 140.18],

    ['晴栄窯',
     '陶芸・七宝焼工房・少人数制創作陶芸',
     '千葉県',
     '',
     'https://seieigama.jimdofree.com/',
     35.61, 140.12],

    ['陶芸家 神谷紀雄',
     '鉄絵銅彩・伝統工芸陶芸作品',
     '千葉県',
     '',
     'https://kamiyanorio.com/',
     35.61, 140.12],

    ['林寧彦の津田沼陶芸教室',
     '陶芸体験・陶芸教室（初心者歓迎）',
     '千葉県習志野市',
     '',
     'https://www.ne.jp/asahi/yasuhiko/hayashi/tsudanuma/studio.htm',
     35.68, 140.02],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '陶・木・金属';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebCeramics 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【木工】WEB収集データ追記
// ============================================================
function importWebWoodcraft() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['木工房六地蔵',
     '注文家具・雑貨・玩具・木工教室・オーダーメイド家具',
     '千葉県長生郡長柄町',
     '',
     'https://6jizoumokkou.wixsite.com/mysite',
     35.41, 140.24],

    ['千葉のオーダー家具工房 wood 凪',
     'オーダーメイド家具・無垢材木工品',
     '千葉県千葉市花見川区',
     '',
     'https://wood-nagi.com/',
     35.65, 140.06],

    ['WOOD STUDIO KUZE\'S',
     '家具工房・木工教室・オリジナル木製品',
     '千葉県',
     '',
     'https://kuze-s.com/',
     35.61, 140.12],

    ['家具工房Tabineko',
     'オーダーメイド家具・無垢材テーブル・木工クラフト',
     '千葉県山武市小松866-3',
     '',
     'https://www.tabinekokagu.com/',
     35.53, 140.35],

    ['秋元木工',
     '一枚板テーブル・オーダー家具・屋久杉・国産材家具',
     '千葉県君津市笹1782-10',
     '',
     'https://akimotomokko.com/',
     35.33, 139.94],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = '陶・木・金属';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebWoodcraft 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// 【ウェブデザイン・制作】WEB収集データ追記
// ============================================================
function importWebDesigners() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_NAME); return; }

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[COL.NAME]).trim()));

  // [屋号, 代表作・制作物, 住所, メール, URL, 緯度, 経度]
  const DATA = [
    ['フィット株式会社',
     'ホームページ制作・SEO対策・WEBコンテンツ制作',
     '千葉県印西市',
     '',
     'https://fit-jp.com/',
     35.82, 140.13],

    ['リアサポートデザインオフィス千葉本店',
     '格安ホームページ制作・WEBリニューアル・LP制作',
     '千葉県千葉市',
     '',
     'https://arm-ls.com/',
     35.61, 140.12],

    ['エアリーWEB事業部',
     'ECサイト・ホームページ制作・WEBデザイン',
     '千葉県千葉市',
     '',
     'https://www.airily.co.jp/web/',
     35.61, 140.12],

    ['WebClimb',
     'ホームページ制作・WEB制作・中小企業向けWEBデザイン',
     '千葉県',
     '',
     'https://www.webclimb.co.jp/hp-chiba/',
     35.61, 140.12],
  ];

  let addCount = 0;
  const today = new Date();

  for (const d of DATA) {
    const name = d[0];
    if (existingNames.has(name)) {
      Logger.log('スキップ（重複）: ' + name);
      continue;
    }
    const row = new Array(13).fill('');
    row[COL.TIMESTAMP]   = today;
    row[COL.EMAIL_AUTO]  = '';
    row[COL.NAME]        = name;
    row[COL.CATEGORY]    = 'IT・Web';
    row[COL.WORKS]       = d[1];
    row[COL.DESCRIPTION] = '';
    row[COL.ADDRESS]     = d[2];
    row[COL.EMAIL]       = d[3];
    row[COL.URL]         = d[4];
    row[COL.CONSENT]     = '';
    row[COL.LAT]         = d[5];
    row[COL.LNG]         = d[6];
    row[COL.APPROVED]    = true;

    sheet.appendRow(row);
    existingNames.add(name);
    addCount++;
    Logger.log('追加: ' + name);
  }

  Logger.log('=== importWebDesigners 完了: ' + addCount + '件追加 ===');
}


// ============================================================
// フォーム送信時トリガー：住所→緯度経度自動取得 + 同意でTRUE設定
// Apps Scriptのトリガー設定：onFormSubmit → フォーム送信時
// ============================================================
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  Logger.log('lastRow: ' + lastRow);

  // 住所を取得してジオコーディング
  const address = sheet.getRange(lastRow, COL.ADDRESS + 1).getValue();
  Logger.log('住所: ' + address);

  if (address) {
    const url = 'https://maps.googleapis.com/maps/api/geocode/json'
              + '?address=' + encodeURIComponent(address)
              + '&key=' + MAPS_API_KEY
              + '&language=ja';
    try {
      const res  = UrlFetchApp.fetch(url);
      const json = JSON.parse(res.getContentText());
      Logger.log('ジオコード結果: ' + JSON.stringify(json.status));
      if (json.results && json.results.length > 0) {
        const loc = json.results[0].geometry.location;
        Logger.log('緯度: ' + loc.lat + ' 経度: ' + loc.lng);
        sheet.getRange(lastRow, COL.LAT + 1).setValue(loc.lat);
        sheet.getRange(lastRow, COL.LNG + 1).setValue(loc.lng);
      } else {
        Logger.log('ジオコード結果なし: ' + json.status);
      }
    } catch(err) {
      Logger.log('ジオコードエラー: ' + err);
    }
  }

  // フォーム登録と同時に即時公開（管理者が不要なものをFALSEにして非掲載処置）
  sheet.getRange(lastRow, COL.APPROVED + 1).setValue(true);
  Logger.log('APPROVED=TRUE を設定しました（即時公開）');
}


// ============================================================
// WEB収集クリエイターデータを一括インポート
// 実行するたびに重複チェックして新規分のみ追加
// 追加後に geocodeAll() を実行して緯度経度を取得すること
// ============================================================
function importWebCreators() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  // 収集データ（WEB検索で確認済み）
  // カテゴリは9分野: 革・レザー／布・糸・手芸／陶・木・金属／ファブ・デジタル工作／
  //                  食・農・発酵／絵・デザイン／映像・写真／IT・Web／その他・複合
  const creators = [
    // ── 革・レザー ──────────────────────────────────────
    {
      name: 'LEATHER工房YANAI',
      category: '革・レザー',
      works: 'オーダーメイド革製品・修理・リメイク',
      description: '革製品のオーダーメイド、修復、リペア、リメイクまで。革専門店ならではの技術を提供。',
      address: '埼玉県所沢市上新井2-1-11',
      email: 'leather.yanai@gmail.com',
      url: 'https://leather-yanai.com/',
    },
    {
      name: '犬山革工房 vinculum leather',
      category: '革・レザー',
      works: 'レザークラフト教室・革製品販売',
      description: '本格コースから体験まで対応するレザークラフト教室と革製品販売。企業向けOEM生産も。',
      address: '愛知県犬山市東古券677',
      email: '',
      url: 'https://www.vinculumweb.net/',
    },
    {
      name: "M's FACTORY（エムズファクトリー）",
      category: '革・レザー',
      works: 'レザークラフト体験・オーダーメイド革製品',
      description: 'アットホームな革工房。プロ用工具を使ったレザークラフト体験やオーダーメイド革製品を提供。',
      address: '兵庫県姫路市石倉54-7 ざっぱ村',
      email: '',
      url: 'https://www.ms-leatherschool.com/',
    },
    // ── 陶・木・金属 ─────────────────────────────────────
    {
      name: '愚陶庵',
      category: '陶・木・金属',
      works: '陶芸作品制作・販売・体験教室',
      description: '西多摩の自然に囲まれた陶工房。作品の制作・販売と陶芸体験教室を開催。',
      address: '東京都西多摩郡日の出町平井1161-2',
      email: '',
      url: 'https://gutouan.com/',
    },
    {
      name: '家具工房 kinome',
      category: '陶・木・金属',
      works: '無垢材オーダーメイド家具',
      description: '栃木県那須町から無垢材のオーダーメイド家具をお届け。自然素材にこだわった一点もの家具を製作。',
      address: '栃木県那須郡那須町高久丙1147-251',
      email: '',
      url: 'https://www.kinome.jp/',
    },
    {
      name: 'TIDAFIKA WOOD STUDIO TOKYO',
      category: '陶・木・金属',
      works: '無垢材オーダーメイド家具',
      description: '20年以上の木工経験を持つ職人によるやさしい無垢材のオーダーメイド家具工房。',
      address: '東京都日野市大字上田472-1',
      email: '',
      url: 'https://tidafika.com/',
    },
    // ── 食・農・発酵 ──────────────────────────────────────
    {
      name: '株式会社 寺田本家',
      category: '食・農・発酵',
      works: '自然酒・発酵食品の醸造',
      description: '微生物との共生を通じて自然酒を醸造する千葉の蔵元。自然酒五人娘・発芽玄米酒むすひなどを製造。',
      address: '千葉県香取郡神崎町神崎本宿1964',
      email: '',
      url: 'https://www.teradahonke.co.jp/',
    },
    // ── 映像・写真 ────────────────────────────────────────
    {
      name: 'ウスダフォトスタジオ',
      category: '映像・写真',
      works: '写真撮影・フォトスタジオ',
      description: '1947年創業。大田区で写真館を長年営む街のホームフォトグラファー。家族の成長を大切に記録。',
      address: '東京都大田区中央1-21-1',
      email: '',
      url: 'https://usuda-photo.com/',
    },
    // ── 絵・デザイン ──────────────────────────────────────
    {
      name: '株式会社アトリエｍ',
      category: '絵・デザイン',
      works: 'グラフィックデザイン・各種印刷物制作',
      description: '横浜のグラフィックデザイン事務所。広告・パンフレット・ロゴなど各種デザインを手がける。',
      address: '神奈川県横浜市南区永田北1-5-9',
      email: 'info@atelier-m-design.com',
      url: 'https://atelier-m-design.com/',
    },
    {
      name: 'Atelier Grove（アトリエ・グローヴ）',
      category: '絵・デザイン',
      works: '透明水彩・アクリル絵画・商業イラスト',
      description: '心に残る風景を水彩で描くイラストレーター森シンジのアトリエ。広告・パッケージ等の商業アートも。',
      address: '',  // 住所非公開
      email: '',
      url: 'https://grove22.com/',
    },
  ];

  // 既存データとの重複チェック（屋号+URLで判定）
  const existingData = sheet.getDataRange().getValues();
  const existingKeys = new Set();
  for (let i = 1; i < existingData.length; i++) {
    const name = String(existingData[i][COL.NAME] || '').trim();
    const url  = String(existingData[i][COL.URL]  || '').trim();
    if (name) existingKeys.add(name + '|' + url);
  }

  let addCount = 0;
  for (const c of creators) {
    const key = c.name.trim() + '|' + c.url.trim();
    if (existingKeys.has(key)) {
      Logger.log('スキップ（重複）: ' + c.name);
      continue;
    }
    existingKeys.add(key);

    const newRow = new Array(13).fill('');
    newRow[COL.TIMESTAMP]   = new Date();
    newRow[COL.EMAIL_AUTO]  = '';
    newRow[COL.NAME]        = c.name;
    newRow[COL.CATEGORY]    = c.category;
    newRow[COL.WORKS]       = c.works;
    newRow[COL.DESCRIPTION] = c.description;
    newRow[COL.ADDRESS]     = c.address;
    newRow[COL.EMAIL]       = c.email;
    newRow[COL.URL]         = c.url;
    newRow[COL.CONSENT]     = '';
    newRow[COL.LAT]         = '';
    newRow[COL.LNG]         = '';
    newRow[COL.APPROVED]    = true;

    sheet.appendRow(newRow);
    addCount++;
    Logger.log('追加: ' + c.name + ' [' + c.category + ']');
  }

  Logger.log('=== 完了: ' + addCount + '件追加 / スキップ: ' + (creators.length - addCount) + '件 ===');
  if (addCount > 0) Logger.log('次に geocodeAll() を実行して緯度経度を取得してください。');
}


// ============================================================
// 既存データの緯度経度を一括再取得（初期セットアップ時に手動実行）
// ============================================================
function geocodeAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAME);
    return;
  }
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const address = data[i][COL.ADDRESS];
    const lat     = data[i][COL.LAT];
    if (!address || lat) continue; // 住所なし・取得済みはスキップ

    const url = 'https://maps.googleapis.com/maps/api/geocode/json'
              + '?address=' + encodeURIComponent(address)
              + '&key=' + MAPS_API_KEY
              + '&language=ja';
    try {
      const res  = UrlFetchApp.fetch(url);
      const json = JSON.parse(res.getContentText());
      if (json.results && json.results.length > 0) {
        const loc = json.results[0].geometry.location;
        sheet.getRange(i + 1, COL.LAT + 1).setValue(loc.lat);
        sheet.getRange(i + 1, COL.LNG + 1).setValue(loc.lng);
        Logger.log('行' + (i+1) + ': ' + address + ' → ' + loc.lat + ', ' + loc.lng);
      } else {
        Logger.log('行' + (i+1) + ' 結果なし: ' + json.status);
      }
      Utilities.sleep(300); // API制限対策
    } catch(err) {
      Logger.log('行' + (i+1) + ' エラー: ' + address + ' / ' + err);
    }
  }
  Logger.log('geocodeAll 完了');
}


// ============================================================
// Webアプリのエントリーポイント
// ============================================================
function doGet(e) {
  // ?mode=json のとき → JSON返却（Safari対応fetch方式）
  if (e && e.parameter && e.parameter.mode === 'json') {
    const data = getCreators();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 通常アクセス → HTMLを返す
  const template = HtmlService.createTemplateFromFile('index');
  template.apiKey = MAPS_API_KEY;
  return template.evaluate()
    .setTitle('千葉ニュータウン クリエイターズマップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// クリエイターデータをJSONで返す（index.htmlから呼び出し）
// APPROVED列がTRUE、または空欄の場合は公開する
// ============================================================
function getCreators() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const creators = [];
  const now30DaysAgo = new Date();
  now30DaysAgo.setDate(now30DaysAgo.getDate() - 30);

  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const lat      = row[COL.LAT];
    const lng      = row[COL.LNG];
    const approved = row[COL.APPROVED];

    // 緯度経度なし、またはFALSEで非承認のものは除外
    if (!lat || !lng) continue;
    if (approved === false || String(approved).toUpperCase() === 'FALSE') continue;

    const ts      = row[COL.TIMESTAMP];
    const isNew   = (ts instanceof Date) && (ts > now30DaysAgo);
    const address = String(row[COL.ADDRESS] || '');
    const city    = address.includes('印西市') ? '印西市'
                  : address.includes('白井市') ? '白井市'
                  : address.includes('栄町')   ? '栄町'
                  : 'その他';

    creators.push({
      name        : String(row[COL.NAME]        || '').trim(),
      category    : String(row[COL.CATEGORY]    || 'その他・複合').trim(),
      works       : String(row[COL.WORKS]        || '').trim(),
      description : String(row[COL.DESCRIPTION] || '').trim(),
      url         : String(row[COL.URL]          || '').trim(),
      lat         : lat,
      lng         : lng,
      isNew       : isNew,
      city        : city,
    });
  }

  return creators;
}


// ============================================================
// OSM（OpenStreetMap）からクリエイター候補を自動収集
// Apps Scriptエディタから手動実行 → APPROVED=falseで追加 → 管理者が確認・公開
//
// 【使い方】
//   importFromOSM()      → 印西市 or 千葉県 など単一エリアで検索
//   importFromOSMJapan() → 全国を都道府県ごとに分割して検索（時間がかかる）
// ============================================================

// ── 単一エリア検索（印西市 / 千葉県 / 関東） ──────────────────
// AREA_MODE: 'inzai'=印西市, 'chiba'=千葉県, 'kanto'=関東地方
function importFromOSM() {
  const AREA_MODE = 'chiba'; // ← ここを変更

  let areaQueryStr = '';
  let af = '';
  switch (AREA_MODE) {
    case 'inzai':
      areaQueryStr = 'area["name"="\u5370\u897f\u5e02"]["admin_level"="7"]->.a;';
      af = '(area.a)'; break;
    case 'chiba':
      areaQueryStr = 'area["name"="\u5343\u8449\u770c"]["admin_level"="4"]->.a;';
      af = '(area.a)'; break;
    case 'kanto':
      af = '(34.8,138.3,36.9,141.0)'; break;
    default:
      areaQueryStr = 'area["name"="\u5370\u897f\u5e02"]["admin_level"="7"]->.a;';
      af = '(area.a)';
  }

  Logger.log('=== importFromOSM / AREA_MODE: ' + AREA_MODE + ' ===');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  const existingKeys = buildExistingKeys_(sheet);
  const elements = fetchOSMElements_(areaQueryStr, af);
  const added = appendOSMRows_(sheet, elements, existingKeys);
  Logger.log('=== 完了: ' + added + '件追加 ===');
}


// ── 全国検索（都道府県ごとに分割） ────────────────────────────
// Apps Scriptの最大実行時間6分に注意。途中で止まった場合は再実行してください。
// 重複チェックが入るため、再実行しても同じデータは2重追加されません。
function importFromOSMJapan() {
  const PREFECTURES = [
    '北海道',
    '青森県','岩手県','宮城県','秋田県','山形県','福島県',
    '茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県',
    '新潟県','富山県','石川県','福井県','山梨県','長野県','岐阜県','静岡県','愛知県',
    '三重県','滋賀県','京都府','大阪府','兵庫県','奈良県','和歌山県',
    '鳥取県','島根県','岡山県','広島県','山口県',
    '徳島県','香川県','愛媛県','高知県',
    '福岡県','佐賀県','長崎県','熊本県','大分県','宮崎県','鹿児島県','沖縄県',
  ];

  Logger.log('=== importFromOSMJapan 開始（' + PREFECTURES.length + '都道府県） ===');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  let totalAdded = 0;
  const existingKeys = buildExistingKeys_(sheet);

  for (let i = 0; i < PREFECTURES.length; i++) {
    const pref = PREFECTURES[i];
    const areaQueryStr = 'area["name"="' + pref + '"]["admin_level"="4"]->.a;';
    Logger.log('[' + (i+1) + '/' + PREFECTURES.length + '] ' + pref + ' 検索中...');

    const elements = fetchOSMElements_(areaQueryStr, '(area.a)');
    const added = appendOSMRows_(sheet, elements, existingKeys);
    totalAdded += added;
    Logger.log('  → ' + elements.length + '件取得 / ' + added + '件追加');

    Utilities.sleep(3000); // Overpass API負荷軽減（3秒待機）
  }

  Logger.log('=== importFromOSMJapan 完了: 合計' + totalAdded + '件追加 ===');
}


// ── Overpass APIへ問い合わせ（内部関数） ─────────────────────
// 名前キーワードを1つの正規表現にまとめて高速化
function fetchOSMElements_(areaQueryStr, af) {
  // 全キーワードを1本の正規表現に統合（11クエリ→1クエリで大幅短縮）
  const nameRegex = [
    '\u5de5\u623f','\u30a2\u30c8\u30ea\u30a8','\u30b9\u30bf\u30b8\u30aa',
    '\u5de5\u82b8','\u9640\u82b8','\u9640\u5668',
    '\u6728\u5de5','\u6728\u5f6b','\u91d1\u5de5','\u935b\u51b6','\u5f6b\u91d1',
    '\u30ec\u30b6\u30fc','\u9769\u7d30\u5de5','\u76ae\u9769',
    '\u30cf\u30f3\u30c9\u30e1\u30a4\u30c9','\u624b\u82b8','\u5237\u7e94','\u7de8\u307f\u7269','\u30cb\u30c3\u30c8','\u67d3\u8272',
    '\u30d5\u30a1\u30d6\u30e9\u30dc','\u30e1\u30a4\u30ab\u30fc',
    '\u30c7\u30b6\u30a4\u30f3','\u30af\u30ea\u30a8\u30a4\u30bf\u30fc','\u4f5c\u5bb6','\u30a2\u30fc\u30c6\u30a3\u30b9\u30c8',
    '\u5199\u771f\u9928','\u30d5\u30a9\u30c8\u30b9\u30bf\u30b8\u30aa',
    '\u30ae\u30e3\u30e9\u30ea\u30fc','\u753b\u5ec8','\u7d75\u753b','\u30a4\u30e9\u30b9\u30c8','\u7f8e\u8853',
    '\u6620\u50cf','\u52d5\u753b','\u30d3\u30c7\u30aa',
    '\u91b8\u9020','\u767a\u9175','\u5473\u564c','\u9187\u6cb9','\u9152\u84b5',
    '\u8fb2\u5712','\u8fb2\u5834','\u8fb2\u5bb6','\u6709\u6a5f\u8fb2',
  ].join('|');

  const query = [
    '[out:json][timeout:120];',
    areaQueryStr,
    '(',
    'node["name"~"' + nameRegex + '"]' + af + ';',
    'node["craft"~"leather_worker|tailor|ceramics|pottery|carpenter|blacksmith|jeweller|brewery|painter|sculptor"]' + af + ';',
    'node["shop"~"art|photo|pottery|farm|organic|leather|fabric"]' + af + ';',
    'node["tourism"="gallery"]' + af + ';',
    'node["amenity"~"arts_centre|makerspace|studio|gallery"]' + af + ';',
    ');',
    'out body;',
  ].join('\n');

  try {
    const res = UrlFetchApp.fetch('https://overpass-api.de/api/interpreter', {
      method     : 'post',
      payload    : 'data=' + encodeURIComponent(query),
      contentType: 'application/x-www-form-urlencoded',
      muteHttpExceptions: true,
    });
    const statusCode = res.getResponseCode();
    if (statusCode !== 200) {
      Logger.log('HTTP エラー: ' + statusCode);
      return [];
    }
    const json = JSON.parse(res.getContentText());
    if (json.remark) Logger.log('Overpass remark: ' + json.remark);
    return json.elements || [];
  } catch (err) {
    Logger.log('fetchOSMElements_ エラー: ' + err);
    return [];
  }
}


// ── 既存データのキーセットを作成（重複チェック用） ──────────────
function buildExistingKeys_(sheet) {
  const data = sheet.getDataRange().getValues();
  const keys = new Set();
  for (let i = 1; i < data.length; i++) {
    const n = String(data[i][COL.NAME] || '').trim();
    if (n) keys.add(n + '|' + data[i][COL.LAT] + '|' + data[i][COL.LNG]);
  }
  return keys;
}


// ── OSM取得データをシートに追記（内部関数） ───────────────────
function appendOSMRows_(sheet, elements, existingKeys) {
  let addCount = 0;
  for (const el of elements) {
    const tags = el.tags || {};
    const name = (tags.name || '').trim();
    if (!name) continue;
    const lat = el.lat, lon = el.lon;
    if (!lat || !lon) continue;

    const key = name + '|' + lat + '|' + lon;
    if (existingKeys.has(key)) continue;
    existingKeys.add(key);

    const addrParts = [
      tags['addr:prefecture'] || tags['addr:province'] || '',
      tags['addr:city']       || '',
      tags['addr:quarter']    || tags['addr:suburb']   || '',
      tags['addr:street']     || '',
      tags['addr:housenumber']|| '',
    ].filter(Boolean);

    const newRow = new Array(13).fill('');
    newRow[COL.TIMESTAMP]   = new Date();
    newRow[COL.NAME]        = name;
    newRow[COL.CATEGORY]    = detectOSMCategory_(tags);
    newRow[COL.DESCRIPTION] = tags['description:ja'] || tags['description'] || '';
    newRow[COL.ADDRESS]     = addrParts.join('');
    newRow[COL.EMAIL]       = tags['email'] || tags['contact:email'] || '';
    newRow[COL.URL]         = tags['website'] || tags['contact:website'] || tags['url'] || '';
    newRow[COL.CONSENT]     = '';
    newRow[COL.LAT]         = lat;
    newRow[COL.LNG]         = lon;
    newRow[COL.APPROVED]    = true;

    sheet.appendRow(newRow);
    addCount++;
  }
  return addCount;
}


// ── OSMタグからカテゴリを自動判定（内部関数） ──────────────────
function detectOSMCategory_(tags) {
  const name    = tags.name    || '';
  const craft   = tags.craft   || '';
  const shop    = tags.shop    || '';
  const amenity = tags.amenity || '';
  const studio  = tags.studio  || '';
  const office  = tags.office  || '';

  if (craft === 'leather_worker' || shop === 'leather' ||
      /\u30ec\u30b6\u30fc|\u9769\u7d30\u5de5|\u76ae\u9769|\u9769\u5de5\u623f/.test(name)) return '\u9769\u30fb\u30ec\u30b6\u30fc';

  if (['tailor','embroidery'].includes(craft) || ['fabric','wool','sewing'].includes(shop) ||
      /\u624b\u82b8|\u5237\u7e94|\u7de8\u307f\u7269|\u30cb\u30c3\u30c8|\u67d3\u8272/.test(name)) return '\u5e03\u30fb\u7cf8\u30fb\u624b\u82b8';

  if (['ceramics','pottery','carpenter','wood_carver','blacksmith','jeweller'].includes(craft) ||
      shop === 'pottery' ||
      /\u9640\u82b8|\u9640\u5668|\u6728\u5de5|\u6728\u5f6b|\u91d1\u5de5|\u935b\u51b6|\u5f6b\u91d1/.test(name)) return '\u9640\u30fb\u6728\u30fb\u91d1\u5c5e';

  if (['makerspace','hackerspace'].includes(amenity) ||
      /\u30d5\u30a1\u30d6|\u30e1\u30a4\u30ab\u30fc|\u30c7\u30b8\u30bf\u30eb\u9020\u5f62/.test(name)) return '\u30d5\u30a1\u30d6\u30fb\u30c7\u30b8\u30bf\u30eb\u5de5\u4f5c';

  if (['brewery','winery','distillery'].includes(craft) || ['farm','organic'].includes(shop) ||
      /\u91b8\u9020|\u767a\u9175|\u5473\u564c|\u9187\u6cb9|\u9152\u84b5|\u8fb2\u5712|\u8fb2\u5834|\u8fb2\u5bb6|\u6709\u6a5f\u8fb2/.test(name)) return '\u98df\u30fb\u8fb2\u30fb\u767a\u9175';

  if (shop === 'photo' || studio === 'photo' ||
      /\u5199\u771f\u9928|\u30d5\u30a9\u30c8|\u6620\u50cf|\u52d5\u753b|\u30d3\u30c7\u30aa/.test(name)) return '\u6620\u50cf\u30fb\u5199\u771f';

  if (['painter','sculptor'].includes(craft) || shop === 'art' ||
      ['arts_centre','gallery'].includes(amenity) || tags.tourism === 'gallery' ||
      /\u30ae\u30e3\u30e9\u30ea\u30fc|\u753b\u5ec8|\u7d75\u753b|\u30a4\u30e9\u30b9\u30c8|\u7f8e\u8853|\u30c7\u30b6\u30a4\u30f3|\u30af\u30ea\u30a8\u30a4\u30bf\u30fc|\u4f5c\u5bb6/.test(name)) return '\u7d75\u30fb\u30c7\u30b6\u30a4\u30f3';

  if (['software','it'].includes(office)) return 'IT\u30fbWeb';

  return '\u305d\u306e\u4ed6\u30fb\u8907\u5408';
}

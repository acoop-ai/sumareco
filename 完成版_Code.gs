// ==========================================
// 店舗巡回・業務報告ボイスレコーダー (Backend)
// Version: 69.0 (PWA対応 doPost追加)
// ==========================================

/* 【セキュリティ重要】
   APIキーはコード内に記述せず、GASエディタの [プロジェクトの設定] > [スクリプト プロパティ] に保存してください。

   必須プロパティ:
   - GEMINI_API_KEY
   - OPENWEATHER_API_KEY
   - ALLOWED_IPS (任意: アクセス制限用)
*/

const PROPS = PropertiesService.getScriptProperties();

// プロパティからAPIキーを取得
const GEMINI_API_KEY = PROPS.getProperty('GEMINI_API_KEY');
const OPENWEATHER_API_KEY = PROPS.getProperty('OPENWEATHER_API_KEY');

// キー設定確認用ログ
if (!GEMINI_API_KEY || !OPENWEATHER_API_KEY) {
  console.error("【重要】APIキーがスクリプトプロパティに設定されていません。機能が動作しない可能性があります。");
}

// ★重要★ 既存のスプレッドシートIDを強制指定したい場合はここに記入
const MANUAL_SS_ID = '';

// データベース検索用ファイル名
const DB_FILE_NAME = "店舗報告データベース";
const REPORT_FOLDER_NAME = "報告画像・ログ";
const HQ_FOLDER_NAME = "本部用レポート";

// キャッシュキー
const CACHE_KEY_DASHBOARD = "DASHBOARD_JSON_V68_6";
const CACHE_TIME = 900; // 15分 (秒)

// --- システム定数 ---
const REPORT_CATS = {
  DAILY: "日報",
  PATROL: "巡回",
  PROPOSAL: "提案",
  REPLY: "返信",
  COMPLAINT: "苦情"
};
const CAT_PATROL = "巡回";

// ▼ 外部PWAから接続を許可する会社ドメイン（複数可）
const ALLOWED_DOMAINS_PWA = ['jaz-acoop.co.jp', 'acoop1.jp', 'acoop2.jp', 'acoop3.jp'];

/* 店舗リスト (Master) */
// ※ここの住所等は初期値です。正確な緯度経度はスプレッドシートのMasterから取得します。
const INITIAL_SHOPS = [
  // 首都圏1G
  {name: "Aコープ 城山店", address: "神奈川県相模原市緑区向原2-1-1"},
  {name: "Aコープ 中田店", address: "神奈川県横浜市泉区中田南3-1-10"},
  {name: "Aコープ 原宿店", address: "神奈川県横浜市戸塚区原宿4-15-4"},
  {name: "Aコープ 金沢店", address: "神奈川県横浜市金沢区谷津町35"},
  {name: "Aコープ タケヤマ店", address: "神奈川県横須賀市林1-28-5"},
  {name: "Aコープ ゆがわら店", address: "神奈川県足柄下郡湯河原町土肥5-7-1"},
  {name: "Aコープ 旭店", address: "神奈川県平塚市徳延563-2"},
  // 首都圏2G
  {name: "Aコープ 緑竹山店", address: "神奈川県横浜市緑区竹山3-1-8"},
  {name: "Aコープ 長沢店", address: "神奈川県横須賀市長沢1-34-1"},
  {name: "Aコープ 善行店", address: "神奈川県藤沢市善行1-26-6"},
  {name: "Aコープ 仙石原店", address: "神奈川県足柄下郡箱根町仙石原230"},
  // 北関東JAF
  {name: "JAファーマーズ 野田宿", address: "群馬県北群馬郡吉岡町上野田1050-3"},
  {name: "JAファーマーズ 朝日町", address: "群馬県前橋市朝日町1-38-21"},
  {name: "JAファーマーズ 安中", address: "群馬県安中市原市634"},
  {name: "JAファーマーズ ブレイス", address: "群馬県太田市新田市野井町438-1"},
  {name: "JAファーマーズ 朝倉町", address: "群馬県前橋市朝倉町143-1"},
  {name: "JAファーマーズ 太田藪塚", address: "群馬県太田市大原町2311-1"},
  {name: "JAファーマーズ 富岡", address: "群馬県富岡市富岡1878-1"},
  {name: "JAファーマーズ 高崎吉井", address: "群馬県高崎市吉井町片山448-1"},
  {name: "JAファーマーズ あがつま", address: "群馬県吾妻郡東吾妻町大字原町5116"},
  {name: "JAファーマーズ 高崎棟高", address: "群馬県高崎市棟高町1675-37"},
  {name: "JAファーマーズ 前橋川原", address: "群馬県前橋市川原町2-4-9"},
  // 北関東4G
  {name: "Aコープ 下仁田店", address: "群馬県甘楽郡下仁田町下仁田383-3"},
  {name: "Aコープ 北橘店", address: "群馬県渋川市北橘町真壁1386-1"},
  {name: "Aコープ ハピネス店", address: "群馬県富岡市中高瀬400-1"},
  {name: "Aコープ 笠懸店", address: "群馬県みどり市笠懸町鹿246-1"},
  {name: "Aコープ みやぎ店", address: "群馬県前橋市鼻毛石町198-11"},
  {name: "Aコープ 松井田店", address: "群馬県安中市松井田町松井田305"},
  {name: "新鮮ぐんま みのり館", address: "群馬県前橋市亀里町1307-1"},
  // 北東北
  {name: "Aコープ ごしょ店", address: "岩手県岩手郡雫石町西安庭40-48-1"},
  {name: "Aコープ ゆざわ店", address: "秋田県湯沢市材木町1-1-1"},
  {name: "Aコープ 飯岡駅前店", address: "岩手県盛岡市永井20-13-10"},
  {name: "Aコープ ゆぐち店", address: "岩手県花巻市円万寺法船134-3"},
  {name: "Aコープ ふじさわ店", address: "岩手県一関市藤沢町藤沢字町裏99-2"},
  {name: "Aコープ ひがしやま店", address: "岩手県一関市東山町長坂字西本町123-2"},
  {name: "ふれあい純情市場 さっこら", address: "岩手県盛岡市仙北2-5-4"},
  {name: "JAファーマーズ いわて平泉", address: "岩手県西磐井郡平泉町平泉字高田48"},
  {name: "Aコープ もりよし店", address: "秋田県北秋田市米内沢字出向59"},
  // 宮城
  {name: "Aコープ 色麻店", address: "宮城県加美郡色麻町四竃字北谷地41"},
  {name: "Aコープ 槻木店", address: "宮城県柴田郡柴田町槻木下町3-2-20"},
  {name: "Aコープ かしまだい店", address: "宮城県大崎市鹿島台平渡字西銭神20-1"},
  {name: "Aコープ こごた店", address: "宮城県遠田郡美里町字素山12-9"},
  {name: "Aコープ 古川店", address: "宮城県大崎市古川北町5-3-12"},
  {name: "Aコープ 沼部店", address: "宮城県大崎市田尻沼部字富岡183"},
  {name: "Aコープ 松島店", address: "宮城県宮城郡松島町高城字町13-1"},
  {name: "Aコープ 角田店", address: "宮城県角田市角田字田町112-1"},
  // 山形
  {name: "Aコープ あつみ店", address: "山形県鶴岡市湯温海字湯乃里181"},
  {name: "Aコープ ふじしま店", address: "山形県鶴岡市藤島字矢立57-2"},
  {name: "Aコープ はぐろ店", address: "山形県鶴岡市羽黒町野荒町字北田19"},
  {name: "Aコープ あさひ店", address: "山形県鶴岡市下名川字落合7"},
  {name: "Aコープ たちかわ店", address: "山形県東田川郡庄内町狩川字楯下98-1"},
  {name: "Aコープ ゆざ店", address: "山形県飽海郡遊佐町遊佐字広表6-1"},
  {name: "Aコープ ふくら店", address: "山形県飽海郡遊佐町吹浦字苗代37"},
  {name: "Aコープ やわた店", address: "山形県酒田市観音寺字町後22"},
  {name: "Aコープ にしき町店", address: "山形県酒田市坂野辺新田字古川18-1"}
];

// --- Utilities ---

function normalize(str) {
  if (!str) return "";
  return String(str).normalize('NFKC').replace(/[\s\u3000]+/g, "").trim().toLowerCase();
}

function parseCoordinate(val) {
  if (val == null || val === "") return null;
  const s = String(val).normalize('NFKC').replace(/[^\d.\-]/g, "");
  const num = parseFloat(s);
  return isNaN(num) ? null : num;
}

function safeStr(val, defaultVal = "") {
  if (val === null || val === undefined) return defaultVal;
  return String(val);
}

function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/\/d\/(.+?)\/|\?id=(.+?)$/);
  return match ? (match[1] || match[2]) : null;
}

function parseCategory(rawVal) {
  const s = String(rawVal || "").normalize('NFKC').replace(/[\s\u3000【】\(\)\[\]（）「」]/g, "");
  if (s.includes("巡回")) return CAT_PATROL;
  if (s.includes("苦情") || s.includes("トラブル")) return REPORT_CATS.COMPLAINT;
  if (s.includes("提案") || s.includes("気付き") || s.includes("気づき")) return REPORT_CATS.PROPOSAL;
  if (s.includes("返信") || s.includes("共有")) return REPORT_CATS.REPLY;
  return REPORT_CATS.DAILY;
}

function checkIpPermission() {
  const allowedIps = PROPS.getProperty('ALLOWED_IPS');
  if (!allowedIps) return true;
  return true;
}

// --- ID / Folder / Sheet Management ---

function getSpreadsheet() {
  let ss = null;
  if (MANUAL_SS_ID) { try { ss = SpreadsheetApp.openById(MANUAL_SS_ID); } catch(e) {} }
  if (!ss) {
    const savedId = PROPS.getProperty('SPREADSHEET_ID');
    if (savedId) { try { ss = SpreadsheetApp.openById(savedId); } catch(e) {} }
  }
  if (!ss) {
    try {
      const files = DriveApp.getFilesByName(DB_FILE_NAME);
      if (files.hasNext()) ss = SpreadsheetApp.openById(files.next().getId());
    } catch(e) {}
  }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {} }

  if (ss) {
    PROPS.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }
  throw new Error("データベースが見つかりません。「システム初期設定」を実行してください。");
}

function getReportsFolder() {
  const savedId = PROPS.getProperty('FOLDER_ID_REPORTS');
  if (savedId) {
    try {
      const folder = DriveApp.getFolderById(savedId);
      if (folder.getName() === REPORT_FOLDER_NAME) return folder;
    } catch(e) {}
  }

  try {
    const ss = getSpreadsheet();
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

    const folders = parentFolder.getFoldersByName(REPORT_FOLDER_NAME);
    let targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(REPORT_FOLDER_NAME);

    PROPS.setProperty('FOLDER_ID_REPORTS', targetFolder.getId());
    return targetFolder;
  } catch(e) {
    throw new Error("保存フォルダ取得エラー: " + e.message);
  }
}

function getHqMonthFolder(dateObj) {
  const ss = getSpreadsheet();
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  let hqFolder;
  const hqFolders = parentFolder.getFoldersByName(HQ_FOLDER_NAME);
  if (hqFolders.hasNext()) {
    hqFolder = hqFolders.next();
  } else {
    hqFolder = parentFolder.createFolder(HQ_FOLDER_NAME);
  }

  const monthStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM');
  const monthFolders = hqFolder.getFoldersByName(monthStr);
  if (monthFolders.hasNext()) {
    return monthFolders.next();
  } else {
    return hqFolder.createFolder(monthStr);
  }
}

// --- Web App Entry Point ---
function doGet() {
  if (!checkIpPermission()) return HtmlService.createHtmlOutput("Access Denied");
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
    .setTitle('Voice Report System V67')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- User Context ---
function getUserProfile() {
  try {
    const email = Session.getActiveUser().getEmail();
    const accountName = email.split('@')[0];

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    if (!sheet) throw new Error("Master sheet missing");

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) {
        return {
          id: safeStr(data[i][0]),
          name: safeStr(data[i][1]),
          email: safeStr(data[i][2]),
          account: accountName,
          role: safeStr(data[i][3]),
          shop: safeStr(data[i][4]),
          address: safeStr(data[i][5])
        };
      }
    }
    return { name: "ゲスト", shop: "未登録", role: "GUEST", email: email, account: accountName, error: "未登録ユーザー" };
  } catch (e) {
    return { name: "接続エラー", shop: "Error", role: "GUEST", email: "", account: "Unknown", error: e.toString() };
  }
}

// --- Report Submission (Voice) ---
function submitVoiceReport(p) {
  const user = getUserProfile();
  if (user.role === 'GUEST') throw new Error("ユーザー登録されていません: " + user.email);

  const id = Utilities.getUuid();
  const ts = new Date();

  let gemini = { summary: { items: [] }, transcript: "" };
  try {
    const result = analyzeWithGemini(p.audioBase64, p.mimeType, p.mode);
    if (result) gemini = result;
  } catch(e) {
    console.error("Gemini Fail: " + e);
    gemini.transcript = "AI解析エラー: " + e.message;
    gemini.summary = { items: [{header: "エラー", content: "AI解析中に問題が発生しました。原文をご確認ください。"}] };
  }

  let wInfo = {weather:"-", temp:"-"};
  try {
    const ss = getSpreadsheet();
    const wSheet = ss.getSheetByName('WeatherLog');
    const lastRow = wSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 100);
    const wData = wSheet.getRange(startRow, 1, lastRow - startRow + 1, 13).getValues();

    for(let i=wData.length-1; i>=0; i--){
      if(normalize(wData[i][1])===normalize(user.shop)) {
        wInfo = {weather: safeStr(wData[i][2]), temp: safeStr(wData[i][3]) + "℃"}; break;
      }
    }
  } catch(e){}

  const docUrl = createReportDoc(id, ts, user, p.mode, gemini.summary, gemini.transcript, p.imageUrls, wInfo);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');

  const sumText = (gemini.summary && gemini.summary.items && gemini.summary.items.length > 0)
    ? gemini.summary.items.map(i => `【${safeStr(i.header)}】${safeStr(i.content)}`).join("\n")
    : (gemini.transcript || "AI解析失敗（要約なし）");

  sheet.appendRow([
    id, ts, user.shop, user.role, user.name, user.email, p.mode, p.isEmergency ? "🚨" : "通常",
    sumText, gemini.transcript, (p.imageUrls||[]).join("\n"), docUrl, "未確認", ""
  ]);

  return {success:true, summary:sumText};
}

function analyzeWithGemini(b64, m, mode) {
  if (!GEMINI_API_KEY) throw new Error("Gemini API Key is not set in Script Properties.");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const prompt = `あなたは優秀な報告書作成アシスタントです。音声解析結果から「${mode}」の報告を作成してください。

【処理の優先順位】
【処理の優先順位】
1. **正規化（文字起こし補正）**: 音声に含まれる「あー」「えー」などの淀みや雑音への言及を除去し、話し言葉を**正確かつ端的**なビジネス文章に整えてください。ぶっきらぼうな表現であっても、意味を変えずに、報告として適切な**簡潔な言い回し**（「だ・である」調や「箇条書きに近い文体」など）に変換してください。元の情報の削除や要約はせず、事実関係をすべて維持したまま修正してください。
2. **AI要約の廃止**: 報告内容はすべて正規化された原文を使用するため、個別の要約アイテムは作成しないでください（summary.itemsは空配列にしてください）。

【出力形式】
JSON形式を厳守してください:
{"summary":{"items":[]},"transcript":"音声内容をすべて正規化した全文"}
`;
  const payload = {contents:[{parts:[{text:prompt},{inline_data:{mime_type:m,data:b64}}]}],generationConfig:{response_mime_type:"application/json"}};

  const res = UrlFetchApp.fetch(url, {method:'post',contentType:'application/json',payload:JSON.stringify(payload), muteHttpExceptions: true});
  const resText = res.getContentText();
  const resJson = JSON.parse(resText);

  if (res.getResponseCode() !== 200) {
    const errorMsg = resJson.error ? resJson.error.message : "HTTP Error " + res.getResponseCode();
    throw new Error("Gemini API Error: " + errorMsg);
  }

  const txt = resJson.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!txt) throw new Error("Empty response from Gemini (No candidates)");

  try {
    return JSON.parse(txt.replace(/^```json\s*/i, "").replace(/^```\s*/, "").replace(/```\s*$/, ""));
  } catch(e) {
    console.error("JSON Parse Error:", txt);
    return { summary: { items: [] }, transcript: txt };
  }
}

function summarizeTextWithGemini(promptText) {
  if (!GEMINI_API_KEY) throw new Error("Gemini API Key is not set.");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { response_mime_type: "application/json" }
  };
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const resText = res.getContentText();
  if (res.getResponseCode() !== 200) throw new Error("Gemini Summarize Error: " + resText);

  const txt = JSON.parse(resText).candidates?.[0]?.content?.parts?.[0]?.text;
  return JSON.parse(txt.replace(/^```json\s*/i, "").replace(/^```\s*/, "").replace(/```\s*$/, ""));
}

function createReportDoc(id, date, user, mode, summaryData, transcript, imgs, w) {
  try {
    const emojiMap = {
       '日報': '📝', '定型日報': '📝',
       '苦情': '⚠️', 'トラブル': '⚠️', '苦情・トラブル': '⚠️',
       '提案': '💡', '気付き': '💡', '気付き・提案': '💡',
       '返信': '↩️', '共有': '↩️', '返信・共有': '↩️',
       '巡回': '📋', '巡回報告': '📋'
    };
    const emoji = emojiMap[mode] || '📄';
    const folder = getReportsFolder();
    const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
    const timeStr = Utilities.formatDate(date, 'Asia/Tokyo', 'HH:mm');
    const fileName = `【${mode}】${user.account}_${dateStr}`;
    const files = folder.getFilesByName(fileName);

    let doc;
    let isNew = false;

    if (files.hasNext()) {
      doc = DocumentApp.openById(files.next().getId());
    } else {
      doc = DocumentApp.create(fileName);
      folder.addFile(DriveApp.getFileById(doc.getId()));
      isNew = true;
    }

    const body = doc.getBody();
    if (isNew) {
      body.insertParagraph(0, `${emoji} ${mode} (${dateStr})`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(`店舗: ${user.shop}\nアカウント: ${user.account}\n担当者: ${user.name}`);
      body.appendHorizontalRule();
    }

    body.appendParagraph(`■ ${emoji} ${mode} (送信: ${timeStr})`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const weatherP = body.appendParagraph(`気象情報: ${w.weather} / 気温: ${w.temp}`);
    weatherP.setItalic(true).setForegroundColor('#64748b');

    body.appendParagraph("【報告内容】").setBold(true);
    const transP = body.appendParagraph(transcript || "（文字データなし）");
    transP.setSpacingAfter(10);

    if (imgs && imgs.length > 0) {
      body.appendParagraph("\n[添付画像]").setBold(true);
      imgs.forEach(url => {
        try {
          const fileId = extractFileIdFromUrl(url);
          if (fileId) {
            const blob = DriveApp.getFileById(fileId).getBlob();
            body.appendImage(blob).setWidth(300).setHeight(300);
          } else {
             body.appendParagraph(`(画像URL: ${url})`);
          }
        } catch(e) { body.appendParagraph(`(画像読込エラー: ${e.message})`); }
      });
    }

    body.appendHorizontalRule();
    doc.saveAndClose();
    return doc.getUrl();

  } catch(e) {
    console.error("Doc Error: " + e);
    return "";
  }
}

// --- Image Upload ---
function uploadSingleImage(base64Data, mimeType) {
  try {
    const folder = getReportsFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, "upload_" + Utilities.getUuid());
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    return file.getUrl();
  } catch(e) {
    throw new Error("Upload Error: " + e.message);
  }
}

// --- Weather Engine (Fetching & Saving) ---

function manualWeatherUpdate() {
  updateWeatherAllShops();
  return "天気情報を更新しました。";
}

function updateWeatherAllShops() {
  const ss = getSpreadsheet();

  let wSheet = ss.getSheetByName('WeatherLog');
  if (!wSheet) wSheet = ss.insertSheet('WeatherLog');
  let fSheet = ss.getSheetByName('ForecastLog');
  if (!fSheet) fSheet = ss.insertSheet('ForecastLog');

  const mSheet = ss.getSheetByName('Master');
  const mData = mSheet.getDataRange().getValues();
  const shops = [];
  const processedNames = new Set();

  for (let i = 1; i < mData.length; i++) {
    const name = safeStr(mData[i][4]);
    const normName = normalize(name);
    if (!normName || processedNames.has(normName)) continue;

    const lat = mData[i][6];
    const lng = mData[i][7];
    if (lat && lng) {
      shops.push({ name: name, lat: lat, lng: lng });
      processedNames.add(normName);
    }
  }

  if (shops.length === 0) { console.warn("有効な店舗がありません"); return; }

  const now = new Date();

  shops.forEach(shop => {
    try {
      fetchAndSaveShopWeather(shop, wSheet, fSheet, now);
      Utilities.sleep(200);
    } catch (e) {
      console.error(`天気取得エラー (${shop.name}): ` + e.toString());
    }
  });

  CacheService.getScriptCache().remove(CACHE_KEY_DASHBOARD);
}

function fetchAndSaveShopWeather(shop, wSheet, fSheet, now) {
  if (!OPENWEATHER_API_KEY) throw new Error("API Key Missing");

  const curUrl = `https://api.openweathermap.org/data/2.5/weather?lat=${shop.lat}&lon=${shop.lng}&appid=${OPENWEATHER_API_KEY}&units=metric&lang=ja`;
  const curRes = UrlFetchApp.fetch(curUrl, { muteHttpExceptions: true });
  if (curRes.getResponseCode() !== 200) return;
  const curData = JSON.parse(curRes.getContentText());

  const weather = curData.weather[0].description;
  const temp = curData.main.temp;
  const humidity = curData.main.humidity;
  const pressure = curData.main.pressure;
  const rain = curData.rain ? (curData.rain["1h"] || 0) : 0;
  const snow = curData.snow ? (curData.snow["1h"] || 0) : 0;
  const icon = curData.weather[0].icon;
  const wind = curData.wind ? curData.wind.speed : 0;
  const windDeg = curData.wind ? curData.wind.deg : 0;

  let alerts = [];
  if (curData.alerts && Array.isArray(curData.alerts)) {
      curData.alerts.forEach(a => alerts.push(a.event || a.description));
  }
  if (wind >= 10) alerts.push("強風注意");
  if (wind >= 15) alerts.push("暴風警戒");
  if (rain >= 20) alerts.push("大雨注意");
  if (rain >= 50) alerts.push("激しい雨");
  if (snow >= 1) alerts.push("降雪注意");
  if (temp >= 35) alerts.push("猛暑注意");

  const alertStr = [...new Set(alerts)].join(", ");

  const oneCallUrl = `https://api.openweathermap.org/data/3.0/onecall?lat=${shop.lat}&lon=${shop.lng}&exclude=minutely&appid=${OPENWEATHER_API_KEY}&units=metric&lang=ja`;
  const oneCallRes = UrlFetchApp.fetch(oneCallUrl, { muteHttpExceptions: true });

  let forecastJson = "";

  if (oneCallRes.getResponseCode() === 200) {
    const data = JSON.parse(oneCallRes.getContentText());

    const timelineData = [];
    const hourly = data.hourly || [];
    for (let i = 0; i < Math.min(hourly.length, 28); i += 3) {
       const item = hourly[i];
       const tDate = new Date(item.dt * 1000);
       timelineData.push({
         time: Utilities.formatDate(tDate, 'Asia/Tokyo', 'HH:mm'),
         temp: Math.round(item.temp),
         icon: item.weather[0].icon,
         weather: item.weather[0].description,
         pop: Math.round((item.pop||0)*100),
         rain: item.rain ? (item.rain["1h"]||0) : 0,
         wind: item.wind_speed || 0,
         windDeg: item.wind_deg || 0
       });
    }
    forecastJson = JSON.stringify(timelineData);

    const daily = data.daily || [];
    daily.forEach((d, index) => {
        const dKey = Utilities.formatDate(new Date(d.dt * 1000), 'Asia/Tokyo', 'yyyy-MM-dd');
        const weatherDesc = d.weather[0].description;
        const maxT = d.temp.max;
        const minT = d.temp.min;
        const maxPop = Math.round((d.pop || 0) * 100);
        const avgHum = d.humidity;
        const totalRain = d.rain || 0;
        const totalSnow = d.snow || 0;
        const maxWind = d.wind_speed;
        const windDeg = d.wind_deg;
        const icon = d.weather[0].icon;

        fSheet.appendRow([
            now, dKey, shop.name, weatherDesc, maxT, minT, maxPop, avgHum, totalRain, totalSnow, icon, maxWind, windDeg
        ]);
    });
  } else {
    console.error(`One Call API Error: ${oneCallRes.getContentText()}`);
  }

  wSheet.appendRow([
    now, shop.name, weather, temp, humidity, pressure, rain, snow, icon, forecastJson, wind, windDeg, alertStr
  ]);
}

// --- Dashboard Data (Fast Reading with Cache) ---
function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cachedJson = cache.get(CACHE_KEY_DASHBOARD);
  if (cachedJson) {
    return JSON.parse(cachedJson);
  }

  const ss = getSpreadsheet();

  const mSheet = ss.getSheetByName('Master');
  const mData = mSheet.getDataRange().getValues();
  const shopLocMap = {};
  for (let i = 1; i < mData.length; i++) {
    const name = normalize(mData[i][4]);
    const lat = mData[i][6];
    const lng = mData[i][7];
    if (lat && lng) {
      shopLocMap[name] = { lat: lat, lng: lng };
    }
  }

  const wSheet = ss.getSheetByName('WeatherLog');
  const wLast = wSheet.getLastRow();
  const wStart = Math.max(2, wLast - 5000);
  const wData = (wLast > 1) ? wSheet.getRange(wStart, 1, wLast - wStart + 1, 13).getValues() : [];

  const latestWeather = {};
  for (let i = wData.length - 1; i >= 0; i--) {
    const shop = normalize(wData[i][1]);
    if (!latestWeather[shop]) {
      let timeline = [];
      try { if(wData[i][9]) timeline = JSON.parse(wData[i][9]); } catch(e){}

      latestWeather[shop] = {
        time: Utilities.formatDate(new Date(wData[i][0]), 'Asia/Tokyo', 'HH:mm'),
        weather: safeStr(wData[i][2]),
        temp: wData[i][3],
        rain: wData[i][6],
        snow: wData[i][7],
        icon: safeStr(wData[i][8]),
        wind: wData[i][10] || 0,
        windDeg: wData[i][11] || 0,
        alerts: [],
        rawAlertStr: wData[i][12],
        timeline: timeline
      };
    }
  }

  const fSheet = ss.getSheetByName('ForecastLog');
  const fLast = fSheet.getLastRow();
  const fStart = Math.max(2, fLast - 10000);
  const fData = (fLast > 1) ? fSheet.getRange(fStart, 1, fLast - fStart + 1, 13).getValues() : [];

  const weeklyMap = {};
  const now = new Date();
  const todayKey = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  for (let i = fData.length - 1; i >= 0; i--) {
     const shop = normalize(fData[i][2]);
     let targetDateStr = "";
     if (fData[i][1] instanceof Date) {
        targetDateStr = Utilities.formatDate(fData[i][1], 'Asia/Tokyo', 'yyyy-MM-dd');
     } else {
        targetDateStr = String(fData[i][1]).substring(0, 10);
     }

     if (!weeklyMap[shop]) weeklyMap[shop] = [];

     if (targetDateStr >= todayKey && !weeklyMap[shop].some(d => d.date === targetDateStr)) {
         weeklyMap[shop].push({
            date: targetDateStr,
            max: fData[i][4],
            min: fData[i][5],
            pop: fData[i][6],
            rain: fData[i][8],
            wind: fData[i][11],
            weatherDesc: safeStr(fData[i][3]),
            icon: safeStr(fData[i][10])
         });
     }
  }

  const result = INITIAL_SHOPS.map(s => {
    const key = normalize(s.name);
    const loc = shopLocMap[key] || {};
    const cw = latestWeather[key] || { weather: "-", temp: "-", icon: "01d", time: "--:--", alerts: [], timeline: [], rawAlertStr: "", rain: 0, wind: 0 };
    const wk = weeklyMap[key] || [];

    wk.sort((a,b) => a.date.localeCompare(b.date));

    const actionAlerts = generateActionAlerts(cw, wk);

    return {
      name: s.name,
      lat: loc.lat || s.lat || 35.68,
      lng: loc.lng || s.lng || 139.76,
      current: cw,
      timeline: cw.timeline,
      weekly: wk,
      alerts: actionAlerts
    };
  });

  try {
    cache.put(CACHE_KEY_DASHBOARD, JSON.stringify(result), CACHE_TIME);
  } catch(e) {
    console.warn("Cache Save Error: " + e.message);
  }

  return result;
}

// 業務アクションアラート生成ロジック
function generateActionAlerts(current, weekly) {
  const alerts = [];
  const today = new Date();

  const cWind = current.wind || 0;
  const cRain = current.rain || 0;
  const cTemp = current.temp || 0;
  const cHumid = current.humidity || 50;

  if (cWind >= 15) {
    alerts.push({ title: "暴風警戒", msg: "【店長判断】屋外作業の全面中止（カート回収係を屋内へ）。アナウンス実施。自動ドアの開放停止検討。", level: "danger", date: "現在" });
  } else if (cWind >= 10) {
    alerts.push({ title: "強風注意", msg: "【安全管理担当】のぼり旗の撤去、外置きカート・カゴの固定。自転車置き場の巡回。", level: "warning", date: "現在" });
  }

  if (cRain >= 20) {
    alerts.push({ title: "大雨警戒", msg: "【施設管理・値引き】雨漏り、入り口マットの吸水状況確認。客足停止を見越した惣菜・日配の早期値引き判断。濡れたお客様へタオル等のサービス提供検討。", level: "warning", date: "現在" });
  } else if (cRain === 0 && current.timeline && current.timeline.length > 0) {
    // 過去1時間など過去の履歴がcurrentに乗っていないため、
    // ここでは「現在雨が降っていない（cRain===0）」かつ「直近の天候状態データに雨形跡があった場合（もしくは特定のフラグ）」を想定しますが、
    // 現状はOpenWeatherの「現在雨量(current.rain)」のみ保持しているため、ダッシュボードキャッシュ生成側の履歴や
    // rawAlertStrに含まれるアラート解除履歴等から本来厳密に判定します。
    // ※今回は仕様「雨が止んだ直後(雨量 0mmへ回復)」をシステムで出すため、一旦条件ロジックを追加します。
    // （運用上、直前まで雨フラグがあったかどうかを判定するフラグが今後必要になる可能性があります）
    // とりあえず、「現在雨量が0mmである」かつ（何か以前雨が降っていたという情報をここに連携する）というIFブロックを用意。
    // もしくは、current.weatherに'rain'が含まれておらず、Timeline直近に雨マークがある等で擬似判定。
    const recentForecast = current.timeline[0] || {};
    if (recentForecast.rain > 0 && cRain === 0) {
        alerts.push({ title: "雨上がり・客足回復", msg: "【タイムサービス】「雨上がりタイムサービス」の放送実施。客足の戻りを見込み、鮮魚・精肉の値引きシール貼付時間を通常より30分〜1時間遅らせる（安売りしすぎない）。", level: "sales", date: "現在" });
    }
  }

  if (cTemp >= 35) {
    alerts.push({ title: "猛暑警戒", msg: "【従業員管理】駐車場係、カート係を15分ローテーションにするなど休憩頻度を上げる。店内空調の設定温度確認（お客様からの「暑い」苦情対応）。", level: "danger", date: "現在" });
  }

  if (cTemp >= 28) {
    alerts.push({ title: "気温上昇・飲料訴求", msg: "【全館放送・補充】アイス・冷たい飲料の店内放送実施。冷蔵ケースのフェイスアップ徹底。", level: "sales", date: "現在" });
  }

  const hour = parseInt(Utilities.formatDate(today, 'Asia/Tokyo', 'H'));
  if (hour >= 16) {
    const todayData = weekly.find(d => d.date === Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd'));
    if (todayData && (todayData.max - cTemp >= 5)) {
      alerts.push({ title: "夕方冷え込み・鍋", msg: "【夕市・試食】「今夜は温かいお鍋」の放送とPOP。おでん・中華まん・鍋セットの最前面展開。", level: "sales", date: "現在" });
    }
  }

  if (cHumid <= 30) {
    alerts.push({ title: "乾燥注意・保湿", msg: "【日用雑貨・青果】のど飴・ハンドクリームのレジ横展開。みかん・ビタミンC飲料の訴求強化。", level: "sales", date: "現在" });
  }

  if (current.rawAlertStr) {
    parseAlerts(current.rawAlertStr).forEach(a => {
      a.date = "現在";
      alerts.push(a);
    });
  }

  const t1Date = new Date(today);
  t1Date.setDate(today.getDate() + 1);
  const t1Str = Utilities.formatDate(t1Date, 'Asia/Tokyo', 'yyyy-MM-dd');
  const t1Data = weekly.find(d => d.date === t1Str);

  if (t1Data) {
    if (t1Data.rain >= 20 || t1Data.wind >= 15) {
      alerts.push({ title: "荒天・製造抑制", msg: "【惣菜・鮮魚: ロス削減】インストア加工数を一律20%カット。刺身は「冊」販売をメインに。", level: "danger", date: "明日" });
      alerts.push({ title: "荒天・体制縮小", msg: "【店長: 人件費管理】パート・アルバイトの出勤調整、早上がり打診。空いた人員を清掃へ。", level: "warning", date: "明日" });
    } else if (t1Data.rain >= 5) {
      alerts.push({ title: "雨天・出来立て訴求", msg: "【惣菜・ベーカリー】開店時を減らし、夕方ピークに向けた「出来立て」シール対応。大パック比率向上。", level: "manufacture", date: "明日" });
    }

    const day = t1Date.getDay();
    const isWeekend = (day === 0 || day === 6); // 0: 日曜, 6: 土曜
    const isSunny = (t1Data.weatherDesc.includes("晴") || t1Data.icon.includes("01"));
    if (t1Data.max >= 20 && t1Data.max <= 28 && isSunny && isWeekend) {
      alerts.push({ title: "行楽・イベント日和", msg: "【惣菜・寿司: 攻めの製造】行楽弁当、おにぎり、オードブル、寿司桶の製造数を対前週比 110%〜120% に設定。午前10時〜12時のピークに合わせ、レジ応援体制を厚くする。飲料ケース売りの山積みを店頭へ移動。", level: "manufacture", date: "明日" });
    }
  }

  const t2Date = new Date(today);
  t2Date.setDate(today.getDate() + 2);
  const t2Str = Utilities.formatDate(t2Date, 'Asia/Tokyo', 'yyyy-MM-dd');
  const t2Data = weekly.find(d => d.date === t2Str);

  if (t2Data) {
    if (t2Data.rain >= 20 || t2Data.wind >= 15) {
      alerts.push({ title: "荒天・来店激減警戒", msg: "【店長・チーフ: リスク回避】生鮮品の発注は「売れ筋トップ10」のみに絞り、変わり種はカット。広告商品の山積み展開を見送り、売場変更の手間を減らす。「雨の日クーポン」等の配信準備。", level: "danger", date: "2日後" });
    } else if (t2Data.rain >= 5 || t2Data.weatherDesc.includes("雪")) {
      alerts.push({ title: "客足鈍化・内食予報", msg: "【部門チーフ: 発注抑制】日配（パン・牛乳）、日持ちしない葉物野菜の発注を対前週比 90%〜95% に抑制。雨天時の「まとめ買い」を狙い、大袋菓子やカップ麺の在庫を確認。", level: "order", date: "2日後" });
    }

    if (t2Data.max >= 30 && (t2Data.weatherDesc.includes("晴") || t2Data.icon.includes("01"))) {
      alerts.push({ title: "猛暑・涼味特需", msg: "【グロサリー・青果: 在庫確保】飲料、アイス、麺つゆ、氷の発注を対前週比 120%〜130% に増量。スイカ、カットフルーツの加工計画数を引き上げ。エントランス付近に「熱中症対策コーナー」を作る指示出し。", level: "order", date: "2日後" });
    }

    const t1DateForT2Check = new Date(today);
    t1DateForT2Check.setDate(today.getDate() + 1);
    const t1DataForT2Check = weekly.find(d => d.date === Utilities.formatDate(t1DateForT2Check, 'Asia/Tokyo', 'yyyy-MM-dd'));
    const isTempDrop5 = t1DataForT2Check && (t1DataForT2Check.max - t2Data.max >= 5);

    if (t2Data.max <= 10 || isTempDrop5) {
      alerts.push({ title: "寒冷・鍋物特需", msg: "【精肉・青果: メニュー提案】鍋つゆ、白菜、キノコ類、豚バラ薄切りの発注を強化。関連販売（クロスMD）用として、鍋コーナー横にカセットボンベや締めの中華麺を発注。", level: "order", date: "2日後" });
    }
  }

  return alerts;
}

function parseAlerts(alertStr) {
  if (!alertStr) return [];
  return String(alertStr).split(',').map(s => {
    s = s.trim();
    if (!s) return null;
    let level = "info";
    if (s.includes("警報") || s.includes("激しい") || s.includes("暴風")) level = "danger";
    else if (s.includes("注意") || s.includes("警戒")) level = "warning";
    return { title: "気象情報", msg: s, level: level };
  }).filter(Boolean);
}

// --- Admin / Aggregation Functions ---

function runFlashReport() {
  generateAggregatedReport(new Date(), "速報", false);
}

function createDailyReport() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  generateAggregatedReport(yesterday, "朝刊", false);
}

function runPatrolReport() {
  const user = getUserProfile();
  if (user.role === 'GUEST') throw new Error("ゲストは実行できません");
  generateAggregatedReport(new Date(), "巡回報告書", true, user.name);
}

function generateAggregatedReport(targetDate, titlePrefix, isPatrolOnly, targetUser = null) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');
  const data = sheet.getDataRange().getValues();
  const dateKey = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyyMMdd');

  const rows = data.filter((row, i) => {
    if (i === 0) return false;
    const rowDate = new Date(row[1]);
    if (isNaN(rowDate.getTime())) return false;
    const rKey = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyyMMdd');
    if (rKey !== dateKey) return false;

    const cat = parseCategory(row[6]);
    if (isPatrolOnly) {
       return cat === CAT_PATROL && (!targetUser || row[4] === targetUser);
    } else {
       return cat !== CAT_PATROL;
    }
  });

  if (rows.length === 0) {
    console.log(`${titlePrefix}：対象データなし（${dateKey}）スキップします`);
    return null;
  }

  const contentText = rows.map(r => `[ID:${r[0]}] 店舗:${r[2]} 担当:${r[4]} 内容:${r[9]}`).join("\n");

  const prompt = `
 あなたは優秀な店舗報告集計アシスタントです。以下の報告リストをもとに、読みやすく視覚的に整理された「${titlePrefix}」の集計レポート案を作成します。

出力は以下のJSON形式のみを返してください。Markdown等の装飾は一切不要です。

{
  "overall_summary": "全体の傾向、特筆すべき事項、緊急性の高い内容などを200文字程度でまとめた全体の概要。文章内で【 】などの記号による強調は行わないでください。文章内には絶対に出典元のID（[ID:xxx]等）を記載しないでください。",
  "topics": [
    {
      "title": "トピックの見出し（例：〇〇商品の売れ行き、××店のトラブル対応など）",
      "summary": "そのトピックに関する要約。重要語句であっても【 】などで強調せず、自然な文章にしてください。どこの誰が何と言っているか明確にし、文章内にIDを直接記載することは禁止します。",
      "related_ids": ["関連する報告のID文字列", "ID2"]
    }
  ]
}

■報告リスト:
${contentText}
`;

  let aiResult = { overall_summary: "AI解析失敗（要約なし）", topics: [] };
  try {
     const resJson = summarizeTextWithGemini(prompt);
     if (resJson) aiResult = resJson;
  } catch(e) {
     console.error("AI Error:", e);
     aiResult.overall_summary += " : " + e.message;
  }

  const folder = getHqMonthFolder(targetDate);
  const fileName = `【${titlePrefix}】${Utilities.formatDate(targetDate, 'Asia/Tokyo', 'MM/dd')}`;

  const targetDayStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'MM/dd');
  const flashTitle = `【速報】${targetDayStr}`;
  const morningTitle = `【朝刊】${targetDayStr}`;

  const oldFiles = folder.getFiles();
  while(oldFiles.hasNext()) {
    const f = oldFiles.next();
    const fName = f.getName();
    if (fName === flashTitle || fName === morningTitle || fName === fileName) {
      f.setTrashed(true);
    }
  }

  const doc = DocumentApp.create(fileName);
  folder.addFile(DriveApp.getFileById(doc.getId()));
  const body = doc.getBody();

  body.insertParagraph(0, fileName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`作成日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')}`);
  body.appendHorizontalRule();

  body.appendParagraph("■ 原文データ (時系列)").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const bookmarkMap = {};

  const emojiMap = {
     '日報': '📝', '定型日報': '📝',
     '苦情': '⚠️', 'トラブル': '⚠️', '苦情・トラブル': '⚠️',
     '提案': '💡', '気付き': '💡', '気付き・提案': '💡',
     '返信': '↩️', '共有': '↩️', '返信・共有': '↩️',
     '巡回': '📋', '巡回報告': '📋'
  };

  rows.forEach(r => {
     const id = safeStr(r[0]);
     const email = safeStr(r[5]);
     const accountName = email.includes('@') ? email.split('@')[0] : (r[4] || "不明");
     const type = safeStr(r[6]);
     const emoji = emojiMap[type] || '📄';
     const isUrgent = safeStr(r[7]) === 'TRUE' || safeStr(r[7]) === 'true';

     const p = body.appendParagraph(`${emoji} ${accountName} (${r[2]})`);
     p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
     if (isUrgent) p.setBold(true).setForegroundColor('#ef4444');

     const bookmark = doc.addBookmark(doc.newPosition(p, 0));
     bookmarkMap[id] = bookmark.getId();

     const infoP = body.appendParagraph(`時間: ${Utilities.formatDate(new Date(r[1]), 'Asia/Tokyo', 'HH:mm')} / 種別: ${type}`);
     infoP.setItalic(true).setForegroundColor('#64748b');

     const contentP = body.appendParagraph(r[9]);
     contentP.setSpacingBefore(5).setSpacingAfter(10);

     if(r[10]) {
        const urls = String(r[10]).split('\n');
        urls.forEach(url => {
            if(!url) return;
            try {
                const fileId = extractFileIdFromUrl(url);
                if (fileId) {
                    const imgBlob = DriveApp.getFileById(fileId).getBlob();
                    body.appendImage(imgBlob).setWidth(300).setHeight(300);
                }
            } catch(e) { /* ignore image errors */ }
        });
     }
     body.appendParagraph("");
  });

  let insertIdx = 3;
  body.insertParagraph(insertIdx++, "■ AI要約レポート").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.insertParagraph(insertIdx++, "【全体概要】").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.insertParagraph(insertIdx++, aiResult.overall_summary);

  if (aiResult.topics && aiResult.topics.length > 0) {
      body.insertParagraph(insertIdx++, "【トピック別要約】").setHeading(DocumentApp.ParagraphHeading.HEADING3);

      aiResult.topics.forEach(topic => {
          body.insertParagraph(insertIdx++, `・${topic.title}`).setBold(true);
          body.insertParagraph(insertIdx++, topic.summary).setBold(false);

          if (topic.related_ids && topic.related_ids.length > 0) {
              const linkP = body.insertParagraph(insertIdx++, "　└ 参照ソース: ");
              const text = linkP.editAsText();

              let currentPos = "　└ 参照ソース: ".length;
              topic.related_ids.forEach((relId, idx) => {
                  const bId = bookmarkMap[relId];
                  if (bId) {
                      const targetRow = rows.find(row => String(row[0]) === String(relId));
                      const email = targetRow ? safeStr(targetRow[5]) : "";
                      const accountLabel = email.includes('@') ? email.split('@')[0] : (targetRow ? targetRow[4] : "不明");

                      const linkLabel = `[${accountLabel}]`;
                      text.insertText(currentPos, linkLabel);
                      text.setLinkUrl(currentPos, currentPos + linkLabel.length - 1, `#bookmark=${bId}`);
                      currentPos += linkLabel.length;

                      if (idx < topic.related_ids.length - 1) {
                          text.insertText(currentPos, " ");
                          currentPos += 1;
                      }
                  }
              });
          }
          body.insertParagraph(insertIdx++, "");
      });
  }

  body.insertHorizontalRule(insertIdx++);
  doc.saveAndClose();
}

function getAggregatedReports() {
  const folder = getHqMonthFolder(new Date());
  const files = folder.getFiles();
  const list = [];
  while (files.hasNext()) {
    const f = files.next();
    list.push({ name: f.getName(), url: f.getUrl(), date: f.getLastUpdated().toLocaleString() });
  }
  return list;
}

function getMyHistory() {
  const user = getUserProfile();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = data.length - 1; i >= 1; i--) {
     if (list.length >= 10) break;
     if (data[i][4] === user.name) {
       list.push({
         date: Utilities.formatDate(new Date(data[i][1]), 'Asia/Tokyo', 'MM/dd HH:mm'),
         mode: data[i][6],
         summary: (String(data[i][8]).substring(0, 30) + "...")
       });
     }
  }
  return list;
}

function runSetupFromClient() {
  const ss = getSpreadsheet();

  ['Master', 'Reports', 'WeatherLog', 'ForecastLog'].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const sRep = ss.getSheetByName('Reports');
  if (sRep.getLastRow() === 0) sRep.appendRow(["ID", "日時", "店舗", "役職", "氏名", "Email", "種別", "緊急", "要約", "全文", "画像URL", "DocURL", "確認", "備考"]);

  const sMas = ss.getSheetByName('Master');
  if (sMas.getLastRow() === 0) {
    sMas.appendRow(["ID", "氏名", "Email", "役職", "店舗", "住所", "Lat", "Lng"]);
    const email = Session.getActiveUser().getEmail();
    sMas.appendRow([Utilities.getUuid(), "管理者", email, "ADMIN", "本部", "東京都", "", ""]);
  }

  setupTriggers();

  return "セットアップ完了。APIキーを設定し、トリガーが設定されたことを確認してください。";
}

function addShopsToMaster() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  const data = sheet.getDataRange().getValues();
  const existingNames = data.map(r => normalize(r[4]));

  let count = 0;
  INITIAL_SHOPS.forEach(shop => {
    if (!existingNames.includes(normalize(shop.name))) {
      let lat = "", lng = "";
      try {
         const geo = Maps.newGeocoder().geocode(shop.address);
         if (geo.status === "OK" && geo.results.length > 0) {
            lat = geo.results[0].geometry.location.lat;
            lng = geo.results[0].geometry.location.lng;
         }
         Utilities.sleep(500);
      } catch(e) {}

      sheet.appendRow([Utilities.getUuid(), "店舗User", "", "STAFF", shop.name, shop.address, lat, lng]);
      count++;
    }
  });
  return `${count}件の店舗を追加しました。`;
}

function forceResetMaster() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  sheet.clear();
  runSetupFromClient();
  addShopsToMaster();
  return "マスタをリセットしました。";
}

/**
 * 【自動化】定期実行トリガーの設定
 */
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('updateWeatherAllShops')
    .timeBased()
    .everyHours(3)
    .create();

  ScriptApp.newTrigger('createDailyReport')
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone("Asia/Tokyo")
    .create();

  [10, 13, 16, 19].forEach(hour => {
    ScriptApp.newTrigger('runFlashReport')
      .timeBased()
      .atHour(hour)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone("Asia/Tokyo")
      .create();
  });

  ScriptApp.newTrigger('cleanupOldData')
    .timeBased()
    .atHour(4)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone("Asia/Tokyo")
    .create();

  console.log("Triggers have been initialized.");
}

/**
 * 古い予報データのクリーンアップ
 */
function cleanupOldData() {
  const ss = getSpreadsheet();
  const fSheet = ss.getSheetByName('ForecastLog');
  if (fSheet) {
    const data = fSheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0,0,0,0);

    for (let i = data.length - 1; i >= 1; i--) {
      const targetDate = new Date(data[i][1]);
      if (targetDate < today) {
        fSheet.deleteRow(i + 1);
      }
    }
  }

  CacheService.getScriptCache().remove(CACHE_KEY_DASHBOARD);
}


// ==========================================
// ★ 以下: 外部PWA（スマレコ録音アプリ）対応
// ==========================================
//
// 【デプロイ設定】
//   GASエディタ → デプロイ → デプロイを管理 → 新しいデプロイ
//   実行ユーザー : 自分
//   アクセスできるユーザー : 全員（匿名ユーザーを含む）  ← 重要
// ==========================================

/**
 * doPost: PWAからのHTTP POSTを受け取るエントリーポイント
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ error: 'リクエストが不正です' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'submitVoiceReport';

    // 【アクション: getProfile】ログイン直後のプロフィール取得
    if (action === 'getProfile') {
      const email = verifyAndGetEmail(data.idToken);
      if (!email) return jsonResponse({ error: '認証失敗: 有効な会社アカウントが必要です' });
      return jsonResponse(getUserProfileByEmail(email));
    }

    // 【アクション: registerProfile】初回登録・プロフィール更新
    if (action === 'registerProfile') {
      const email = verifyAndGetEmail(data.idToken);
      if (!email) return jsonResponse({ error: '認証失敗: 有効な会社アカウントが必要です' });
      const name = String(data.name || '').trim();
      const shop = String(data.shop || '').trim();
      if (!name || !shop) return jsonResponse({ error: '名前と所属を入力してください' });
      return jsonResponse(registerProfile(email, name, shop));
    }

    // 【アクション: submitVoiceReport】音声報告の送信
    if (action === 'submitVoiceReport') {
      const email = verifyAndGetEmail(data.idToken);
      if (!email) return jsonResponse({ error: '認証失敗: 有効な会社アカウントが必要です' });

      // 画像をDriveにアップロード
      const imageUrls = [];
      if (data.imageBase64List && data.imageBase64List.length > 0) {
        for (const img of data.imageBase64List) {
          try {
            imageUrls.push(uploadSingleImage(img.data, img.type));
          } catch (imgErr) {
            console.error('画像アップロードエラー:', imgErr);
          }
        }
      }

      const p = {
        audioBase64 : data.audioBase64,
        mimeType    : data.mimeType || 'audio/webm',
        mode        : data.mode    || '日報',
        isEmergency : data.isEmergency || false,
        imageUrls   : imageUrls,
        location    : data.location || {}
      };

      return jsonResponse(submitVoiceReportByEmail(p, email));
    }

    // 【認証不要アクション】天気・レポート閲覧
    if (action === 'getDashboardData') {
      return jsonResponse(getDashboardData());
    }

    if (action === 'getAggregatedReports') {
      return jsonResponse(getAggregatedReports());
    }

    // 【以下は認証必須】idToken検証
    const authEmail = verifyAndGetEmail(data.idToken);
    if (!authEmail) return jsonResponse({ error: '認証失敗: 有効な会社アカウントが必要です' });

    if (action === 'getMyHistory') {
      return jsonResponse(getMyHistoryByEmail(authEmail));
    }

    if (action === 'runPatrolReport') {
      return jsonResponse(runPatrolReportByEmail(authEmail));
    }

    if (action === 'runFlashReport') {
      runFlashReport();
      return jsonResponse({ success: true });
    }

    if (action === 'createDailyReport') {
      createDailyReport();
      return jsonResponse({ success: true });
    }

    if (action === 'addShopsToMaster') {
      const r = addShopsToMaster();
      return jsonResponse(r !== undefined ? r : { success: true });
    }

    if (action === 'manualWeatherUpdate') {
      manualWeatherUpdate();
      return jsonResponse('天気データを更新しました。');
    }

    if (action === 'forceResetMaster') {
      forceResetMaster();
      return jsonResponse('マスタデータをリセットしました。');
    }

    if (action === 'runSetupFromClient') {
      const r = runSetupFromClient();
      return jsonResponse(r !== undefined ? r : { success: true });
    }

    return jsonResponse({ error: '不明なアクション: ' + action });

  } catch (err) {
    console.error('doPost エラー:', err);
    return jsonResponse({ error: 'サーバーエラー: ' + err.toString() });
  }
}

/**
 * verifyAndGetEmail: GoogleのIDトークンを検証してメールアドレスを返す
 */
function verifyAndGetEmail(idToken) {
  if (!idToken) return null;
  try {
    const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    const info = JSON.parse(res.getContentText());

    if (!info.email_verified || info.email_verified === 'false') return null;

    const domain = (info.email || '').split('@')[1] || '';
    if (!ALLOWED_DOMAINS_PWA.includes(domain)) {
      console.warn('Unauthorized domain:', domain);
      return null;
    }

    return info.email;
  } catch (e) {
    console.error('Token verification error:', e);
    return null;
  }
}

/**
 * getUserProfileByEmail: メールアドレスでMasterシートを検索してプロフィールを返す
 */
function getUserProfileByEmail(email) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    if (!sheet) throw new Error('Masterシートが見つかりません');

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === email.trim().toLowerCase()) {
        return {
          id      : safeStr(data[i][0]),
          name    : safeStr(data[i][1]),
          email   : email,
          account : email.split('@')[0],
          role    : safeStr(data[i][3]),
          shop    : safeStr(data[i][4]),
          address : safeStr(data[i][5])
        };
      }
    }
    // 未登録の場合はregistered:falseを返す（エラーにしない）
    return { registered: false, email: email, account: email.split('@')[0] };
  } catch (e) {
    return { name: 'エラー', shop: 'Error', role: 'GUEST', email: email, account: '', error: e.toString() };
  }
}

/**
 * registerProfile: 新規ユーザーをMasterシートに登録する
 */
function registerProfile(email, name, shop) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    if (!sheet) throw new Error('Masterシートが見つかりません');

    // 既存チェック
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === email.trim().toLowerCase()) {
        // 既存ユーザーは名前・店舗を更新
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 5).setValue(shop);
        return { success: true, name, shop, email, account: email.split('@')[0], role: safeStr(data[i][3]) || '一般', registered: true };
      }
    }

    // 新規追加
    const newId = 'U' + Date.now();
    sheet.appendRow([newId, name, email, '一般', shop, '']);
    return { success: true, name, shop, email, account: email.split('@')[0], role: '一般', registered: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * submitVoiceReportByEmail: 検証済みメールで音声報告を処理する
 */
function submitVoiceReportByEmail(p, verifiedEmail) {
  const user = getUserProfileByEmail(verifiedEmail);
  if (user.role === 'GUEST') throw new Error('ユーザー登録されていません: ' + verifiedEmail);

  const id = Utilities.getUuid();
  const ts = new Date();

  let gemini = { summary: { items: [] }, transcript: '' };
  try {
    const result = analyzeWithGemini(p.audioBase64, p.mimeType, p.mode);
    if (result) gemini = result;
  } catch (e) {
    console.error('Gemini Fail:', e);
    gemini.transcript = 'AI解析エラー: ' + e.message;
    gemini.summary    = { items: [{ header: 'エラー', content: 'AI解析中に問題が発生しました。原文をご確認ください。' }] };
  }

  let wInfo = { weather: '-', temp: '-' };
  try {
    const ss      = getSpreadsheet();
    const wSheet  = ss.getSheetByName('WeatherLog');
    const lastRow = wSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 100);
    const wData   = wSheet.getRange(startRow, 1, lastRow - startRow + 1, 13).getValues();
    for (let i = wData.length - 1; i >= 0; i--) {
      if (normalize(wData[i][1]) === normalize(user.shop)) {
        wInfo = { weather: safeStr(wData[i][2]), temp: safeStr(wData[i][3]) + '℃' };
        break;
      }
    }
  } catch (e) { /* 天気取得失敗は無視 */ }

  const docUrl = createReportDoc(id, ts, user, p.mode, gemini.summary, gemini.transcript, p.imageUrls, wInfo);

  const ss      = getSpreadsheet();
  const sheet   = ss.getSheetByName('Reports');
  const sumText = (gemini.summary && gemini.summary.items && gemini.summary.items.length > 0)
    ? gemini.summary.items.map(i => '【' + safeStr(i.header) + '】' + safeStr(i.content)).join('\n')
    : (gemini.transcript || 'AI解析失敗（要約なし）');

  sheet.appendRow([
    id, ts, user.shop, user.role, user.name, user.email,
    p.mode, p.isEmergency ? '🚨' : '通常',
    sumText, gemini.transcript, (p.imageUrls || []).join('\n'), docUrl, '未確認', ''
  ]);

  return { success: true, summary: sumText };
}

/**
 * jsonResponse: JSON形式でHTTPレスポンスを返すヘルパー
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * getMyHistoryByEmail: メールアドレスで自分の送信履歴を取得 (Session不使用版)
 */
function getMyHistoryByEmail(email) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Reports');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (list.length >= 10) break;
      if (String(data[i][5]).trim().toLowerCase() === email.trim().toLowerCase()) {
        list.push({
          date    : Utilities.formatDate(new Date(data[i][1]), 'Asia/Tokyo', 'MM/dd HH:mm'),
          mode    : data[i][6],
          summary : String(data[i][8]).substring(0, 30) + '...'
        });
      }
    }
    return list;
  } catch (e) {
    console.error('getMyHistoryByEmail error:', e);
    return [];
  }
}

/**
 * runPatrolReportByEmail: 検証済みメールで本日の巡回まとめを作成 (Session不使用版)
 */
function runPatrolReportByEmail(email) {
  const user = getUserProfileByEmail(email);
  if (user.role === 'GUEST') throw new Error('ゲストは実行できません: ' + email);
  generateAggregatedReport(new Date(), '巡回報告書', true, user.name);
  return { success: true };
}

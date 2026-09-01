/**
 * 單字練習網站 — Google Apps Script 後端
 *
 * 使用方式：
 * 1. 建立一份 Google 試算表，依下面「需要的工作表」建好 4 個分頁與表頭。
 * 2. 該試算表「擴充功能 → Apps Script」，把這個檔案整個貼進去（取代預設的 Code.gs）。
 * 3. 「部署 → 新增部署作業 → 網頁應用程式」：
 *      執行身分：我
 *      誰可以存取：知道連結的任何人
 *    取得 .../exec 網址，貼到 app.js 最上面的 APPS_SCRIPT_URL。
 *
 * 需要的工作表（表頭列必須完全比照，順序不拘）：
 *
 * 「單字庫」：題號 | 英文 | 中文
 *   → 對應 單字庫.csv，可直接匯入。中文欄可含詞性標籤，如「真相、真理、真實性(n)」。
 *
 * 「學生」：帳號 | 密碼 | 姓名 | 起始題號 | 每週題數 | 開課週一日期 | 累計分數
 *   → 開課週一日期格式 YYYY-MM-DD。每週題數＝每天背幾個字（週一~週六，預設 10）。
 *   → 累計分數初始請填 0（數字），之後由系統自動累加，不用手動改。
 *
 * 「作答紀錄」：時間 | 帳號 | 題號 | 英文 | 中文 | 測驗類型 | 學生作答 | 是否正確
 *   → 每答一題新增一列，不更新舊列。
 *
 * 「錯題本」：帳號 | 題號 | 英文 | 中文 | 測驗類型 | 累計錯誤次數 | 最後錯誤時間
 *   → 用「帳號＋題號＋測驗類型」當唯一鍵：答錯累加次數，答對整列刪除。
 *
 * 計分規則：同一個帳號對同一個題號＋同一種測驗類型（中文卷/英文卷），
 * 只有「有史以來第一次作答」才計分（答對 +1、答錯 -0.5），判斷方式是看
 * 「作答紀錄」裡有沒有這個帳號＋題號＋測驗類型的舊紀錄。之後不管是錯題本
 * 重測、還是週日整週複習又碰到同一個字，都不會重複計分。
 */

const SHEET_WORDS = "單字庫";
const SHEET_STUDENTS = "學生";
const SHEET_LOG = "作答紀錄";
const SHEET_WRONG = "錯題本";

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ ok: false, error: "無法解析請求內容" });
  }

  const action = body.action;
  try {
    switch (action) {
      case "login": return jsonOutput(handleLogin(body));
      case "getAssignment": return jsonOutput(handleGetAssignment(body));
      case "submitAnswer": return jsonOutput(handleSubmitAnswer(body));
      case "getWrongBank": return jsonOutput(handleGetWrongBank(body));
      default: return jsonOutput({ ok: false, error: "未知的 action：" + action });
    }
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("找不到工作表：" + name);
  return sheet;
}

// 把一個工作表讀成「表頭 → 值」的物件陣列
function readTable_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: values[0] || [], rows: [] };
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1)
    .filter(r => r.some(cell => cell !== "" && cell !== null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
  return { headers, rows };
}

// ------------------------------------------------------------------------
// 單字庫（快取在 CacheService，避免每次都整張表重讀，5 分鐘過期）
// ------------------------------------------------------------------------
function getWordBank_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("wordBank");
  if (cached) return JSON.parse(cached);

  const { rows } = readTable_(SHEET_WORDS);
  const words = rows.map(r => ({
    id: Number(r["題號"]),
    en: String(r["英文"]),
    zh: String(r["中文"]),
  })).sort((a, b) => a.id - b.id);

  cache.put("wordBank", JSON.stringify(words), 300);
  return words;
}

function getWordsByRange_(words, start, end) {
  return words.filter(w => w.id >= start && w.id <= end);
}

// ------------------------------------------------------------------------
// 登入
// ------------------------------------------------------------------------
function handleLogin({ account, password }) {
  const { rows } = readTable_(SHEET_STUDENTS);
  const student = rows.find(r => String(r["帳號"]) === String(account));
  if (!student || String(student["密碼"]) !== String(password)) {
    return { ok: false, error: "帳號或密碼錯誤" };
  }
  return { ok: true, account: student["帳號"], name: student["姓名"] };
}

// ------------------------------------------------------------------------
// 取得指派與今日/本週單字
// ------------------------------------------------------------------------
function getStudentConfig_(account) {
  const { rows } = readTable_(SHEET_STUDENTS);
  const student = rows.find(r => String(r["帳號"]) === String(account));
  if (!student) throw new Error("查無此帳號的指派設定：" + account);
  return {
    programStartNum: Number(student["起始題號"]) || 1,
    dailyCount: Number(student["每週題數"]) || 10,
    programStartMonday: student["開課週一日期"],
    score: Number(student["累計分數"]) || 0,
  };
}

function computeAssignment_(config, now) {
  const start = (config.programStartMonday instanceof Date)
    ? config.programStartMonday
    : new Date(config.programStartMonday);
  if (isNaN(start.getTime())) {
    throw new Error("「學生」工作表的「開課週一日期」讀不到有效日期，目前讀到的值是：" +
      JSON.stringify(config.programStartMonday) + "。請確認欄位標題完全是「開課週一日期」，" +
      "且該學生列的值是日期格式（YYYY-MM-DD）。");
  }
  const daysSince = Math.floor((now - start) / (1000 * 60 * 60 * 24));
  const weekIndex = Math.max(0, Math.floor(daysSince / 7));
  const weekCount = config.dailyCount * 6;
  const weekStart = config.programStartNum + weekIndex * weekCount;
  const weekEnd = weekStart + weekCount - 1;

  const weekdayRaw = ((daysSince % 7) + 7) % 7; // 0=開課當週的週一
  const weekday = weekdayRaw === 6 ? 7 : weekdayRaw + 1; // 1~6=週一~週六, 7=週日

  let todayStart, todayEnd;
  if (weekday === 7) {
    todayStart = weekStart; todayEnd = weekEnd;
  } else {
    todayStart = weekStart + (weekday - 1) * config.dailyCount;
    todayEnd = todayStart + config.dailyCount - 1;
  }
  return { weekIndex, weekStart, weekEnd, weekday, todayStart, todayEnd };
}

function handleGetAssignment({ account }) {
  const config = getStudentConfig_(account);
  const a = computeAssignment_(config, new Date());
  const words = getWordBank_();
  const maxId = words.length ? words[words.length - 1].id : 0;

  return {
    ok: true,
    weekIndex: a.weekIndex,
    weekStart: a.weekStart,
    weekEnd: Math.min(a.weekEnd, maxId),
    weekday: a.weekday,
    todayStart: a.todayStart,
    todayEnd: Math.min(a.todayEnd, maxId),
    todayWords: getWordsByRange_(words, a.todayStart, Math.min(a.todayEnd, maxId)),
    weekWords: getWordsByRange_(words, a.weekStart, Math.min(a.weekEnd, maxId)),
    score: config.score,
  };
}

// ------------------------------------------------------------------------
// 送出一題作答
// ------------------------------------------------------------------------
const MODE_LABEL = { zh: "中文卷", en: "英文卷" };

function handleSubmitAnswer({ account, wordId, en, zh, mode, given, correct }) {
  const modeLabel = MODE_LABEL[mode] || mode;
  const isCorrect = !!correct;
  const firstAttempt = !hasPriorAttempt_(account, wordId, modeLabel);

  const logSheet = getSheet_(SHEET_LOG);
  logSheet.appendRow([new Date(), account, wordId, en, zh, modeLabel, given || "", isCorrect]);

  updateWrongBank_({ account, wordId, en, zh, mode, correct: isCorrect });

  let delta = 0;
  let newScore = null;
  if (firstAttempt) {
    delta = isCorrect ? 1 : -0.5;
    newScore = addStudentScore_(account, delta);
  }

  return { ok: true, firstAttempt, delta, newScore };
}

// 有沒有這個帳號＋題號＋測驗類型的舊作答紀錄（用來判斷是不是「有史以來第一次作答」）
function hasPriorAttempt_(account, wordId, modeLabel) {
  const { rows } = readTable_(SHEET_LOG);
  return rows.some(r =>
    String(r["帳號"]) === String(account) &&
    Number(r["題號"]) === Number(wordId) &&
    String(r["測驗類型"]) === modeLabel
  );
}

// 把 delta 加進「學生」工作表該帳號的「累計分數」欄，回傳更新後的分數
function addStudentScore_(account, delta) {
  const sheet = getSheet_(SHEET_STUDENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const accCol = headers.indexOf("帳號");
  const scoreCol = headers.indexOf("累計分數");
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][accCol]) === String(account)) {
      const updated = (Number(values[i][scoreCol]) || 0) + delta;
      sheet.getRange(i + 1, scoreCol + 1).setValue(updated);
      return updated;
    }
  }
  throw new Error("查無此帳號可更新分數：" + account);
}

function updateWrongBank_({ account, wordId, en, zh, mode, correct }) {
  const sheet = getSheet_(SHEET_WRONG);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = (name) => headers.indexOf(name);
  const modeLabel = MODE_LABEL[mode] || mode;

  let foundRow = -1;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[col("帳號")]) === String(account) &&
        Number(row[col("題號")]) === Number(wordId) &&
        String(row[col("測驗類型")]) === modeLabel) {
      foundRow = i + 1; // 1-based sheet row
      break;
    }
  }

  if (correct) {
    if (foundRow > 0) sheet.deleteRow(foundRow);
    return;
  }

  if (foundRow > 0) {
    const countCol = col("累計錯誤次數") + 1;
    const timeCol = col("最後錯誤時間") + 1;
    const current = Number(sheet.getRange(foundRow, countCol).getValue()) || 0;
    sheet.getRange(foundRow, countCol).setValue(current + 1);
    sheet.getRange(foundRow, timeCol).setValue(new Date());
  } else {
    sheet.appendRow([account, wordId, en, zh, modeLabel, 1, new Date()]);
  }
}

// ------------------------------------------------------------------------
// 取得錯題本
// ------------------------------------------------------------------------
const MODE_KEY = { "中文卷": "zh", "英文卷": "en" };

function handleGetWrongBank({ account }) {
  const { rows } = readTable_(SHEET_WRONG);
  const items = rows
    .filter(r => String(r["帳號"]) === String(account))
    .map(r => ({
      wordId: Number(r["題號"]),
      en: r["英文"],
      zh: r["中文"],
      mode: MODE_KEY[r["測驗類型"]] || r["測驗類型"],
      count: Number(r["累計錯誤次數"]) || 0,
    }));
  return { ok: true, items };
}

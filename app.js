/* =========================================================================
   單字練習網站 — 前端邏輯

   資料來源：words-data.js 裡的 WORD_BANK_FULL（由 Excel 轉出的 1316 個單字，
   1~1316 連續、無缺漏，中文欄位保留原始詞性標籤如「(n)」「(vt)」）。

   後端：APPS_SCRIPT_URL 設定好之後，登入／取得指派／送出作答／錯題本
   都會呼叫 Google Apps Script（見 Code.gs）。在還沒部署 Apps Script、
   APPS_SCRIPT_URL 還是空字串的時候，會自動退回「本機示範模式」
   （帳號 demo／密碼 1234，資料存在瀏覽器 localStorage），方便先預覽畫面。
   ========================================================================= */

// 部署好 Google Apps Script 之後，把 /exec 網址貼在這裡即可切換成正式後端。
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwaRpGKZApqAe5coEokf6aKdQ_hwMWgJNX-Ai1TXBHEjWuC3axWVNNL41i28S2oV8hm/exec";

const DAILY_COUNT = 10;          // 每天背幾個字（週一~週六）
const WEEK_COUNT = DAILY_COUNT * 6; // 一週總字數 = 60
const DAY_NAMES = ["", "週一", "週二", "週三", "週四", "週五", "週六", "週日"];

// ------------------------------------------------------------------------
// 中文答案批改：去除詞性標籤／括號註記後，用「、,，;；」與空白切成多個可接受答案
// 例："真相、真理、真實性(n)" -> ["真相","真理","真實性"]
// ------------------------------------------------------------------------
function acceptedZhAnswers(zh) {
  const core = zh.replace(/\([^()]*\)/g, "");
  const parts = core.split(/[、,，;；\s]+/).map(s => s.trim()).filter(Boolean);
  return [...new Set(parts)];
}

function checkAnswer(word, mode, raw) {
  const answer = raw.trim();
  if (mode === "en") return answer.toLowerCase() === word.en.trim().toLowerCase();
  return acceptedZhAnswers(word.zh).includes(answer);
}

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// ------------------------------------------------------------------------
// 本機示範模式（APPS_SCRIPT_URL 未設定時使用）
// ------------------------------------------------------------------------
const DEMO_STORE_KEY = "vocabapp_demo_v1";

function demoLoadStore() {
  try {
    return JSON.parse(localStorage.getItem(DEMO_STORE_KEY)) || { wrongBank: [], history: [] };
  } catch { return { wrongBank: [], history: [] }; }
}
function demoSaveStore(s) { localStorage.setItem(DEMO_STORE_KEY, JSON.stringify(s)); }

function demoGetWordsByRange(start, end) {
  return WORD_BANK_FULL.filter(w => w.id >= start && w.id <= end);
}

// 示範模式的指派邏輯：假設從題號 1 開始、每週 60 題，遇到題庫尾端就繞回開頭，
// 讓示範帳號可以一直往後測試、不會因為題庫用完而卡住。正式後端（Code.gs）
// 不會繞回開頭，而是照 SPEC.md 的邏輯往後累加。
function demoGetAssignment() {
  const totalWeeks = Math.floor(WORD_BANK_FULL.length / WEEK_COUNT);
  const start = new Date(state.demoProgramMonday);
  const now = new Date();
  const daysSince = Math.floor((now - start) / (1000 * 60 * 60 * 24));
  const weekIndex = ((Math.floor(daysSince / 7) % totalWeeks) + totalWeeks) % totalWeeks;
  const weekStart = 1 + weekIndex * WEEK_COUNT;
  const weekEnd = weekStart + WEEK_COUNT - 1;
  const weekday = state.simDay || (() => { const d = now.getDay(); return d === 0 ? 7 : d; })();

  let todayStart, todayEnd, todayWords;
  if (weekday === 7) {
    todayStart = weekStart; todayEnd = weekEnd;
    todayWords = demoGetWordsByRange(weekStart, weekEnd);
  } else {
    todayStart = weekStart + (weekday - 1) * DAILY_COUNT;
    todayEnd = todayStart + DAILY_COUNT - 1;
    todayWords = demoGetWordsByRange(todayStart, todayEnd);
  }
  return {
    ok: true,
    weekIndex,
    weekStart, weekEnd, weekday, todayStart, todayEnd,
    todayWords,
    weekWords: demoGetWordsByRange(weekStart, weekEnd),
    score: demoLoadStore().score || 0,
    dailyCount: DAILY_COUNT,
    programStartNum: 1,
    wordBankMax: WORD_BANK_FULL.length,
  };
}

// 本機示範模式的「學習進度」：掃 localStorage 裡的作答歷史，找出「答對過」的題號
// （不管是中文卷還是英文卷，只要對過一次就算精熟），邏輯跟 Code.gs 的 handleGetProgress 一致
function demoGetProgress() {
  const s = demoLoadStore();
  const mastered = new Set();
  (s.history || []).forEach(h => { if (h.correct) mastered.add(h.wordId); });
  return { ok: true, masteredIds: [...mastered] };
}

const RESCORE_COOLDOWN_MS = 30 * 24 * 60 * 60 * 1000; // 跟 Code.gs 的 RESCORE_COOLDOWN_DAYS 一致

// 計分規則：同一個單字＋同一種測驗模式（zh/en），真正第一次作答、或是距離上一次
// 作答這題已經超過 30 天，才會計分（答對 +1、答錯 -0.5）——讓學生一個月後回頭複習
// 舊字仍然能得分，對抗學習曲線下滑，但同一題不能無限刷分數。分數會持續累加，
// 不會因為練習次數增加而被稀釋或歸零。
function demoSubmitAnswer({ word, mode, correct }) {
  const s = demoLoadStore();
  if (typeof s.score !== "number") s.score = 0;

  const priorAttempts = s.history.filter(h => h.wordId === word.id && h.mode === mode);
  const lastAttempt = priorAttempts.length ? priorAttempts[priorAttempts.length - 1] : null;
  const firstAttempt = !lastAttempt ||
    (Date.now() - new Date(lastAttempt.time).getTime()) >= RESCORE_COOLDOWN_MS;
  const delta = firstAttempt ? (correct ? 1 : -0.5) : 0;
  s.score += delta;

  s.history.push({
    time: new Date().toISOString(), user: state.user,
    wordId: word.id, en: word.en, zh: word.zh, mode, correct,
  });
  const key = `${word.id}-${mode}`;
  if (!correct) {
    const existing = s.wrongBank.find(w => w.key === key);
    if (existing) existing.count += 1;
    else s.wrongBank.push({ key, wordId: word.id, en: word.en, zh: word.zh, mode, count: 1 });
  } else {
    s.wrongBank = s.wrongBank.filter(w => w.key !== key);
  }
  demoSaveStore(s);
  return { ok: true, firstAttempt, delta, newScore: s.score };
}

function demoGetWrongBank() {
  return demoLoadStore().wrongBank.map(w => ({
    wordId: w.wordId, en: w.en, zh: w.zh, mode: w.mode, count: w.count,
  }));
}

// ------------------------------------------------------------------------
// API 層：有設定 APPS_SCRIPT_URL 就打真正的後端，否則走本機示範模式
// ------------------------------------------------------------------------
async function callApi(action, payload) {
  // 用 text/plain 送出可避開 Apps Script 對 CORS preflight 的限制
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, ...payload }),
  });
  return res.json();
}

async function apiLogin(account, password) {
  if (!APPS_SCRIPT_URL) {
    if (account.trim() === "demo" && password === "1234") {
      return { ok: true, account: "demo", name: "王小明" };
    }
    return { ok: false, error: "帳號或密碼錯誤" };
  }
  return callApi("login", { account: account.trim(), password });
}

async function apiGetAssignment(account) {
  if (!APPS_SCRIPT_URL) return demoGetAssignment();
  return callApi("getAssignment", { account });
}

async function apiSubmitAnswer({ word, mode, correct }) {
  if (!APPS_SCRIPT_URL) return demoSubmitAnswer({ word, mode, correct });
  return callApi("submitAnswer", {
    account: state.user, wordId: word.id, en: word.en, zh: word.zh,
    mode, given: state.lastGiven || "", correct,
  });
}

async function apiGetWrongBank() {
  if (!APPS_SCRIPT_URL) return { ok: true, items: demoGetWrongBank() };
  return callApi("getWrongBank", { account: state.user });
}

async function apiGetProgress() {
  if (!APPS_SCRIPT_URL) return demoGetProgress();
  return callApi("getProgress", { account: state.user });
}

// ------------------------------------------------------------------------
// 狀態管理
// ------------------------------------------------------------------------
const state = {
  user: null,
  simDay: null,            // 示範用：手動指定星期幾（1~7），只在本機示範模式生效
  demoProgramMonday: (() => {
    // 示範模式固定用「本週一」當第一週起點，讓雛形一直顯示第 1 週方便預覽
    const now = new Date();
    const day = now.getDay();
    const diffToMonday = day === 0 ? -6 : 1 - day;
    const monday = new Date(now);
    monday.setDate(now.getDate() + diffToMonday);
    return monday.toISOString().slice(0, 10);
  })(),
  assignment: null,        // 最近一次 apiGetAssignment() 的結果
  wrongBank: [],           // 最近一次 apiGetWrongBank() 的結果
  masteredIds: new Set(),  // 最近一次 apiGetProgress() 的結果：答對過的題號（判斷完成度／進度用）
  viewWeekIndex: null,     // 目前「本週進度」面板正在看第幾週（null＝目前實際那一週）
  studyWords: [],          // 目前正在背誦／即將測驗的單字（今日／本週／點選的某一天）
  studyIndex: 0,
  studyFlipped: {},
  quizMode: null,          // "zh" or "en"
  quizWords: [],
  quizIndex: 0,
  quizAnswers: [],
  pendingQuizWords: [],
  quizSource: "daily",     // "daily" | "week" | "wrong"
};

function findWord(id) { return WORD_BANK_FULL.find(w => w.id === id); }

// 分數顯示：整數就不顯示小數點，半分（如 -0.5 累加後的 12.5）才顯示到小數第一位
function formatScore(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(1);
}
function updateTopbarScore(score) {
  state.score = score;
  document.getElementById("topbarScore").textContent = `累計 ${formatScore(score)} 分`;
}

// ------------------------------------------------------------------------
// 畫面切換
// ------------------------------------------------------------------------
function showView(id) {
  document.querySelectorAll(".view").forEach(el => el.classList.remove("active"));
  document.getElementById(id)?.classList.add("active");
}

// ------------------------------------------------------------------------
// 登入
// ------------------------------------------------------------------------
document.getElementById("loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const user = document.getElementById("loginUser").value;
  const pass = document.getElementById("loginPass").value;
  const submitBtn = e.target.querySelector("button[type=submit]");
  const userInput = document.getElementById("loginUser");
  const passInput = document.getElementById("loginPass");
  // 登入要打後端 API，網路慢的時候畫面容易看起來像沒反應：鎖住輸入框、
  // 按鈕顯示「登入中…」加上轉圈動畫，讓學生清楚知道系統正在處理、不要一直重按
  function setLoginLoading(loading) {
    submitBtn.disabled = loading;
    submitBtn.classList.toggle("is-loading", loading);
    submitBtn.innerHTML = loading ? '<span class="spinner"></span>登入中…' : "登入";
    userInput.disabled = loading;
    passInput.disabled = loading;
  }

  setLoginLoading(true);
  document.getElementById("loginError").textContent = "";

  const result = await apiLogin(user, pass);
  if (!result.ok) {
    setLoginLoading(false);
    document.getElementById("loginError").textContent = result.error || "帳號或密碼錯誤，請再試一次";
    return;
  }

  state.user = result.account;
  document.getElementById("studentName").textContent = result.name + " 同學";
  // 密碼驗證通過後還要再打幾支 API 準備首頁資料（指派範圍／錯題本／進度），
  // 這段期間繼續停在登入畫面顯示「登入中…」，等首頁資料都準備好才一次切過去，
  // 避免中間出現「畫面空白、看起來像沒登入成功」的空檔
  await renderHome();
  document.querySelector(".app").classList.add("active");
  document.getElementById("screen-login").classList.remove("active");
  setLoginLoading(false);
  showView("view-home");
});

document.getElementById("logoutBtn").addEventListener("click", () => {
  state.user = null;
  document.querySelector(".app").classList.remove("active");
  document.getElementById("screen-login").classList.add("active");
});

// ------------------------------------------------------------------------
// 首頁
// ------------------------------------------------------------------------
async function renderHome() {
  const a = await apiGetAssignment(state.user);
  state.assignment = a;
  document.getElementById("assignRange").textContent = `第 ${a.weekStart}~${a.weekEnd} 題`;
  updateTopbarScore(a.score || 0);

  const sundayHint = document.getElementById("sundayHint");
  if (a.weekday === 7) {
    document.getElementById("todayTag").textContent = "週日 · 本週總複習";
    document.getElementById("todayRange").textContent = `題號 ${a.weekStart}~${a.weekEnd}（全部混合）`;
    document.getElementById("todayHeading").textContent = "今天把這週的單字全部混合再看一次";
    document.getElementById("goStudyBtn").textContent = "看本週單字";
    sundayHint.textContent = "今天測驗會從整週題目隨機出題";
  } else {
    document.getElementById("todayTag").textContent = `${DAY_NAMES[a.weekday]} · 每日十字`;
    document.getElementById("todayRange").textContent = `題號 ${a.todayStart}~${a.todayEnd}`;
    document.getElementById("todayHeading").textContent = "先看過今天的十個單字，背熟再測驗";
    document.getElementById("goStudyBtn").textContent = "看今日單字";
    sundayHint.textContent = "";
  }

  const wb = await apiGetWrongBank();
  state.wrongBank = wb.items || [];
  document.getElementById("wrongCount").textContent = `目前 ${state.wrongBank.length} 題待複習`;

  const prog = await apiGetProgress();
  state.masteredIds = new Set(prog.masteredIds || []);

  state.courseProgress = computeCourseProgress();
  state.viewWeekIndex = a.weekIndex; // 每次回首頁都重設回「目前實際那一週」
  renderPaceBanner();
  renderWeekPanel();
}

// ------------------------------------------------------------------------
// 週次／每日題號範圍換算（純前端算式）：programStartNum/dailyCount/wordBankMax
// 由後端 getAssignment 回傳一次，之後任何一週都能直接算出來，不用再打 API。
// 「目前實際那一週」的單字內容用後端剛給的 weekWords（即時對照 Google Sheets）；
// 往前／往後翻看別週時，用 words-data.js 內建的完整題庫切，兩者同源資料。
// ------------------------------------------------------------------------
function weekRangeFor(weekIndex) {
  const a = state.assignment;
  const weekCount = a.dailyCount * 6;
  const start = a.programStartNum + weekIndex * weekCount;
  const end = Math.min(start + weekCount - 1, a.wordBankMax);
  return { start, end };
}

function dayRangeFor(weekIndex, day) {
  const a = state.assignment;
  const { start: weekStart, end: weekEnd } = weekRangeFor(weekIndex);
  const start = weekStart + (day - 1) * a.dailyCount;
  const end = Math.min(start + a.dailyCount - 1, weekEnd);
  return { start, end };
}

function wordsInRange(start, end) {
  if (start > end) return [];
  const a = state.assignment;
  if (start >= a.weekStart && end <= a.weekEnd) {
    return (a.weekWords || []).filter(w => w.id >= start && w.id <= end);
  }
  return WORD_BANK_FULL.filter(w => w.id >= start && w.id <= end);
}

// 這個範圍內的題號是不是「全部都答對過」（不分中文卷／英文卷，對過一次就算）
function isRangeMastered(start, end) {
  if (start > end) return false;
  for (let id = start; id <= end; id++) {
    if (!state.masteredIds.has(id)) return false;
  }
  return true;
}

// 課程從第 0 週第 1 天算起、累計到第 weekIndex 週第 day 天，是第幾天
// （週日複習跟週六用同一個序號——複習的是同一批字，不算「之後」的新進度）
function daySeq(weekIndex, day) {
  return weekIndex * 6 + Math.min(day, 6);
}

// ------------------------------------------------------------------------
// 課程進度總表：從課程第一天掃到「今天」，看有沒有全部完成（caughtUp，決定
// 能不能點開「今天之後」的進度）、還有幾天沒完成（missedDays，催促用）。
// 「之後」的進度一定要先把之前跟今天的都完成才能點開，回頭補之前的天數則
// 永遠不鎖。
// ------------------------------------------------------------------------
function computeCourseProgress() {
  const a = state.assignment;
  const todayDay = a.weekday === 7 ? 6 : a.weekday;
  let missedDays = 0;
  for (let wi = 0; wi <= a.weekIndex; wi++) {
    const lastDay = wi < a.weekIndex ? 6 : todayDay;
    for (let day = 1; day <= lastDay; day++) {
      const { start, end } = dayRangeFor(wi, day);
      if (start > a.wordBankMax) continue;
      if (!isRangeMastered(start, end)) missedDays++;
    }
  }
  return { caughtUp: missedDays === 0, missedDays, todaySeq: daySeq(a.weekIndex, todayDay) };
}

// ------------------------------------------------------------------------
// 進度提示：還沒把「之前＋今天」的進度全部完成就催促（不限本週，整個課程累計）；
// 已經全部完成、還把本週後面的天數提前做完了，就鼓勵
// ------------------------------------------------------------------------
function renderPaceBanner() {
  const a = state.assignment;
  const { missedDays } = state.courseProgress;
  let aheadDays = 0;
  if (state.courseProgress.caughtUp) {
    for (let day = 1; day <= 6; day++) {
      const isFuture = a.weekday !== 7 && day > a.weekday;
      if (!isFuture) continue;
      const { start, end } = dayRangeFor(a.weekIndex, day);
      if (start > a.wordBankMax) continue;
      if (isRangeMastered(start, end)) aheadDays++;
    }
  }

  const banner = document.getElementById("paceBanner");
  if (missedDays > 0) {
    banner.className = "pace-banner warn";
    banner.textContent = `⏰ 目前還有 ${missedDays} 天的進度沒完成，要先補完之前跟今天的進度，才能解鎖之後的內容，加油！`;
  } else if (aheadDays > 0) {
    banner.className = "pace-banner success";
    banner.textContent = `🎉 太棒了！已經超前完成本週 ${aheadDays} 天的進度！`;
  } else {
    banner.className = "pace-banner hidden";
    banner.textContent = "";
  }
}

// ------------------------------------------------------------------------
// 本週進度面板：列出目前所選那一週週一~週六六天的題號範圍與完成狀態，
// 不限「今天」，也可以用左右箭頭往前往後翻週，回補進度或超前進度都可以
// ------------------------------------------------------------------------
function renderWeekPanel() {
  const a = state.assignment;
  const wi = state.viewWeekIndex;
  const { start: weekStart, end: weekEnd } = weekRangeFor(wi);
  const rel = wi - a.weekIndex;
  const label = rel === 0 ? "本週進度" : rel === -1 ? "上週進度" : rel === 1 ? "下週進度"
    : `第 ${wi + 1} 週進度`;

  document.getElementById("weekPanelTitle").textContent = label;
  document.getElementById("weekPanelSub").textContent =
    weekStart > a.wordBankMax ? "這一週已經超出題庫範圍了" : `題號 ${weekStart}~${weekEnd}`;
  document.getElementById("weekPrevBtn").disabled = wi <= 0;
  document.getElementById("weekNextBtn").disabled = weekRangeFor(wi + 1).start > a.wordBankMax;

  const todaySeq = state.courseProgress.todaySeq;
  const caughtUp = state.courseProgress.caughtUp;

  const list = document.getElementById("weekDaysList");
  const items = [];
  for (let day = 1; day <= 6; day++) {
    const { start, end } = dayRangeFor(wi, day);
    const empty = start > a.wordBankMax || start > weekEnd;
    const isToday = rel === 0 && a.weekday === day;
    const done = !empty && isRangeMastered(start, end);
    const isPastDue = rel < 0 || (rel === 0 && (a.weekday === 7 || day < a.weekday));
    const locked = !empty && daySeq(wi, day) > todaySeq && !caughtUp;
    let tag = "";
    if (locked) tag = '<span class="week-day-tag locked">🔒 未開放</span>';
    else if (done) tag = '<span class="week-day-tag done">✓ 已完成</span>';
    else if (isToday) tag = '<span class="week-day-tag today">今天</span>';
    else if (isPastDue) tag = '<span class="week-day-tag pending">尚未完成</span>';
    items.push(`
      <li class="week-day-item${isToday ? " is-today" : ""}${done ? " is-done" : ""}${(empty || locked) ? " is-locked" : ""}" data-week="${wi}" data-day="${day}">
        <span class="week-day-label">${DAY_NAMES[day]}</span>
        <span class="week-day-range">${empty ? "超出題庫範圍" : `題號 ${start}~${end}`}</span>
        ${tag}
      </li>`);
  }

  // 週日：這週全部混合再刷一次，跟週六用同一個解鎖時機（複習同一批字，不算新進度）
  {
    const sunEmpty = weekStart > a.wordBankMax;
    const sunDone = !sunEmpty && isRangeMastered(weekStart, weekEnd);
    const sunToday = rel === 0 && a.weekday === 7;
    const sunLocked = !sunEmpty && daySeq(wi, 6) > todaySeq && !caughtUp;
    let sunTag = "";
    if (sunLocked) sunTag = '<span class="week-day-tag locked">🔒 未開放</span>';
    else if (sunDone) sunTag = '<span class="week-day-tag done">✓ 已完成</span>';
    else if (sunToday) sunTag = '<span class="week-day-tag today">今天</span>';
    items.push(`
      <li class="week-day-item week-day-sunday${sunToday ? " is-today" : ""}${sunDone ? " is-done" : ""}${(sunEmpty || sunLocked) ? " is-locked" : ""}" data-week="${wi}" data-day="7">
        <span class="week-day-label">週日</span>
        <span class="week-day-range">${sunEmpty ? "超出題庫範圍" : `本週混合複習 題號 ${weekStart}~${weekEnd}`}</span>
        ${sunTag}
      </li>`);
  }
  list.innerHTML = items.join("");
}

document.getElementById("weekPrevBtn").addEventListener("click", () => {
  if (state.viewWeekIndex <= 0) return;
  state.viewWeekIndex--;
  renderWeekPanel();
});
document.getElementById("weekNextBtn").addEventListener("click", () => {
  if (weekRangeFor(state.viewWeekIndex + 1).start > state.assignment.wordBankMax) return;
  state.viewWeekIndex++;
  renderWeekPanel();
});

document.getElementById("weekDaysList").addEventListener("click", (e) => {
  const li = e.target.closest(".week-day-item");
  if (!li || li.classList.contains("is-locked")) return;
  startDayStudy(Number(li.dataset.week), Number(li.dataset.day));
});

function startDayStudy(weekIndex, day) {
  const { start, end } = day === 7 ? weekRangeFor(weekIndex) : dayRangeFor(weekIndex, day);
  const words = wordsInRange(start, end);
  if (!words.length) return;
  state.studyWords = words;
  state.studyIndex = 0;
  state.studyFlipped = {};
  const rel = weekIndex - state.assignment.weekIndex;
  const weekLabel = rel === 0 ? "" : rel === -1 ? "上週" : rel === 1 ? "下週" : `第${weekIndex + 1}週`;
  const dayLabel = day === 7 ? "週日混合複習" : `${DAY_NAMES[day]}單字`;
  document.getElementById("studyTitle").textContent = `${weekLabel}${dayLabel}`;
  renderStudyDots();
  renderStudyCard();
  showView("view-study");
}

// 雛形示範控制列：只影響本機示範模式（APPS_SCRIPT_URL 未設定時）
document.getElementById("demoDay").addEventListener("change", async (e) => {
  state.simDay = Number(e.target.value);
  await renderHome();
});
document.getElementById("demoReset").addEventListener("click", async () => {
  localStorage.removeItem(DEMO_STORE_KEY);
  await renderHome();
  renderWrongList();
});
if (APPS_SCRIPT_URL) {
  document.getElementById("demoBar").classList.add("hidden");
}

// ------------------------------------------------------------------------
// 背誦（單字卡）
// ------------------------------------------------------------------------
function currentStudyWords() {
  return state.studyWords;
}

document.getElementById("goStudyBtn").addEventListener("click", () => {
  state.studyWords = state.assignment.weekday === 7 ? state.assignment.weekWords : state.assignment.todayWords;
  state.studyIndex = 0;
  state.studyFlipped = {};
  document.getElementById("studyTitle").textContent = state.assignment.weekday === 7 ? "本週單字總覽" : "今日單字";
  renderStudyDots();
  renderStudyCard();
  showView("view-study");
});

function renderStudyDots() {
  const words = currentStudyWords();
  const wrap = document.getElementById("studyDots");
  wrap.innerHTML = words.map((_, i) =>
    `<span class="${state.studyFlipped[i] ? "done" : ""}"></span>`).join("");
}

function renderStudyCard() {
  const words = currentStudyWords();
  const w = words[state.studyIndex];
  document.getElementById("cardIndex").textContent = `${state.studyIndex + 1} / ${words.length}`;
  document.getElementById("cardEn").textContent = w.en;
  const zhEl = document.getElementById("cardZh");
  const flipped = !!state.studyFlipped[state.studyIndex];
  zhEl.textContent = w.zh;
  zhEl.classList.toggle("hidden", !flipped);
  document.getElementById("cardFlipHint").textContent = flipped ? "" : "點卡片看中文意思";
  document.getElementById("prevCardBtn").disabled = state.studyIndex === 0;
  const isLast = state.studyIndex === words.length - 1;
  document.getElementById("nextCardBtn").textContent = isLast ? "完成" : "下一個";
  document.getElementById("studyDoneBtn").classList.toggle("hidden",
    !Object.keys(state.studyFlipped).length || !isLast);
}

document.getElementById("flashcard").addEventListener("click", () => {
  state.studyFlipped[state.studyIndex] = true;
  renderStudyDots();
  renderStudyCard();
});
document.getElementById("prevCardBtn").addEventListener("click", () => {
  if (state.studyIndex > 0) { state.studyIndex--; renderStudyCard(); }
});
document.getElementById("nextCardBtn").addEventListener("click", () => {
  const words = currentStudyWords();
  if (state.studyIndex < words.length - 1) { state.studyIndex++; renderStudyCard(); }
});
// 背完卡片按「去測驗」：不管背的是今天、整週、還是點某一天的進度，都拿剛剛背的那份單字去測驗
document.getElementById("studyDoneBtn").addEventListener("click", () => openQuizPick(state.studyWords, "daily"));

// ------------------------------------------------------------------------
// 測驗：選擇中文卷／英文卷
// ------------------------------------------------------------------------
function openQuizPick(words, source) {
  state.quizSource = source;
  // 錯題本重測：每一題都用「當初答錯的那個模式」出題，不用再選一次中文卷／英文卷
  if (source === "wrong") {
    state.quizMode = null;
    state.pendingQuizWords = state.wrongBank
      .map(w => ({ word: findWord(w.wordId), mode: w.mode }))
      .filter(item => item.word);
    startQuiz();
    return;
  }
  state.pendingWords = words;
  document.getElementById("quizPickDesc").textContent = `將測驗 ${words.length} 個單字`;
  showView("view-quizpick");
}
document.getElementById("goQuizBtn").addEventListener("click", () => {
  const source = state.assignment.weekday === 7 ? "week" : "daily";
  const words = source === "week" ? state.assignment.weekWords : state.assignment.todayWords;
  openQuizPick(words, source);
});
document.getElementById("retestWrongBtn").addEventListener("click", () => openQuizPick([], "wrong"));

document.querySelectorAll(".pick-card").forEach(btn => {
  btn.addEventListener("click", () => {
    state.quizMode = btn.dataset.mode;
    state.pendingQuizWords = state.pendingWords.map(w => ({ word: w, mode: state.quizMode }));
    startQuiz();
  });
});

// ------------------------------------------------------------------------
// 測驗進行
// ------------------------------------------------------------------------
function startQuiz() {
  state.quizWords = shuffle(state.pendingQuizWords);
  state.quizIndex = 0;
  state.quizAnswers = [];
  setQuizBusy(false);
  document.getElementById("quizTitle").textContent =
    state.quizSource === "wrong" ? "錯題重測" : (state.quizMode === "zh" ? "中文卷" : "英文卷");
  renderQuizDots();
  renderQuizQuestion();
  showView("view-quiz");
}

function renderQuizDots() {
  const wrap = document.getElementById("quizDots");
  wrap.innerHTML = state.quizWords.map((_, i) => {
    const ans = state.quizAnswers[i];
    const cls = ans === undefined ? "" : (ans.correct ? "done" : "wrong-dot");
    return `<span class="${cls}"></span>`;
  }).join("");
}

function renderQuizQuestion() {
  const { word: w, mode } = state.quizWords[state.quizIndex];
  document.getElementById("quizIndex").textContent = `第 ${state.quizIndex + 1} / ${state.quizWords.length} 題`;
  document.getElementById("quizPrompt").textContent = mode === "zh" ? w.en : w.zh;
  const input = document.getElementById("quizInput");
  input.value = "";
  input.disabled = false;
  input.focus();
  document.getElementById("quizFeedback").classList.add("hidden");
  document.getElementById("quizSubmitBtn").classList.remove("hidden");
  document.getElementById("quizNextBtn").classList.add("hidden");
}

document.getElementById("quizSubmitBtn").addEventListener("click", submitQuizAnswer);

// 防止「送出答案」「下一題」被連續誤觸（連按 Enter、雙擊）：兩次動作之間至少間隔
// 400ms，避免答完題按太快，下一題還沒渲染好就把 Enter 又吃成「送出空白答案」
let lastQuizActionAt = 0;
function quizActionAllowed() {
  const now = Date.now();
  if (now - lastQuizActionAt < 400) return false;
  lastQuizActionAt = now;
  return true;
}

// 送出答案期間（等後端批改／寫入 Google Sheets 這段網路來回）整個測驗卡片鎖住、
// 蓋上「批改中…」提示，讓學生清楚知道系統在處理、不會誤以為沒反應而一直按
function setQuizBusy(busy) {
  state.quizLocked = busy;
  document.getElementById("quizCard").classList.toggle("is-busy", busy);
  document.getElementById("quizBusyOverlay").classList.toggle("hidden", !busy);
}

// 測驗中按 Enter：還沒作答就送出答案，已經作答完就直接跳下一題（不用滑鼠點）
document.addEventListener("keydown", (e) => {
  if (e.key !== "Enter" || e.repeat) return;
  if (!document.getElementById("view-quiz").classList.contains("active")) return;
  if (state.quizLocked) { e.preventDefault(); return; }
  if (!document.getElementById("quizNextBtn").classList.contains("hidden")) {
    e.preventDefault();
    document.getElementById("quizNextBtn").click();
  } else if (!document.getElementById("quizSubmitBtn").classList.contains("hidden")) {
    e.preventDefault();
    submitQuizAnswer();
  }
});

async function submitQuizAnswer() {
  if (state.quizLocked || !quizActionAllowed()) return;
  const input = document.getElementById("quizInput");
  const submitBtn = document.getElementById("quizSubmitBtn");
  // 先立刻鎖住輸入框、按鈕、整張測驗卡片，避免網路回應還沒回來前，重複按 Enter／
  // 按鈕造成同一題送出兩次、或是下一題打字時被判成上一題的答案
  input.disabled = true;
  submitBtn.disabled = true;
  setQuizBusy(true);

  const { word: w, mode } = state.quizWords[state.quizIndex];
  const given = input.value.trim();
  const correct = checkAnswer(w, mode, given);
  state.lastGiven = given;
  const scoreResult = await apiSubmitAnswer({ word: w, mode, correct });
  state.quizAnswers[state.quizIndex] = {
    word: w, mode, correct, given,
    scored: !!scoreResult.firstAttempt, delta: scoreResult.delta || 0,
  };
  if (typeof scoreResult.newScore === "number") updateTopbarScore(scoreResult.newScore);

  const fb = document.getElementById("quizFeedback");
  fb.classList.remove("hidden", "correct", "wrong");
  fb.classList.add(correct ? "correct" : "wrong");
  fb.textContent = correct
    ? "答對了！"
    : `答錯了。正確答案：${mode === "zh" ? w.zh : w.en}`;

  setQuizBusy(false);
  submitBtn.disabled = false;
  submitBtn.classList.add("hidden");
  document.getElementById("quizNextBtn").classList.remove("hidden");
  renderQuizDots();
}

document.getElementById("quizNextBtn").addEventListener("click", async () => {
  if (state.quizLocked || !quizActionAllowed()) return;
  if (state.quizIndex < state.quizWords.length - 1) {
    state.quizIndex++;
    renderQuizQuestion();
  } else {
    state.quizLocked = true; // 結算頁要重新抓一次首頁資料，鎖住避免這段期間誤觸
    await renderQuizResult();
    state.quizLocked = false;
    showView("view-result");
  }
});

// ------------------------------------------------------------------------
// 結果
// ------------------------------------------------------------------------
async function renderQuizResult() {
  const total = state.quizAnswers.length;
  const correctN = state.quizAnswers.filter(a => a.correct).length;
  document.getElementById("resultScore").textContent = `${correctN} / ${total}`;
  document.getElementById("resultSub").textContent =
    `答對 ${correctN} 題，答錯 ${total - correctN} 題已記入錯題本`;

  const earned = state.quizAnswers.reduce((sum, a) => sum + (a.delta || 0), 0);
  const pointsEl = document.getElementById("resultPoints");
  if (state.quizAnswers.some(a => a.scored)) {
    pointsEl.textContent = `這次測驗獲得 ${earned > 0 ? "+" : ""}${formatScore(earned)} 分`;
  } else {
    pointsEl.textContent = "這是重複練習，不計分";
  }

  const list = document.getElementById("resultList");
  list.innerHTML = state.quizAnswers.map(a => {
    const pointTag = a.scored
      ? `<span class="point-tag earned">${a.delta > 0 ? "+" : ""}${formatScore(a.delta)} 分</span>`
      : `<span class="point-tag">不計分</span>`;
    // 錯題重測是混合中文卷／英文卷出題，額外標示這題當時考的是哪一種
    const modeTag = state.quizSource === "wrong"
      ? `<span class="point-tag">${a.mode === "zh" ? "中文卷" : "英文卷"}</span>`
      : "";
    return `
    <li class="${a.correct ? "correct" : "wrong"}">
      <span>${a.word.en}／${a.word.zh}</span>
      <span class="answer-col">
        <span>${a.correct ? "✓" : "✕ 你答：" + (a.given || "（空白）")}</span>
        ${modeTag}
        ${pointTag}
      </span>
    </li>`;
  }).join("");
  await renderHome();
}
document.getElementById("backHomeBtn").addEventListener("click", () => showView("view-home"));

// ------------------------------------------------------------------------
// 錯題本
// ------------------------------------------------------------------------
function renderWrongList() {
  const bank = state.wrongBank;
  const list = document.getElementById("wrongList");
  if (!bank.length) {
    list.innerHTML = `<li style="justify-content:center;color:var(--ink-soft)">目前沒有錯題，太棒了！</li>`;
  } else {
    list.innerHTML = bank.map(w => `
      <li>
        <div>
          <div class="wrong-en">${w.en}</div>
          <div class="wrong-zh">${w.zh} · ${w.mode === "zh" ? "中文卷" : "英文卷"}</div>
        </div>
        <span class="wrong-count">錯 ${w.count} 次</span>
      </li>`).join("");
  }
  document.getElementById("retestWrongBtn").disabled = !bank.length;
}
document.getElementById("goWrongBtn").addEventListener("click", async () => {
  const wb = await apiGetWrongBank();
  state.wrongBank = wb.items || [];
  renderWrongList();
  showView("view-wrong");
});

// ------------------------------------------------------------------------
// 返回鍵（統一處理）
// ------------------------------------------------------------------------
document.querySelectorAll("[data-back]").forEach(btn => {
  btn.addEventListener("click", () => showView("view-home"));
});

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
    weekStart, weekEnd, weekday, todayStart, todayEnd,
    todayWords,
    weekWords: demoGetWordsByRange(weekStart, weekEnd),
    score: demoLoadStore().score || 0,
  };
}

// 計分規則：同一個單字＋同一種測驗模式（zh/en），只有「這帳號有史以來第一次作答」
// 會計分（答對 +1、答錯 -0.5）；之後不管是錯題本重測、還是週日整週複習又碰到同一個字，
// 都不會重複計分。分數會持續累加，不會因為練習次數增加而被稀釋或歸零。
function demoSubmitAnswer({ word, mode, correct }) {
  const s = demoLoadStore();
  if (typeof s.score !== "number") s.score = 0;

  const firstAttempt = !s.history.some(h => h.wordId === word.id && h.mode === mode);
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
  submitBtn.disabled = true;
  const result = await apiLogin(user, pass);
  submitBtn.disabled = false;
  if (!result.ok) {
    document.getElementById("loginError").textContent = result.error || "帳號或密碼錯誤，請再試一次";
    return;
  }
  state.user = result.account;
  document.getElementById("loginError").textContent = "";
  document.getElementById("studentName").textContent = result.name + " 同學";
  document.querySelector(".app").classList.add("active");
  document.getElementById("screen-login").classList.remove("active");
  await renderHome();
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

  renderWeekDays();
}

// ------------------------------------------------------------------------
// 本週進度：列出週一~週六六天各自的題號範圍，不限「今天」，方便回補或超前
// ------------------------------------------------------------------------
function assignmentDailyCount() {
  const a = state.assignment;
  return Math.max(1, Math.round((a.weekEnd - a.weekStart + 1) / 6));
}

function weekDayRange(day) {
  const a = state.assignment;
  const dc = assignmentDailyCount();
  const start = a.weekStart + (day - 1) * dc;
  const end = Math.min(start + dc - 1, a.weekEnd);
  return { start, end };
}

function weekDayWords(day) {
  const { start, end } = weekDayRange(day);
  return (state.assignment.weekWords || []).filter(w => w.id >= start && w.id <= end);
}

function renderWeekDays() {
  const a = state.assignment;
  const list = document.getElementById("weekDaysList");
  const items = [];
  for (let day = 1; day <= 6; day++) {
    const { start, end } = weekDayRange(day);
    const empty = start > a.weekEnd;
    const isToday = a.weekday === day;
    items.push(`
      <li class="week-day-item${isToday ? " is-today" : ""}${empty ? " is-empty" : ""}" data-day="${day}">
        <span class="week-day-label">${DAY_NAMES[day]}</span>
        <span class="week-day-range">${empty ? "本週題庫已用完" : `題號 ${start}~${end}`}</span>
        ${isToday ? '<span class="week-day-tag today">今天</span>' : ""}
      </li>`);
  }
  list.innerHTML = items.join("");
}

document.getElementById("weekDaysList").addEventListener("click", (e) => {
  const li = e.target.closest(".week-day-item");
  if (!li || li.classList.contains("is-empty")) return;
  startDayStudy(Number(li.dataset.day));
});

function startDayStudy(day) {
  const words = weekDayWords(day);
  if (!words.length) return;
  state.studyWords = words;
  state.studyIndex = 0;
  state.studyFlipped = {};
  document.getElementById("studyTitle").textContent = `${DAY_NAMES[day]}單字`;
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

// 測驗中按 Enter：還沒作答就送出答案，已經作答完就直接跳下一題（不用滑鼠點）
document.addEventListener("keydown", (e) => {
  if (e.key !== "Enter") return;
  if (!document.getElementById("view-quiz").classList.contains("active")) return;
  if (!document.getElementById("quizNextBtn").classList.contains("hidden")) {
    e.preventDefault();
    document.getElementById("quizNextBtn").click();
  } else if (!document.getElementById("quizSubmitBtn").classList.contains("hidden")) {
    e.preventDefault();
    submitQuizAnswer();
  }
});

async function submitQuizAnswer() {
  const input = document.getElementById("quizInput");
  const submitBtn = document.getElementById("quizSubmitBtn");
  // 先立刻鎖住輸入框與按鈕，避免網路回應還沒回來前，重複按 Enter／按鈕造成同一題送出兩次、
  // 或是下一題打字時被判成上一題的答案（送出後到後端回應之間有網路延遲，這段時間一定要鎖住）
  if (input.disabled) return;
  input.disabled = true;
  submitBtn.disabled = true;
  submitBtn.textContent = "送出中…";

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

  submitBtn.disabled = false;
  submitBtn.textContent = "送出答案";
  submitBtn.classList.add("hidden");
  document.getElementById("quizNextBtn").classList.remove("hidden");
  renderQuizDots();
}

document.getElementById("quizNextBtn").addEventListener("click", async () => {
  if (state.quizIndex < state.quizWords.length - 1) {
    state.quizIndex++;
    renderQuizQuestion();
  } else {
    await renderQuizResult();
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

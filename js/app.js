/* =========================================================
   AI 簡報生成器 — 主程式
   純前端：呼叫使用者自帶的 LLM API → 產生大綱 → PptxGenJS 生成 PPTX
   ========================================================= */

/* =========================================================
   部署設定（管理者可在此調整，讓同事開箱即用）
   defaultProvider：第一次開啟時預設的 AI 供應商
   DEFAULT_MODELS ：各供應商的預設模型（同事不必自己填）
   注意：基於安全，請勿在此填入 API 金鑰並推到「公開」repo。
   ========================================================= */
const APP_CONFIG = {
  defaultProvider: "anthropic"   // 可選："anthropic" | "openai" | "gemini"
};
const DEFAULT_MODELS = {
  anthropic: "claude-opus-4-8",
  openai: "gpt-4o",
  gemini: "gemini-1.5-flash"
};
// 供應商可選模型（下拉），使用者仍可自行輸入其他名稱
const MODEL_PRESETS = {
  anthropic: ["claude-opus-4-8", "claude-sonnet-4-6", "claude-haiku-4-5-20251001"],
  openai: ["gpt-4o", "gpt-4o-mini", "gpt-4.1"],
  gemini: ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"]
};
const PROVIDER_HINTS = {
  anthropic: "在 console.anthropic.com 取得金鑰。瀏覽器直連已啟用 anthropic-dangerous-direct-browser-access。",
  openai: "在 platform.openai.com 取得金鑰。常用模型：gpt-4o、gpt-4o-mini。",
  gemini: "在 aistudio.google.com 取得金鑰。常用模型：gemini-1.5-flash、gemini-1.5-pro。"
};

const TWO_COL_THRESHOLD = 6;   // 重點超過此數 → 自動雙欄
const STORE_KEY = "ppt_workstate";

const state = {
  purpose: null,
  topic: "", direction: "", audience: "", slideCount: 10, lang: "繁體中文", tone: "清楚易懂",
  deck: null,          // { title, subtitle, slides:[{title,bullets:[],notes,layout,chartType}] }
  theme: THEMES[0].id,
  customAccent: null,  // 企業主色（覆蓋風格主色）
  logo: null,          // { data, w, h } 品牌 Logo（已縮圖的 dataURL）
  step: 1,
  previewIdx: 0
};

/* ---------- 工具 ---------- */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
function toast(msg, isErr = false) {
  const t = $("#toast");
  t.textContent = msg; t.hidden = false;
  t.classList.toggle("err", isErr);
  clearTimeout(t._t); t._t = setTimeout(() => (t.hidden = true), 3200);
}
function loading(show, text = "AI 思考中…") {
  $("#loadingText").textContent = text;
  $("#loadingOverlay").hidden = !show;
}
// localStorage 安全存取（隱私模式/停用時不會讓程式壞掉）
function lsGet(k){ try { return localStorage.getItem(k); } catch { return null; } }
function lsSet(k,v){ try { localStorage.setItem(k,v); return true; } catch { return false; } }
function getSettings() {
  return {
    provider: lsGet("ppt_provider") || APP_CONFIG.defaultProvider,
    model: lsGet("ppt_model") || "",
    key: lsGet("ppt_key") || ""
  };
}
function escapeHtml(s){ return (s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

/* ---------- 版型與配色工具（專業商務排版用） ---------- */
const LAYOUTS = [
  { id: "bullets", name: "條列重點", hint: "" },
  { id: "section", name: "章節分隔", hint: "此頁作為章節封面：標題即章節名，第一個重點作為簡短說明。" },
  { id: "metrics", name: "數據亮點 (KPI)", hint: "每個重點請用「數值 | 說明」，例如：180% | 三年 ROI。" },
  { id: "chart", name: "圖表 (長條/圓餅)", hint: "每個重點用「標籤 | 數值」，數值需為數字，例如：北部 | 45。" },
  { id: "quote", name: "金句引言", hint: "第一個重點 = 金句，第二個重點 = 出處/署名。" }
];
const LAYOUT_IDS = LAYOUTS.map(l => l.id);
const CHART_TYPES = [
  { id: "bar", name: "長條圖" },
  { id: "pie", name: "圓餅圖" },
  { id: "line", name: "折線圖" }
];
const CHART_IDS = CHART_TYPES.map(c => c.id);

function hx(c){ return (c || "").replace("#", ""); }
function _rgb(h){ h = hx(h); return [parseInt(h.slice(0,2),16), parseInt(h.slice(2,4),16), parseInt(h.slice(4,6),16)]; }
function _hex(a){ return a.map(v => Math.max(0,Math.min(255,Math.round(v))).toString(16).padStart(2,"0")).join("").toUpperCase(); }
function mix(a, b, t){ const A=_rgb(a), B=_rgb(b); return _hex(A.map((v,i)=>v+(B[i]-v)*t)); }
function lighten(a, t){ return mix(a, "FFFFFF", t); }
function darken(a, t){ return mix(a, "000000", t); }
function isDark(a){ const [r,g,b]=_rgb(a); return (0.299*r+0.587*g+0.114*b) < 140; }

// 解析 KPI：每列「數值 | 說明」
function parseMetrics(bullets){
  const out = [];
  (bullets || []).forEach(b => {
    const parts = String(b).split(/\s*[|｜:：\t]\s*/);
    if (parts.length >= 2 && parts[0].trim()) out.push({ value: parts[0].trim(), label: parts.slice(1).join(" ").trim() });
  });
  return out;
}
// 目前簡報的 eyebrow 小標（用簡報目的名稱）
function kickerText(){ const p = PURPOSES.find(x => x.id === state.purpose); return p ? p.name : ""; }

// 解析圖表資料：每列「標籤 | 數值」，數值取數字（去除 %、逗號、貨幣符號）
function parseChartData(bullets){
  const labels = [], values = [];
  (bullets || []).forEach(b => {
    const parts = String(b).split(/\s*[|｜:：\t]\s*/);
    if (parts.length < 2) return;
    const num = parseFloat(parts[1].replace(/[^0-9.\-]/g, ""));
    if (!isNaN(num)) { labels.push(parts[0].trim()); values.push(num); }
  });
  return { labels, values };
}
// 由主色/輔色推導的圖表色盤
function chartPalette(t, n){
  const base = [t.accent, t.accent2, lighten(t.accent, 0.28), darken(t.accent, 0.2), lighten(t.accent2, 0.2), mix(t.accent, t.accent2, 0.5)];
  const out = [];
  for (let i = 0; i < n; i++) out.push(base[i % base.length]);
  return out;
}

// 企業主色：在所選風格上套用自訂主色，重算相關顏色（保留中性底色/文字）
function customizeTheme(base, hex){
  const c = hx(hex);
  const onDarkBg = isDark(base.bg);
  const coverText = isDark(c) ? "FFFFFF" : "1A1A1A";
  return Object.assign({}, base, {
    id: base.id + "-custom",
    accent: c,
    accent2: lighten(c, 0.35),
    bullet: c,
    title: onDarkBg ? lighten(c, 0.5) : darken(c, 0.12),
    coverBg: c,
    coverText: coverText,
    coverAccent: isDark(c) ? lighten(c, 0.55) : darken(c, 0.28)
  });
}
// 目前實際套用的主題（含自訂主色）
function activeTheme(){
  const base = THEMES.find(x => x.id === state.theme) || THEMES[0];
  return state.customAccent ? customizeTheme(base, state.customAccent) : base;
}

/* ---------- 自動存檔 / 還原 ---------- */
let _saveTimer = null;
function saveState() {
  lsSet(STORE_KEY, JSON.stringify({
    purpose: state.purpose, topic: state.topic, direction: state.direction,
    audience: state.audience, slideCount: state.slideCount, lang: state.lang,
    tone: state.tone, deck: state.deck, theme: state.theme, step: state.step,
    customAccent: state.customAccent, logo: state.logo
  }));
}
function collectAll() {
  collectInputs();
  if ($$("#slidesEditor .slide-card").length) harvestOutline();
  saveState();
}
function debouncedSave() {
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(collectAll, 600);
}
function clearState() { try { localStorage.removeItem(STORE_KEY); } catch {} }
function restoreState() {
  let saved;
  try { saved = JSON.parse(lsGet(STORE_KEY) || "null"); } catch { saved = null; }
  if (!saved || (!saved.topic && !saved.deck)) return false;
  Object.assign(state, {
    purpose: saved.purpose || null, topic: saved.topic || "", direction: saved.direction || "",
    audience: saved.audience || "", slideCount: saved.slideCount || 10, lang: saved.lang || "繁體中文",
    tone: saved.tone || "清楚易懂", deck: saved.deck || null, theme: saved.theme || THEMES[0].id,
    customAccent: saved.customAccent || null, logo: saved.logo || null
  });
  applyBrandUI();
  // 回填表單
  $("#topicInput").value = state.topic;
  $("#directionInput").value = state.direction;
  $("#audienceInput").value = state.audience;
  $("#slideCount").value = state.slideCount;
  $("#langSelect").value = state.lang;
  $("#toneSelect").value = state.tone;
  // 回填目的與風格選取
  if (state.purpose) {
    const pc = $(`.purpose-card[data-id="${state.purpose}"]`);
    if (pc) { pc.classList.add("selected"); $("#toStep2").disabled = false; }
  }
  renderThemes();
  // 續接到上次步驟
  const step = saved.step || 1;
  if (state.deck && step >= 3) { renderOutline(); goStep(step); }
  else if (step === 2) goStep(2);
  toast("已還原上次進度");
  return true;
}

/* ---------- 初始化 UI ---------- */
function renderPurposes() {
  const grid = $("#purposeGrid");
  grid.innerHTML = PURPOSES.map(p => `
    <button class="purpose-card" data-id="${p.id}">
      <div class="ico">${p.icon}</div>
      <h3>${p.name}</h3>
      <p>${p.desc}</p>
    </button>`).join("");
  grid.querySelectorAll(".purpose-card").forEach(card => {
    card.addEventListener("click", () => {
      grid.querySelectorAll(".purpose-card").forEach(c => c.classList.remove("selected"));
      card.classList.add("selected");
      state.purpose = card.dataset.id;
      $("#toStep2").disabled = false;
      saveState();
    });
  });
}

function renderThemes() {
  const grid = $("#themeGrid");
  grid.innerHTML = THEMES.map(t => `
    <div class="theme-card ${t.id === state.theme ? "selected" : ""}" data-id="${t.id}">
      <div class="theme-preview" style="background:#${t.bg};color:#${t.text}">
        <div>
          <div class="tp-title" style="color:#${t.title}">簡報標題</div>
          <div class="tp-sub" style="color:#${t.accent}">副標題 · 講者</div>
        </div>
        <div class="tp-bullets" style="color:#${t.text}">
          <span>• 重點一</span><span>• 重點二</span><span>• 重點三</span>
        </div>
        <div class="tp-bar" style="background:#${t.accent}"></div>
      </div>
      <div class="tc-name">${t.name}<small>${t.subtitle}</small></div>
    </div>`).join("");
  grid.querySelectorAll(".theme-card").forEach(card => {
    card.addEventListener("click", () => {
      grid.querySelectorAll(".theme-card").forEach(c => c.classList.remove("selected"));
      card.classList.add("selected");
      state.theme = card.dataset.id;
      saveState();
    });
  });
}

/* ---------- 步驟導航 ---------- */
function goStep(n) {
  if (n === 5) { buildSummary(); state.previewIdx = 0; renderPreview(); }
  state.step = n;
  $$(".panel").forEach(p => p.classList.remove("active"));
  $(`#panel-${n}`).classList.add("active");
  if (n === 3) requestAnimationFrame(growAll);   // 面板顯示後再計算高度
  $$(".step").forEach(s => {
    const sn = +s.dataset.step;
    s.classList.toggle("active", sn === n);
    s.classList.toggle("done", sn < n);
  });
  saveState();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ---------- 收集步驟2輸入 ---------- */
function collectInputs() {
  state.topic = $("#topicInput").value.trim();
  state.direction = $("#directionInput").value.trim();
  state.audience = $("#audienceInput").value.trim();
  state.slideCount = +$("#slideCount").value;
  state.lang = $("#langSelect").value;
  state.tone = $("#toneSelect").value;
}

/* ---------- 建立 AI 提示詞 ---------- */
function buildPrompt() {
  const p = PURPOSES.find(x => x.id === state.purpose) || PURPOSES[PURPOSES.length - 1];
  return `你是一位專業的簡報設計顧問。請為以下需求設計一份簡報的完整大綱與內容。

【簡報目的】${p.name}
【建議架構參考】${p.structure}
【主題】${state.topic}
${state.direction ? `【重點方向 / 既有素材】${state.direction}` : ""}
${state.audience ? `【目標聽眾】${state.audience}` : ""}
【內容投影片張數】約 ${state.slideCount} 張（不含封面）
【語言】${state.lang}
【語氣】${state.tone}

要求：
1. 內容要具體、有資訊量，避免空話；每張投影片 3-5 個重點，每個重點為精煉短句（非整段文字）。
2. 依「目的」與「建議架構」安排合理的投影片順序。
3. 為每張投影片提供 notes（口頭講稿提示，1-3 句）。
4. 為每張投影片指定 layout（版型），讓簡報更專業、有節奏：
   - "section"：章節分隔頁，放在每個主要段落開始前（標題=章節名，bullets 放 1 句簡短說明）。
   - "metrics"：關鍵數據頁（如財務、KPI、成效）；此時每個 bullet 用「數值 | 說明」格式，例如 "180% | 三年投資報酬率"、"NT$5M | 首年投入"。
   - "chart"：需要以圖表呈現的量化比較（如各區占比、逐年趨勢）；此時每個 bullet 用「標籤 | 數值」格式，數值必須是純數字，例如 "北部 | 45"、"2024 | 120"。並加上 "chartType"，可為 "bar"（長條）、"pie"（圓餅）、"line"（折線）。
   - "quote"：金句 / 願景 / 客戶見證頁（bullet1 = 金句，bullet2 = 出處或署名）。
   - "bullets"：一般條列重點（預設）。
   適度穿插 section、metrics 與 chart，讓整份簡報有起承轉合；不要每張都一樣。
5. 只回傳 JSON，不要加任何解說文字或 markdown 標記。

JSON 格式（嚴格遵守）：
{
  "title": "簡報主標題",
  "subtitle": "副標題或一句話定位（可含講者/日期占位）",
  "slides": [
    { "title": "投影片標題", "layout": "bullets", "chartType": "bar", "bullets": ["重點1", "重點2", "重點3"], "notes": "講稿提示" }
  ]
}`;
}

/* ---------- 呼叫 LLM ---------- */
async function callLLM(prompt) {
  const { provider, model, key } = getSettings();
  const mdl = model || DEFAULT_MODELS[provider];
  if (!key) throw new Error("NO_KEY");

  if (provider === "anthropic") {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify({ model: mdl, max_tokens: 4000, messages: [{ role: "user", content: prompt }] })
    });
    if (!r.ok) throw new Error(`Anthropic ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.content.map(c => c.text || "").join("");
  }
  if (provider === "openai") {
    const r = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "content-type": "application/json", "Authorization": `Bearer ${key}` },
      body: JSON.stringify({ model: mdl, temperature: 0.7, messages: [{ role: "user", content: prompt }] })
    });
    if (!r.ok) throw new Error(`OpenAI ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.choices[0].message.content;
  }
  if (provider === "gemini") {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${mdl}:generateContent?key=${key}`;
    const r = await fetch(url, {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
    });
    if (!r.ok) throw new Error(`Gemini ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.candidates[0].content.parts.map(p => p.text || "").join("");
  }
  throw new Error("未知的供應商");
}

/* ---------- 解析 JSON ---------- */
function extractJson(raw) {
  let txt = (raw || "").trim().replace(/```json/gi, "").replace(/```/g, "").trim();
  const s = txt.indexOf("{"), e = txt.lastIndexOf("}");
  if (s === -1 || e === -1) throw new Error("找不到 JSON 內容");
  return JSON.parse(txt.slice(s, e + 1));
}
function parseDeck(raw) {
  const obj = extractJson(raw);
  if (!obj.slides || !Array.isArray(obj.slides)) throw new Error("JSON 缺少 slides 陣列");
  obj.slides = obj.slides.map(sl => ({
    title: sl.title || "未命名投影片",
    bullets: Array.isArray(sl.bullets) ? sl.bullets : (sl.content ? [sl.content] : []),
    notes: sl.notes || "",
    layout: LAYOUT_IDS.includes(sl.layout) ? sl.layout : "bullets",
    chartType: CHART_IDS.includes(sl.chartType) ? sl.chartType : "bar"
  }));
  return { title: obj.title || state.topic, subtitle: obj.subtitle || "", slides: obj.slides };
}

/* ---------- 產生大綱（主流程） ---------- */
async function generateOutline() {
  collectInputs();
  if (!state.topic) { toast("請先輸入簡報主題", true); return; }
  const prompt = buildPrompt();
  const { key } = getSettings();
  if (!key) { openManual(prompt); return; }

  loading(true, "AI 正在撰寫大綱與內容…");
  try {
    const raw = await callLLM(prompt);
    state.deck = parseDeck(raw);
    renderOutline();
    saveState();
    loading(false);
    goStep(3);
  } catch (err) {
    loading(false);
    if (err.message === "NO_KEY") { openManual(prompt); return; }
    console.error(err);
    toast("產生失敗：" + err.message.slice(0, 120), true);
  }
}

/* ---------- 單張重寫 ---------- */
async function regenerateSlide(card) {
  const { key } = getSettings();
  if (!key) { toast("單張重寫需要 API 金鑰，請先到「⚙ API 設定」設定", true); return; }
  harvestOutline();
  const idx = [...$$("#slidesEditor .slide-card")].indexOf(card);
  const cur = state.deck.slides[idx];
  const p = PURPOSES.find(x => x.id === state.purpose) || PURPOSES[PURPOSES.length - 1];
  const prompt = `這是一份「${p.name}」簡報，主題「${state.deck.title}」，語言${state.lang}、語氣${state.tone}。
請只重寫其中一張投影片，讓內容更精煉具體。原標題：「${cur.title}」，原重點：${JSON.stringify(cur.bullets)}。
只回傳 JSON：{ "title": "...", "bullets": ["...","..."], "notes": "講稿提示" }`;
  loading(true, "重寫這張投影片中…");
  try {
    const obj = extractJson(await callLLM(prompt));
    const nu = {
      title: obj.title || cur.title,
      bullets: Array.isArray(obj.bullets) ? obj.bullets : cur.bullets,
      notes: obj.notes || cur.notes,
      layout: cur.layout || "bullets",
      chartType: cur.chartType || "bar"
    };
    state.deck.slides[idx] = nu;
    const fresh = slideCardEl(nu, idx);
    card.replaceWith(fresh);
    renumber(); requestAnimationFrame(growAll); saveState();
    loading(false); toast("已重寫這張投影片");
  } catch (err) { loading(false); toast("重寫失敗：" + err.message.slice(0, 100), true); }
}

/* ---------- 渲染可編輯大綱 ---------- */
function renderOutline() {
  $("#deckTitle").value = state.deck.title;
  $("#deckSubtitle").value = state.deck.subtitle;
  const ed = $("#slidesEditor");
  ed.innerHTML = "";
  state.deck.slides.forEach((sl, i) => ed.appendChild(slideCardEl(sl, i)));
}

function slideCardEl(sl, i) {
  const card = document.createElement("div");
  card.className = "slide-card";
  card.innerHTML = `
    <div class="sc-head">
      <div class="num">${i + 1}</div>
      <input class="sc-title" value="${escapeHtml(sl.title)}" />
      <div class="sc-ctrls">
        <button class="sc-up" title="上移">▲</button>
        <button class="sc-down" title="下移">▼</button>
        <button class="sc-regen" title="用 AI 重寫這張">↻</button>
        <button class="sc-del" title="刪除這張">✕</button>
      </div>
    </div>
    <div class="sc-layout-row">
      <label>版型</label>
      <select class="sc-layout">${LAYOUTS.map(l => `<option value="${l.id}" ${l.id === (sl.layout || "bullets") ? "selected" : ""}>${l.name}</option>`).join("")}</select>
      <select class="sc-charttype" hidden>${CHART_TYPES.map(c => `<option value="${c.id}" ${c.id === (sl.chartType || "bar") ? "selected" : ""}>${c.name}</option>`).join("")}</select>
      <span class="sc-hint"></span>
    </div>
    <div class="bullets"></div>
    <button class="add-bullet">＋ 新增重點</button>
    <div class="sc-notes">
      <button class="notes-toggle">📝 講稿備註${sl.notes ? "（已有內容）" : ""}</button>
      <textarea class="notes-area" placeholder="口頭講稿提示…" hidden>${escapeHtml(sl.notes)}</textarea>
    </div>`;
  const bl = card.querySelector(".bullets");
  (sl.bullets.length ? sl.bullets : [""]).forEach(b => bl.appendChild(bulletRow(b)));
  card.querySelector(".add-bullet").addEventListener("click", () => { bl.appendChild(bulletRow("")); saveState(); });
  card.querySelector(".sc-del").addEventListener("click", () => { card.remove(); renumber(); collectAll(); });
  card.querySelector(".sc-up").addEventListener("click", () => moveCard(card, -1));
  card.querySelector(".sc-down").addEventListener("click", () => moveCard(card, 1));
  card.querySelector(".sc-regen").addEventListener("click", () => regenerateSlide(card));
  card.querySelector(".notes-toggle").addEventListener("click", () => {
    const ta = card.querySelector(".notes-area");
    ta.hidden = !ta.hidden;
    if (!ta.hidden) { autoGrow(ta); ta.focus(); }
  });
  const lay = card.querySelector(".sc-layout");
  const chartSel = card.querySelector(".sc-charttype");
  const hint = card.querySelector(".sc-hint");
  const updateHint = () => {
    hint.textContent = (LAYOUTS.find(l => l.id === lay.value) || {}).hint || "";
    chartSel.hidden = lay.value !== "chart";
  };
  updateHint();
  lay.addEventListener("change", () => { updateHint(); saveState(); });
  chartSel.addEventListener("change", saveState);
  return card;
}
function bulletRow(text) {
  const row = document.createElement("div");
  row.className = "bullet-row";
  row.innerHTML = `<span class="dot">●</span>
    <textarea rows="1">${escapeHtml(text)}</textarea>
    <button class="b-del" title="刪除">✕</button>`;
  const ta = row.querySelector("textarea");
  autoGrow(ta); ta.addEventListener("input", () => autoGrow(ta));
  row.querySelector(".b-del").addEventListener("click", () => { row.remove(); saveState(); });
  return row;
}
function moveCard(card, dir) {
  if (dir < 0 && card.previousElementSibling) card.parentNode.insertBefore(card, card.previousElementSibling);
  if (dir > 0 && card.nextElementSibling) card.parentNode.insertBefore(card.nextElementSibling, card);
  renumber(); collectAll();
}
function autoGrow(ta){ ta.style.height="auto"; if(ta.scrollHeight) ta.style.height=ta.scrollHeight+"px"; }
function growAll(){ $$("#slidesEditor textarea").forEach(autoGrow); }
function renumber(){ $$("#slidesEditor .slide-card .num").forEach((n,i)=>n.textContent=i+1); }

/* 從 DOM 讀回大綱 */
function harvestOutline() {
  const slides = [];
  $$("#slidesEditor .slide-card").forEach(card => {
    const title = card.querySelector(".sc-title").value.trim();
    const bullets = [...card.querySelectorAll(".bullet-row textarea")].map(t => t.value.trim()).filter(Boolean);
    const notes = card.querySelector(".notes-area").value.trim();
    const layout = card.querySelector(".sc-layout").value;
    const chartType = card.querySelector(".sc-charttype").value;
    slides.push({ title, bullets, notes, layout, chartType });
  });
  state.deck = {
    title: $("#deckTitle").value.trim() || state.topic,
    subtitle: $("#deckSubtitle").value.trim(),
    slides
  };
}

/* ---------- 摘要 ---------- */
function buildSummary() {
  harvestOutline();
  const p = PURPOSES.find(x => x.id === state.purpose);
  const t = activeTheme();
  $("#summaryBox").innerHTML = `
    <div class="row"><span class="k">簡報目的</span><span class="v">${p ? p.name : "—"}</span></div>
    <div class="row"><span class="k">主題</span><span class="v">${escapeHtml(state.deck.title)}</span></div>
    <div class="row"><span class="k">設計風格</span><span class="v">${t.name}（${t.subtitle}）</span></div>
    <div class="row"><span class="k">投影片數</span><span class="v">封面 + ${state.deck.slides.length} 張內容 + 結尾頁</span></div>
    <div class="row"><span class="k">語言 / 語氣</span><span class="v">${state.lang} · ${state.tone}</span></div>`;
}

/* ---------- 即時預覽 ---------- */
function previewModel() {
  const arr = [{ type: "cover" }];
  state.deck.slides.forEach((_, i) => arr.push({ type: "content", i }));
  arr.push({ type: "closing" });
  return arr;
}
function splitCols(bullets) {
  if (bullets.length <= TWO_COL_THRESHOLD) return [bullets];
  const mid = Math.ceil(bullets.length / 2);
  return [bullets.slice(0, mid), bullets.slice(mid)];
}
function sectionNumberFor(idx) {
  let n = 0;
  for (let k = 0; k <= idx; k++) if ((state.deck.slides[k].layout || "bullets") === "section") n++;
  return n;
}
function logoImg(cls) {
  return state.logo && state.logo.data ? `<img class="pv-logo ${cls || ""}" src="${state.logo.data}" alt="logo" />` : "";
}
// 預覽用的輕量 SVG 圖表（與 PPTX 圖表對應）
function pvChartSvg(labels, values, type, t) {
  const max = Math.max(...values, 1), n = values.length;
  const col = chartPalette(t, type === "pie" ? n : 1);
  if (type === "pie") {
    const total = values.reduce((a, b) => a + b, 0) || 1;
    let ang = -Math.PI / 2; const cx = 26, cy = 30, r = 23; let paths = "";
    values.forEach((v, i) => {
      const a2 = ang + (v / total) * 2 * Math.PI;
      const x1 = cx + r * Math.cos(ang), y1 = cy + r * Math.sin(ang);
      const x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
      const large = (a2 - ang) > Math.PI ? 1 : 0;
      paths += `<path d="M${cx},${cy} L${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${large} 1 ${x2.toFixed(2)},${y2.toFixed(2)} Z" fill="#${col[i]}" stroke="#${t.bg}" stroke-width="0.6"/>`;
      ang = a2;
    });
    const legend = labels.map((l, i) => `<g transform="translate(58,${12 + i * 7})"><rect width="4" height="4" rx="0.6" fill="#${col[i]}"/><text x="6" y="3.6" font-size="3.3" fill="#${t.text}">${escapeHtml(l)} (${values[i]})</text></g>`).join("");
    return `<svg viewBox="0 0 100 60" class="pv-svg">${paths}${legend}</svg>`;
  }
  const gap = 4, bw = (100 - gap * (n + 1)) / n, base = 52;
  if (type === "line") {
    const pts = values.map((v, i) => [gap + i * (bw + gap) + bw / 2, base - (v / max) * 44]);
    const poly = pts.map(p => p.map(z => z.toFixed(2)).join(",")).join(" ");
    const dots = pts.map((p, i) => `<circle cx="${p[0].toFixed(2)}" cy="${p[1].toFixed(2)}" r="1.2" fill="#${col[0]}"/><text x="${p[0].toFixed(2)}" y="${(p[1] - 2).toFixed(2)}" font-size="3" fill="#${t.text}" text-anchor="middle">${values[i]}</text>`).join("");
    const cats = labels.map((l, i) => `<text x="${(gap + i * (bw + gap) + bw / 2).toFixed(2)}" y="58" font-size="3" fill="#${t.footer}" text-anchor="middle">${escapeHtml(l)}</text>`).join("");
    return `<svg viewBox="0 0 100 60" class="pv-svg"><line x1="2" y1="${base}" x2="98" y2="${base}" stroke="#${t.footer}" stroke-width="0.4"/><polyline points="${poly}" fill="none" stroke="#${col[0]}" stroke-width="1"/>${dots}${cats}</svg>`;
  }
  const bars = values.map((v, i) => {
    const h = (v / max) * 44, x = gap + i * (bw + gap), y = base - h;
    return `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${bw.toFixed(2)}" height="${h.toFixed(2)}" rx="0.6" fill="#${col[0]}"/>`
      + `<text x="${(x + bw / 2).toFixed(2)}" y="${(y - 1.2).toFixed(2)}" font-size="3" fill="#${t.text}" text-anchor="middle">${v}</text>`
      + `<text x="${(x + bw / 2).toFixed(2)}" y="58" font-size="3" fill="#${t.footer}" text-anchor="middle">${escapeHtml(labels[i])}</text>`;
  }).join("");
  return `<svg viewBox="0 0 100 60" class="pv-svg"><line x1="2" y1="${base}" x2="98" y2="${base}" stroke="#${t.footer}" stroke-width="0.4"/>${bars}</svg>`;
}
function renderPreview() {
  if (!state.deck) return;
  const t = activeTheme();
  const model = previewModel();
  state.previewIdx = Math.max(0, Math.min(state.previewIdx, model.length - 1));
  const item = model[state.previewIdx];
  const stage = $("#previewStage");
  $("#previewCounter").textContent = `${state.previewIdx + 1} / ${model.length}`;

  const foot = i => `<div class="pv-foot">
      <span style="color:#${t.footer}">${escapeHtml(state.deck.title)}</span>
      <span style="color:#${t.footer}">${i + 1} / ${state.deck.slides.length}</span></div>`;
  const header = title => `
      <div class="pv-strip" style="background:#${t.accent}"></div>
      ${logoImg()}
      <div class="pv-kicker" style="color:#${t.accent}">${escapeHtml(kickerText())}</div>
      <div class="pv-title" style="color:#${t.title}">${escapeHtml(title)}</div>
      <div class="pv-underline" style="background:#${t.accent}"></div>`;

  if (item.type === "cover") {
    stage.style.background = "#" + t.coverBg;
    stage.innerHTML = `
      <div class="pv-sidebar" style="background:#${t.coverAccent}"></div>
      ${logoImg("pv-logo-cover")}
      <div class="pv-cover-kicker" style="color:#${t.coverText}">${escapeHtml(kickerText())}</div>
      <div class="pv-cover-title" style="color:#${t.coverText}">${escapeHtml(state.deck.title)}</div>
      <div class="pv-cover-accent" style="background:#${t.coverAccent}"></div>
      <div class="pv-cover-sub" style="color:#${t.coverText}">${escapeHtml(state.deck.subtitle)}</div>`;
    return;
  }
  if (item.type === "closing") {
    stage.style.background = "#" + t.coverBg;
    stage.innerHTML = `
      <div class="pv-sidebar" style="background:#${t.coverAccent}"></div>
      <div class="pv-end-title" style="color:#${t.coverText}">謝謝聆聽</div>
      <div class="pv-cover-bar" style="background:#${t.coverAccent}"></div>`;
    return;
  }

  const sl = state.deck.slides[item.i];
  const layout = sl.layout || "bullets";

  if (layout === "section") {
    stage.style.background = "#" + t.coverBg;
    const no = String(sectionNumberFor(item.i)).padStart(2, "0");
    stage.innerHTML = `
      <div class="pv-sidebar" style="background:#${t.coverAccent}"></div>
      <div class="pv-sec-no" style="color:#${t.coverAccent}">${no}</div>
      <div class="pv-sec-bar" style="background:#${t.coverAccent}"></div>
      <div class="pv-sec-title" style="color:#${t.coverText}">${escapeHtml(sl.title)}</div>
      ${sl.bullets[0] ? `<div class="pv-sec-desc" style="color:#${t.coverText}">${escapeHtml(sl.bullets[0])}</div>` : ""}`;
    return;
  }
  if (layout === "chart") {
    const cd = parseChartData(sl.bullets);
    if (cd.labels.length) {
      stage.style.background = "#" + t.bg;
      stage.innerHTML = header(sl.title) + `<div class="pv-chart">${pvChartSvg(cd.labels, cd.values, sl.chartType || "bar", t)}</div>` + foot(item.i);
      return;
    }
  }
  if (layout === "metrics") {
    const ms = parseMetrics(sl.bullets);
    if (ms.length) {
      stage.style.background = "#" + t.bg;
      const cardBg = mix(t.accent, t.bg, 0.9), cardLine = mix(t.accent, t.bg, 0.72);
      const cards = ms.slice(0, 4).map(m => `
        <div class="pv-card" style="background:#${cardBg};border-color:#${cardLine}">
          <div class="pv-card-tick" style="background:#${t.accent}"></div>
          <div class="pv-card-val" style="color:#${t.accent}">${escapeHtml(m.value)}</div>
          <div class="pv-card-lbl" style="color:#${t.text}">${escapeHtml(m.label)}</div>
        </div>`).join("");
      stage.innerHTML = header(sl.title) + `<div class="pv-cards">${cards}</div>` + foot(item.i);
      return;
    }
  }
  if (layout === "quote") {
    stage.style.background = "#" + t.bg;
    stage.innerHTML = `
      <div class="pv-sidebar" style="background:#${t.accent}"></div>
      <div class="pv-quote-mark" style="color:#${t.accent}">“</div>
      <div class="pv-quote-text" style="color:#${t.title}">${escapeHtml(sl.bullets[0] || sl.title)}</div>
      ${sl.bullets[1] ? `<div class="pv-quote-attr" style="color:#${t.accent}">— ${escapeHtml(sl.bullets[1])}</div>` : ""}`
      + foot(item.i);
    return;
  }
  // 條列（預設，含自動雙欄）
  stage.style.background = "#" + t.bg;
  const cols = splitCols(sl.bullets).map(col => `
    <div class="pv-col">${col.map(b => `
      <div class="pv-bullet" style="color:#${t.text}">
        <span class="pv-dot" style="background:#${t.bullet}"></span><span>${escapeHtml(b)}</span>
      </div>`).join("")}</div>`).join("");
  stage.innerHTML = header(sl.title)
    + `<div class="pv-rail" style="background:#${t.accent2}"></div>`
    + `<div class="pv-body pv-body-rail">${cols}</div>` + foot(item.i);
}

/* ---------- 生成 PPTX（多版型專業排版） ---------- */
function pptxHeader(pptx, s, t, W, title) {
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.16, fill: { color: t.accent } });
  s.addText(kickerText(), { x: 0.72, y: 0.42, w: W - 1.4, h: 0.3, fontSize: 12, bold: true, color: t.accent, charSpacing: 2, fontFace: t.bodyFont });
  s.addText(title, { x: 0.7, y: 0.74, w: W - 1.4, h: 0.85, fontFace: t.titleFont, fontSize: 28, bold: true, color: t.title, align: "left", valign: "middle" });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.68, w: 1.7, h: 0.07, fill: { color: t.accent } });
}
function pptxFooter(pptx, s, t, W, H, i, total) {
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: H - 0.62, w: W - 1.44, h: 0.012, fill: { color: t.footer, transparency: 45 } });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: H - 0.5, w: 0.14, h: 0.14, fill: { color: t.accent } });
  s.addText(state.deck.title, { x: 0.98, y: H - 0.56, w: W - 3.4, h: 0.3, fontSize: 9.5, color: t.footer, valign: "middle", fontFace: t.bodyFont });
  s.addText(`${i + 1} / ${total}`, { x: W - 1.6, y: H - 0.56, w: 0.88, h: 0.3, fontSize: 9.5, color: t.footer, align: "right", valign: "middle", fontFace: t.bodyFont });
}
function drawLogo(s, W, maxH, y) {
  if (!state.logo || !state.logo.data) return;
  const { data, w, h } = state.logo;
  const hh = maxH, ww = hh * (w / h);
  try { s.addImage({ data, x: W - 0.72 - ww, y, w: ww, h: hh }); } catch (e) { /* 忽略無效圖檔 */ }
}
// 圖表以向量圖形繪製（相容性最佳，不依賴內嵌 Excel 的 addChart）
function pptxChart(pptx, s, t, W, H, sl, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  const { labels, values } = parseChartData(sl.bullets);
  if (!labels.length) { pptxBullets(pptx, s, t, W, H, sl, i, total); return; }
  const ct = sl.chartType || "bar";
  const colors = chartPalette(t, labels.length);
  if (ct === "pie") pptxPie(pptx, s, t, W, H, labels, values, colors);
  else if (ct === "line") pptxLine(pptx, s, t, W, H, labels, values);
  else pptxBar(pptx, s, t, W, H, labels, values, colors);
  pptxFooter(pptx, s, t, W, H, i, total);
}
function pptxBar(pptx, s, t, W, H, labels, values, colors) {
  const x0 = 1.0, plotW = W - 2.0, top = 2.4, baseline = 6.1, plotH = baseline - top;
  const max = Math.max(...values, 1), n = values.length;
  const gap = plotW * 0.03, bw = Math.min(1.6, (plotW - gap * (n + 1)) / n);
  const startX = x0 + (plotW - (bw * n + gap * (n - 1))) / 2;
  s.addShape(pptx.ShapeType.rect, { x: x0, y: baseline, w: plotW, h: 0.015, fill: { color: t.footer, transparency: 30 } });
  labels.forEach((l, k) => {
    const h = Math.max(0.03, values[k] / max * plotH), x = startX + k * (bw + gap), y = baseline - h;
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: bw, h, fill: { color: colors[k] }, rectRadius: 0.04 });
    s.addText(String(values[k]), { x: x - 0.3, y: y - 0.36, w: bw + 0.6, h: 0.32, align: "center", fontSize: 12, bold: true, color: t.title, fontFace: t.bodyFont });
    s.addText(l, { x: x - 0.4, y: baseline + 0.06, w: bw + 0.8, h: 0.4, align: "center", fontSize: 11, color: t.text, fontFace: t.bodyFont });
  });
}
function pptxLine(pptx, s, t, W, H, labels, values) {
  const x0 = 1.0, plotW = W - 2.0, top = 2.55, baseline = 6.1, plotH = baseline - top;
  const max = Math.max(...values, 1), n = values.length, slot = plotW / n;
  const cx = k => x0 + slot * k + slot / 2, cy = v => baseline - v / max * plotH;
  s.addShape(pptx.ShapeType.rect, { x: x0, y: baseline, w: plotW, h: 0.015, fill: { color: t.footer, transparency: 30 } });
  for (let k = 0; k < n - 1; k++) {
    const x1 = cx(k), y1 = cy(values[k]), x2 = cx(k + 1), y2 = cy(values[k + 1]);
    const len = Math.hypot(x2 - x1, y2 - y1), ang = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
    s.addShape(pptx.ShapeType.rect, { x: (x1 + x2) / 2 - len / 2, y: (y1 + y2) / 2 - 0.02, w: len, h: 0.04, fill: { color: t.accent }, rotate: ang });
  }
  labels.forEach((l, k) => {
    const x = cx(k), y = cy(values[k]);
    s.addShape(pptx.ShapeType.ellipse, { x: x - 0.08, y: y - 0.08, w: 0.16, h: 0.16, fill: { color: t.accent } });
    s.addText(String(values[k]), { x: x - 0.4, y: y - 0.42, w: 0.8, h: 0.3, align: "center", fontSize: 11, bold: true, color: t.title, fontFace: t.bodyFont });
    s.addText(l, { x: x - slot / 2, y: baseline + 0.06, w: slot, h: 0.4, align: "center", fontSize: 11, color: t.text, fontFace: t.bodyFont });
  });
}
function pptxPie(pptx, s, t, W, H, labels, values, colors) {
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const d = 3.7, cx = 1.2, cy = 2.5;
  let start = 0;
  values.forEach((v, k) => {
    const end = start + v / total * 360;
    s.addShape(pptx.ShapeType.pie, { x: cx, y: cy, w: d, h: d, fill: { color: colors[k] }, line: { color: t.bg, width: 1 }, angleRange: [start, end] });
    start = end;
  });
  const lx = cx + d + 0.7, ly = cy + 0.15;
  labels.forEach((l, k) => {
    const yy = ly + k * 0.52;
    s.addShape(pptx.ShapeType.rect, { x: lx, y: yy, w: 0.26, h: 0.26, fill: { color: colors[k] } });
    const pct = Math.round(values[k] / total * 100);
    s.addText(`${l}　${values[k]}（${pct}%）`, { x: lx + 0.4, y: yy - 0.06, w: W - lx - 0.9, h: 0.38, fontSize: 13, color: t.text, valign: "middle", fontFace: t.bodyFont });
  });
}
function pptxCover(pptx, t, W, H) {
  const s = pptx.addSlide();
  s.background = { color: t.coverBg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addShape(pptx.ShapeType.rect, { x: W - 3.4, y: H - 3.4, w: 2.6, h: 2.6, fill: { color: t.coverAccent, transparency: 82 } });
  s.addShape(pptx.ShapeType.rect, { x: W - 2.3, y: H - 2.3, w: 2.6, h: 2.6, fill: { color: t.coverAccent, transparency: 90 } });
  s.addText(kickerText(), { x: 0.95, y: 1.75, w: W - 2, h: 0.4, fontSize: 13, bold: true, color: t.coverAccent, charSpacing: 3, fontFace: t.bodyFont });
  s.addText(state.deck.title, { x: 0.9, y: 2.2, w: W - 3, h: 1.9, fontFace: t.titleFont, fontSize: 44, bold: true, color: t.coverText, align: "left", valign: "top" });
  s.addShape(pptx.ShapeType.rect, { x: 0.95, y: 4.2, w: 2.2, h: 0.06, fill: { color: t.coverAccent } });
  if (state.deck.subtitle) s.addText(state.deck.subtitle, { x: 0.95, y: 4.45, w: W - 3, h: 1.0, fontFace: t.bodyFont, fontSize: 18, color: t.coverText, transparency: 10 });
  s.addText("AI 簡報生成器", { x: 0.95, y: H - 0.7, w: 6, h: 0.35, fontSize: 10.5, color: t.coverText, transparency: 40, fontFace: t.bodyFont });
  drawLogo(s, W, 0.7, 0.5);
}
function pptxSection(pptx, s, t, W, H, sl, no) {
  s.background = { color: t.coverBg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addText(String(no).padStart(2, "0"), { x: 0.55, y: 0.1, w: 6, h: 3.4, fontSize: 170, bold: true, color: t.coverAccent, transparency: 80, fontFace: t.titleFont, align: "left", valign: "top" });
  s.addShape(pptx.ShapeType.rect, { x: 0.95, y: 3.95, w: 1.8, h: 0.08, fill: { color: t.coverAccent } });
  s.addText(sl.title, { x: 0.9, y: 4.15, w: W - 1.8, h: 1.3, fontSize: 40, bold: true, color: t.coverText, fontFace: t.titleFont });
  if (sl.bullets[0]) s.addText(sl.bullets[0], { x: 0.95, y: 5.45, w: W - 2, h: 0.9, fontSize: 16, color: t.coverText, transparency: 15, fontFace: t.bodyFont });
}
function pptxBullets(pptx, s, t, W, H, sl, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 2.15, w: 0.09, h: H - 3.1, fill: { color: t.accent2 } });
  const cols = splitCols(sl.bullets);
  const fs = bulletSize(Math.max(...cols.map(c => c.length), 1));
  const colW = (W - 1.9) / cols.length - (cols.length > 1 ? 0.2 : 0);
  cols.forEach((col, ci) => {
    if (!col.length) return;
    s.addText(col.map(b => ({ text: b, options: { bullet: { code: "2022", indent: 16 }, color: t.text, fontSize: fs, paraSpaceAfter: 12 } })), {
      x: 1.0 + ci * (colW + 0.4), y: 2.12, w: colW, h: H - 3.3,
      fontFace: t.bodyFont, valign: "top", lineSpacingMultiple: 1.18, color: t.text
    });
  });
  pptxFooter(pptx, s, t, W, H, i, total);
}
function pptxMetrics(pptx, s, t, W, H, sl, metrics, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  const n = Math.min(metrics.length, 4);
  const gap = 0.35, cardW = (W - 1.44 - gap * (n - 1)) / n, y = 2.55, cardH = 2.7;
  const cardBg = mix(t.accent, t.bg, 0.9), cardLine = mix(t.accent, t.bg, 0.7);
  metrics.slice(0, n).forEach((m, k) => {
    const x = 0.72 + k * (cardW + gap);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cardW, h: cardH, fill: { color: cardBg }, line: { color: cardLine, width: 1 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.3, y: y + 0.4, w: 0.5, h: 0.08, fill: { color: t.accent } });
    s.addText(String(m.value), { x, y: y + 0.55, w: cardW, h: 1.1, align: "center", fontSize: 40, bold: true, color: t.accent, fontFace: t.titleFont });
    s.addText(String(m.label), { x: x + 0.25, y: y + 1.7, w: cardW - 0.5, h: 0.85, align: "center", valign: "top", fontSize: 14, color: t.text, fontFace: t.bodyFont });
  });
  pptxFooter(pptx, s, t, W, H, i, total);
}
function pptxQuote(pptx, s, t, W, H, sl, i, total) {
  s.background = { color: t.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.accent } });
  s.addText("“", { x: 0.55, y: 0.1, w: 3, h: 3, fontSize: 200, bold: true, color: t.accent, transparency: 80, fontFace: "Georgia" });
  s.addText(sl.bullets[0] || sl.title, { x: 1.5, y: 2.1, w: W - 3, h: 2.8, fontSize: 30, italic: true, bold: true, color: t.title, fontFace: t.titleFont, valign: "middle" });
  if (sl.bullets[1]) s.addText("— " + sl.bullets[1], { x: 1.6, y: 5.1, w: W - 3.2, h: 0.6, fontSize: 16, color: t.accent, fontFace: t.bodyFont });
  pptxFooter(pptx, s, t, W, H, i, total);
}
function pptxClosing(pptx, t, W, H) {
  const s = pptx.addSlide();
  s.background = { color: t.coverBg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addShape(pptx.ShapeType.rect, { x: W - 2.8, y: H - 2.8, w: 2.6, h: 2.6, fill: { color: t.coverAccent, transparency: 85 } });
  s.addText("謝謝聆聽", { x: 0.9, y: 2.6, w: W - 1.8, h: 1.3, fontSize: 48, bold: true, color: t.coverText, fontFace: t.titleFont });
  s.addText("THANK YOU", { x: 0.95, y: 3.95, w: 6, h: 0.5, fontSize: 15, color: t.coverText, transparency: 30, charSpacing: 4, fontFace: t.bodyFont });
  if (state.deck.title) s.addText(state.deck.title, { x: 0.95, y: H - 1.0, w: W - 2, h: 0.5, fontSize: 13, color: t.coverText, transparency: 30, fontFace: t.bodyFont });
  drawLogo(s, W, 0.6, 0.5);
}

function generatePptx() {
  harvestOutline();
  if (!state.deck || !state.deck.slides.length) { toast("沒有可生成的內容", true); return; }
  const t = activeTheme();
  loading(true, "正在生成 PPTX 檔案…");
  try {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "W16x9", width: 13.333, height: 7.5 });
    pptx.layout = "W16x9";
    pptx.author = "AI 簡報生成器";
    pptx.title = state.deck.title;
    const W = 13.333, H = 7.5;

    pptxCover(pptx, t, W, H);
    const total = state.deck.slides.length;
    let sectionNo = 0;
    state.deck.slides.forEach((sl, i) => {
      const layout = sl.layout || "bullets";
      const s = pptx.addSlide();
      s.background = { color: t.bg };
      const isCoverLike = layout === "section";
      if (layout === "section") { sectionNo++; pptxSection(pptx, s, t, W, H, sl, sectionNo); }
      else if (layout === "metrics") { const m = parseMetrics(sl.bullets); m.length ? pptxMetrics(pptx, s, t, W, H, sl, m, i, total) : pptxBullets(pptx, s, t, W, H, sl, i, total); }
      else if (layout === "chart") pptxChart(pptx, s, t, W, H, sl, i, total);
      else if (layout === "quote") pptxQuote(pptx, s, t, W, H, sl, i, total);
      else pptxBullets(pptx, s, t, W, H, sl, i, total);
      drawLogo(s, W, isCoverLike ? 0.6 : 0.4, isCoverLike ? 0.5 : 0.34);
      if (sl.notes) s.addNotes(sl.notes);
    });
    pptxClosing(pptx, t, W, H);

    const fname = (state.deck.title || "簡報").replace(/[\\/:*?"<>|]/g, "").slice(0, 40) || "presentation";
    pptx.writeFile({ fileName: `${fname}.pptx` }).then(() => {
      loading(false); toast("✅ PPTX 已開始下載！");
    }).catch(e => { loading(false); toast("下載失敗：" + e.message, true); });
  } catch (err) {
    loading(false); console.error(err); toast("生成失敗：" + err.message, true);
  }
}
function bulletSize(n){ return n <= 3 ? 20 : n <= 5 ? 18 : 15; }

/* ---------- 匯出 / 匯入 / Markdown ---------- */
function downloadFile(name, text, mime) {
  const blob = new Blob([text], { type: mime });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
function exportJson() {
  harvestOutline();
  downloadFile((state.deck.title || "outline") + ".json", JSON.stringify(state.deck, null, 2), "application/json");
  toast("已匯出 JSON");
}
function deckToMarkdown() {
  harvestOutline();
  let md = `# ${state.deck.title}\n`;
  if (state.deck.subtitle) md += `\n*${state.deck.subtitle}*\n`;
  state.deck.slides.forEach((s, i) => {
    md += `\n## ${i + 1}. ${s.title}\n\n`;
    s.bullets.forEach(b => md += `- ${b}\n`);
    if (s.notes) md += `\n> 講稿：${s.notes}\n`;
  });
  return md;
}
function importJsonFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      state.deck = parseDeck(reader.result);
      renderOutline(); saveState(); goStep(3);
      toast("已匯入大綱");
    } catch (e) { toast("匯入失敗：" + e.message, true); }
  };
  reader.readAsText(file);
}

/* ---------- 手動模式 ---------- */
function openManual(prompt) {
  $("#manualPrompt").value = prompt;
  $("#manualJson").value = "";
  $("#manualModal").hidden = false;
}

/* ---------- 設定 Modal ---------- */
function fillModelDatalist(provider) {
  const dl = $("#modelPresets");
  if (dl) dl.innerHTML = (MODEL_PRESETS[provider] || []).map(m => `<option value="${m}">`).join("");
}
function openSettings() {
  const s = getSettings();
  $("#providerSelect").value = s.provider;
  fillModelDatalist(s.provider);
  $("#modelInput").value = s.model || DEFAULT_MODELS[s.provider];
  $("#apiKeyInput").value = s.key;
  $("#providerHint").textContent = PROVIDER_HINTS[s.provider];
  $("#settingsModal").hidden = false;
}

/* ---------- 企業品牌客製（主色 + Logo） ---------- */
function applyBrandUI() {
  if (state.customAccent) {
    const h = "#" + hx(state.customAccent);
    $("#accentPicker").value = h; $("#accentHex").value = h;
    $("#accentNote").textContent = "已套用企業主色 " + h;
  } else {
    $("#accentNote").textContent = "目前使用風格預設色";
  }
  const hasLogo = state.logo && state.logo.data;
  $("#logoPreview").hidden = !hasLogo;
  $("#clearLogo").hidden = !hasLogo;
  if (hasLogo) $("#logoPreview").src = state.logo.data;
}
function applyAccentFromInput() {
  let v = ($("#accentHex").value.trim() || $("#accentPicker").value).replace(/[^#0-9a-fA-F]/g, "");
  if (!/^#?[0-9a-fA-F]{6}$/.test(v)) { toast("請輸入有效的 6 位色碼，如 #2563EB", true); return; }
  state.customAccent = hx(v).toUpperCase();
  applyBrandUI(); saveState(); toast("已套用企業主色，於預覽 / 下載生效");
}
function handleLogoFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      const max = 260, scale = Math.min(1, max / Math.max(img.width || max, img.height || max));
      const w = Math.max(1, Math.round((img.width || max) * scale)), h = Math.max(1, Math.round((img.height || max) * scale));
      let data;
      try { const cv = document.createElement("canvas"); cv.width = w; cv.height = h; cv.getContext("2d").drawImage(img, 0, 0, w, h); data = cv.toDataURL("image/png"); }
      catch (e) { data = reader.result; }
      state.logo = { data, w, h };
      applyBrandUI(); saveState(); toast("Logo 已上傳");
    };
    img.onerror = () => toast("圖片讀取失敗", true);
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

/* ---------- 全部重設 ---------- */
function resetAll() {
  if (!confirm("確定要清除目前的主題、大綱與進度，重新開始嗎？（API 設定會保留）")) return;
  clearState();
  Object.assign(state, { purpose: null, topic: "", direction: "", audience: "", slideCount: 10, lang: "繁體中文", tone: "清楚易懂", deck: null, theme: THEMES[0].id, customAccent: null, logo: null, step: 1, previewIdx: 0 });
  ["topicInput", "directionInput", "audienceInput"].forEach(id => $("#" + id).value = "");
  $("#slideCount").value = "10"; $("#langSelect").value = "繁體中文"; $("#toneSelect").value = "清楚易懂";
  $("#slidesEditor").innerHTML = "";
  $("#toStep2").disabled = true;
  renderPurposes(); renderThemes(); applyBrandUI();
  goStep(1);
  toast("已重設");
}

/* ---------- 事件綁定 ---------- */
function bindEvents() {
  renderPurposes();
  renderThemes();

  $("#toStep2").addEventListener("click", () => goStep(2));
  $("#generateOutlineBtn").addEventListener("click", generateOutline);
  $("#toStep4").addEventListener("click", () => { harvestOutline(); goStep(4); });
  $("#toStep5").addEventListener("click", () => goStep(5));
  $("#downloadPptxBtn").addEventListener("click", generatePptx);
  $("#addSlideBtn").addEventListener("click", () => {
    $("#slidesEditor").appendChild(slideCardEl({ title: "新投影片", bullets: ["重點"], notes: "", layout: "bullets", chartType: "bar" }, $$("#slidesEditor .slide-card").length));
    renumber(); requestAnimationFrame(growAll); collectAll();
  });
  $("#regenBtn").addEventListener("click", generateOutline);
  $$("[data-goto]").forEach(b => b.addEventListener("click", () => goStep(+b.dataset.goto)));

  // 匯出 / 匯入 / Markdown
  $("#exportJsonBtn").addEventListener("click", exportJson);
  $("#importJsonBtn").addEventListener("click", () => $("#importFile").click());
  $("#importFile").addEventListener("change", e => { if (e.target.files[0]) importJsonFile(e.target.files[0]); e.target.value = ""; });
  $("#copyMdBtn").addEventListener("click", () => navigator.clipboard.writeText(deckToMarkdown()).then(() => toast("Markdown 已複製")));

  // 預覽
  $("#prevSlideBtn").addEventListener("click", () => { state.previewIdx--; renderPreview(); });
  $("#nextSlideBtn").addEventListener("click", () => { state.previewIdx++; renderPreview(); });

  // 企業品牌客製
  $("#applyAccent").addEventListener("click", applyAccentFromInput);
  $("#accentPicker").addEventListener("input", () => { $("#accentHex").value = $("#accentPicker").value.toUpperCase(); });
  $("#accentHex").addEventListener("change", () => { const v = $("#accentHex").value.trim(); if (/^#?[0-9a-fA-F]{6}$/.test(v)) $("#accentPicker").value = "#" + hx(v); });
  $("#clearAccent").addEventListener("click", () => { state.customAccent = null; applyBrandUI(); saveState(); toast("已還原風格預設色"); });
  $("#logoBtn").addEventListener("click", () => $("#logoFile").click());
  $("#logoFile").addEventListener("change", e => { if (e.target.files[0]) handleLogoFile(e.target.files[0]); e.target.value = ""; });
  $("#clearLogo").addEventListener("click", () => { state.logo = null; applyBrandUI(); saveState(); toast("已移除 Logo"); });

  // 重設
  $("#resetAllBtn").addEventListener("click", resetAll);

  // 設定
  $("#settingsBtn").addEventListener("click", openSettings);
  $("#closeSettings").addEventListener("click", () => $("#settingsModal").hidden = true);
  $("#providerSelect").addEventListener("change", e => {
    fillModelDatalist(e.target.value);
    $("#modelInput").value = DEFAULT_MODELS[e.target.value];
    $("#providerHint").textContent = PROVIDER_HINTS[e.target.value];
  });
  $("#saveSettings").addEventListener("click", () => {
    const ok = lsSet("ppt_provider", $("#providerSelect").value)
      && lsSet("ppt_model", $("#modelInput").value.trim())
      && lsSet("ppt_key", $("#apiKeyInput").value.trim());
    $("#settingsModal").hidden = true;
    toast(ok ? "設定已儲存" : "瀏覽器停用了本機儲存，設定僅在本次有效", !ok);
  });

  // 手動模式
  $("#closeManual").addEventListener("click", () => $("#manualModal").hidden = true);
  $("#copyPromptBtn").addEventListener("click", () => navigator.clipboard.writeText($("#manualPrompt").value).then(() => toast("提示詞已複製")));
  $("#applyManual").addEventListener("click", () => {
    try {
      state.deck = parseDeck($("#manualJson").value);
      $("#manualModal").hidden = true;
      renderOutline(); saveState(); goStep(3);
      toast("大綱已套用");
    } catch (e) { toast("JSON 解析失敗：" + e.message, true); }
  });

  // 點背景關閉 modal
  $$(".modal-backdrop").forEach(m => m.addEventListener("click", e => { if (e.target === m) m.hidden = true; }));

  // 自動存檔：表單與大綱有變動就儲存（金鑰欄不納入工作存檔）
  ["topicInput", "directionInput", "audienceInput", "slideCount", "langSelect", "toneSelect"].forEach(id =>
    $("#" + id).addEventListener("input", debouncedSave));
  $("#slidesEditor").addEventListener("input", debouncedSave);
  $("#deckTitle").addEventListener("input", debouncedSave);
  $("#deckSubtitle").addEventListener("input", debouncedSave);

  // 還原上次進度
  restoreState();
}

document.addEventListener("DOMContentLoaded", bindEvents);

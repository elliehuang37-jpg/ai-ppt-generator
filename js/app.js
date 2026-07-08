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
const PROVIDER_HINTS_EN = {
  anthropic: "Get a key at console.anthropic.com. Direct browser access is enabled via anthropic-dangerous-direct-browser-access.",
  openai: "Get a key at platform.openai.com. Common models: gpt-4o, gpt-4o-mini.",
  gemini: "Get a key at aistudio.google.com. Common models: gemini-1.5-flash, gemini-1.5-pro."
};
function providerHint(p) { return (UI_LANG === "en" ? PROVIDER_HINTS_EN : PROVIDER_HINTS)[p]; }

const TWO_COL_THRESHOLD = 6;   // 重點超過此數 → 自動雙欄
const STORE_KEY = "ppt_workstate";

const state = {
  purpose: null,
  topic: "", direction: "", audience: "", slideCount: 10, lang: "繁體中文", tone: "清楚易懂",
  deck: null,          // { title, subtitle, slides:[{title,bullets:[],notes,layout,chartType}] }
  theme: THEMES[0].id,
  customAccent: null,  // 企業主色（覆蓋風格主色）
  logo: null,          // { data, w, h } 品牌 Logo（已縮圖的 dataURL）
  coverStyle: "left",  // 封面版式：left | center | full
  step: 1,
  previewIdx: 0
};
const COVER_STYLES = [
  { id: "left", name: "左置經典", nameEn: "Left classic" },
  { id: "center", name: "置中對稱", nameEn: "Centered" },
  { id: "full", name: "滿版橫幅", nameEn: "Full banner" }
];

/* ---------- 工具 ---------- */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
function toast(msg, isErr = false) {
  const t = $("#toast");
  t.textContent = tMsg(msg); t.hidden = false;
  t.classList.toggle("err", isErr);
  clearTimeout(t._t); t._t = setTimeout(() => (t.hidden = true), 3200);
}
function loading(show, text = "AI 思考中…") {
  $("#loadingText").textContent = tMsg(text);
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
  { id: "bullets", name: "條列重點", nameEn: "Bullets", hint: "", hintEn: "" },
  { id: "section", name: "章節分隔", nameEn: "Section divider", hint: "此頁作為章節封面：標題即章節名，第一個重點作為簡短說明。", hintEn: "Section cover: title = section name, first bullet = short description." },
  { id: "metrics", name: "數據亮點 (KPI)", nameEn: "KPI metrics", hint: "每個重點請用「數值 | 說明」，例如：180% | 三年 ROI。", hintEn: "Each bullet as “value | label”, e.g. 180% | 3-yr ROI." },
  { id: "chart", name: "圖表 (長條/圓餅)", nameEn: "Chart (bar/pie)", hint: "每列「標籤 | 數值」；多系列用「標籤 | 值1 | 值2」，並可加一列「#系列 | 名稱1 | 名稱2」定義圖例。", hintEn: "Each row “label | value”; multi-series “label | v1 | v2”, and a “#series | name1 | name2” row for the legend." },
  { id: "quote", name: "金句引言", nameEn: "Quote", hint: "第一個重點 = 金句，第二個重點 = 出處/署名。", hintEn: "First bullet = quote, second = attribution." }
];
const LAYOUT_IDS = LAYOUTS.map(l => l.id);
const CHART_TYPES = [
  { id: "bar", name: "長條圖", nameEn: "Bar" },
  { id: "pie", name: "圓餅圖", nameEn: "Pie" },
  { id: "line", name: "折線圖", nameEn: "Line" }
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

// 視覺質感：陰影 + 漸層背景（canvas 產生圖片，跨投影片以快取重用）
const SHADOW = { type: "outer", angle: 90, blur: 9, offset: 3, color: "1A2740", opacity: 0.22 };
function rgbaStr(hex, a){ const [r, g, b] = _rgb(hex); return `rgba(${r},${g},${b},${a})`; }
const _bgCache = {};
function makeGradient(kind, t){
  const key = kind + t.coverBg + t.coverAccent;
  if (_bgCache[key]) return _bgCache[key];
  const W = 1280, H = 720, cv = document.createElement("canvas");
  cv.width = W; cv.height = H;
  const g = cv.getContext("2d");
  const b1 = kind === "section" ? darken(t.coverBg, 0.12) : t.coverBg;
  const b2 = kind === "section" ? mix(t.coverBg, t.coverAccent, 0.12) : darken(t.coverBg, 0.42);
  const lin = g.createLinearGradient(0, 0, W * 0.65, H);
  lin.addColorStop(0, "#" + b1); lin.addColorStop(1, "#" + b2);
  g.fillStyle = lin; g.fillRect(0, 0, W, H);
  const blob = (cx, cy, r, a) => {
    const rg = g.createRadialGradient(cx, cy, 0, cx, cy, r);
    rg.addColorStop(0, rgbaStr(t.coverAccent, a)); rg.addColorStop(1, rgbaStr(t.coverAccent, 0));
    g.fillStyle = rg; g.fillRect(0, 0, W, H);
  };
  blob(W * 0.83, H * 0.16, 560, 0.32); blob(W * 0.04, H * 0.97, 520, 0.16);
  let d; try { d = cv.toDataURL("image/jpeg", 0.85); } catch (e) { d = null; }
  _bgCache[key] = d; return d;
}
// 預覽用的漸層 CSS（與 makeGradient 對應）
function gradientCss(kind, t){
  const b1 = kind === "section" ? darken(t.coverBg, 0.12) : t.coverBg;
  const b2 = kind === "section" ? mix(t.coverBg, t.coverAccent, 0.12) : darken(t.coverBg, 0.42);
  return `radial-gradient(circle at 83% 16%, ${rgbaStr(t.coverAccent, 0.32)}, transparent 42%),`
    + `radial-gradient(circle at 4% 97%, ${rgbaStr(t.coverAccent, 0.16)}, transparent 40%),`
    + `linear-gradient(135deg, #${b1}, #${b2})`;
}
// 由主色/輔色推導的圖表色盤
function chartPalette(t, n){
  const base = t.clean
    ? [t.accent, "9AA0A6", "5F6368", "C8CCD1", "E5A0B4", "3C4043"]   // Avon Red + 中性色（品牌規範）
    : [t.accent, t.accent2, lighten(t.accent, 0.28), darken(t.accent, 0.2), lighten(t.accent2, 0.2), mix(t.accent, t.accent2, 0.5)];
  const out = [];
  for (let i = 0; i < n; i++) out.push(base[i % base.length]);
  return out;
}
// 座標軸「漂亮」上限（1/2/5 ×10^n）
function niceMax(v){
  if (v <= 0) return 1;
  const exp = Math.floor(Math.log10(v)), f = v / Math.pow(10, exp);
  const nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * Math.pow(10, exp);
}
// 解析圖表資料（支援多系列）：
//  一般列「標籤 | 值1 | 值2 …」；系列名稱列以 # / 系列 / 圖例 開頭「#系列 | 名稱1 | 名稱2」
function parseChartSeries(bullets){
  const SEP = /\s*[|｜\t]\s*/;
  let seriesNames = null;
  const labels = [], rows = [];
  (bullets || []).forEach(b => {
    const parts = String(b).trim().split(SEP);
    if (parts.length < 2) return;
    const head = parts[0].trim(), rest = parts.slice(1).map(s => s.trim());
    if (/^(#|系列|圖例|legend)/i.test(head)) { seriesNames = rest; return; }
    const nums = rest.map(s => parseFloat(s.replace(/[^0-9.\-]/g, "")));
    const clean = nums.filter(x => !isNaN(x));
    if (!clean.length) return;
    labels.push(head);
    rows.push(nums.map(x => isNaN(x) ? 0 : x));
  });
  if (!labels.length) return null;
  const cols = Math.max(...rows.map(r => r.length));
  const series = [];
  for (let c = 0; c < cols; c++) {
    series.push({ name: (seriesNames && seriesNames[c]) || (cols > 1 ? `系列${c + 1}` : ""), values: labels.map((_, i) => rows[i][c] ?? 0) });
  }
  return { labels, series };
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
    customAccent: state.customAccent, logo: state.logo, coverStyle: state.coverStyle
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
    customAccent: saved.customAccent || null, logo: saved.logo || null, coverStyle: saved.coverStyle || "left"
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
    <button class="purpose-card ${p.id === state.purpose ? "selected" : ""}" data-id="${p.id}">
      <div class="ico">${p.icon}</div>
      <h3>${tLang(p.name, p.nameEn)}</h3>
      <p>${tLang(p.desc, p.descEn)}</p>
    </button>`).join("");
  if (state.purpose) $("#toStep2").disabled = false;
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
      <div class="tc-name">${tLang(t.name, t.nameEn)}<small>${tLang(t.subtitle, t.subtitleEn)}</small></div>
    </div>`).join("");
  grid.querySelectorAll(".theme-card").forEach(card => {
    card.addEventListener("click", () => {
      grid.querySelectorAll(".theme-card").forEach(c => c.classList.remove("selected"));
      card.classList.add("selected");
      state.theme = card.dataset.id;
      // Avon 品牌軟預設：若語言仍為預設繁中，切換為中英雙語
      if (state.theme === "avon" && state.lang === "繁體中文") {
        state.lang = "English + 繁體中文（雙語並列）";
        const ls = $("#langSelect"); if (ls) ls.value = state.lang;
        toast("Avon 品牌預設：已切換為中英雙語（可於步驟 2 改回）");
      }
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

/* ---------- 匯入文案 → 可編輯簡報 ---------- */
// 純本機解析：Markdown / 純文字 → deck（不需 API）
function textToDeck(text) {
  const lines = String(text).replace(/\r/g, "").split("\n");
  let deckTitle = "", deckSub = "";
  const slides = [];
  let cur = null, blank = false;
  const push = () => { if (cur) slides.push(cur); cur = null; };
  const open = title => { push(); cur = { title: title || "", bullets: [], notes: "", layout: "bullets", chartType: "bar" }; };
  lines.forEach(raw => {
    const line = raw.trim();
    if (!line) { blank = true; return; }
    const h1 = line.match(/^#\s+(.*)/);
    const h2 = line.match(/^#{2,6}\s+(.*)/);
    const b = line.match(/^(?:[-*•·◦]|\d+[.)、]|[（(]\d+[)）])\s+(.*)/);
    if (h1) {
      if (!deckTitle && slides.length === 0 && !cur) deckTitle = h1[1];
      else open(h1[1]);
    } else if (h2) {
      open(h2[1]);
    } else if (b) {
      if (!cur) open("");
      cur.bullets.push(b[1]);
    } else {
      if (!cur) open(line);
      else if (!cur.title) cur.title = line;
      else if (blank && cur.bullets.length) open(line);
      else cur.bullets.push(line);
    }
    blank = false;
  });
  push();
  slides.forEach((s, i) => { if (!s.title) s.title = `投影片 ${i + 1}`; });
  if (!deckTitle) deckTitle = slides.length ? slides[0].title : "未命名簡報";
  return { title: deckTitle, subtitle: deckSub, slides };
}
// 讀取 .docx（用 PptxGenJS 內建的 JSZip 解壓 word/document.xml）
async function docxToText(buf) {
  if (typeof JSZip === "undefined") throw new Error("缺少解壓元件，請改貼純文字");
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file("word/document.xml");
  if (!f) throw new Error("這不是有效的 .docx");
  let xml = await f.async("string");
  xml = xml.replace(/<\/w:p>/g, "\n").replace(/<w:tab[^>]*\/?>/g, "\t").replace(/<[^>]+>/g, "");
  return xml.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}
function readImportFile(file) {
  $("#importFname").textContent = file.name;
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const r = new FileReader();
  if (ext === "docx") {
    r.onload = async () => { try { $("#importText").value = (await docxToText(r.result)).trim(); toast("已讀取 Word 文字"); } catch (e) { toast("Word 解析失敗：" + e.message, true); } };
    r.readAsArrayBuffer(file);
  } else {
    r.onload = () => { $("#importText").value = String(r.result); toast("已讀取檔案"); };
    r.readAsText(file);
  }
}
function buildImportPrompt(content) {
  const p = PURPOSES.find(x => x.id === state.purpose) || PURPOSES[PURPOSES.length - 1];
  return `你是專業簡報設計顧問。以下是使用者「已經寫好的文案／內容」，請忠實整理成一份結構清楚、可直接上台的簡報。務必以原文為準，不要杜撰未提供的資訊。

【簡報目的】${p.name}
【語言】${state.lang}
【語氣】${state.tone}

要求：
1. 依內容合理分頁；每頁 3-5 個精煉重點（短句）。
2. 適當指定 layout：section（章節分隔）、metrics（數據，重點用「數值 | 說明」）、chart（圖表，重點用「標籤 | 數值」並加 chartType）、quote（金句，bullet1 金句 bullet2 出處）、bullets（預設）。
3. 每頁加 notes（講稿提示）。
4. 只回傳 JSON，不要加任何解說或 markdown 標記。格式：
{ "title":"...", "subtitle":"...", "slides":[ { "title":"...", "layout":"bullets", "chartType":"bar", "bullets":["..."], "notes":"..." } ] }

【文案內容】
"""
${String(content).slice(0, 12000)}
"""`;
}
async function importContent(mode) {
  const text = $("#importText").value.trim();
  if (!text) { toast("請先貼上或上傳文案", true); return; }
  if (!state.purpose) { state.purpose = "general"; }
  collectInputs();
  // Avon 品牌軟預設：雙語
  if (state.theme === "avon" && state.lang === "繁體中文") { state.lang = "English + 繁體中文（雙語並列）"; const ls = $("#langSelect"); if (ls) ls.value = state.lang; }

  if (mode === "direct") {
    const deck = textToDeck(text);
    if (!deck.slides.length) { toast("無法解析出投影片，請檢查格式", true); return; }
    state.deck = deck; renderOutline(); saveState(); goStep(3); showImportHints();
    return;
  }
  const { key } = getSettings();
  const prompt = buildImportPrompt(text);
  if (!key) { openManual(prompt); return; }
  loading(true, "AI 正在整理你的文案…");
  try {
    state.deck = parseDeck(await callLLM(prompt));
    renderOutline(); saveState(); loading(false); goStep(3); showImportHints();
  } catch (err) {
    loading(false);
    if (err.message === "NO_KEY") { openManual(prompt); return; }
    toast("整理失敗：" + err.message.slice(0, 120), true);
  }
}

/* ---------- 匯入後的智慧建議 ---------- */
function analyzeDeck(deck) {
  const dataSlides = [], longSlides = [];
  deck.slides.forEach((s, i) => {
    if ((s.layout || "bullets") !== "bullets") return;
    const bl = s.bullets || [];
    const piped = bl.filter(b => /[|｜]/.test(b)).length;
    const numeric = bl.filter(b => /[\d０-９]/.test(b)).length;
    if (bl.length >= 2 && piped >= Math.ceil(bl.length * 0.6)) dataSlides.push({ i, kind: "chart" });
    else if (bl.length >= 2 && numeric >= Math.ceil(bl.length * 0.7) && bl.every(b => b.length <= 22)) dataSlides.push({ i, kind: "metrics" });
    if (bl.length > 7) longSlides.push({ i });
  });
  return { n: deck.slides.length, dataSlides, longSlides };
}
function showImportHints() {
  const box = $("#importHints");
  if (!box || !state.deck) return;
  const a = analyzeDeck(state.deck);
  const sep = UI_LANG === "en" ? ", " : "、";
  const msgs = [t("hint_n", a.n)];
  if (a.dataSlides.length) {
    const chartN = a.dataSlides.filter(d => d.kind === "chart").map(d => d.i + 1);
    const metricN = a.dataSlides.filter(d => d.kind === "metrics").map(d => d.i + 1);
    if (chartN.length) msgs.push(t("hint_chart", chartN.join(sep)));
    if (metricN.length) msgs.push(t("hint_metrics", metricN.join(sep)));
  }
  if (a.longSlides.length) msgs.push(t("hint_long", a.longSlides.map(d => d.i + 1).join(sep)));
  box.innerHTML = `<div class="hints-msgs">${msgs.map(m => `<div>${m}</div>`).join("")}</div>
    <div class="hints-actions">
      ${a.dataSlides.length ? `<button class="ghost-btn small" id="applyHintsBtn">${t("apply_hints_btn")}</button>` : ""}
      <button class="ghost-btn small" id="dismissHintsBtn">${t("dismiss_hints")}</button>
    </div>`;
  box.hidden = false;
  if (a.dataSlides.length) $("#applyHintsBtn").addEventListener("click", () => applyHintSuggestions(a.dataSlides));
  $("#dismissHintsBtn").addEventListener("click", () => { box.hidden = true; });
}
function applyHintSuggestions(dataSlides) {
  harvestOutline();
  let firstChart = -1;
  dataSlides.forEach(d => {
    if (state.deck.slides[d.i]) { state.deck.slides[d.i].layout = d.kind; if (firstChart < 0 && d.kind === "chart") firstChart = d.i; }
  });
  renderOutline(); requestAnimationFrame(growAll); saveState();
  toast("已套用建議版型，帶你看效果");
  // 自動跳到預覽並定位到第一個圖表頁（沒有圖表則定位第一個數據頁）
  const target = firstChart >= 0 ? firstChart : (dataSlides[0] ? dataSlides[0].i : -1);
  if (target >= 0) {
    goStep(5);
    state.previewIdx = target + 1;   // 預覽模型：封面在最前，內容頁索引 = i + 1
    renderPreview();
    setTimeout(() => { const pw = $(".preview-wrap"); if (pw) pw.scrollIntoView({ behavior: "smooth", block: "center" }); }, 80);
  }
}

/* ---------- 內建本機大綱生成器（無金鑰也能生成） ---------- */
// 依「段落名稱」產生範本重點；T=主題、K=關鍵詞陣列
const SECTION_TPLS = [
  { re: /執行摘要|重點摘要|摘要|吸睛開場|封面/, mk: (T, K) => [`主題：${T}`, K.length ? `聚焦重點：${K.slice(0, 3).join("、")}` : "背景、方案與效益一頁掌握", "關鍵結論與建議行動"] },
  { re: /市場|背景|現況|動機/, mk: (T, K) => [`${T} 的背景與現況`, "市場規模與成長趨勢", K[0] ? `與「${K[0]}」相關的驅動因素` : "主要驅動因素與外部環境", "目前面臨的核心問題"] },
  { re: /機會|挑戰|痛點/, mk: (T, K) => ["最大的機會點與切入時機", K[1] ? `機會領域：${K[1]}` : "未被滿足的需求", "主要挑戰與限制", "為什麼是現在（Why now）"] },
  { re: /策略|主軸|建議方案|解決方案|核心概念/, mk: (T, K) => (K.length ? K.slice(0, 5).map(k => `策略重點：${k}`) : ["策略主軸一：聚焦核心客群", "策略主軸二：強化差異化", "策略主軸三：擴大通路布局"]) },
  { re: /方案比較|可選方案/, mk: () => ["方案 A：優點與成本", "方案 B：優點與成本", "方案 C：維持現狀的風險", "建議採用方案與理由"] },
  { re: /行動|實施|計畫|步驟|方法|練習/, mk: (T, K) => ["第一階段：規劃與資源盤點", K[0] ? `第二階段：啟動「${K[0]}」` : "第二階段：試點執行", "第三階段：全面推展與優化", "所需資源與分工"] },
  { re: /財務|投資|效益|ROI|數據|成果|指標|達成/, mk: () => ["投入成本 | （請填金額）", "預期營收 | （請填金額）", "投資回收期 | （請填月數）", "關鍵 KPI | （請填指標）"], layout: "metrics" },
  { re: /時程|路線圖|里程碑|議程/, mk: () => ["Q1：籌備與資源到位", "Q2：試點與修正", "Q3：擴大執行", "Q4：檢視成效與下年度規劃"] },
  { re: /風險|對策|問題/, mk: () => ["風險一：市場變化 → 對策：滾動式調整", "風險二：資源不足 → 對策：分階段投入", "風險三：執行落差 → 對策：週期檢核"] },
  { re: /案例|見證|示範/, mk: (T) => [`成功案例：與 ${T} 情境相近的實例`, "做法重點與可複製之處", "成果數字與客戶回饋"] },
  { re: /價格|方案與價格|需求|the ask/i, mk: () => ["建議方案與內容", "投資金額與期程", "本次需要的決策與支持"] },
  { re: /結論|下一步|總結|回顧|心得|反思/, mk: (T, K) => [`重申核心：${T}`, K.length ? `三個帶走的重點：${K.slice(0, 3).join("、")}` : "三個帶走的重點", "下一步行動與時間點"] },
  { re: /參考|文獻|資源/, mk: () => ["資料來源一（請補充）", "資料來源二（請補充）", "延伸閱讀與工具"] }
];
function localOutline() {
  const p = PURPOSES.find(x => x.id === state.purpose) || PURPOSES[PURPOSES.length - 1];
  const T = state.topic;
  const K = state.direction.split(/[、,，;；\/\n]+/).map(s => s.trim()).filter(s => s && s.length <= 30);
  // 取目的架構的段落（去掉封面類），依張數裁切
  let secs = p.structure.split(/、/).map(s => s.trim()).filter(s => s && !/^標題頁|^封面/.test(s));
  if (p.id === "general") secs = ["開場與背景", "現況分析", "重點內容", "案例與說明", "效益與價值", "結論與下一步"];
  const n = Math.min(Math.max(state.slideCount, 4), secs.length + 4);
  if (secs.length > n) {
    // 保留頭尾，均勻抽掉中間段
    while (secs.length > n) secs.splice(Math.ceil(secs.length / 2), 1);
  }
  const slides = secs.map(sec => {
    const tpl = SECTION_TPLS.find(x => x.re.test(sec));
    const bullets = tpl ? tpl.mk(T, K) : [`${sec}重點一（請補充）`, `${sec}重點二（請補充）`, `${sec}重點三（請補充）`];
    return {
      title: sec, bullets,
      notes: `此頁說明「${sec}」：先講結論，再補一個例子或數字。`,
      layout: (tpl && tpl.layout) || "bullets", chartType: "bar"
    };
  });
  // 關鍵詞若很多，補一頁「重點方向總覽」
  if (K.length >= 4 && slides.length < state.slideCount + 2) {
    slides.splice(1, 0, { title: "重點方向總覽", bullets: K.slice(0, 6), notes: "快速帶過本次涵蓋的面向。", layout: "bullets", chartType: "bar" });
  }
  return { title: T, subtitle: state.audience ? `對象：${state.audience}` : "（副標題／講者／日期）", slides };
}
// 本機生成後的提示（提供升級到 AI 的路徑）
function showLocalGenNotice(prompt) {
  const box = $("#importHints");
  if (!box) return;
  box.innerHTML = `<div class="hints-msgs"><div>${t("local_gen_msg")}</div></div>
    <div class="hints-actions">
      <button class="ghost-btn small" id="lgSettingsBtn">${t("local_gen_key")}</button>
      <button class="ghost-btn small" id="lgManualBtn">${t("local_gen_manual")}</button>
      <button class="ghost-btn small" id="dismissHintsBtn">${t("dismiss_hints")}</button>
    </div>`;
  box.hidden = false;
  $("#lgSettingsBtn").addEventListener("click", openSettings);
  $("#lgManualBtn").addEventListener("click", () => openManual(prompt));
  $("#dismissHintsBtn").addEventListener("click", () => { box.hidden = true; });
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
   - "chart"：需要以圖表呈現的量化比較（如各區占比、逐年趨勢）；此時每個 bullet 用「標籤 | 數值」格式，數值必須是純數字，例如 "北部 | 45"、"2024 | 120"。並加上 "chartType"，可為 "bar"（長條）、"pie"（圓餅）、"line"（折線）。若要比較多組數列（如「實際 vs 目標」），用「標籤 | 值1 | 值2」，並在第一個 bullet 放「#系列 | 名稱1 | 名稱2」定義圖例名稱（圓餅圖僅取第一組）。
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

  // 沒有金鑰 → 用內建生成器直接產出大綱（不再卡在手動模式）
  if (!key) {
    state.deck = localOutline();
    renderOutline(); saveState(); goStep(3);
    showLocalGenNotice(prompt);
    toast("已用內建範本生成大綱，可直接編輯");
    return;
  }

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
    console.error(err);
    // AI 呼叫失敗 → 退回內建生成器，不讓使用者卡住
    state.deck = localOutline();
    renderOutline(); saveState(); goStep(3);
    showLocalGenNotice(prompt);
    toast("AI 連線失敗，已改用內建範本生成：" + err.message.slice(0, 80), true);
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
  const hb = $("#importHints"); if (hb) hb.hidden = true;   // 預設隱藏，匯入時才顯示
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
        <button class="sc-up" title="${t("sc_up")}">▲</button>
        <button class="sc-down" title="${t("sc_down")}">▼</button>
        <button class="sc-regen" title="${t("sc_regen")}">↻</button>
        <button class="sc-del" title="${t("sc_del")}">✕</button>
      </div>
    </div>
    <div class="sc-layout-row">
      <label>${t("sc_layout")}</label>
      <select class="sc-layout">${LAYOUTS.map(l => `<option value="${l.id}" ${l.id === (sl.layout || "bullets") ? "selected" : ""}>${tLang(l.name, l.nameEn)}</option>`).join("")}</select>
      <select class="sc-charttype" hidden>${CHART_TYPES.map(c => `<option value="${c.id}" ${c.id === (sl.chartType || "bar") ? "selected" : ""}>${tLang(c.name, c.nameEn)}</option>`).join("")}</select>
      <span class="sc-hint"></span>
    </div>
    <div class="bullets"></div>
    <button class="add-bullet">${t("add_point")}</button>
    <div class="sc-notes">
      <button class="notes-toggle">${t("notes_toggle")}${sl.notes ? t("notes_has") : ""}</button>
      <textarea class="notes-area" placeholder="${t("notes_ph")}" hidden>${escapeHtml(sl.notes)}</textarea>
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
    const L = LAYOUTS.find(l => l.id === lay.value) || {};
    hint.textContent = tLang(L.hint || "", L.hintEn || "");
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
    <button class="b-del" title="${t("b_del")}">✕</button>`;
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
  const th = activeTheme();
  $("#summaryBox").innerHTML = `
    <div class="row"><span class="k">${t("sum_purpose")}</span><span class="v">${p ? tLang(p.name, p.nameEn) : "—"}</span></div>
    <div class="row"><span class="k">${t("sum_topic")}</span><span class="v">${escapeHtml(state.deck.title)}</span></div>
    <div class="row"><span class="k">${t("sum_style")}</span><span class="v">${tLang(th.name, th.nameEn)}（${tLang(th.subtitle, th.subtitleEn)}）</span></div>
    <div class="row"><span class="k">${t("sum_count")}</span><span class="v">${t("sum_count_val", state.deck.slides.length)}</span></div>
    <div class="row"><span class="k">${t("sum_lang")}</span><span class="v">${state.lang} · ${state.tone}</span></div>`;
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
function logoImg(cls, allowBrand) {
  const logo = effectiveLogo(allowBrand);
  return logo && logo.data ? `<img class="pv-logo ${cls || ""}" src="${logo.data}" alt="logo" />` : "";
}
// 圖例（SVG）
function pvLegendSvg(series, col, t, x0) {
  if (series.length < 2 || !series[0].name) return "";
  let x = x0;
  return series.map((sr, i) => {
    const g = `<rect x="${x.toFixed(1)}" y="0.6" width="2.4" height="2.4" rx="0.3" fill="#${col[i]}"/><text x="${(x + 3.2).toFixed(1)}" y="2.7" font-size="2.7" fill="#${t.text}">${escapeHtml(sr.name)}</text>`;
    x += 6 + sr.name.length * 1.7;
    return g;
  }).join("");
}
// 預覽用的輕量 SVG 圖表（多系列 + 座標軸，與 PPTX 對應）
function pvChartSvg(labels, series, type, t) {
  if (type === "pie") {
    const values = series[0].values, col = chartPalette(t, labels.length);
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
    const legend = labels.map((l, i) => `<g transform="translate(58,${12 + i * 7})"><rect width="4" height="4" rx="0.6" fill="#${col[i]}"/><text x="6" y="3.6" font-size="3.3" fill="#${t.text}">${escapeHtml(l)} (${Math.round(values[i] / total * 100)}%)</text></g>`).join("");
    return `<svg viewBox="0 0 100 60" class="pv-svg">${paths}${legend}</svg>`;
  }
  const col = chartPalette(t, Math.max(series.length, 1));
  const L = 10, R = 2, T = 4, B = 8, vw = 100, vh = 62, pw = vw - L - R, ph = vh - T - B, base = T + ph;
  const max = niceMax(Math.max(...series.flatMap(s => s.values), 1)), ticks = 4;
  let grid = "";
  for (let k = 0; k <= ticks; k++) {
    const y = base - ph * k / ticks;
    grid += `<line x1="${L}" y1="${y.toFixed(1)}" x2="${L + pw}" y2="${y.toFixed(1)}" stroke="#${t.footer}" stroke-width="${k === 0 ? 0.4 : 0.2}" opacity="${k === 0 ? 0.65 : 0.35}"/>`;
    grid += `<text x="${L - 1}" y="${(y + 1).toFixed(1)}" font-size="2.6" fill="#${t.footer}" text-anchor="end">${Math.round(max * k / ticks)}</text>`;
  }
  const n = labels.length, slot = pw / n;
  const cats = labels.map((l, k) => `<text x="${(L + slot * (k + 0.5)).toFixed(1)}" y="${base + 4}" font-size="2.7" fill="#${t.footer}" text-anchor="middle">${escapeHtml(l)}</text>`).join("");
  const legend = pvLegendSvg(series, col, t, L);
  if (type === "line") {
    let body = "";
    series.forEach((sr, si) => {
      const pts = sr.values.map((v, k) => [L + slot * (k + 0.5), base - v / max * ph]);
      body += `<polyline points="${pts.map(p => p.map(z => z.toFixed(1)).join(",")).join(" ")}" fill="none" stroke="#${col[si]}" stroke-width="1"/>`;
      body += pts.map(p => `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="1" fill="#${col[si]}"/>`).join("");
    });
    return `<svg viewBox="0 0 100 62" class="pv-svg">${grid}${body}${cats}${legend}</svg>`;
  }
  const nSer = series.length, gGap = slot * 0.16, gW = slot - gGap, bw = gW / nSer;
  let bars = "";
  labels.forEach((l, ci) => {
    const gx = L + slot * ci + gGap / 2;
    series.forEach((sr, si) => {
      const v = sr.values[ci], h = v / max * ph, x = gx + si * bw, y = base - h;
      bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${(bw * 0.86).toFixed(1)}" height="${Math.max(0, h).toFixed(1)}" rx="0.4" fill="#${col[si]}"/>`;
      if (nSer <= 2) bars += `<text x="${(x + bw * 0.43).toFixed(1)}" y="${(y - 1).toFixed(1)}" font-size="2.5" fill="#${t.text}" text-anchor="middle">${v}</text>`;
    });
  });
  return `<svg viewBox="0 0 100 62" class="pv-svg">${grid}${bars}${cats}${legend}</svg>`;
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
      ${logoImg("", true)}
      <div class="pv-kicker" style="color:#${t.accent}">${escapeHtml(kickerText())}</div>
      <div class="pv-title" style="color:#${t.title}">${escapeHtml(title)}</div>
      <div class="pv-underline" style="background:#${t.accent}"></div>`;

  if (item.type === "cover") {
    stage.style.background = t.clean ? "#" + t.coverBg : gradientCss("cover", t);
    const cs = state.coverStyle || "left";
    if (cs === "center") {
      stage.innerHTML = `
        <div class="pv-cband pv-cband-top" style="background:#${t.coverAccent}"></div>
        <div class="pv-cband pv-cband-bot" style="background:#${t.coverAccent}"></div>
        ${logoImg("pv-logo-cover")}
        <div class="pv-cvC-kicker" style="color:#${t.coverText}">${escapeHtml(kickerText())}</div>
        <div class="pv-cvC-title" style="color:#${t.coverText}">${escapeHtml(state.deck.title)}</div>
        <div class="pv-cvC-rule" style="background:#${t.coverAccent}"></div>
        <div class="pv-cvC-sub" style="color:#${t.coverText}">${escapeHtml(state.deck.subtitle)}</div>`;
    } else if (cs === "full") {
      const bandText = isDark(t.coverAccent) ? "FFFFFF" : "1A1A1A";
      stage.innerHTML = `
        ${logoImg("pv-logo-cover")}
        <div class="pv-cvF-kicker" style="color:#${t.coverAccent}">${escapeHtml(kickerText())}</div>
        <div class="pv-cvF-band" style="background:#${t.coverAccent}"><span class="pv-cvF-title" style="color:#${bandText}">${escapeHtml(state.deck.title)}</span></div>
        <div class="pv-cvF-sub" style="color:#${t.coverText}">${escapeHtml(state.deck.subtitle)}</div>`;
    } else {
      stage.innerHTML = `
        <div class="pv-sidebar" style="background:#${t.coverAccent}"></div>
        ${logoImg("pv-logo-cover")}
        <div class="pv-cover-kicker" style="color:#${t.coverText}">${escapeHtml(kickerText())}</div>
        <div class="pv-cover-title" style="color:#${t.coverText}">${escapeHtml(state.deck.title)}</div>
        <div class="pv-cover-accent" style="background:#${t.coverAccent}"></div>
        <div class="pv-cover-sub" style="color:#${t.coverText}">${escapeHtml(state.deck.subtitle)}</div>`;
    }
    return;
  }
  if (item.type === "closing") {
    stage.style.background = t.clean ? "#" + t.coverBg : gradientCss("cover", t);
    stage.innerHTML = `
      <div class="pv-end-title" style="color:#${t.coverText}">謝謝聆聽</div>
      <div class="pv-end-sub" style="color:#${t.coverText}">THANK YOU</div>
      <div class="pv-cover-bar" style="background:#${t.coverAccent}"></div>`;
    return;
  }

  const sl = state.deck.slides[item.i];
  const layout = sl.layout || "bullets";

  if (layout === "section") {
    stage.style.background = t.clean ? "#" + t.coverBg : gradientCss("section", t);
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
    const cd = parseChartSeries(sl.bullets);
    if (cd) {
      stage.style.background = "#" + t.bg;
      stage.innerHTML = header(sl.title) + `<div class="pv-chart">${pvChartSvg(cd.labels, cd.series, sl.chartType || "bar", t)}</div>` + foot(item.i);
      return;
    }
  }
  if (layout === "metrics") {
    const ms = parseMetrics(sl.bullets);
    if (ms.length) {
      stage.style.background = "#" + t.bg;
      const cardBg = t.clean ? "FFFFFF" : mix(t.accent, t.bg, 0.9), cardLine = t.clean ? "E6E6E6" : mix(t.accent, t.bg, 0.72);
      const cards = ms.slice(0, 4).map(m => `
        <div class="pv-card${t.clean ? " pv-card-flat" : ""}" style="background:#${cardBg};border-color:#${cardLine}">
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
  // 條列
  stage.style.background = "#" + t.bg;
  let idx = 0;
  if (t.clean) {
    // 乾淨模式：紅色編號 + 文字 + 細分隔線
    const cols = splitCols(sl.bullets).map(col => `
      <div class="pv-col">${col.map(b => { idx++; return `
        <div class="pv-crow">
          <span class="pv-cnum" style="color:#${t.accent}">${String(idx).padStart(2, "0")}</span>
          <span class="pv-btext" style="color:#${t.text}">${escapeHtml(b)}</span>
        </div>`; }).join("")}</div>`).join("");
    stage.innerHTML = header(sl.title) + `<div class="pv-body pv-cleanlist">${cols}</div>` + foot(item.i);
    return;
  }
  const panel = mix(t.accent, t.bg, 0.92), chipText = isDark(t.accent) ? "FFFFFF" : "111111";
  const cols = splitCols(sl.bullets).map(col => `
    <div class="pv-col">${col.map(b => { idx++; return `
      <div class="pv-brow" style="background:#${panel}">
        <span class="pv-bnum" style="background:#${t.accent};color:#${chipText}">${String(idx).padStart(2, "0")}</span>
        <span class="pv-btext" style="color:#${t.text}">${escapeHtml(b)}</span>
      </div>`; }).join("")}</div>`).join("");
  stage.innerHTML = header(sl.title) + `<div class="pv-body pv-cardlist">${cols}</div>` + foot(item.i);
}

/* ---------- 生成 PPTX（多版型專業排版） ---------- */
function pptxHeader(pptx, s, t, W, title) {
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.16, fill: { color: t.accent } });
  s.addText(kickerText(), { x: 0.72, y: 0.42, w: W - 1.4, h: 0.3, fontSize: 12, bold: true, color: t.accent, charSpacing: 2, fontFace: t.bodyFont });
  s.addText(title, { x: 0.7, y: 0.74, w: W - 1.4, h: 0.85, fontFace: t.titleFont, fontSize: 28, bold: true, color: t.title, align: "left", valign: "middle" });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.68, w: 1.7, h: 0.07, fill: { color: t.accent } });
  s.addShape(pptx.ShapeType.rect, { x: 2.5, y: 1.705, w: W - 3.22, h: 0.014, fill: { color: t.accent2, transparency: 55 } });
}
function pptxFooter(pptx, s, t, W, H, i, total) {
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: H - 0.62, w: W - 1.44, h: 0.012, fill: { color: t.footer, transparency: 45 } });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: H - 0.5, w: 0.14, h: 0.14, fill: { color: t.accent } });
  s.addText(state.deck.title, { x: 0.98, y: H - 0.56, w: W - 3.4, h: 0.3, fontSize: 9.5, color: t.footer, valign: "middle", fontFace: t.bodyFont });
  s.addText(`${i + 1} / ${total}`, { x: W - 1.6, y: H - 0.56, w: 0.88, h: 0.3, fontSize: 9.5, color: t.footer, align: "right", valign: "middle", fontFace: t.bodyFont });
}
function drawLogo(s, W, maxH, y, allowBrand) {
  const logo = effectiveLogo(allowBrand);
  if (!logo || !logo.data) return;
  const { data, w, h } = logo;
  const hh = maxH, ww = hh * (w / h);
  try { s.addImage({ data, x: W - 0.72 - ww, y, w: ww, h: hh }); } catch (e) { /* 忽略無效圖檔 */ }
}
// 漸層背景（滿版圖片）；乾淨模式或產圖失敗則退回純色
function slideBgImage(s, W, H, t, kind) {
  if (t.clean) { s.background = { color: t.coverBg }; return; }
  const data = makeGradient(kind, t);
  if (data) s.addImage({ data, x: 0, y: 0, w: W, h: H });
  else s.background = { color: t.coverBg };
}
// 目前有效的 Logo：使用者上傳優先；否則若為品牌主題且允許，用官方品牌 Logo
function effectiveLogo(allowBrand) {
  if (state.logo && state.logo.data) return state.logo;
  if (allowBrand) { const at = activeTheme(); if (at.brand === "avon" && typeof AVON_LOGO !== "undefined") return AVON_LOGO; }
  return null;
}
// 圖表以向量圖形繪製（相容性最佳，不依賴內嵌 Excel 的 addChart）
function pptxChart(pptx, s, t, W, H, sl, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  const parsed = parseChartSeries(sl.bullets);
  if (!parsed) { pptxBullets(pptx, s, t, W, H, sl, i, total); return; }
  const { labels, series } = parsed;
  const ct = sl.chartType || "bar";
  if (ct === "pie") pptxPie(pptx, s, t, W, H, labels, series[0].values, chartPalette(t, labels.length));
  else {
    const colors = chartPalette(t, Math.max(series.length, 1));
    (ct === "line" ? pptxLine : pptxBar)(pptx, s, t, W, H, labels, series, colors);
  }
  pptxFooter(pptx, s, t, W, H, i, total);
}
// Y 軸刻度 + 水平格線
function pptxAxis(pptx, s, t, x0, top, baseline, plotW, max) {
  const ticks = 4;
  for (let k = 0; k <= ticks; k++) {
    const y = baseline - (baseline - top) * k / ticks;
    s.addShape(pptx.ShapeType.rect, { x: x0, y, w: plotW, h: k === 0 ? 0.014 : 0.008, fill: { color: t.footer, transparency: k === 0 ? 25 : 68 } });
    s.addText(String(Math.round(max * k / ticks)), { x: x0 - 1.0, y: y - 0.15, w: 0.85, h: 0.3, align: "right", valign: "middle", fontSize: 9, color: t.footer, fontFace: t.bodyFont });
  }
}
// 圖例（多系列）
function pptxLegend(pptx, s, t, series, colors, W) {
  if (series.length < 2 || !series[0].name) return;
  let cx = W - 0.72;
  for (let i = series.length - 1; i >= 0; i--) {
    const label = series[i].name, wText = Math.min(2.4, 0.16 * label.length + 0.45);
    cx -= wText;
    s.addText(label, { x: cx + 0.3, y: 1.9, w: wText - 0.3, h: 0.3, fontSize: 11, color: t.text, valign: "middle", fontFace: t.bodyFont });
    s.addShape(pptx.ShapeType.rect, { x: cx, y: 1.96, w: 0.2, h: 0.2, fill: { color: colors[i] } });
    cx -= 0.2;
  }
}
function pptxBar(pptx, s, t, W, H, labels, series, colors) {
  const x0 = 1.3, plotW = W - 2.3, top = 2.5, baseline = 6.1, plotH = baseline - top;
  const max = niceMax(Math.max(...series.flatMap(sr => sr.values), 1));
  pptxAxis(pptx, s, t, x0, top, baseline, plotW, max);
  pptxLegend(pptx, s, t, series, colors, W);
  const nCat = labels.length, nSer = series.length;
  const gGap = plotW * 0.04, groupW = (plotW - gGap * (nCat + 1)) / nCat;
  const bGap = nSer > 1 ? groupW * 0.06 : 0, bw = (groupW - bGap * (nSer - 1)) / nSer;
  labels.forEach((l, ci) => {
    const gx = x0 + gGap + ci * (groupW + gGap);
    series.forEach((sr, si) => {
      const v = sr.values[ci], h = Math.max(0.02, v / max * plotH), x = gx + si * (bw + bGap), y = baseline - h;
      s.addShape(pptx.ShapeType.roundRect, { x, y, w: bw, h, fill: { color: colors[si] }, rectRadius: 0.03 });
      if (nSer <= 2) s.addText(String(v), { x: x - 0.25, y: y - 0.32, w: bw + 0.5, h: 0.3, align: "center", fontSize: 10, bold: true, color: t.title, fontFace: t.bodyFont });
    });
    s.addText(l, { x: gx - 0.2, y: baseline + 0.06, w: groupW + 0.4, h: 0.4, align: "center", fontSize: 11, color: t.text, fontFace: t.bodyFont });
  });
}
function pptxLine(pptx, s, t, W, H, labels, series, colors) {
  const x0 = 1.3, plotW = W - 2.3, top = 2.6, baseline = 6.1, plotH = baseline - top;
  const max = niceMax(Math.max(...series.flatMap(sr => sr.values), 1));
  pptxAxis(pptx, s, t, x0, top, baseline, plotW, max);
  pptxLegend(pptx, s, t, series, colors, W);
  const n = labels.length, slot = plotW / n;
  const cx = k => x0 + slot * k + slot / 2, cy = v => baseline - v / max * plotH;
  series.forEach((sr, si) => {
    const col = colors[si];
    for (let k = 0; k < n - 1; k++) {
      const x1 = cx(k), y1 = cy(sr.values[k]), x2 = cx(k + 1), y2 = cy(sr.values[k + 1]);
      const len = Math.hypot(x2 - x1, y2 - y1), ang = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
      s.addShape(pptx.ShapeType.rect, { x: (x1 + x2) / 2 - len / 2, y: (y1 + y2) / 2 - 0.018, w: len, h: 0.036, fill: { color: col }, rotate: ang });
    }
    sr.values.forEach((v, k) => {
      s.addShape(pptx.ShapeType.ellipse, { x: cx(k) - 0.07, y: cy(v) - 0.07, w: 0.14, h: 0.14, fill: { color: col } });
      if (series.length <= 2) s.addText(String(v), { x: cx(k) - 0.4, y: cy(v) - 0.4, w: 0.8, h: 0.28, align: "center", fontSize: 10, bold: true, color: t.title, fontFace: t.bodyFont });
    });
  });
  labels.forEach((l, k) => s.addText(l, { x: cx(k) - slot / 2, y: baseline + 0.06, w: slot, h: 0.4, align: "center", fontSize: 11, color: t.text, fontFace: t.bodyFont }));
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
  slideBgImage(s, W, H, t, "cover");
  const style = state.coverStyle || "left";
  const titleShadow = t.clean ? undefined : { type: "outer", angle: 90, blur: 10, offset: 3, color: "000000", opacity: 0.25 };

  if (style === "center") {
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.18, fill: { color: t.coverAccent } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.18, w: W, h: 0.18, fill: { color: t.coverAccent } });
    s.addText(kickerText(), { x: 1, y: 2.05, w: W - 2, h: 0.4, fontSize: 13, bold: true, color: t.coverAccent, charSpacing: 3, align: "center", fontFace: t.bodyFont });
    s.addText(state.deck.title, { x: 1, y: 2.5, w: W - 2, h: 1.7, fontFace: t.titleFont, fontSize: 44, bold: true, color: t.coverText, align: "center", valign: "middle", shadow: titleShadow });
    s.addShape(pptx.ShapeType.rect, { x: (W - 2.2) / 2, y: 4.35, w: 2.2, h: 0.06, fill: { color: t.coverAccent } });
    if (state.deck.subtitle) s.addText(state.deck.subtitle, { x: 1.5, y: 4.6, w: W - 3, h: 0.9, fontFace: t.bodyFont, fontSize: 18, color: t.coverText, align: "center", transparency: 10 });
    drawLogo(s, W, 0.7, 0.5, false);
    return;
  }
  if (style === "full") {
    // 滿版橫幅：標題落在整條強調色帶上
    const bandY = 2.55, bandH = 2.1;
    s.addText(kickerText(), { x: 1, y: bandY - 0.55, w: W - 2, h: 0.4, fontSize: 13, bold: true, color: t.coverAccent, charSpacing: 3, fontFace: t.bodyFont });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: bandY, w: W, h: bandH, fill: { color: t.coverAccent } });
    const bandText = isDark(t.coverAccent) ? "FFFFFF" : "1A1A1A";
    s.addText(state.deck.title, { x: 0.9, y: bandY, w: W - 1.8, h: bandH, fontFace: t.titleFont, fontSize: 46, bold: true, color: bandText, align: "left", valign: "middle" });
    if (state.deck.subtitle) s.addText(state.deck.subtitle, { x: 0.95, y: bandY + bandH + 0.25, w: W - 3, h: 0.9, fontFace: t.bodyFont, fontSize: 18, color: t.coverText, transparency: 8 });
    s.addText("AI 簡報生成器", { x: 0.95, y: H - 0.7, w: 6, h: 0.35, fontSize: 10.5, color: t.coverText, transparency: 40, fontFace: t.bodyFont });
    drawLogo(s, W, 0.7, 0.5, false);
    return;
  }
  // left（經典）
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addShape(pptx.ShapeType.rect, { x: W - 3.4, y: H - 3.4, w: 2.6, h: 2.6, fill: { color: t.coverAccent, transparency: 82 } });
  s.addShape(pptx.ShapeType.rect, { x: W - 2.3, y: H - 2.3, w: 2.6, h: 2.6, fill: { color: t.coverAccent, transparency: 90 } });
  s.addText(kickerText(), { x: 0.95, y: 1.75, w: W - 2, h: 0.4, fontSize: 13, bold: true, color: t.coverAccent, charSpacing: 3, fontFace: t.bodyFont });
  s.addText(state.deck.title, { x: 0.9, y: 2.2, w: W - 3, h: 1.9, fontFace: t.titleFont, fontSize: 44, bold: true, color: t.coverText, align: "left", valign: "top", shadow: titleShadow });
  s.addShape(pptx.ShapeType.rect, { x: 0.95, y: 4.2, w: 2.2, h: 0.06, fill: { color: t.coverAccent } });
  if (state.deck.subtitle) s.addText(state.deck.subtitle, { x: 0.95, y: 4.45, w: W - 3, h: 1.0, fontFace: t.bodyFont, fontSize: 18, color: t.coverText, transparency: 10 });
  s.addText("AI 簡報生成器", { x: 0.95, y: H - 0.7, w: 6, h: 0.35, fontSize: 10.5, color: t.coverText, transparency: 40, fontFace: t.bodyFont });
  drawLogo(s, W, 0.7, 0.5, false);
}
function pptxSection(pptx, s, t, W, H, sl, no) {
  slideBgImage(s, W, H, t, "section");
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addText(String(no).padStart(2, "0"), { x: 0.55, y: 0.1, w: 6, h: 3.4, fontSize: 170, bold: true, color: t.coverAccent, transparency: 80, fontFace: t.titleFont, align: "left", valign: "top" });
  s.addShape(pptx.ShapeType.rect, { x: 0.95, y: 3.95, w: 1.8, h: 0.08, fill: { color: t.coverAccent } });
  s.addText(sl.title, { x: 0.9, y: 4.15, w: W - 1.8, h: 1.3, fontSize: 40, bold: true, color: t.coverText, fontFace: t.titleFont });
  if (sl.bullets[0]) s.addText(sl.bullets[0], { x: 0.95, y: 5.45, w: W - 2, h: 0.9, fontSize: 16, color: t.coverText, transparency: 15, fontFace: t.bodyFont });
}
function pptxBulletsClean(pptx, s, t, W, H, bullets) {
  // 乾淨模式：紅色編號 + 文字 + 細分隔線，充足留白（Avon 品牌規範）
  const cols = splitCols(bullets), top = 2.0, availH = H - 2.85, colGap = 0.55;
  const colW = (W - 1.44 - (cols.length - 1) * colGap) / cols.length;
  const maxN = Math.max(...cols.map(c => c.length), 1);
  const fs = maxN <= 4 ? 18 : maxN <= 6 ? 16 : 14;
  let idx = 0;
  cols.forEach((col, ci) => {
    const x = 0.72 + ci * (colW + colGap), n = col.length, rowH = availH / n;
    col.forEach((b, k) => {
      idx++; const y = top + k * rowH;
      s.addText(String(idx).padStart(2, "0"), { x, y, w: 0.62, h: rowH, valign: "middle", fontSize: fs + 3, bold: true, color: t.accent, fontFace: t.titleFont });
      s.addText(b, { x: x + 0.72, y, w: colW - 0.72, h: rowH, valign: "middle", fontSize: fs, color: t.text, fontFace: t.bodyFont, lineSpacingMultiple: 1.08 });
      if (k < n - 1) s.addShape(pptx.ShapeType.rect, { x: x + 0.04, y: y + rowH - 0.008, w: colW - 0.08, h: 0.008, fill: { color: "E8E8E8" } });
    });
  });
}
function pptxBullets(pptx, s, t, W, H, sl, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  const bullets = sl.bullets.length ? sl.bullets : [""];
  if (t.clean) { pptxBulletsClean(pptx, s, t, W, H, bullets); pptxFooter(pptx, s, t, W, H, i, total); return; }
  const cols = splitCols(bullets);
  const top = 1.98, availH = H - 2.78, colGap = 0.35;
  const colW = (W - 1.44 - (cols.length - 1) * colGap) / cols.length;
  const panel = mix(t.accent, t.bg, 0.9), chipText = isDark(t.accent) ? "FFFFFF" : "111111";
  const maxN = Math.max(...cols.map(c => c.length), 1);
  const fs = maxN <= 4 ? 17 : maxN <= 6 ? 15 : 13;
  let idx = 0;
  cols.forEach((col, ci) => {
    const x = 0.72 + ci * (colW + colGap);
    const n = col.length, gap = 0.16, rowH = Math.min(1.15, (availH - gap * (n - 1)) / n);
    const cs = Math.min(0.52, rowH * 0.62);
    col.forEach((b, k) => {
      idx++;
      const yy = top + k * (rowH + gap);
      s.addShape(pptx.ShapeType.roundRect, { x, y: yy, w: colW, h: rowH, fill: { color: panel }, line: { width: 0 }, rectRadius: 0.07, shadow: SHADOW });
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.18, y: yy + (rowH - cs) / 2, w: cs, h: cs, fill: { color: t.accent }, line: { width: 0 }, rectRadius: 0.1 });
      s.addText(String(idx).padStart(2, "0"), { x: x + 0.18, y: yy + (rowH - cs) / 2, w: cs, h: cs, align: "center", valign: "middle", fontSize: 12, bold: true, color: chipText, fontFace: t.titleFont });
      s.addText(b, { x: x + 0.18 + cs + 0.22, y: yy, w: colW - (0.18 + cs + 0.22) - 0.2, h: rowH, valign: "middle", fontSize: fs, color: t.text, fontFace: t.bodyFont, lineSpacingMultiple: 1.05 });
    });
  });
  pptxFooter(pptx, s, t, W, H, i, total);
}
function pptxMetrics(pptx, s, t, W, H, sl, metrics, i, total) {
  pptxHeader(pptx, s, t, W, sl.title);
  const n = Math.min(metrics.length, 4);
  const gap = 0.35, cardW = (W - 1.44 - gap * (n - 1)) / n, y = 2.55, cardH = 2.7;
  const cardBg = t.clean ? "FFFFFF" : mix(t.accent, t.bg, 0.9);
  const cardLine = t.clean ? "E6E6E6" : mix(t.accent, t.bg, 0.7);
  metrics.slice(0, n).forEach((m, k) => {
    const x = 0.72 + k * (cardW + gap);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cardW, h: cardH, fill: { color: cardBg }, line: { color: cardLine, width: 1 }, rectRadius: t.clean ? 0.05 : 0.1, shadow: t.clean ? undefined : SHADOW });
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
  slideBgImage(s, W, H, t, "cover");
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.32, h: H, fill: { color: t.coverAccent } });
  s.addText("謝謝聆聽", { x: 0.9, y: 2.6, w: W - 1.8, h: 1.3, fontSize: 48, bold: true, color: t.coverText, fontFace: t.titleFont, shadow: t.clean ? undefined : { type: "outer", angle: 90, blur: 10, offset: 3, color: "000000", opacity: 0.25 } });
  s.addText("THANK YOU", { x: 0.95, y: 3.95, w: 6, h: 0.5, fontSize: 15, color: t.coverText, transparency: 30, charSpacing: 4, fontFace: t.bodyFont });
  if (state.deck.title) s.addText(state.deck.title, { x: 0.95, y: H - 1.0, w: W - 2, h: 0.5, fontSize: 13, color: t.coverText, transparency: 30, fontFace: t.bodyFont });
  drawLogo(s, W, 0.6, 0.5, false);
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
      drawLogo(s, W, isCoverLike ? 0.6 : 0.4, isCoverLike ? 0.5 : 0.34, !isCoverLike);
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
  $("#providerHint").textContent = providerHint(s.provider);
  $("#settingsModal").hidden = false;
}

/* ---------- 企業品牌客製（主色 + Logo） ---------- */
function applyBrandUI() {
  if (state.customAccent) {
    const h = "#" + hx(state.customAccent);
    $("#accentPicker").value = h; $("#accentHex").value = h;
    $("#accentNote").textContent = tLang("已套用企業主色 " + h, "Brand color applied " + h);
  } else {
    $("#accentNote").textContent = tLang("目前使用風格預設色", "Using the style's default color");
  }
  const hasLogo = state.logo && state.logo.data;
  $("#logoPreview").hidden = !hasLogo;
  $("#clearLogo").hidden = !hasLogo;
  if (hasLogo) $("#logoPreview").src = state.logo.data;
  $$("#coverStyleSeg .seg-btn").forEach(b => b.classList.toggle("active", b.dataset.cover === (state.coverStyle || "left")));
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
  if (!confirm(t("reset_confirm"))) return;
  clearState();
  Object.assign(state, { purpose: null, topic: "", direction: "", audience: "", slideCount: 10, lang: "繁體中文", tone: "清楚易懂", deck: null, theme: THEMES[0].id, customAccent: null, logo: null, coverStyle: "left", step: 1, previewIdx: 0 });
  ["topicInput", "directionInput", "audienceInput"].forEach(id => $("#" + id).value = "");
  $("#slideCount").value = "10"; $("#langSelect").value = "繁體中文"; $("#toneSelect").value = "清楚易懂";
  $("#slidesEditor").innerHTML = "";
  $("#toStep2").disabled = true;
  renderPurposes(); renderThemes(); applyBrandUI();
  goStep(1);
  toast("已重設");
}

/* ---------- 事件綁定 ---------- */
/* ---------- 介面語言切換 ---------- */
function applyUiLang() {
  document.documentElement.lang = UI_LANG === "en" ? "en" : "zh-Hant";
  document.title = t("doc_title");
  $$("[data-i18n]").forEach(el => { el.textContent = t(el.getAttribute("data-i18n")); });
  $$("[data-i18n-html]").forEach(el => { el.innerHTML = t(el.getAttribute("data-i18n-html")); });
  $$("[data-i18n-ph]").forEach(el => { el.setAttribute("placeholder", t(el.getAttribute("data-i18n-ph"))); });
  const b = $("#uiLangBtn"); if (b) b.textContent = UI_LANG === "en" ? "中文" : "EN";
  renderPurposes(); renderThemes(); applyBrandUI();
  if (state.deck && $$("#slidesEditor .slide-card").length) { harvestOutline(); renderOutline(); requestAnimationFrame(growAll); }
  if ($("#providerHint")) $("#providerHint").textContent = providerHint(getSettings().provider);
  if (state.deck && state.step === 5) { buildSummary(); renderPreview(); }
}

function bindEvents() {
  renderPurposes();
  renderThemes();

  $("#uiLangBtn").addEventListener("click", () => {
    UI_LANG = UI_LANG === "en" ? "zh" : "en";
    try { localStorage.setItem("ppt_uilang", UI_LANG); } catch (e) {}
    applyUiLang();
  });

  $("#toStep2").addEventListener("click", () => goStep(2));
  $("#generateOutlineBtn").addEventListener("click", generateOutline);
  $("#toStep4").addEventListener("click", () => { harvestOutline(); goStep(4); });
  $("#toStep5").addEventListener("click", () => goStep(5));
  $("#downloadPptxBtn").addEventListener("click", generatePptx);
  $("#addSlideBtn").addEventListener("click", () => {
    $("#slidesEditor").appendChild(slideCardEl({ title: t("new_slide"), bullets: [t("new_point")], notes: "", layout: "bullets", chartType: "bar" }, $$("#slidesEditor .slide-card").length));
    renumber(); requestAnimationFrame(growAll); collectAll();
  });
  $("#regenBtn").addEventListener("click", generateOutline);
  $$("[data-goto]").forEach(b => b.addEventListener("click", () => goStep(+b.dataset.goto)));

  // 步驟 1 直接進入匯入
  $("#haveContentBtn").addEventListener("click", () => {
    if (!state.purpose) state.purpose = "general";
    goStep(2);
    setTimeout(() => {
      const box = $(".import-box");
      if (box) { box.scrollIntoView({ behavior: "smooth", block: "center" }); box.classList.add("flash"); setTimeout(() => box.classList.remove("flash"), 1200); }
      $("#importText").focus();
    }, 80);
  });

  // 匯入文案
  $("#importFileBtn").addEventListener("click", () => $("#importDocFile").click());
  $("#importDocFile").addEventListener("change", e => { if (e.target.files[0]) readImportFile(e.target.files[0]); e.target.value = ""; });
  $("#importDirectBtn").addEventListener("click", () => importContent("direct"));
  $("#importAiBtn").addEventListener("click", () => importContent("ai"));

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
  $$("#coverStyleSeg .seg-btn").forEach(b => b.addEventListener("click", () => {
    state.coverStyle = b.dataset.cover;
    $$("#coverStyleSeg .seg-btn").forEach(x => x.classList.toggle("active", x === b));
    saveState();
  }));

  // 重設
  $("#resetAllBtn").addEventListener("click", resetAll);

  // 設定
  $("#settingsBtn").addEventListener("click", openSettings);
  $("#closeSettings").addEventListener("click", () => $("#settingsModal").hidden = true);
  $("#providerSelect").addEventListener("change", e => {
    fillModelDatalist(e.target.value);
    $("#modelInput").value = DEFAULT_MODELS[e.target.value];
    $("#providerHint").textContent = providerHint(e.target.value);
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
  // 套用介面語言（放最後，確保所有動態內容都本地化）
  applyUiLang();
}

document.addEventListener("DOMContentLoaded", bindEvents);

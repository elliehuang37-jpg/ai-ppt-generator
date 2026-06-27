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
const PROVIDER_HINTS = {
  anthropic: "在 console.anthropic.com 取得金鑰。瀏覽器直連已啟用 anthropic-dangerous-direct-browser-access。",
  openai: "在 platform.openai.com 取得金鑰。常用模型：gpt-4o、gpt-4o-mini。",
  gemini: "在 aistudio.google.com 取得金鑰。常用模型：gemini-1.5-flash、gemini-1.5-pro。"
};

const state = {
  purpose: null,
  topic: "", direction: "", audience: "", slideCount: 10, lang: "繁體中文", tone: "清楚易懂",
  deck: null,          // { title, subtitle, slides:[{title,bullets:[],notes}] }
  theme: THEMES[0].id,
  step: 1
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
function getSettings() {
  return {
    provider: localStorage.getItem("ppt_provider") || APP_CONFIG.defaultProvider,
    model: localStorage.getItem("ppt_model") || "",
    key: localStorage.getItem("ppt_key") || ""
  };
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
    });
  });
}

/* ---------- 步驟導航 ---------- */
function goStep(n) {
  if (n === 5) buildSummary();
  state.step = n;
  $$(".panel").forEach(p => p.classList.remove("active"));
  $(`#panel-${n}`).classList.add("active");
  if (n === 3) requestAnimationFrame(growAll);   // 面板顯示後再計算高度
  $$(".step").forEach(s => {
    const sn = +s.dataset.step;
    s.classList.toggle("active", sn === n);
    s.classList.toggle("done", sn < n);
  });
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
4. 只回傳 JSON，不要加任何解說文字或 markdown 標記。

JSON 格式（嚴格遵守）：
{
  "title": "簡報主標題",
  "subtitle": "副標題或一句話定位（可含講者/日期占位）",
  "slides": [
    { "title": "投影片標題", "bullets": ["重點1", "重點2", "重點3"], "notes": "講稿提示" }
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
      body: JSON.stringify({
        model: mdl, max_tokens: 4000,
        messages: [{ role: "user", content: prompt }]
      })
    });
    if (!r.ok) throw new Error(`Anthropic ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.content.map(c => c.text || "").join("");
  }

  if (provider === "openai") {
    const r = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "content-type": "application/json", "Authorization": `Bearer ${key}` },
      body: JSON.stringify({
        model: mdl, temperature: 0.7,
        messages: [{ role: "user", content: prompt }]
      })
    });
    if (!r.ok) throw new Error(`OpenAI ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.choices[0].message.content;
  }

  if (provider === "gemini") {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${mdl}:generateContent?key=${key}`;
    const r = await fetch(url, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
    });
    if (!r.ok) throw new Error(`Gemini ${r.status}: ${await r.text()}`);
    const d = await r.json();
    return d.candidates[0].content.parts.map(p => p.text || "").join("");
  }
  throw new Error("未知的供應商");
}

/* ---------- 解析 JSON ---------- */
function parseDeck(raw) {
  let txt = (raw || "").trim();
  txt = txt.replace(/```json/gi, "").replace(/```/g, "").trim();
  const s = txt.indexOf("{"), e = txt.lastIndexOf("}");
  if (s === -1 || e === -1) throw new Error("找不到 JSON 內容");
  const obj = JSON.parse(txt.slice(s, e + 1));
  if (!obj.slides || !Array.isArray(obj.slides)) throw new Error("JSON 缺少 slides 陣列");
  obj.slides = obj.slides.map(sl => ({
    title: sl.title || "未命名投影片",
    bullets: Array.isArray(sl.bullets) ? sl.bullets : (sl.content ? [sl.content] : []),
    notes: sl.notes || ""
  }));
  return { title: obj.title || state.topic, subtitle: obj.subtitle || "", slides: obj.slides };
}

/* ---------- 產生大綱（主流程） ---------- */
async function generateOutline() {
  collectInputs();
  if (!state.topic) { toast("請先輸入簡報主題", true); return; }
  const prompt = buildPrompt();
  const { key } = getSettings();

  if (!key) { openManual(prompt); return; }   // 無金鑰 → 手動模式

  loading(true, "AI 正在撰寫大綱與內容…");
  try {
    const raw = await callLLM(prompt);
    state.deck = parseDeck(raw);
    renderOutline();
    loading(false);
    goStep(3);
  } catch (err) {
    loading(false);
    if (err.message === "NO_KEY") { openManual(prompt); return; }
    console.error(err);
    toast("產生失敗：" + err.message.slice(0, 120), true);
  }
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
      <button class="sc-del" title="刪除這張">✕</button>
    </div>
    <div class="bullets"></div>
    <button class="add-bullet">＋ 新增重點</button>`;
  const bl = card.querySelector(".bullets");
  sl.bullets.forEach(b => bl.appendChild(bulletRow(b)));
  card.querySelector(".add-bullet").addEventListener("click", () => bl.appendChild(bulletRow("")));
  card.querySelector(".sc-del").addEventListener("click", () => { card.remove(); renumber(); });
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
  row.querySelector(".b-del").addEventListener("click", () => row.remove());
  return row;
}
function autoGrow(ta){ ta.style.height="auto"; if(ta.scrollHeight) ta.style.height=ta.scrollHeight+"px"; }
function growAll(){ $$("#slidesEditor textarea").forEach(autoGrow); }
function renumber(){ $$("#slidesEditor .slide-card .num").forEach((n,i)=>n.textContent=i+1); }
function escapeHtml(s){ return (s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

/* 從 DOM 讀回大綱 */
function harvestOutline() {
  const slides = [];
  $$("#slidesEditor .slide-card").forEach((card, i) => {
    const title = card.querySelector(".sc-title").value.trim();
    const bullets = [...card.querySelectorAll(".bullet-row textarea")]
      .map(t => t.value.trim()).filter(Boolean);
    const notes = state.deck.slides[i] ? state.deck.slides[i].notes : "";
    slides.push({ title, bullets, notes });
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
  const t = THEMES.find(x => x.id === state.theme);
  $("#summaryBox").innerHTML = `
    <div class="row"><span class="k">簡報目的</span><span class="v">${p ? p.name : "—"}</span></div>
    <div class="row"><span class="k">主題</span><span class="v">${escapeHtml(state.deck.title)}</span></div>
    <div class="row"><span class="k">設計風格</span><span class="v">${t.name}（${t.subtitle}）</span></div>
    <div class="row"><span class="k">投影片數</span><span class="v">封面 + ${state.deck.slides.length} 張內容</span></div>
    <div class="row"><span class="k">語言 / 語氣</span><span class="v">${state.lang} · ${state.tone}</span></div>`;
}

/* ---------- 生成 PPTX ---------- */
function generatePptx() {
  harvestOutline();
  if (!state.deck || !state.deck.slides.length) { toast("沒有可生成的內容", true); return; }
  const t = THEMES.find(x => x.id === state.theme);
  loading(true, "正在生成 PPTX 檔案…");

  try {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "W16x9", width: 13.333, height: 7.5 });
    pptx.layout = "W16x9";
    pptx.author = "AI 簡報生成器";
    pptx.title = state.deck.title;

    const W = 13.333, H = 7.5;

    /* === 封面 === */
    const cover = pptx.addSlide();
    cover.background = { color: t.coverBg };
    cover.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.9, w: W, h: 0.9, fill: { color: t.coverAccent } });
    cover.addShape(pptx.ShapeType.rect, { x: 0.9, y: 2.9, w: 2.0, h: 0.16, fill: { color: t.coverAccent } });
    cover.addText(state.deck.title, {
      x: 0.9, y: 2.0, w: W - 1.8, h: 1.6,
      fontFace: t.titleFont, fontSize: 40, bold: true, color: t.coverText, align: "left", valign: "bottom"
    });
    if (state.deck.subtitle) {
      cover.addText(state.deck.subtitle, {
        x: 0.9, y: 3.4, w: W - 1.8, h: 1.0,
        fontFace: t.bodyFont, fontSize: 18, color: t.coverText, align: "left", transparency: 12
      });
    }

    /* === 內容頁 === */
    state.deck.slides.forEach((sl, i) => {
      const s = pptx.addSlide();
      s.background = { color: t.bg };
      // 頂部細條
      s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.16, fill: { color: t.accent } });
      // 標題
      s.addText(sl.title, {
        x: 0.7, y: 0.55, w: W - 1.4, h: 0.9,
        fontFace: t.titleFont, fontSize: 27, bold: true, color: t.title, align: "left", valign: "middle"
      });
      // 標題底線
      s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.5, w: 1.6, h: 0.07, fill: { color: t.accent2 } });
      // 重點
      if (sl.bullets.length) {
        s.addText(sl.bullets.map(b => ({
          text: b,
          options: { bullet: { code: "2022", indent: 18 }, color: t.text, fontSize: bulletSize(sl.bullets.length), paraSpaceAfter: 10 }
        })), {
          x: 0.85, y: 1.85, w: W - 1.7, h: H - 2.7,
          fontFace: t.bodyFont, valign: "top", lineSpacingMultiple: 1.15, color: t.text
        });
      }
      // 頁碼 + 標題
      s.addText(`${i + 1} / ${state.deck.slides.length}`, {
        x: W - 1.6, y: H - 0.55, w: 1.2, h: 0.35,
        fontFace: t.bodyFont, fontSize: 10, color: t.footer, align: "right"
      });
      s.addText(state.deck.title, {
        x: 0.7, y: H - 0.55, w: W - 2.4, h: 0.35,
        fontFace: t.bodyFont, fontSize: 10, color: t.footer, align: "left"
      });
      if (sl.notes) s.addNotes(sl.notes);
    });

    const fname = (state.deck.title || "簡報").replace(/[\\/:*?"<>|]/g, "").slice(0, 40) || "presentation";
    pptx.writeFile({ fileName: `${fname}.pptx` }).then(() => {
      loading(false);
      toast("✅ PPTX 已開始下載！");
    }).catch(e => { loading(false); toast("下載失敗：" + e.message, true); });
  } catch (err) {
    loading(false);
    console.error(err);
    toast("生成失敗：" + err.message, true);
  }
}
function bulletSize(n){ return n <= 3 ? 20 : n <= 5 ? 18 : 15; }

/* ---------- 手動模式 ---------- */
function openManual(prompt) {
  $("#manualPrompt").value = prompt;
  $("#manualJson").value = "";
  $("#manualModal").hidden = false;
}

/* ---------- 設定 Modal ---------- */
function openSettings() {
  const s = getSettings();
  $("#providerSelect").value = s.provider;
  $("#modelInput").value = s.model || DEFAULT_MODELS[s.provider];
  $("#apiKeyInput").value = s.key;
  $("#providerHint").textContent = PROVIDER_HINTS[s.provider];
  $("#settingsModal").hidden = false;
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
    $("#slidesEditor").appendChild(slideCardEl({ title: "新投影片", bullets: ["重點"], notes: "" }, $$("#slidesEditor .slide-card").length));
    renumber();
  });
  $("#regenBtn").addEventListener("click", generateOutline);
  $$("[data-goto]").forEach(b => b.addEventListener("click", () => goStep(+b.dataset.goto)));

  // 設定
  $("#settingsBtn").addEventListener("click", openSettings);
  $("#closeSettings").addEventListener("click", () => $("#settingsModal").hidden = true);
  $("#providerSelect").addEventListener("change", e => {
    $("#modelInput").value = DEFAULT_MODELS[e.target.value];
    $("#providerHint").textContent = PROVIDER_HINTS[e.target.value];
  });
  $("#saveSettings").addEventListener("click", () => {
    localStorage.setItem("ppt_provider", $("#providerSelect").value);
    localStorage.setItem("ppt_model", $("#modelInput").value.trim());
    localStorage.setItem("ppt_key", $("#apiKeyInput").value.trim());
    $("#settingsModal").hidden = true;
    toast("設定已儲存");
  });

  // 手動模式
  $("#closeManual").addEventListener("click", () => $("#manualModal").hidden = true);
  $("#copyPromptBtn").addEventListener("click", () => {
    navigator.clipboard.writeText($("#manualPrompt").value).then(() => toast("提示詞已複製"));
  });
  $("#applyManual").addEventListener("click", () => {
    try {
      state.deck = parseDeck($("#manualJson").value);
      $("#manualModal").hidden = true;
      renderOutline();
      goStep(3);
      toast("大綱已套用");
    } catch (e) { toast("JSON 解析失敗：" + e.message, true); }
  });

  // 點背景關閉 modal
  $$(".modal-backdrop").forEach(m => m.addEventListener("click", e => { if (e.target === m) m.hidden = true; }));
}

document.addEventListener("DOMContentLoaded", bindEvents);

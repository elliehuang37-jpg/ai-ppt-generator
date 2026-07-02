# AI 簡報生成器 · 商務 & 學校報告投影片自動生成

輸入主題 → AI 產出**大綱與內容** → 確認/編輯 → 選擇**設計風格** → 一鍵生成可下載的 **PPTX**。

純前端（HTML/CSS/JS），無需後端伺服器，可直接部署在 **GitHub Pages**。

---

## ✨ 功能

- **8 種簡報目的**：商業策略、商業提案 (Business Case)、銷售推廣、教育訓練、學校報告、研究發表、工作匯報，以及「一般主題」自動安排架構。
- **AI 自動產生大綱**：先給你結構化的草稿，可逐張**編輯標題、增刪重點、上下排序、編輯講稿備註、單張 AI 重寫**，確認後再進設計。
- **專業多版型設計系統**：AI 依內容自動編排，並可每張自選版型——**封面、章節分隔頁（超大編號）、條列重點、KPI 數據卡、圖表頁、金句引言、結尾致謝**；長重點自動雙欄。
- **真・數據圖表**：KPI/圖表頁支援**長條圖、圓餅圖、折線圖**（以向量圖形繪製，相容性最佳）；重點用「標籤 | 數值」即可。
- **8 種設計風格**：Avon 雅芳品牌、企業專業藍、優雅玫瑰金、簡約黑白、活力漸層、學術沉穩、科技深色、暖陽橙；每種有即時預覽。
- **企業品牌客製**：可套用**專屬企業主色**（自動推導深淺色階）與上傳**品牌 Logo**（自動放到封面與各頁頁首）。
- **即時投影片預覽**：生成前逐張預覽，所見即所得。
- **自動存檔 / 還原**：主題、大綱、風格自動存在瀏覽器，關掉再開可續接。
- **匯出 / 匯入**：大綱可匯出 / 匯入 JSON、複製 Markdown。
- **一鍵生成 PPTX**：使用 [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) 在瀏覽器本機生成 16:9 投影片（含封面、頁碼、講稿備註 notes），直接下載。
- **隱私**：API 金鑰只存在你瀏覽器的 `localStorage`，檔案也在本機生成，**不會上傳任何伺服器**。

## 🔑 使用前：設定 AI 金鑰

點右上角「⚙ API 設定」，選擇供應商並貼上你自己的金鑰：

| 供應商 | 取得金鑰 | 常用模型 |
| --- | --- | --- |
| Anthropic（Claude） | console.anthropic.com | `claude-opus-4-8` |
| OpenAI（GPT） | platform.openai.com | `gpt-4o`、`gpt-4o-mini` |
| Google（Gemini） | aistudio.google.com | `gemini-1.5-flash`、`gemini-1.5-pro` |

> **沒有金鑰也能用**：不填金鑰時按「產生大綱」會給你一段提示詞，貼到任何 AI 聊天工具，再把回傳的 JSON 貼回即可繼續設計與下載。

## 🚀 部署到 GitHub Pages

1. 建立一個 GitHub repo，把本資料夾所有檔案推上去：
   ```bash
   git init
   git add .
   git commit -m "AI PPT generator"
   git branch -M main
   git remote add origin https://github.com/<你的帳號>/<repo>.git
   git push -u origin main
   ```
2. 到 repo 的 **Settings → Pages**，Source 選 `main` 分支、`/ (root)` 資料夾，儲存。
3. 等一會兒，網站會出現在 `https://<你的帳號>.github.io/<repo>/`。

## 🖥️ 本機測試

直接用瀏覽器打開 `index.html` 即可；或啟動簡易伺服器：

```bash
python -m http.server 8000   # 然後開 http://localhost:8000
```

## 📁 檔案結構

```
ppt-generator/
├─ index.html        介面與步驟流程
├─ css/style.css     樣式
├─ js/themes.js      簡報目的 + 設計風格定義（可自行擴充）
├─ js/app.js         主程式：呼叫 AI、編輯大綱、生成 PPTX
└─ README.md
```

## 🎨 自訂

- **設定預設 AI 供應商 / 模型（部署前）**：打開 `js/app.js` 最上方的 `APP_CONFIG` 與 `DEFAULT_MODELS`，把 `defaultProvider` 改成你們要的供應商、`DEFAULT_MODELS` 改成對應模型。這樣同事第一次打開時就已選好供應商與模型，**只需貼上自己的金鑰**即可使用。
  > ⚠️ 安全提醒：請勿把 API 金鑰寫進程式碼後推到「公開」repo，否則金鑰會外洩被盜用。金鑰一律由每位使用者在「⚙ API 設定」自行輸入（只存在各自瀏覽器）。
- **新增設計風格**：在 `js/themes.js` 的 `THEMES` 陣列加一筆（複製一筆改顏色即可）。第一個 = 預設風格。
- **新增簡報目的 / 架構**：在 `PURPOSES` 陣列加一筆，`structure` 會餵給 AI 當架構參考。

---

純前端工具，PPTX 由 PptxGenJS 生成。

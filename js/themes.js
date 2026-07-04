/* =========================================================
   簡報「目的」定義
   每個目的提供：圖示、說明、以及給 AI 的架構提示
   ========================================================= */
const PURPOSES = [
  {
    id: "strategy",
    icon: "🎯",
    name: "商業策略簡報",
    desc: "向管理層說明策略方向、市場與成長計畫",
    structure: "標題頁、執行摘要、市場與背景、機會與挑戰、策略主軸、行動方案、財務預估、時程路線圖、風險與對策、結論與下一步"
  },
  {
    id: "businesscase",
    icon: "📊",
    name: "商業提案 / Business Case",
    desc: "提出方案、論證可行性並爭取資源",
    structure: "標題頁、背景與問題、機會評估、可選方案比較、建議方案、投資與效益（ROI）、實施計畫、風險評估、結論與需求（the ask）"
  },
  {
    id: "sales",
    icon: "🚀",
    name: "銷售 / 產品推廣",
    desc: "向客戶介紹產品價值、打動人心並促成行動",
    structure: "吸睛開場、客戶痛點、解決方案、產品亮點與差異化、成功案例 / 見證、方案與價格、行動呼籲（CTA）"
  },
  {
    id: "training",
    icon: "🎓",
    name: "教育訓練 / 課程",
    desc: "教學、培訓或工作坊使用的教材簡報",
    structure: "封面、課程目標、議程大綱、核心概念、步驟與方法、實例示範、練習 / 互動、重點回顧、延伸資源"
  },
  {
    id: "schoolreport",
    icon: "📚",
    name: "學校報告 / 專題",
    desc: "課堂報告、專題成果或讀書心得",
    structure: "封面（題目/組別/姓名）、研究動機與目的、背景介紹、研究方法、內容分析、發現與成果、結論、心得與反思、參考資料"
  },
  {
    id: "research",
    icon: "🔬",
    name: "研究成果發表",
    desc: "學術或研究導向的正式發表",
    structure: "標題頁、研究背景、研究問題與假設、文獻回顧、研究方法、數據與結果、討論、結論與貢獻、未來方向、參考文獻"
  },
  {
    id: "workreport",
    icon: "📈",
    name: "工作匯報 / 月報",
    desc: "週報、月報、季度回顧與成果彙整",
    structure: "封面、本期重點摘要、目標達成狀況、關鍵數據與成果、進行中項目、遇到的問題與對策、下期計畫、需要的支援"
  },
  {
    id: "general",
    icon: "💡",
    name: "一般主題簡報",
    desc: "不確定？讓 AI 依主題自動安排最合適的架構",
    structure: "由 AI 依主題自行設計最合適的開場、主體段落與結論結構"
  }
];

/* =========================================================
   設計「風格」定義
   colors 用於 PPTX 與預覽；font 為 PptxGenJS 字型名稱
   ========================================================= */
const THEMES = [
  {
    // Avon 官方品牌規範（avon-brand-guidelines skill）：
    // 官方色 Avon Red #E5004B（取自官方 logo）、字體 Lato、乾淨結構、避免漸層與過度陰影、官方 logo。
    id: "avon",
    name: "Avon 官方品牌",
    subtitle: "Avon Red · Lato · 官方規範乾淨版",
    clean: true, brand: "avon",   // 乾淨模式：無漸層、無陰影、官方 logo
    bg: "FFFFFF", titleBg: "E5004B", title: "1A1A1A", text: "2B2B2B",
    accent: "E5004B", accent2: "E5004B", bullet: "E5004B", footer: "8A8A8A",
    titleFont: "Lato", bodyFont: "Lato",
    coverBg: "E5004B", coverText: "FFFFFF", coverAccent: "FFFFFF"
  },
  {
    id: "corporate",
    name: "企業專業藍",
    subtitle: "穩重 · 可信賴 · 商務首選",
    bg: "FFFFFF", titleBg: "0B2A5B", title: "0B2A5B", text: "1F2937",
    accent: "2563EB", accent2: "60A5FA", bullet: "2563EB", footer: "94A3B8",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "0B2A5B", coverText: "FFFFFF", coverAccent: "60A5FA"
  },
  {
    id: "rosegold",
    name: "優雅玫瑰金",
    subtitle: "柔和 · 精緻 · 美妝 / 品牌",
    bg: "FFF9F7", titleBg: "8C5B5B", title: "7A3B3B", text: "4A3636",
    accent: "C08457", accent2: "E0A899", bullet: "C0857A", footer: "C9B3AE",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "7A3B3B", coverText: "FFF4EE", coverAccent: "E6B98F"
  },
  {
    id: "minimal",
    name: "簡約黑白",
    subtitle: "極簡 · 大氣 · 內容為王",
    bg: "FFFFFF", titleBg: "111111", title: "111111", text: "333333",
    accent: "111111", accent2: "888888", bullet: "111111", footer: "BBBBBB",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "111111", coverText: "FFFFFF", coverAccent: "FACC15"
  },
  {
    id: "vibrant",
    name: "活力漸層",
    subtitle: "年輕 · 吸睛 · 行銷 / 提案",
    bg: "FFFFFF", titleBg: "6D28D9", title: "5B21B6", text: "2A2440",
    accent: "7C3AED", accent2: "EC4899", bullet: "DB2777", footer: "C4B5FD",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "6D28D9", coverText: "FFFFFF", coverAccent: "F472B6"
  },
  {
    id: "academic",
    name: "學術沉穩",
    subtitle: "嚴謹 · 清晰 · 報告 / 研究",
    bg: "FCFCF9", titleBg: "1E3A34", title: "14532D", text: "27313A",
    accent: "0F766E", accent2: "5EAD9C", bullet: "0F766E", footer: "A7B0AE",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "14532D", coverText: "F4F7F2", coverAccent: "86C7B6"
  },
  {
    id: "techdark",
    name: "科技深色",
    subtitle: "前衛 · 數據 · 科技 / 新創",
    bg: "0E1525", titleBg: "0E1525", title: "F1F5FF", text: "CBD5E1",
    accent: "38BDF8", accent2: "818CF8", bullet: "38BDF8", footer: "475569",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "0B1020", coverText: "F1F5FF", coverAccent: "38BDF8"
  },
  {
    id: "warmsun",
    name: "暖陽橙",
    subtitle: "親切 · 正面 · 教學 / 內部",
    bg: "FFFCF5", titleBg: "9A3412", title: "9A3412", text: "44342B",
    accent: "EA580C", accent2: "FB923C", bullet: "EA580C", footer: "D8C3A5",
    titleFont: "Microsoft JhengHei", bodyFont: "Microsoft JhengHei",
    coverBg: "9A3412", coverText: "FFF7ED", coverAccent: "FDBA74"
  }
];

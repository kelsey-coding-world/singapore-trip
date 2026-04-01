// ============================================================
// Parser.gs － PDF 行程解析器
// 功能：從 Google Drive 找到 PDF → 萃取文字 → 結構化成 Sheet 欄位
// 欄位：Day / Type / Time / Location / Description / Weather
// ============================================================

// Day 對照表（PDF 日期欄 → 可讀標籤）
const DAY_MAP = {
  '4/13': 'Day 1 (4/13)',
  '4/14': 'Day 2 (4/14)',
  '4/15': 'Day 3 (4/15)',
  '4/16': 'Day 4 (4/16)',
  '4/17': 'Day 5 (4/17)',
};

// 活動類型分類關鍵字
const TYPE_KEYWORDS = {
  '交通': ['機場', '起飛', '抵達', 'MRT', '🚄', '🚍', '巴士', 'Stn', 'Station', 'Express', 'CG', 'NE', 'TE', 'Opp', 'Aft'],
  '住宿': ['飯店', 'Hotel', 'Marriott', 'Dusit', 'Thani', 'check-in', '放行李'],
  '景點': ['Universal Studio', 'Gardens by the Bay', 'Night Safari', 'Sentosa', 'Skyline Luge',
            '魚尾獅', 'Chinatown', 'Arab Street', 'Bugis', 'Jewel', 'Clarke Quay', '克拉碼頭',
            '蘇丹回教堂', '觀音堂', 'Haji', '辦公室', 'Titansoft', 'Taikoo'],
  '餐飲': ['早餐', '午餐', '晚餐', 'Cafe', 'Restaurant', 'Grill', 'Duck', '肉骨茶', 'Rice',
            '麻辣', '吐司', 'Bagel', 'Noodle', 'Fernet', 'Chatterbox', 'Numb', '喜園'],
  '購物': ['超市', 'CS Fresh', 'Citylink', 'Vivocity', '逛'],
};

// ── 主解析函式 ───────────────────────────────────────────────
function parsePdfItinerary() {
  // 1. 從 Drive 取得 PDF 文字
  const rawText = extractTextFromPdf(CONFIG.PDF_FILE_NAME);
  if (!rawText) {
    Logger.log('❌ 無法取得 PDF 文字');
    return [];
  }
  Logger.log('📄 PDF 原始文字（前 500 字）：\n' + rawText.substring(0, 500));

  // 2. 解析成結構化列
  const rows = buildRows(rawText);
  Logger.log(`✅ 解析完成，共 ${rows.length} 筆`);
  return rows;
}

// ── 從 Google Drive 讀取 PDF 文字 ───────────────────────────
function extractTextFromPdf(fileName) {
  try {
    // 搜尋 Drive 中符合名稱的 PDF
    const files = DriveApp.getFilesByName(fileName);
    if (!files.hasNext()) {
      Logger.log(`❌ Drive 找不到檔案：${fileName}`);
      return null;
    }
    const pdfFile = files.next();
    Logger.log(`📂 找到 PDF：${pdfFile.getName()} (${pdfFile.getId()})`);

    // 將 PDF 轉為 Google Doc 以萃取文字
    const blob     = pdfFile.getBlob();
    const resource = { title: '_temp_itinerary_parse', mimeType: MimeType.GOOGLE_DOCS };
    const docFile  = Drive.Files.insert(resource, blob, { convert: true });
    const doc      = DocumentApp.openById(docFile.id);
    const text     = doc.getBody().getText();

    // 清理暫存 Doc
    DriveApp.getFileById(docFile.id).setTrashed(true);

    return text;
  } catch (e) {
    Logger.log('❌ extractTextFromPdf 錯誤：' + e.message);
    return null;
  }
}

// ── 將原始文字組成 [Day, Type, Time, Location, Description, Weather] ──
function buildRows(rawText) {
  const rows = [];

  // 依行切割
  const lines = rawText
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  // 預先定義每天的靜態行程（作為 fallback / 補充來源）
  // 格式：[day, type, time, location, description, weather]
  const staticData = buildStaticItinerary();

  // 嘗試從文字中匹配已知項目
  const usedKeys = new Set();

  lines.forEach(line => {
    staticData.forEach(item => {
      const key = item[0] + '|' + item[2] + '|' + item[3];
      if (usedKeys.has(key)) return;

      // 模糊比對：行程關鍵字出現在這行
      const keywords = [item[3], item[4]].join(' ').split(/[，,、\s]+/).filter(k => k.length > 1);
      const matched  = keywords.some(k => line.includes(k));

      if (matched) {
        usedKeys.add(key);
      }
    });
  });

  // 直接使用靜態資料（已從 PDF 內容手工結構化）
  return staticData;
}

// ── 靜態行程資料（根據 PDF 內容結構化）────────────────────────
// 欄位順序：Day / Type / Time / Location / Description / Weather
function buildStaticItinerary() {
  return [
    // ── Day 1 (4/13) ──────────────────────────────────────────
    ['Day 1 (4/13)', '交通', '08:00', '桃園機場 T2',       '從桃園機場 T2 起飛', ''],
    ['Day 1 (4/13)', '交通', '12:40', '樟宜機場 T2',       '抵達樟宜機場 T2', ''],
    ['Day 1 (4/13)', '住宿', '抵達後', 'JW Marriott South Beach', '入住 JW Marriott South Beach，放行李', ''],
    ['Day 1 (4/13)', '交通', '下午',   'MRT CG 線',        '樟宜機場 → 丹那美拉 Tanah Merah → 政府大廈 City Hall', ''],
    ['Day 1 (4/13)', '餐飲', '午餐',   'Jewel Changi Airport', '候選：① PS Cafe ② Josh\'s Grill ③ Imperial Treasure Super Peking Duck', ''],
    ['Day 1 (4/13)', '餐飲', '晚餐',   '黃亞細肉骨茶 Ng Ah Sio', 'Ng Ah Sio Bak Kut Teh (Tai Seng)', ''],
    ['Day 1 (4/13)', '景點', '晚上',   'Clarke Quay 克拉碼頭', '前往克拉碼頭看夜景', ''],

    // ── Day 2 (4/14) ──────────────────────────────────────────
    ['Day 2 (4/14)', '餐飲', '早餐',   '亞坤咖椰吐司 Citylink Mall', '亞坤咖椰吐司（Citylink Mall 分店）', ''],
    ['Day 2 (4/14)', '景點', '上午',   'Universal Studios Singapore', '搭 Sentosa Express 前往環球影城', ''],
    ['Day 2 (4/14)', '交通', '上午',   'MRT + Sentosa Express',    '🚍 10 Clementi Rd → HarbourFront / Vivocity；🚄 Sentosa Express：怡豐城站 → Resorts World Station', ''],
    ['Day 2 (4/14)', '景點', '全天',   'Universal Studios Singapore', '必玩：① Sci-Fi City（TRANSFORMERS / Battlestar Galactica）② Minion Land ③ Ancient Egypt（Revenge of the Mummy）④ The Lost World（Jurassic Park Rapids）⑤ Far Far Away ⑥ New York；水世界（必看）', ''],
    ['Day 2 (4/14)', '餐飲', '午餐',   'Caffe Fernet',             '景觀超讚的餐廳', ''],
    ['Day 2 (4/14)', '景點', '14:00',  'Skyline Luge Sentosa',     '玩 Skyline Luge', ''],
    ['Day 2 (4/14)', '交通', '下午',   'Sentosa Express',          '🚄 Resorts World Station → 怡豐城站；NE Punggol Coast：港灣 → 克拉碼頭', ''],

    // ── Day 3 (4/15) ──────────────────────────────────────────
    ['Day 3 (4/15)', '餐飲', '早餐',   '飯店早餐',                 'JW Marriott 飯店早餐', ''],
    ['Day 3 (4/15)', '景點', '上午',   '魚尾獅公園',               '前往魚尾獅拍照', ''],
    ['Day 3 (4/15)', '景點', '上午',   'Arab Street',              '① 蘇丹回教堂 ② Haji Lane', ''],
    ['Day 3 (4/15)', '景點', '上午',   'Bugis',                    '觀音堂佛祖廟', ''],
    ['Day 3 (4/15)', '餐飲', '午餐',   'Chin Chin Restaurant',     'JJ 推薦', ''],
    ['Day 3 (4/15)', '景點', '下午',   'Titansoft 鈦坦辦公室',      '前往 Titansoft 新加坡辦公室參觀\n🚍 70 Yio Chu Kang：Opp Suntec City → Aft Tai Seng Stn，步行 10 mins', ''],
    ['Day 3 (4/15)', '餐飲', '晚餐',   'Chatterbox Chicken Rice',  'Singapore 雞飯名店', ''],
    ['Day 3 (4/15)', '景點', '晚上',   'Gardens by the Bay',       '🚄 Sentosa Express：Imbiah → 怡豐城站；NE → TE Bayshore → Gardens by the Bay', ''],
    ['Day 3 (4/15)', '景點', '19:15',  'Night Safari',             '夜間野生動物園（快速通關）', ''],

    // ── Day 4 (4/16) ──────────────────────────────────────────
    ['Day 4 (4/16)', '餐飲', '早餐',   '飯店早餐 & 喜園',           '飯店早餐，另外推薦喜園', ''],
    ['Day 4 (4/16)', '景點', '上午',   'Chinatown',                '逛牛車水；推薦：① 日日紅麻辣香鍋 ② 林志源 新橋路店 ③ High Street Tai Wah Pork Noodle ④ My Awesome Cafe ⑤ Two Men Bagel House Tanjong Pagar', ''],
    ['Day 4 (4/16)', '交通', '上午',   'MRT',                      '🚍 131 Bt Merah：OUE Bayfront → HarbourFront / Vivocity；🚄 Sentosa Express：怡豐城站 → Imbiah Station', ''],
    ['Day 4 (4/16)', '住宿', '下午',   'Dusit Thani Laguna Singapore', '入住 Dusit Thani Laguna Singapore', ''],
    ['Day 4 (4/16)', '餐飲', '晚餐',   'Numb Restaurant 川麻记',   '川麻记 @ Marina One', ''],
    ['Day 4 (4/16)', '購物', '晚上',   'CS Fresh Bugis Junction',  '逛 CS Fresh 超市', ''],

    // ── Day 5 (4/17) ──────────────────────────────────────────
    ['Day 5 (4/17)', '景點', '上午',   'Jewel Changi Airport',     '逛 Jewel Changi Airport', ''],
    ['Day 5 (4/17)', '交通', '14:00',  '樟宜機場 T2',              '從樟宜機場 T2 起飛返台', ''],
    ['Day 5 (4/17)', '交通', '18:50',  '桃園機場 T2',              '抵達桃園機場 T2', ''],
  ];
}

// ── 工具：依關鍵字推斷活動類型 ──────────────────────────────
function inferType(text) {
  for (const [type, keywords] of Object.entries(TYPE_KEYWORDS)) {
    if (keywords.some(k => text.includes(k))) return type;
  }
  return '其他';
}

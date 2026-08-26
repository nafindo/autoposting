/**
 * ============================================================
 * NAFINDO BLOGGER AUTOMATION API v3.5
 * Auto Keyword Generation | Nightly Refresh | Smart Queue
 * ============================================================
 */

const DEFAULT_CONFIG = {
  BLOG_ID: '5323802325028284266',
  GROQ_API_KEY: '',
  WHATSAPP_NUMBER: '081232797271',
  COMPANY_NAME: 'CV Nafindo Group',
  GROQ_MODEL: 'openai/gpt-oss-120b',
  MIN_POST_PER_DAY: 25,
  MAX_POST_PER_DAY: 35,
  MIN_DELAY_MINUTES: 17,
  MAX_DELAY_MINUTES: 35,
  POSTS_PER_BATCH: 1,
  TRIGGER_INTERVAL_MINUTES: 25,
  JAM_AKTIF_MULAI: 8,
  JAM_AKTIF_SELESAI: 22,
  FOTO_MODE: 'drive',
  DRIVE_FOLDER_ID: '',
  SPREADSHEET_ID: '',
  KEYWORD_REFRESH_HOUR: 0  // Jam 00:00 refresh keyword
};

// ==================== WEB APP ====================

function doGet(e) {
  const action = e.parameter.action;
  let result = { success: false, error: 'Unknown action: ' + action };
  
  try {
    switch(action) {
      case 'getConfig': result = getConfig(); break;
      case 'getStats': result = getStats(); break;
      case 'getStatsByDate': result = getStatsByDate(e.parameter.date); break;
      case 'getWeeklyStats': result = getWeeklyStats(parseInt(e.parameter.offset) || 0); break;
      case 'getMonthlyStats': result = getMonthlyStats(parseInt(e.parameter.offset) || 0); break;
      case 'getProdukList': result = { success: true, data: getProdukList() }; break;
      case 'getProdukData': result = getProdukData(e.parameter.sheetName); break;
      case 'getLog': result = { success: true, data: getLog(parseInt(e.parameter.limit) || 50) }; break;
      case 'getLogByDate': result = getLogByDate(e.parameter.date); break;
      case 'getAIKeys': result = getAIKeys(); break;
      case 'getDriveImages': result = getDriveImages(e.parameter.folderId); break;
      case 'getTriggerStatus': result = getTriggerStatus(); break;
      case 'previewPost': result = previewPost(e.parameter.produk, e.parameter.keyword, parseInt(e.parameter.row)); break;
      case 'getMasterKeywords': result = getMasterKeywords(e.parameter.produkId); break;
      case 'getQueueStatus': result = getQueueStatus(); break;
      case 'checkBloggerAuth': result = authorizeBlogger(); break;
      case 'checkDriveAuth': result = testDriveAuth(); break;
      case 'getAllLabels': result = { success: true, labels: getAllLabels() }; break;
    }
  } catch(err) {
    result = { success: false, error: err.toString() };
  }
  
  return jsonResponse(result);
}

function doPost(e) {
  let result = { success: false, error: 'Unknown action' };
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'No POST data received' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'saveConfig': result = saveConfig(data.config); break;
      case 'addProduk': result = addProduk(data.nama, data.deskripsi, data.prioritas, data.masterKeywords, data.heroImages, data.varian || data.galleryVarian, data.gallery || data.saranWarna, data.specs, data.labels); break;
      case 'updateProduk': result = updateProduk(data.id, data.nama, data.deskripsi, data.status, data.prioritas, data.heroImages, data.varian || data.galleryVarian, data.gallery || data.saranWarna, data.specs, data.sheetName, data.labels); break;
      case 'updateProdukData': result = updateProdukData(data.sheetName, data.heroImages, data.varian || data.galleryVarian, data.gallery || data.saranWarna, data.specs); break;
      case 'deleteProduk': result = deleteProduk(data.id); break;
      case 'addKeyword': result = addKeyword(data.sheetName, data.keyword, data.label, data.spec); break;
      case 'importKeywords': result = importKeywords(data.sheetName, data.keywords); break;
      case 'deleteKeyword': result = deleteKeyword(data.sheetName, data.row); break;
      case 'updateKeyword': result = updateKeyword(data.sheetName, data.row, data.data); break;
      case 'postManual': result = postManual(data.produkNama, data.keyword, data.row); break;
      case 'saveAIKeys': result = saveAIKeys(data.keys); break;
      case 'aktifkanTrigger': result = aktifkanTrigger(data.minMinutes, data.maxMinutes); break;
      case 'nonaktifkanTrigger': result = nonaktifkanTrigger(); break;
      case 'testPosting': result = testPosting(); break;
      case 'resetStatusGagal': result = resetStatusGagal(data.sheetName); break;
      case 'saveMasterKeywords': result = saveMasterKeywords(data.produkId, data.keywords); break;
      case 'generateKeywords': result = generateKeywordsFromMaster(data.produkId, data.count); break;
      case 'refreshAllKeywords': result = refreshAllKeywords(); break;
      case 'uploadImage': result = uploadImage(data.name, data.type, data.base64, data.folderId); break;
      default: result = { success: false, error: 'Aksi tidak dikenali: ' + action }; break;
    }
  } catch(err) {
    result = { success: false, error: err.toString() };
  }
  
  return jsonResponse(result);
}
function doOptions() {
  return ContentService.createTextOutput('');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== SPREADSHEET HELPER ====================

function getSS() {
  const SPREADSHEET_ID = '1xUqdNRjrZqjO8QclecvER4dEL8z5cbWvvtGf7UGTTKo';
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ==================== KONFIGURASI ====================

function getConfig() {
  const ss = getSS();
  let sheet = ss.getSheetByName('CONFIG');
  
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG');
    sheet.appendRow(['Key', 'Value']);
    Object.entries(DEFAULT_CONFIG).forEach(([k, v]) => sheet.appendRow([k, v]));
  }
  
  const data = sheet.getDataRange().getValues();
  const config = { ...DEFAULT_CONFIG };
  const numFields = ['MIN_POST_PER_DAY', 'MAX_POST_PER_DAY', 'JAM_AKTIF_MULAI', 'JAM_AKTIF_SELESAI', 'TRIGGER_INTERVAL_MINUTES', 'KEYWORD_REFRESH_HOUR', 'MIN_DELAY_MINUTES', 'MAX_DELAY_MINUTES', 'POSTS_PER_BATCH'];

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    let val = data[i][1];
    if (numFields.includes(key) && !isNaN(val) && val !== '') {
      val = Number(val);
    } else {
      val = String(val || '').trim();
    }
    config[key] = val;
  }
  
  // Perbaiki Blog ID jika terpotong akibat format angka
  if (config.BLOG_ID === '5323802325028284000' || !config.BLOG_ID) {
    config.BLOG_ID = DEFAULT_CONFIG.BLOG_ID || '5323802325028284266';
  }

  if (!config.GROQ_MODEL || config.GROQ_MODEL.toLowerCase() === 'groq' || !config.GROQ_MODEL.includes('/')) {
    config.GROQ_MODEL = 'openai/gpt-oss-120b';
  }

  return config;
}

function getTodayTargetPosts() {
  const config = getConfig();
  const minPost = Math.max(1, Number(config.MIN_POST_PER_DAY || 25));
  const maxPost = Math.max(minPost, Number(config.MAX_POST_PER_DAY || 35));
  
  const props = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const storedDate = props.getProperty('TODAY_TARGET_DATE');
  let target = Number(props.getProperty('TODAY_TARGET_POSTS'));
  
  // Jika hari baru atau belum ada target acak atau nilai target di luar rentang min-max baru
  if (storedDate !== todayStr || !target || target < minPost || target > maxPost) {
    target = Math.floor(Math.random() * (maxPost - minPost + 1)) + minPost;
    props.setProperty('TODAY_TARGET_DATE', todayStr);
    props.setProperty('TODAY_TARGET_POSTS', String(target));
    console.log(`🎲 Target kuota posting acak hari ini (${todayStr}): ${target} postingan (Rentang: ${minPost}-${maxPost})`);
  }
  
  return {
    target: target,
    minPost: minPost,
    maxPost: maxPost,
    date: todayStr
  };
}

function saveConfig(configObj) {
  const ss = getSS();
  let sheet = ss.getSheetByName('CONFIG');
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG');
    sheet.appendRow(['Key', 'Value']);
  }
  
  sheet.clear();
  sheet.appendRow(['Key', 'Value']);
  Object.entries(configObj).forEach(([k, v]) => {
    // Tulis sebagai plain text jika ID atau nomor WA
    if (k === 'BLOG_ID' || k === 'WHATSAPP_NUMBER' || k === 'SPREADSHEET_ID') {
      sheet.appendRow([k, "'" + String(v).replace(/^'+/, '')]);
    } else {
      sheet.appendRow([k, v]);
    }
  });
  
  // Sinkronkan juga nilai jeda acak & kuota target ke PropertiesService
  const props = PropertiesService.getScriptProperties();
  if (configObj.MIN_DELAY_MINUTES) {
    props.setProperty('MIN_DELAY_MINUTES', String(configObj.MIN_DELAY_MINUTES));
  }
  if (configObj.MAX_DELAY_MINUTES) {
    props.setProperty('MAX_DELAY_MINUTES', String(configObj.MAX_DELAY_MINUTES));
  }
  if (configObj.MIN_POST_PER_DAY) {
    props.setProperty('MIN_POST_PER_DAY', String(configObj.MIN_POST_PER_DAY));
  }
  if (configObj.MAX_POST_PER_DAY) {
    props.setProperty('MAX_POST_PER_DAY', String(configObj.MAX_POST_PER_DAY));
  }
  
  // Re-generate target hari ini agar sesuai dengan rentang baru
  props.deleteProperty('TODAY_TARGET_POSTS');
  getTodayTargetPosts();
  
  // Jika trigger auto-posting sedang aktif, jadwalkan ulang dengan jeda baru sekarang juga
  const mode = props.getProperty('TRIGGER_MODE');
  if (mode === 'random') {
    jadwalkanPostinganBerikutnya();
  }
  
  return { success: true };
}

// ==================== AI KEY ROTATION ====================

function getAIKeys() {
  const ss = getSS();
  let sheet = ss.getSheetByName('AI_KEYS');
  if (!sheet) return { success: true, keys: [] };
  
  const data = sheet.getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < data.length; i++) {
    keys.push({
      key: data[i][0],
      status: data[i][1] || 'Aktif',
      usage: data[i][2] || 0,
      lastUsed: data[i][3] || ''
    });
  }
  return { success: true, keys };
}

function saveAIKeys(keys) {
  const ss = getSS();
  let sheet = ss.getSheetByName('AI_KEYS');
  if (!sheet) {
    sheet = ss.insertSheet('AI_KEYS');
    sheet.appendRow(['API Key', 'Status', 'Usage Count', 'Last Used']);
  } else {
    sheet.clear();
    sheet.appendRow(['API Key', 'Status', 'Usage Count', 'Last Used']);
  }
  
  keys.forEach(k => {
    sheet.appendRow([k.key, k.status || 'Aktif', k.usage || 0, k.lastUsed || '']);
  });
  
  return { success: true };
}

function getAIKeyRotation() {
  const ss = getSS();
  let sheet = ss.getSheetByName('AI_KEYS');
  if (!sheet) {
    const config = getConfig();
    return config.GROQ_API_KEY || '';
  }
  
  const data = sheet.getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1] !== 'Nonaktif') {
      keys.push({ key: data[i][0], usage: data[i][2] || 0, row: i + 1 });
    }
  }
  
  if (keys.length === 0) {
    const config = getConfig();
    return config.GROQ_API_KEY || '';
  }
  
  keys.sort((a, b) => a.usage - b.usage);
  const selected = keys[0];
  
  sheet.getRange(selected.row, 3).setValue(selected.usage + 1);
  sheet.getRange(selected.row, 4).setValue(new Date());
  
  return selected.key;
}

// ==================== MASTER KEYWORD ====================

function getMasterKeywords(produkId) {
  const ss = getSS();
  let sheet = ss.getSheetByName('MASTER_KEYWORD');
  if (!sheet) return { success: true, keywords: [] };
  
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === produkId) {
      result.push({
        keyword: data[i][1],
        created: data[i][2]
      });
    }
  }
  return { success: true, keywords: result };
}

function saveMasterKeywords(produkId, keywords) {
  const ss = getSS();
  let sheet = ss.getSheetByName('MASTER_KEYWORD');
  if (!sheet) {
    sheet = ss.insertSheet('MASTER_KEYWORD');
    sheet.appendRow(['Produk ID', 'Keyword Master', 'Created']);
  }
  
  // Hapus yang lama untuk produk ini
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === produkId) {
      sheet.deleteRow(i + 1);
    }
  }
  
  // Tambah yang baru
  keywords.forEach(kw => {
    if (kw && String(kw).trim()) {
      sheet.appendRow([produkId, String(kw).trim(), new Date()]);
    }
  });
  
  return { success: true };
}

// ==================== AUTO GENERATE KEYWORD LONG TAIL ====================

/**
 * Generate keyword long tail dari master keyword menggunakan AI
 */
function generateKeywordsFromMaster(produkId, count) {
  const ss = getSS();
  const produkSheet = ss.getSheetByName('MASTER_PRODUK');
  if (!produkSheet) return { success: false, error: 'Produk tidak ditemukan' };
  
  const produkData = produkSheet.getDataRange().getValues();
  let produkNama = '';
  let produkDeskripsi = '';
  let defaultLabel = '';
  for (let i = 1; i < produkData.length; i++) {
    if (produkData[i][0] === produkId) {
      produkNama = produkData[i][1];
      produkDeskripsi = produkData[i][2] || '';
      defaultLabel = produkData[i][6] || '';
      break;
    }
  }
  
  if (!produkNama) return { success: false, error: 'Produk tidak ditemukan' };
  
  // Ambil master keywords
  const masterRes = getMasterKeywords(produkId);
  const masterKeywords = masterRes.keywords.map(k => k.keyword);
  if (masterKeywords.length === 0) return { success: false, error: 'Tidak ada keyword master' };
  
  // Ambil semua keyword yang sudah ada (untuk cek duplicate) & spesifikasi produk
  const sheetName = 'PRODUK_' + produkNama.replace(/\s+/g, '_').substring(0, 20);
  const pSheet = ss.getSheetByName(sheetName);
  const existingKeywords = new Set();
  let specsStr = '';
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      if (pData[i][0]) existingKeywords.add(String(pData[i][0]).toLowerCase().trim());
      if (pData[i][8] && !specsStr) specsStr = String(pData[i][8]).trim();
    }
  }
  
  // Generate via AI dengan pemahaman produk & SEO Intent yang ketat
  const newKeywords = generateKeywordsAI(produkNama, masterKeywords, count || 50, existingKeywords, produkDeskripsi, specsStr);
  
  if (newKeywords.length === 0) return { success: false, error: 'Gagal generate keyword' };
  
  // Insert ke sheet produk
  if (!pSheet) {
    const ns = ss.insertSheet(sheetName);
    ns.appendRow(['Kata Kunci', 'Judul', 'Hero Image', 'Varian', 'Keterangan', 'Label', 'Status', 'Tgl Selesai', 'Spesifikasi', 'Gallery', 'URL Post']);
  }
  
  const targetSheet = ss.getSheetByName(sheetName);
  let added = 0;
  newKeywords.forEach(kw => {
    if (!existingKeywords.has(kw.toLowerCase().trim())) {
      targetSheet.appendRow([kw, '', '', '', '', defaultLabel || '', '', '', '', '', '']);
      existingKeywords.add(kw.toLowerCase().trim());
      added++;
    }
  });
  
  return { success: true, generated: newKeywords.length, added: added, keywords: newKeywords };
}

function generateKeywordsAI(produkNama, masterKeywords, count, existingSet, produkDeskripsi, specs) {
  const apiKey = getAIKeyRotation();
  const url = 'https://api.groq.com/openai/v1/chat/completions';
  
  const prompt = `Kamu adalah Pakar Riset Keyword SEO & Search Intent Specialist untuk Google Indonesia.

INFORMASI PRODUK:
- Nama Produk: ${produkNama}
- Keyword Master: ${masterKeywords.join(', ')}
${produkDeskripsi ? `- Deskripsi Produk: ${produkDeskripsi}` : ''}
${specs ? `- Spesifikasi / Material: ${specs}` : ''}

TUGAS UTAMA:
Buat ${count} Long Tail Keywords SEO bahasa Indonesia yang 100% RELEVAN KHUSUS untuk produk "${produkNama}".
Fokuskan pada High Commercial & Transactional Search Intent (kata kunci yang sering diketik calon pembeli saat ingin mencari, membeli, atau menyewa jasa di Google).

POLA RUMUS SEO SEARCH INTENT (Kombinasikan secara variatif & alami):
1. Pola Transaksi / Pembelian:
   - "jual [keyword] [kota]"
   - "harga [keyword] per meter / per m2 / per box [kota]"
   - "distributor / supplier / agen [keyword] [kota]"
   - "toko [keyword] terdekat [kota]"
   - "grosir [keyword] termurah"
2. Pola Jasa & Pemasangan:
   - "jasa pasang [keyword] [kota]"
   - "biaya pasang [keyword] per m2 [kota]"
   - "tukang pasang [keyword] terpercaya [kota]"
3. Pola Spesifikasi & Keunggulan Produk:
   - "[keyword] motif kayu jati / minimalis / modern"
   - "[keyword] 2mm anti rayap dan tahan air"
   - "[keyword] berkualitas awet"
4. Pola Aplikasi & Kebutuhan Ruangan:
   - "[keyword] untuk kamar tidur / ruang tamu / dapur / kantor / cafe / hotel"
   - "[keyword] untuk lantai rumah minimalis"
5. Variasi Target Kota Indonesia (Local SEO):
   - Jakarta, Surabaya, Bandung, Medan, Semarang, Makassar, Tangerang, Bekasi, Depok, Bogor, Bali, Yogyakarta, Solo, Malang, Balikpapan, Palembang, dll.

⚠️ ATURAN KETAT (STRICT BOUNDARIES):
1. FOKUS LINGKUP PRODUK: Semua keyword HARUS berada dalam lingkup produk "${produkNama}" dan Keyword Master yang diberikan.
2. DILARANG KERAS mencampuradukkan dengan material / produk lain yang BEDA KATEGORI.
   - Contoh: Jika produknya "Vinyl / Vinyl Plank", DILARANG SEKALI-KALI memasukkan kata "SPC", "Parket Kayu Asli", "Keramik", "Granit", "Marmer", atau "Karpet".
   - Jika produknya "Pintu Harmonika", jangan masukkan "Rolling Door" kecuali disebutkan di master keyword.
3. DILARANG menyebut merek kompetitor lain kecuali merek produk ini (${produkNama}).
4. Setiap keyword HARUS unik dan TIDAK BOLEH sama dengan yang sudah ada di list berikut: ${Array.from(existingSet).slice(0, 100).join(', ')}.
5. Huruf kecil semua (lowercase), tanpa nomor urut di depan (seperti 1., 2.), tanpa tanda kutip.
6. Kembalikan HANYA format JSON valid tanpa kata pengantar atau penutup.

OUTPUT FORMAT (JSON):
{"keywords": ["keyword 1", "keyword 2", "keyword 3", ...]}`;

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        messages: [
          { role: 'system', content: 'Kamu adalah pakar riset SEO Google Indonesia. Hasilkan keyword terstruktur 100% relevan dalam JSON valid.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.7,
        max_tokens: 3000,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(res.getContentText());
    if (!json.choices || !json.choices[0]?.message?.content) return [];
    
    const content = JSON.parse(json.choices[0].message.content);
    const rawKeywords = content.keywords || [];
    
    // Double Filter Sanitasi: Menghapus keyword yang melenceng dari jenis produk
    const cleanKeywords = [];
    const forbiddenTerms = [];
    const lowerProd = (produkNama + ' ' + masterKeywords.join(' ')).toLowerCase();
    
    // Jika produknya vinyl murni (bukan SPC), blacklist kata SPC
    if (lowerProd.includes('vinyl') && !lowerProd.includes('spc')) {
      forbiddenTerms.push('spc');
    }
    // Jika produknya SPC murni, blacklist kata vinyl roll atau stiker jika tidak relevan
    if (lowerProd.includes('spc') && !lowerProd.includes('vinyl')) {
      forbiddenTerms.push('vinyl roll');
    }
    
    rawKeywords.forEach(kw => {
      if (typeof kw !== 'string') return;
      const cleanKw = kw.toLowerCase().trim().replace(/^[0-9]+[\.\)\-]\s*/, '');
      if (!cleanKw || cleanKw.length < 5) return;
      
      // Cek apakah mengandung kata terlarang yang melenceng
      const hasForbidden = forbiddenTerms.some(term => cleanKw.includes(term));
      if (!hasForbidden) {
        cleanKeywords.push(cleanKw);
      }
    });
    
    return cleanKeywords;
  } catch(e) {
    console.error('Generate keyword error: ' + e);
    return [];
  }
}

// ==================== REFRESH KEYWORD HARIAN (JAM 00:00) ====================

/**
 * Fungsi ini dipanggil trigger harian jam 00:00
 */
function refreshAllKeywords() {
  console.log('🌙 REFRESH KEYWORD HARIAN: ' + new Date());
  
  const ss = getSS();
  const produkSheet = ss.getSheetByName('MASTER_PRODUK');
  if (!produkSheet) {
    console.log('Master produk tidak ditemukan');
    return;
  }
  
  const produkData = produkSheet.getDataRange().getValues();
  let totalGenerated = 0;
  
  for (let i = 1; i < produkData.length; i++) {
    const produkId = produkData[i][0];
    const produkNama = produkData[i][1];
    const status = produkData[i][3];
    
    if (status === 'Nonaktif') continue;
    
    const sheetName = 'PRODUK_' + produkNama.replace(/\s+/g, '_').substring(0, 20);
    const pSheet = ss.getSheetByName(sheetName);
    
    // Cek berapa keyword yang masih tersedia (belum Selesai/Gagal)
    let availableCount = 0;
    if (pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for (let j = 1; j < pData.length; j++) {
        if (pData[j][0] && pData[j][6] !== 'Selesai' && pData[j][6] !== 'Gagal') {
          availableCount++;
        }
      }
    }
    
    console.log(`📦 ${produkNama}: ${availableCount} keyword tersedia`);
    
    // Kalau sisa keyword < 10, generate 50 baru
    if (availableCount < 10) {
      console.log(`🔄 Generate keyword baru untuk ${produkNama}...`);
      const result = generateKeywordsFromMaster(produkId, 50);
      if (result.success) {
        console.log(`✅ ${produkNama}: ${result.added} keyword baru ditambahkan`);
        totalGenerated += result.added;
      } else {
        console.error(`❌ ${produkNama}: ${result.error}`);
      }
    }
  }
  
  console.log(`🌙 SELESAI. Total keyword baru: ${totalGenerated}`);
}

function aktifkanRefreshKeywordHarian() {
  // Hapus trigger refresh lama
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshAllKeywords') ScriptApp.deleteTrigger(t);
  });
  
  // Buat trigger harian jam 00:00
  ScriptApp.newTrigger('refreshAllKeywords')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(0)
    .create();
  
  return { success: true, message: 'Trigger refresh keyword harian aktif (jam 00:00)' };
}

// ==================== PRODUK ====================

function getProdukList() {
  const ss = getSS();
  let sheet = ss.getSheetByName('MASTER_PRODUK');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const produk = [];
  for (let i = 1; i < data.length; i++) {
    produk.push({
      id: data[i][0],
      nama: data[i][1],
      deskripsi: data[i][2],
      status: data[i][3] || 'Aktif',
      prioritas: data[i][4] || 5,
      created: data[i][5] || '',
      labels: data[i][6] || ''
    });
  }
  return produk;
}

function addProduk(nama, deskripsi, prioritas, masterKeywords, heroImages, varian, gallery, specs, labels) {
  const ss = getSS();
  
  let sheet = ss.getSheetByName('MASTER_PRODUK');
  if (!sheet) {
    sheet = ss.insertSheet('MASTER_PRODUK');
    sheet.appendRow(['ID', 'Nama', 'Deskripsi', 'Status', 'Prioritas', 'Created', 'Labels']);
  } else if (sheet.getMaxColumns() < 7) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, 7).setValue('Labels');
  }
  
  const id = 'PRD_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const labelStr = Array.isArray(labels) ? labels.join(', ') : String(labels || '').trim();
  sheet.appendRow([id, nama, deskripsi || '', 'Aktif', prioritas || 5, new Date(), labelStr]);
  
  // Simpan master keywords
  if (masterKeywords && masterKeywords.length > 0) {
    saveMasterKeywords(id, masterKeywords);
  }
  
  // Buat sheet produk
  const sheetName = 'PRODUK_' + nama.replace(/\s+/g, '_').substring(0, 20);
  let ns = ss.getSheetByName(sheetName);
  if (!ns) {
    ns = ss.insertSheet(sheetName);
    ns.appendRow(['Kata Kunci', 'Judul', 'Hero Image', 'Varian', 'Keterangan', 'Label', 'Status', 'Tgl Selesai', 'Spesifikasi', 'Gallery', 'URL Post']);
  }
  
  // Auto generate 50 keyword pertama
  if (masterKeywords && masterKeywords.length > 0) {
    generateKeywordsFromMaster(id, 50);
  }

  // Update heroImages, varian, gallery, specs jika disediakan
  if (heroImages || varian || gallery || specs) {
    updateProdukData(sheetName, heroImages, varian, gallery, specs);
  }
  
  return { success: true, id, sheetName };
}

function updateProduk(id, nama, deskripsi, status, prioritas, heroImages, varian, gallery, specs, sheetName, labels) {
  const ss = getSS();
  const sheet = ss.getSheetByName('MASTER_PRODUK');
  if (!sheet) return { success: false, error: 'Sheet MASTER_PRODUK tidak ditemukan' };
  
  if (sheet.getMaxColumns() < 7) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, 7).setValue('Labels');
  }

  const labelStr = labels !== undefined ? (Array.isArray(labels) ? labels.join(', ') : String(labels || '').trim()) : undefined;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      if (nama) sheet.getRange(i + 1, 2).setValue(nama);
      if (deskripsi !== undefined) sheet.getRange(i + 1, 3).setValue(deskripsi);
      if (status) sheet.getRange(i + 1, 4).setValue(status);
      if (prioritas !== undefined) sheet.getRange(i + 1, 5).setValue(prioritas);
      if (labelStr !== undefined) sheet.getRange(i + 1, 7).setValue(labelStr);
      
      const targetSheet = sheetName || ('PRODUK_' + (nama || data[i][1]).replace(/\s+/g, '_').substring(0, 20));
      if (heroImages || varian || gallery || specs) {
        updateProdukData(targetSheet, heroImages, varian, gallery, specs);
      }
      return { success: true };
    }
  }
  return { success: false, error: 'Produk ID ' + id + ' tidak ditemukan di MASTER_PRODUK' };
}

function updateProdukData(sheetName, heroImages, varian, gallery, specs) {
  const ss = getSS();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Kata Kunci', 'Judul', 'Hero Image', 'Varian', 'Keterangan', 'Label', 'Status', 'Tgl Selesai', 'Spesifikasi', 'Gallery', 'URL Post']);
  }

  // Pastikan sheet memiliki minimal 11 kolom
  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }

  const heroArr = Array.isArray(heroImages) ? heroImages : parseLinks(heroImages);
  const varianArr = Array.isArray(varian) ? varian : [];
  const galleryArr = Array.isArray(gallery) ? gallery : parseLinks(gallery);
  const specStr = String(specs || '').trim();

  const maxDataRows = Math.max(heroArr.length, varianArr.length, galleryArr.length, 1);
  const currentLastRow = Math.max(sheet.getLastRow(), 1);
  const totalRowsNeeded = Math.max(currentLastRow, maxDataRows + 1);

  // Pastikan baris cukup
  if (sheet.getMaxRows() < totalRowsNeeded) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRowsNeeded - sheet.getMaxRows());
  }

  // Jika baris kosong (hanya header)
  if (sheet.getLastRow() <= 1) {
    const rowsToAdd = [];
    for (let r = 0; r < maxDataRows; r++) {
      const h = heroArr[r] || '';
      const v = varianArr[r] ? (typeof varianArr[r] === 'object' ? (varianArr[r].link || '') : varianArr[r]) : '';
      const vk = varianArr[r] ? (typeof varianArr[r] === 'object' ? (varianArr[r].ket || '') : '') : '';
      const g = galleryArr[r] || '';
      const sp = (r === 0) ? specStr : '';
      rowsToAdd.push(['', '', h, v, vk, '', '', '', sp, g, '']);
    }
    if (rowsToAdd.length > 0) {
      sheet.getRange(2, 1, rowsToAdd.length, 11).setValues(rowsToAdd);
    }
    return { success: true };
  }

  // Baca seluruh data yang ada
  const numRows = totalRowsNeeded - 1;
  const range = sheet.getRange(2, 1, numRows, 11);
  const values = range.getValues();

  for (let r = 0; r < numRows; r++) {
    if (r < heroArr.length) values[r][2] = heroArr[r];
    else if (r < maxDataRows && heroArr.length === 0) values[r][2] = '';

    if (r < varianArr.length) {
      values[r][3] = typeof varianArr[r] === 'object' ? (varianArr[r].link || '') : varianArr[r];
      values[r][4] = typeof varianArr[r] === 'object' ? (varianArr[r].ket || '') : '';
    } else if (r < maxDataRows && varianArr.length === 0) {
      values[r][3] = '';
      values[r][4] = '';
    }

    if (r === 0) {
      values[r][8] = specStr;
    }

    if (r < galleryArr.length) values[r][9] = galleryArr[r];
    else if (r < maxDataRows && galleryArr.length === 0) values[r][9] = '';
  }

  range.setValues(values);
  return { success: true };
}

function deleteProduk(id) {
  const ss = getSS();
  const sheet = ss.getSheetByName('MASTER_PRODUK');
  if (!sheet) return { success: false };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// ==================== KEYWORD ====================

function getProdukData(sheetName) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  const keywords = [];
  const heroImages = [];
  const specs = [];
  const varian = [];
  const gallery = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      keywords.push({
        row: i + 1,
        keyword: data[i][0],
        judul: data[i][1] || '',
        hero: data[i][2] || '',
        varian: data[i][3] || '',
        keterangan: data[i][4] || '',
        label: data[i][5] || '',
        status: data[i][6] || '',
        tglSelesai: data[i][7] || '',
        spec: data[i][8] || '',
        gallery: data[i][9] || '',
        urlPost: data[i][10] || ''
      });
    }
    if (data[i][2]) heroImages.push(...parseLinks(data[i][2]));
    if (data[i][8]) specs.push(String(data[i][8]).trim());
    if (data[i][3]) {
      const links = parseLinks(data[i][3]);
      links.forEach(l => varian.push({ link: l, ket: data[i][4] || '' }));
    }
    if (data[i][9]) gallery.push(...parseLinks(data[i][9]));
  }
  
  // Cari label produk dari MASTER_PRODUK
  let prodLabels = '';
  const mSheet = ss.getSheetByName('MASTER_PRODUK');
  if (mSheet) {
    const mData = mSheet.getDataRange().getValues();
    const cleanSheetName = sheetName.replace(/^PRODUK_/, '').toLowerCase();
    for (let i = 1; i < mData.length; i++) {
      const pName = String(mData[i][1] || '').replace(/\s+/g, '_').substring(0, 20).toLowerCase();
      if (pName === cleanSheetName || String(mData[i][1]).toLowerCase() === cleanSheetName) {
        prodLabels = mData[i][6] || '';
        break;
      }
    }
  }
  if (!prodLabels && data.length > 1 && data[1][5]) {
    prodLabels = data[1][5];
  }

  return {
    success: true,
    keywords,
    labels: prodLabels,
    heroImages: [...new Set(heroImages)],
    specs: [...new Set(specs)].join('\n'),
    varian,
    gallery: [...new Set(gallery)],
    galleryVarian: varian,
    saranWarna: [...new Set(gallery)]
  };
}

function getAllLabels() {
  const ss = getSS();
  const labelSet = new Set(['Katalog Jualan']);
  
  // Ambil dari MASTER_PRODUK
  const mSheet = ss.getSheetByName('MASTER_PRODUK');
  if (mSheet) {
    const mData = mSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      if (mData[i][1]) labelSet.add(String(mData[i][1]).trim());
      if (mData[i][6]) {
        String(mData[i][6]).split(',').forEach(l => {
          if (l.trim()) labelSet.add(l.trim());
        });
      }
    }
  }
  
  // Ambil dari setiap sheet PRODUK_
  const sheets = ss.getSheets();
  sheets.forEach(s => {
    if (s.getName().startsWith('PRODUK_')) {
      const data = s.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][5]) {
          String(data[i][5]).split(',').forEach(l => {
            if (l.trim()) labelSet.add(l.trim());
          });
        }
      }
    }
  });
  
  return Array.from(labelSet).filter(Boolean);
}

function addKeyword(sheetName, keyword, label, spec) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet tidak ditemukan' };
  
  // Cek duplicate
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).toLowerCase().trim() === String(keyword).toLowerCase().trim()) {
      return { success: false, error: 'Keyword sudah ada' };
    }
  }
  
  sheet.appendRow([keyword, '', '', '', '', label || '', '', '', spec || '', '', '']);
  return { success: true };
}

function importKeywords(sheetName, keywordsArray) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false };
  
  // Ambil existing untuk cek duplicate
  const data = sheet.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) existing.add(String(data[i][0]).toLowerCase().trim());
  }
  
  let added = 0;
  keywordsArray.forEach(kw => {
    if (kw && String(kw).trim() && !existing.has(String(kw).toLowerCase().trim())) {
      sheet.appendRow([String(kw).trim(), '', '', '', '', '', '', '', '', '', '']);
      existing.add(String(kw).toLowerCase().trim());
      added++;
    }
  });
  
  return { success: true, count: added };
}

function deleteKeyword(sheetName, row) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false };
  sheet.deleteRow(row);
  return { success: true };
}

function updateKeyword(sheetName, row, data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false };
  
  const range = sheet.getRange(row, 1, 1, 11);
  range.setValues([[ 
    data.keyword, data.judul, data.hero, data.varian || data.gallery, data.keterangan,
    data.label, data.status, data.tglSelesai, data.spec, data.gallery || data.saranWarna, data.urlPost
  ]]);
  return { success: true };
}

function resetStatusGagal(sheetName) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false };
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'Gagal') {
      sheet.getRange(i + 1, 7).setValue('');
      count++;
    }
  }
  return { success: true, count };
}

// ==================== STATISTIK ====================

function getStats() {
  const ss = getSS();
  const config = getConfig();
  const produkSheet = ss.getSheetByName('MASTER_PRODUK');
  const totalProduk = produkSheet ? Math.max(0, produkSheet.getLastRow() - 1) : 0;
  
  let totalKeyword = 0, totalPosted = 0, postedToday = 0, queued = 0;
  const hariIni = new Date().toDateString();
  
  if (produkSheet) {
    const produkData = produkSheet.getDataRange().getValues();
    for (let i = 1; i < produkData.length; i++) {
      const sheetName = 'PRODUK_' + String(produkData[i][1]).replace(/\s+/g, '_').substring(0, 20);
      const pSheet = ss.getSheetByName(sheetName);
      if (pSheet) {
        const pData = pSheet.getDataRange().getValues();
        for (let j = 1; j < pData.length; j++) {
          totalKeyword++;
          if (pData[j][6] === 'Selesai') {
            totalPosted++;
            if (pData[j][7] instanceof Date && pData[j][7].toDateString() === hariIni) postedToday++;
          } else if (pData[j][0] && pData[j][6] !== 'Gagal') {
            queued++;
          }
        }
      }
    }
  }
  
  const triggerStatus = getTriggerStatus();
  const todayTargetInfo = getTodayTargetPosts();
  const now = new Date();
  const jam = now.getHours();
  const jamMulai = Number(config.JAM_AKTIF_MULAI) || 8;
  const jamSelesai = Number(config.JAM_AKTIF_SELESAI) || 22;
  
  return {
    success: true,
    totalProduk, totalKeyword, totalPosted, postedToday, queued,
    targetToday: todayTargetInfo.target,
    minPostPerDay: todayTargetInfo.minPost,
    maxPostPerDay: todayTargetInfo.maxPost,
    maxPerDay: todayTargetInfo.target,
    triggerAktif: triggerStatus.active,
    nextPostTime: triggerStatus.nextPostTime,
    lastInterval: triggerStatus.lastInterval,
    minMinutes: triggerStatus.minMinutes,
    maxMinutes: triggerStatus.maxMinutes,
    mode: triggerStatus.mode,
    jamMulai, jamSelesai,
    dalamJamAktif: jam >= jamMulai && jam < jamSelesai,
    jamSekarang: jam
  };
}

function getQueueStatus() {
  const ss = getSS();
  const produkSheet = ss.getSheetByName('MASTER_PRODUK');
  if (!produkSheet) return { success: true, produk: [] };
  
  const produkData = produkSheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < produkData.length; i++) {
    const nama = produkData[i][1];
    const sheetName = 'PRODUK_' + String(nama).replace(/\s+/g, '_').substring(0, 20);
    const pSheet = ss.getSheetByName(sheetName);
    
    let total = 0, selesai = 0, queue = 0;
    if (pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for (let j = 1; j < pData.length; j++) {
        if (pData[j][0]) {
          total++;
          if (pData[j][6] === 'Selesai') selesai++;
          else if (pData[j][6] !== 'Gagal') queue++;
        }
      }
    }
    
    result.push({
      id: produkData[i][0],
      nama,
      totalKeyword: total,
      posted: selesai,
      queue: queue
    });
  }
  
  return { success: true, produk: result };
}

function getStatsByDate(dateStr) {
  const targetDate = new Date(dateStr).toDateString();
  const logs = getLog(1000);
  const filtered = logs.filter(l => new Date(l.timestamp).toDateString() === targetDate);
  
  return {
    success: true,
    date: dateStr,
    total: filtered.length,
    sukses: filtered.filter(l => l.status.includes('Sukses')).length,
    gagal: filtered.filter(l => l.status.includes('Gagal')).length,
    data: filtered
  };
}

function getWeeklyStats(offset) {
  const now = new Date();
  const day = now.getDay();
  const diff = now.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(now.setDate(diff - (offset * 7)));
  monday.setHours(0,0,0,0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23,59,59,999);
  
  return getStatsRange(monday, sunday, 'week');
}

function getMonthlyStats(offset) {
  const now = new Date();
  const month = now.getMonth() - offset;
  const year = now.getFullYear() + Math.floor(month / 12);
  const adjMonth = ((month % 12) + 12) % 12;
  
  const firstDay = new Date(year, adjMonth, 1);
  const lastDay = new Date(year, adjMonth + 1, 0, 23, 59, 59);
  
  return getStatsRange(firstDay, lastDay, 'month');
}

function getStatsRange(start, end, type) {
  const logs = getLog(2000);
  const filtered = logs.filter(l => {
    const d = new Date(l.timestamp);
    return d >= start && d <= end;
  });
  
  const byDate = {};
  filtered.forEach(l => {
    const d = new Date(l.timestamp).toISOString().split('T')[0];
    if (!byDate[d]) byDate[d] = { sukses: 0, gagal: 0 };
    if (l.status.includes('Sukses')) byDate[d].sukses++;
    else byDate[d].gagal++;
  });
  
  return {
    success: true, type,
    start: start.toISOString(), end: end.toISOString(),
    total: filtered.length,
    sukses: filtered.filter(l => l.status.includes('Sukses')).length,
    gagal: filtered.filter(l => l.status.includes('Gagal')).length,
    byDate,
    data: filtered.slice(0, 100)
  };
}

// ==================== LOG ====================

function getLog(limit) {
  const ss = getSS();
  let sheet = ss.getSheetByName('LOG_POSTING');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const logs = [];
  const start = Math.max(1, data.length - limit);
  
  for (let i = data.length - 1; i >= start; i--) {
    logs.push({
      timestamp: data[i][0],
      produk: data[i][1],
      keyword: data[i][2],
      status: data[i][3],
      url: data[i][4],
      error: data[i][5] || ''
    });
  }
  return logs;
}

function getLogByDate(dateStr) {
  return getStatsByDate(dateStr);
}

function addLog(produk, keyword, status, url, error) {
  const ss = getSS();
  let sheet = ss.getSheetByName('LOG_POSTING');
  if (!sheet) {
    sheet = ss.insertSheet('LOG_POSTING');
    sheet.appendRow(['Timestamp', 'Produk', 'Keyword', 'Status', 'URL', 'Error']);
  }
  sheet.appendRow([new Date(), produk, keyword, status, url || '', error || '']);
}

// ==================== SMART QUEUE (ROTASI PRODUK) ====================

function getQueue() {
  const ss = getSS();
  const produkSheet = ss.getSheetByName('MASTER_PRODUK');
  if (!produkSheet) return [];
  
  const produkData = produkSheet.getDataRange().getValues();
  const queue = [];
  
  // Ambil log hari ini untuk cek produk mana yang sudah sering diposting
  const today = new Date().toDateString();
  const logs = getLog(200);
  const todayPosts = {};
  logs.forEach(l => {
    if (new Date(l.timestamp).toDateString() === today && l.status.includes('Sukses')) {
      todayPosts[l.produk] = (todayPosts[l.produk] || 0) + 1;
    }
  });
  
  for (let i = 1; i < produkData.length; i++) {
    if (produkData[i][3] === 'Nonaktif') continue;
    
    const produkNama = produkData[i][1];
    const sheetName = 'PRODUK_' + String(produkNama).replace(/\s+/g, '_').substring(0, 20);
    const pSheet = ss.getSheetByName(sheetName);
    if (!pSheet) continue;
    
    const pData = pSheet.getDataRange().getValues();
    for (let j = 1; j < pData.length; j++) {
      if (pData[j][0] && pData[j][6] !== 'Selesai' && pData[j][6] !== 'Gagal') {
        queue.push({
          produkId: produkData[i][0],
          produkNama: produkNama,
          sheetName: sheetName,
          row: j + 1,
          keyword: pData[j][0],
          judul: pData[j][1],
          keterangan: pData[j][4] || '',
          labelRaw: pData[j][5] || '',
          prioritas: produkData[i][4] || 5,
          todayCount: todayPosts[produkNama] || 0  // Untuk fair rotation
        });
      }
    }
  }
  
  // Sort: prioritas rendah dulu, lalu yang paling sedikit diposting hari ini (fair rotation)
  queue.sort((a, b) => {
    if (a.todayCount !== b.todayCount) return a.todayCount - b.todayCount;
    return a.prioritas - b.prioritas;
  });
  
  return queue;
}

// ==================== POSTING ENGINE ====================

// ==================== POSTING ENGINE ====================

function buatPostinganOtomatis() {
  const config = getConfig();
  const ss = getSS();
  
  const now = new Date();
  const jam = now.getHours();
  const jamMulai = Number(config.JAM_AKTIF_MULAI) || 8;
  const jamSelesai = Number(config.JAM_AKTIF_SELESAI) || 22;
  
  let currentItem = null;
  
  try {
    if (jam < jamMulai || jam >= jamSelesai) {
      console.log("Di luar jam aktif (" + jam + ":00). Jam aktif: " + jamMulai + ":00 - " + jamSelesai + ":00");
      return;
    }
    
    const todayTargetInfo = getTodayTargetPosts();
    const stats = getStats();
    if (stats.postedToday >= todayTargetInfo.target) {
      console.log(`Target kuota posting acak hari ini (${todayTargetInfo.target} post) sudah tercapai: ${stats.postedToday}/${todayTargetInfo.target}`);
      return;
    }
    
    const queue = getQueue();
    if (queue.length === 0) {
      console.log("Queue kosong, tidak ada keyword untuk diposting");
      return;
    }
    
    currentItem = queue[0];
    const item = currentItem;
    console.log("Memulai posting otomatis: " + item.keyword + " (" + item.produkNama + ")");
    
    const pData = getProdukData(item.sheetName);
    if (!pData.success || pData.heroImages.length === 0) {
      addLog(item.produkNama, item.keyword, 'Gagal', '', 'Foto Hero kosong');
      return;
    }
    
    let labelRaw = item.labelRaw;
    if (!labelRaw) {
      const sheet = ss.getSheetByName(item.sheetName);
      const allData = sheet.getDataRange().getValues();
      for (let prev = item.row - 2; prev >= 1; prev--) {
        if (allData[prev][5] && String(allData[prev][5]).trim()) {
          labelRaw = allData[prev][5];
          break;
        }
      }
    }
    if (!labelRaw && pData.labels) {
      labelRaw = pData.labels;
    }
    const labels = labelRaw ? String(labelRaw).split(',').map(l => l.trim()).filter(Boolean) : [item.produkNama, 'Katalog Jualan'];
    
    const heroLink = pData.heroImages[Math.floor(Math.random() * pData.heroImages.length)];
    const heroThumb = getImageUrlFromLink(heroLink);
    
    const aiData = getAIDescription(item.produkNama, item.keyword, item.keterangan, pData.specs);
    if (!aiData) {
      addLog(item.produkNama, item.keyword, 'Gagal', '', 'AI gagal membuat deskripsi');
      return;
    }
    
    const finalTitle = item.judul || aiData.title;
    const finalHtml = buildModernTemplate({
      heroImg: heroThumb, heroOriginal: heroLink,
      kataKunci: item.keyword, deskripsiAi: aiData.description,
      tabelAi: aiData.table, gallery: pData.gallery || pData.saranWarna, varian: pData.varian || pData.galleryVarian
    });
    
    const resBlogger = postKeBlogger(finalTitle, finalHtml, labels);
    
    if (resBlogger && resBlogger.success) {
      const sheet = ss.getSheetByName(item.sheetName);
      sheet.getRange(item.row, 2).setValue(finalTitle);
      sheet.getRange(item.row, 7).setValue('Selesai');
      sheet.getRange(item.row, 8).setValue(new Date());
      if (resBlogger.url) sheet.getRange(item.row, 11).setValue(resBlogger.url);
      addLog(item.produkNama, item.keyword, 'Sukses', resBlogger.url || '', '');
      console.log("BERHASIL POST: " + finalTitle);
    } else {
      const err = resBlogger ? (resBlogger.error || 'Blogger error') : 'Blogger error';
      addLog(item.produkNama, item.keyword, 'Gagal', '', err);
      console.error("GAGAL POST: " + err);
    }
  } catch(e) {
    console.error("Exception buatPostinganOtomatis: " + e);
    addLog(currentItem ? currentItem.produkNama : '-', currentItem ? currentItem.keyword : '-', 'Gagal', '', e.toString());
  } finally {
    // Jadwalkan postingan acak berikutnya secara otomatis jika mode auto-posting aktif
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('TRIGGER_MODE') === 'random') {
      jadwalkanPostinganBerikutnya();
    }
  }
}

function postManual(produkNama, keyword, row) {
  const config = getConfig();
  const ss = getSS();
  const sheetName = 'PRODUK_' + produkNama.replace(/\s+/g, '_').substring(0, 20);
  
  const pData = getProdukData(sheetName);
  if (!pData || !pData.success) {
    return { success: false, error: 'Sheet produk ' + sheetName + ' tidak ditemukan' };
  }

  const sheet = ss.getSheetByName(sheetName);
  const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
  
  let labelRaw = rowData[5];
  if (!labelRaw) {
    const allData = sheet.getDataRange().getValues();
    for (let prev = row - 2; prev >= 1; prev--) {
      if (allData[prev][5] && String(allData[prev][5]).trim()) {
        labelRaw = allData[prev][5];
        break;
      }
    }
  }
  if (!labelRaw && pData.labels) {
    labelRaw = pData.labels;
  }
  const labels = labelRaw ? String(labelRaw).split(',').map(l => l.trim()).filter(Boolean) : [produkNama, 'Katalog Jualan'];
  
  const heroLink = (pData.heroImages && pData.heroImages.length > 0)
    ? pData.heroImages[Math.floor(Math.random() * pData.heroImages.length)]
    : '';
  const heroThumb = getImageUrlFromLink(heroLink);
  
  const aiData = getAIDescription(produkNama, keyword, rowData[4] || '', pData.specs);
  if (!aiData) return { success: false, error: 'AI gagal membuat deskripsi' };
  
  const finalTitle = rowData[1] || aiData.title;
  const finalHtml = buildModernTemplate({
    heroImg: heroThumb, heroOriginal: heroLink,
    kataKunci: keyword, deskripsiAi: aiData.description,
    tabelAi: aiData.table, gallery: pData.gallery || pData.saranWarna, varian: pData.varian || pData.galleryVarian
  });
  
  const resBlogger = postKeBlogger(finalTitle, finalHtml, labels);
  
  if (resBlogger && resBlogger.success) {
    sheet.getRange(row, 2).setValue(finalTitle);
    sheet.getRange(row, 7).setValue('Selesai');
    sheet.getRange(row, 8).setValue(new Date());
    if (resBlogger.url) sheet.getRange(row, 11).setValue(resBlogger.url);
    addLog(produkNama, keyword, 'Sukses (Manual)', resBlogger.url || '', '');
    return { success: true, title: finalTitle, url: resBlogger.url };
  }
  
  const errDetail = resBlogger ? (resBlogger.error || 'Blogger error') : 'Gagal post ke Blogger';
  addLog(produkNama, keyword, 'Gagal (Manual)', '', errDetail);
  return { success: false, error: errDetail };
}

function previewPost(produkNama, keyword, row) {
  const ss = getSS();
  const sheetName = 'PRODUK_' + produkNama.replace(/\s+/g, '_').substring(0, 20);
  
  const pData = getProdukData(sheetName);
  if (!pData.success) return pData;
  
  const sheet = ss.getSheetByName(sheetName);
  const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
  
  const heroLink = pData.heroImages.length > 0 
    ? pData.heroImages[Math.floor(Math.random() * pData.heroImages.length)] 
    : '';
  const heroThumb = getImageUrlFromLink(heroLink);
  
  const aiData = getAIDescription(produkNama, keyword, rowData[4] || '', pData.specs);
  if (!aiData) return { success: false, error: 'AI gagal' };
  
  const finalHtml = buildModernTemplate({
    heroImg: heroThumb, heroOriginal: heroLink,
    kataKunci: keyword, deskripsiAi: aiData.description,
    tabelAi: aiData.table, gallery: pData.gallery || pData.saranWarna, varian: pData.varian || pData.galleryVarian
  });
  
  return { success: true, html: finalHtml, title: aiData.title };
}

// ==================== TRIGGER JEDA ACAK (RANDOM INTERVAL) ====================

function hapusTriggerPosting() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'buatPostinganOtomatis') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function aktifkanTrigger(minMinutes, maxMinutes) {
  const config = getConfig();
  
  const min = Math.max(5, Number(minMinutes || config.MIN_DELAY_MINUTES || 25));
  const max = Math.max(min, Number(maxMinutes || config.MAX_DELAY_MINUTES || 45));
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TRIGGER_MODE', 'random');
  props.setProperty('MIN_DELAY_MINUTES', String(min));
  props.setProperty('MAX_DELAY_MINUTES', String(max));
  
  // Hapus trigger buatPostinganOtomatis lama
  hapusTriggerPosting();
  
  // Jadwalkan postingan pertama (1 - 2 menit dari sekarang agar segera mulai)
  const firstDelay = Math.floor(Math.random() * 2) + 1;
  const now = new Date();
  const nextRun = new Date(now.getTime() + firstDelay * 60 * 1000);
  
  ScriptApp.newTrigger('buatPostinganOtomatis')
    .timeBased()
    .at(nextRun)
    .create();
    
  props.setProperty('NEXT_POST_TIME', nextRun.toISOString());
  props.setProperty('LAST_TRIGGER_INTERVAL', String(firstDelay));
  
  // Pastikan trigger refresh keyword harian jam 00:00 aktif
  aktifkanRefreshKeywordHarian();
  
  const timeStr = Utilities.formatDate(nextRun, Session.getScriptTimeZone(), 'HH:mm');
  return {
    success: true,
    message: `Auto-Posting Acak AKTIF! Jeda ${min}-${max} menit. Postingan pertama dijadwalkan pukul ${timeStr} WIB.`,
    nextRun: nextRun.toISOString(),
    min,
    max
  };
}

function jadwalkanPostinganBerikutnya() {
  hapusTriggerPosting();
  
  const props = PropertiesService.getScriptProperties();
  const mode = props.getProperty('TRIGGER_MODE');
  if (mode !== 'random') return; // Jika trigger sedang dimatikan, jangan jadwalkan
  
  const config = getConfig();
  const min = Math.max(5, Number(props.getProperty('MIN_DELAY_MINUTES') || config.MIN_DELAY_MINUTES || 25));
  const max = Math.max(min, Number(props.getProperty('MAX_DELAY_MINUTES') || config.MAX_DELAY_MINUTES || 45));
  
  // Interval acak antara min dan max
  const randomMinutes = Math.floor(Math.random() * (max - min + 1)) + min;
  
  const now = new Date();
  let nextRun = new Date(now.getTime() + randomMinutes * 60 * 1000);
  
  const jamMulai = Number(config.JAM_AKTIF_MULAI) || 8;
  const jamSelesai = Number(config.JAM_AKTIF_SELESAI) || 22;
  const targetHour = nextRun.getHours();
  
  const todayTargetInfo = getTodayTargetPosts();
  const stats = getStats();
  const targetReached = stats.postedToday >= todayTargetInfo.target;
  
  // Jika di luar jam kerja ATAU kuota hari ini sudah tercapai, jadwalkan besok pagi pada jamMulai + random 5-20 menit
  if (targetHour >= jamSelesai || targetHour < jamMulai || targetReached) {
    const tomorrow = new Date(now);
    if (targetHour >= jamSelesai || targetReached) {
      tomorrow.setDate(tomorrow.getDate() + 1);
    }
    const morningOffset = Math.floor(Math.random() * 20) + 5; // 08:05 - 08:25
    tomorrow.setHours(jamMulai, morningOffset, 0, 0);
    nextRun = tomorrow;
  }
  
  ScriptApp.newTrigger('buatPostinganOtomatis')
    .timeBased()
    .at(nextRun)
    .create();
    
  props.setProperty('NEXT_POST_TIME', nextRun.toISOString());
  props.setProperty('LAST_TRIGGER_INTERVAL', String(randomMinutes));
  
  console.log(`Jadwal posting acak berikutnya: ${Utilities.formatDate(nextRun, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')} (jeda ${randomMinutes} menit)`);
}

function nonaktifkanTrigger() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TRIGGER_MODE', 'off');
  props.deleteProperty('NEXT_POST_TIME');
  
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'buatPostinganOtomatis' || t.getHandlerFunction() === 'refreshAllKeywords') {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { success: true, message: 'Semua trigger auto-posting telah dinonaktifkan' };
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasPostingTrigger = triggers.some(t => t.getHandlerFunction() === 'buatPostinganOtomatis');
  const props = PropertiesService.getScriptProperties();
  const mode = props.getProperty('TRIGGER_MODE') || (hasPostingTrigger ? 'random' : 'off');
  const nextPostTime = props.getProperty('NEXT_POST_TIME');
  const lastInterval = props.getProperty('LAST_TRIGGER_INTERVAL');
  const config = getConfig();
  const min = props.getProperty('MIN_DELAY_MINUTES') || config.MIN_DELAY_MINUTES || 25;
  const max = props.getProperty('MAX_DELAY_MINUTES') || config.MAX_DELAY_MINUTES || 45;
  
  const isActive = hasPostingTrigger && mode !== 'off';
  
  return {
    success: true,
    active: isActive,
    mode: mode,
    nextPostTime: isActive ? nextPostTime : null,
    lastInterval: isActive ? Number(lastInterval) : null,
    minMinutes: Number(min),
    maxMinutes: Number(max)
  };
}

function testPosting() {
  buatPostinganOtomatis();
  return { success: true };
}

// ==================== DRIVE IMAGES ====================

function getDriveImages(folderId) {
  try {
    if (!folderId) return { success: false, error: 'Folder ID kosong' };
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const images = [];
    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      if (mime.indexOf('image') !== -1 || mime.indexOf('video') !== -1) {
        images.push({
          id: file.getId(),
          name: file.getName(),
          mimeType: mime,
          url: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
          thumb: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400'
        });
      }
    }
    return { success: true, images };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== AI & HELPER ====================

function getAIDescription(nama, kw, tipe, spec) {
  const config = getConfig();
  const apiKey = getAIKeyRotation();
  if (!apiKey) {
    console.error('AI Error: Tidak ada Groq API Key yang aktif.');
    return null;
  }
  
  const url = 'https://api.groq.com/openai/v1/chat/completions';
  const model = (config.GROQ_MODEL && config.GROQ_MODEL.toLowerCase() !== 'groq' && config.GROQ_MODEL.includes('/')) 
    ? config.GROQ_MODEL 
    : 'openai/gpt-oss-120b';
  
  const prompt = `Buat konten penawaran produk profesional.

PRODUK UTAMA: ${nama}
KATA KUNCI/JUDUL: ${kw}
VARIASI TIPE: ${tipe}
DATA SPESIFIKASI ASLI:
${spec || 'Hubungi admin untuk detail spesifikasi.'}

TUGAS:
1. Artikel 3-4 paragraf HTML persuasif, SEO-friendly, bahasa Indonesia. Fokus 100% pada produk "${nama}" dan kata kunci "${kw}".
2. Tabel HTML spesifikasi, header #4285f4 teks putih.
3. JANGAN mengarang data spesifikasi di luar data asli, dan DILARANG menyebut produk/material jenis lain yang berbeda kategori.
4. Judul SEO: kata kunci menarik | produk | kota Indonesia acak | ${config.WHATSAPP_NUMBER}

OUTPUT JSON:
{"title":"...","description":"...","table":"..."}`;

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: model,
        messages: [
          { role: 'system', content: 'Kamu penulis konten marketing profesional Indonesia. Return JSON valid.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.7, max_tokens: 2048,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    });
    
    const text = res.getContentText();
    const json = JSON.parse(text);
    if (!json.choices || !json.choices[0]?.message?.content) {
      console.error('Groq API Error (' + res.getResponseCode() + '): ' + text);
      return null;
    }
    return JSON.parse(json.choices[0].message.content);
  } catch(e) {
    console.error('AI Error: ' + e);
    return null;
  }
}

function buildModernTemplate(d) {
  const config = getConfig();
  let html = `<div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#333;max-width:800px;margin:auto;line-height:1.6;">`;
  
  if (d.heroImg) {
    html += `<div style="text-align:center;margin-bottom:20px;">
      <a href="${d.heroOriginal}" target="_blank"><img src="${d.heroImg}" alt="${d.kataKunci}" style="width:100%;max-width:650px;border-radius:15px;box-shadow:0 10px 20px rgba(0,0,0,0.1);" /></a>
    </div>`;
  }

  html += `<div style="font-size:17px;margin-bottom:35px;text-align:justify;">${d.deskripsiAi}</div>`;

  const galleryList = d.gallery || d.saranWarna || [];
  if (galleryList && galleryList.length > 0) {
    html += `<div style="background:#f8f9fa;border:2px dashed #1a73e8;border-radius:15px;padding:25px;margin:35px 0;">
      <h3 style="margin-top:0;color:#1a73e8;">🖼️ Gallery Produk</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:15px;">`;
    galleryList.forEach(link => {
      const thumb = getImageUrlFromLink(link);
      if (thumb) html += `<div style="background:white;padding:5px;border-radius:10px;border:1px solid #e0e0e0;"><a href="${link}" target="_blank"><img src="${thumb}" style="width:100%;height:130px;object-fit:cover;border-radius:8px;" /></a></div>`;
    });
    html += `</div></div>`;
  }

  html += `<div style="margin:40px 0;">
    <h3 style="border-left:6px solid #1a73e8;padding-left:15px;color:#1a73e8;">Spesifikasi Produk</h3>
    <div style="overflow-x:auto;margin-top:15px;">${d.tabelAi}</div>
  </div>`;

  const varianList = d.varian || d.galleryVarian || [];
  if (varianList && varianList.length > 0) {
    html += `<div style="margin:40px 0;">
      <h3 style="border-left:6px solid #1a73e8;padding-left:15px;color:#1a73e8;">🎨 Varian Produk</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:20px;margin-top:20px;">`;
    varianList.forEach(item => {
      const thumb = getImageUrlFromLink(item.link);
      if (thumb) html += `<div style="border:1px solid #eee;padding:10px;border-radius:12px;background:#fff;text-align:center;"><a href="${item.link}" target="_blank"><img src="${thumb}" style="width:100%;height:180px;object-fit:cover;border-radius:8px;" /></a>${item.ket ? `<p style="margin:10px 0 0;font-size:13px;font-weight:bold;color:#555;">${item.ket}</p>` : ''}</div>`;
    });
    html += `</div></div>`;
  }

  html += `<div style="background:linear-gradient(135deg,#4285f4 0%,#1a73e8 100%);color:white;border-radius:15px;padding:40px 20px;text-align:center;margin-top:50px;">
    <h2 style="margin-top:0;">Pesan Sekarang</h2>
    <p style="font-size:16px;opacity:0.9;">Hubungi tim ahli kami untuk penawaran terbaik.</p>
    <a href="https://wa.me/${String(config.WHATSAPP_NUMBER).replace(/^0/,'62')}" style="background:#fff;color:#1a73e8;padding:18px 40px;text-decoration:none;border-radius:50px;font-weight:bold;font-size:18px;display:inline-block;margin-top:20px;box-shadow:0 4px 15px rgba(0,0,0,0.2);">💬 Hubungi WhatsApp: ${config.WHATSAPP_NUMBER}</a>
    <p style="margin-top:15px;font-size:13px;opacity:0.8;">${config.COMPANY_NAME} - Layanan Pengiriman Seluruh Indonesia</p>
  </div></div>`;
  
  return html;
}

function postKeBlogger(title, content, labels) {
  const config = getConfig();
  const blogId = String(config.BLOG_ID).trim();
  if (!blogId) {
    return { success: false, error: 'BLOG_ID belum diisi di Pengaturan' };
  }

  // 1. Jika Advanced Service Blogger diaktifkan di Apps Script Services
  if (typeof Blogger !== 'undefined' && Blogger.Posts && Blogger.Posts.insert) {
    try {
      const res = Blogger.Posts.insert({
        title: title || 'Untitled',
        content: content || '<p>No content</p>',
        labels: Array.isArray(labels) ? labels : ['Katalog']
      }, blogId);
      if (res && (res.id || res.url)) {
        return { success: true, url: res.url || res.selfLink, id: res.id };
      }
    } catch(e) {
      console.warn('Blogger Advanced Service gagal: ' + e + ', mencoba via REST API UrlFetchApp...');
    }
  }

  // 2. Menggunakan REST API via UrlFetchApp
  const url = `https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts/`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({
        title: title || 'Untitled',
        content: content || '<p>No content</p>',
        labels: Array.isArray(labels) ? labels : ['Katalog']
      }),
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    const text = res.getContentText();
    if (code >= 300) {
      console.error('Blogger API Error ' + code + ': ' + text);
      let errMsg = 'Blogger Error (' + code + ')';
      try {
        const errJson = JSON.parse(text);
        if (errJson.error && errJson.error.message) {
          errMsg += ': ' + errJson.error.message;
          if (code === 401 || code === 403) {
            errMsg += '. Izin akses Blogger belum diaktifkan di Google Apps Script.';
          }
        }
      } catch(e) {
        errMsg += ': ' + text.substring(0, 150);
      }
      return { success: false, error: errMsg };
    }
    
    const json = JSON.parse(text);
    return { success: true, url: json.url || json.selfLink, id: json.id };
  } catch(e) {
    console.error('postKeBlogger Exception: ' + e);
    return { success: false, error: 'Gagal koneksi Blogger: ' + e.toString() };
  }
}

function parseLinks(raw) {
  if (!raw) return [];
  return String(raw).split(/[\n,]/).map(l => l.trim()).filter(l => l.startsWith('http'));
}

function getImageUrlFromLink(url) {
  if (!url) return null;
  url = String(url).trim();
  if (url.includes('googleusercontent.com') || url.includes('blogger.com') || url.includes('bp.blogspot.com') || url.includes('cloudinary.com') || url.includes('imgbb.com') || url.includes('imgur.com')) {
    return url;
  }
  if (url.includes('drive.google.com')) {
    const match = url.match(/[-\w]{25,}/);
    return match ? `https://lh3.googleusercontent.com/d/${match[0]}=s1200` : url;
  }
  return url;
}

function uploadImage(name, type, base64Data, folderId) {
  try {
    const config = getConfig();
    let folder = null;
    const targetFolderId = folderId || config.DRIVE_FOLDER_ID;
    
    if (targetFolderId) {
      try {
        folder = DriveApp.getFolderById(targetFolderId);
      } catch(e) {
        console.warn("Folder ID tidak valid: " + e);
      }
    }
    
    if (!folder) {
      const folders = DriveApp.getFoldersByName("AutoPosting_Uploads");
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder("AutoPosting_Uploads");
      }
    }
    
    // Set sharing folder ke publik agar gambar selalu bisa diakses publik
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}
    
    const cleanBase64 = String(base64Data || '').replace(/^data:image\/[a-zA-Z0-9+]+;base64,/, '');
    const decoded = Utilities.base64Decode(cleanBase64);
    const fileName = name || ('img_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg');
    const blob = Utilities.newBlob(decoded, type || 'image/jpeg', fileName);
    
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}
    
    const fileId = file.getId();
    // Google UserContent CDN link langsung (resmi, cepat, tanpa batasan viewer)
    const directUrl = `https://lh3.googleusercontent.com/d/${fileId}=s1200`;
    const viewUrl = file.getUrl();
    
    return {
      success: true,
      fileId: fileId,
      url: directUrl,
      viewUrl: viewUrl,
      name: fileName
    };
  } catch(err) {
    console.error("Gagal upload foto ke Drive: " + err);
    return {
      success: false,
      error: "Gagal upload ke Google Drive: " + err.toString()
    };
  }
}

function authorizeBlogger() {
  const config = getConfig();
  const blogId = String(config.BLOG_ID || '5323802325028284266').trim();
  console.log("=== MEMULAI TEST OTORISASI BLOGGER ===");
  console.log("Blog ID: " + blogId);
  
  // 1. Tes melalui Advanced Blogger Service (jika diaktifkan di menu Services +)
  if (typeof Blogger !== 'undefined' && Blogger.Blogs && Blogger.Blogs.get) {
    try {
      const blog = Blogger.Blogs.get(blogId);
      console.log("SUKSES: Terhubung via Blogger Advanced Service!");
      console.log("Nama Blog: " + (blog ? blog.name : 'OK'));
      console.log("URL Blog: " + (blog ? blog.url : 'OK'));
      return { success: true, via: 'Advanced Service', blogName: blog ? blog.name : 'OK' };
    } catch(e) {
      console.warn("Blogger Advanced Service: " + e);
    }
  }
  
  // 2. Tes melalui UrlFetchApp REST API
  const url = `https://www.googleapis.com/blogger/v3/blogs/${blogId}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  
  const code = res.getResponseCode();
  const text = res.getContentText();
  console.log("HTTP Response Code: " + code);
  
  if (code === 200) {
    const json = JSON.parse(text);
    console.log("SUKSES: Terhubung via REST API ke Blog: " + json.name + " (" + json.url + ")");
    return { success: true, via: 'REST API', blogName: json.name, blogUrl: json.url };
  } else {
    console.error("GAGAL: " + text);
    return { success: false, code: code, text: text };
  }
}

function testDriveAuth() {
  console.log("=== MEMULAI TEST OTORISASI GOOGLE DRIVE ===");
  try {
    const folders = DriveApp.getFoldersByName("AutoPosting_Uploads");
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("AutoPosting_Uploads");
    }
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Buat file tes untuk memicu popup otorisasi tulis Google Drive
    const blob = Utilities.newBlob("Drive Upload Auth Test", "text/plain", "test_drive_auth.txt");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    file.setTrashed(true); // langsung hapus file tes
    
    console.log("SUKSES: Google Drive terotorisasi dan siap mengupload foto!");
    return { success: true, folderId: folder.getId() };
  } catch(e) {
    console.error("GAGAL OTORISASI DRIVE: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

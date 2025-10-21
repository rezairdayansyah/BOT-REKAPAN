require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Inline parser function (previously in lib/parseAktivasi.js)
function parseAktivasi(text, userRow, username) {
  const lines = text.split('\n').map(l=>l.trim()).filter(l=>l);
  const upper = text.toUpperCase();
  let ao='', workorder='', serviceNo='', customerName='', owner='', workzone='', snOnt='', nikOnt='', stbId='', nikStb='', teknisi='';
  teknisi = (userRow && userRow[1]) ? (userRow[1] || username).replace('@', '') : (username || '');

  function findValue(patterns) {
    for (const pattern of patterns) {
      const matches = text.match(pattern);
      if (matches) {
        if (pattern.global) {
          return matches[matches.length - 1];
        } else if (matches[1]) {
          return matches[1].trim();
        }
      }
    }
    return '';
  }

  function detectOwner(text) {
    const upperText = text.toUpperCase();
    // Prioritize BGES indicators first — some BGES messages include 'DATE CREATED' which previously matched TSEL
    if (upperText.includes('INDIBIZ') || upperText.includes('HSI') || upperText.includes('SC ORDER') || upperText.includes('NCLI') || upperText.includes('MITRA') || upperText.includes('CHANNEL : BS')) {
      return 'BGES';
    }
    // TSEL-specific markers
    if (upperText.includes('CHANNEL : DIGIPOS') || upperText.includes('WORKORDER : WO')) {
      return 'TSEL';
    }
    if (upperText.includes('WMS') || upperText.includes('MWS')) {
      return 'WMS';
    }
    return '';
  }

  owner = detectOwner(text);

  if (owner === 'TSEL') {
    ao = findValue([
      /AO\s*:\s*([A-Za-z0-9]+)/i,
      /AO\s*([A-Za-z0-9]+)/i
    ]);
    workorder = findValue([
      /WORKORDER\s*:\s*([A-Za-z0-9]+)/i
    ]) || ao;
    serviceNo = findValue([
      /SERVICE\s*NO\s*:\s*(\d+)/i
    ]);
    customerName = findValue([
      /CUSTOMER\s*NAME\s*:\s*([A-Z0-9\s]+)/i
    ]);
    workzone = findValue([
      /WORKZONE\s*:\s*([A-Z0-9]+)/i
    ]);
    snOnt = findValue([
      /SN\s*ONT\s*:\s*([A-Z0-9]+)/i,
      /(ZTEGDA[A-Z0-9]+)/i,
      /(HWTC[A-Z0-9]+)/i,
      /(HUAW[A-Z0-9]+)/i,
      /(FHTT[A-Z0-9]+)/i,
      /(FIBR[A-Z0-9]+)/i
    ]);
    nikOnt = findValue([/NIK\s*ONT\s*:\s*(\d+)/i]);
    const stbPattern = /STB\s*ID\s*:\s*([A-Z0-9]+)/i;
    const stbMatch = text.match(stbPattern);
    if (stbMatch && stbMatch[1]) {
      stbId = stbMatch[1];
      nikStb = findValue([/NIK\s*STB\s*:\s*(\d+)/i]);
    }
  } else if (owner === 'BGES' || owner === 'WMS') {
    // BGES/WMS: support many variants from BS/INDIBIZ exports
    // SC ORDER NO -> AO/WORKORDER (prefix with SC if needed)
    const scOrderMatch = text.match(/SC\s*ORDER\s*NO\s*:\s*([0-9A-Za-z\-]+)/i);
    if (scOrderMatch && scOrderMatch[1]) {
      const val = scOrderMatch[1].trim();
      ao = /^SC/i.test(val) ? val : `SC${val}`;
      workorder = ao;
    } else {
      // fallback: try to find SCNNNNNN in AO| lines
      const aoMatches = text.match(/AO\|.*?(SC\d{6,})/g);
      if (aoMatches && aoMatches.length > 0) {
        const lastMatch = aoMatches[aoMatches.length - 1];
        const scMatch = lastMatch.match(/SC(\d{6,})/);
        if (scMatch) {
          ao = `SC${scMatch[1]}`;
          workorder = ao;
        }
      }
    }

    // SERVICE NO: explicit or fallback to long numeric token (11-12 digits)
    const serviceMatch = text.match(/SERVICE\s*NO\s*:\s*([0-9A-Za-z\-]+)/i);
    if (serviceMatch && serviceMatch[1]) {
      serviceNo = serviceMatch[1].trim();
    } else {
      const serviceNoMatches = text.match(/\b\d{11,12}\b/g);
      if (serviceNoMatches && serviceNoMatches.length > 0) {
        serviceNo = serviceNoMatches[serviceNoMatches.length - 1];
      }
    }

    // NCLI
    const ncliMatch = text.match(/NCLI\s*:\s*([0-9A-Za-z\-]+)/i);
    const ncli = ncliMatch && ncliMatch[1] ? ncliMatch[1].trim() : '';

    // CUSTOMER NAME (prefer explicit field)
    const custMatch = text.match(/CUSTOMER\s*NAME\s*:\s*(.+)/i);
    if (custMatch && custMatch[1]) {
      customerName = custMatch[1].trim();
    } else {
      // sometimes embedded in timestamped logs
      const customerMatches = text.match(/\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}\s+\d+\s+([A-Z0-9\s]+?)\s{2,}/g);
      if (customerMatches && customerMatches.length > 0) {
        const nameMatch = customerMatches[0].match(/\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}\s+\d+\s+([A-Z0-9\s]+?)\s{2,}/);
        if (nameMatch && nameMatch[1]) {
          customerName = nameMatch[1].trim();
        }
      }
    }

    // WORKZONE and ODP
    const wzMatch = text.match(/WORKZONE\s*:\s*([A-Z0-9\-\/]+)/i);
    if (wzMatch && wzMatch[1]) {
      workzone = wzMatch[1].trim();
    } else {
      const workzoneMatches = text.match(/AO\|\s+([A-Z]{2,})/g);
      if (workzoneMatches && workzoneMatches.length > 0) {
        const lastWorkzoneMatch = workzoneMatches[workzoneMatches.length - 1];
        const wz2 = lastWorkzoneMatch.match(/AO\|\s+([A-Z]{2,})/);
        if (wz2 && wz2[1]) {
          workzone = wz2[1];
        }
      }
    }
    const odpMatch = text.match(/ODP\s*:\s*([A-Z0-9\-\/]+)/i);
    if (odpMatch && odpMatch[1] && !workzone) {
      workzone = odpMatch[1].trim();
    }

    // ADDRESS, CONTACT PHONE, PAKET, MITRA (partner)
    const addressMatch = text.match(/ADDRESS\s*:\s*(.+)/i);
    const contactMatch = text.match(/CONTACT\s*PHONE\s*:\s*([0-9\+\-\s]+)/i);
    const paketMatch = text.match(/PAKET\s*:\s*(.+)/i);
    const mitraMatch = text.match(/MITRA\s*:\s*(.+)/i);
    const mitra = mitraMatch && mitraMatch[1] ? mitraMatch[1].trim() : '';

    // For BGES messages we want OWNER to be 'BS' (the channel),
    // keep MITRA in a separate variable but do not overwrite OWNER.
    if (owner === 'BGES') {
      owner = 'BS';
    }

    // SN / NIK / STB
    snOnt = findValue([
      /SN\s*ONT[:\s]+([A-Z0-9]+)/i,
      /(ZTEG[A-Z0-9]+)/i,
      /(HWTC[A-Z0-9]+)/i,
      /(HUAW[A-Z0-9]+)/i,
      /(FHTT[A-Z0-9]+)/i,
      /(FHTTC[A-Z0-9]+)/i,
      /(FIBR[A-Z0-9]+)/i
    ]);
    // also try a generic SN pattern near 'SN ONT'
    if (!snOnt) {
      const snNear = text.match(/SN\s*ONT\s*[:\s]*([A-Z0-9\-]+)/i);
      if (snNear && snNear[1]) snOnt = snNear[1].trim();
    }

    nikOnt = findValue([/NIK\s*ONT[:\s]+(\d+)/i]);
    stbId = findValue([/STB\s*ID[:\s]+([A-Z0-9]+)/i]);
    nikStb = findValue([/NIK\s*STB[:\s]+(\d+)/i]);

    // If customer name still empty, fallback to first non-empty ADDRESS segment or MITRA
    if (!customerName) {
      if (addressMatch && addressMatch[1]) {
        // Address might contain pipes - take first segment as fallback name attempt
        const firstSeg = addressMatch[1].split('|')[0].trim();
        if (firstSeg && firstSeg.match(/[A-Z]/i)) customerName = firstSeg;
      }
      if (!customerName && mitra) customerName = mitra;
    }
  } else {
    function getValue(label) {
      const line = lines.find(l => l.toUpperCase().startsWith(label.toUpperCase() + ' :'));
      return line ? line.split(':').slice(1).join(':').trim() : '';
    }
    ao = getValue('AO') || findValue([/AO[:\s]+([A-Z0-9]+)/i]);
    workorder = getValue('WORKORDER') || findValue([/WORKORDER[:\s]+([A-Z0-9-]+)/i]);
    serviceNo = getValue('SERVICE NO') || findValue([/SERVICE\s*NO[:\s]+(\d+)/i]);
    customerName = getValue('CUSTOMER NAME') || findValue([/CUSTOMER\s*NAME[:\s]+(.+)/i]);
    owner = getValue('OWNER') || findValue([/OWNER[:\s]+([A-Z0-9]+)/i]);
    workzone = getValue('WORKZONE') || findValue([/WORKZONE[:\s]+([A-Z0-9]+)/i]);
    snOnt = getValue('SN ONT') || findValue([
      /SN\s*ONT[:\s]+([A-Z0-9]+)/i,
      /(ZTEG[A-Z0-9]+)/i,
      /(HWTC[A-Z0-9]+)/i,
      /(HUAW[A-Z0-9]+)/i,
      /(FHTT[A-Z0-9]+)/i,
      /(FIBR[A-Z0-9]+)/i
    ]);
    nikOnt = getValue('NIK ONT') || findValue([/NIK\s*ONT[:\s]+(\d+)/i]);
    stbId = getValue('STB ID') || findValue([/STB\s*ID[:\s]+([A-Z0-9]+)/i]);
    nikStb = getValue('NIK STB') || findValue([/NIK\s*STB[:\s]+(\d+)/i]);
  }

  return { ao, workorder, serviceNo, customerName, owner, workzone, snOnt, nikOnt, stbId, nikStb, teknisi };
}

// === Konfigurasi dari environment variables ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

// Validasi environment variables
if (!TOKEN) {
  console.error('ERROR: TELEGRAM_TOKEN environment variable is not set!');
  process.exit(1);
}
if (!SHEET_ID) {
  console.error('ERROR: SHEET_ID environment variable is not set!');
  process.exit(1);
}
if (!GOOGLE_SERVICE_ACCOUNT_KEY) {
  console.error('ERROR: GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set!');
  process.exit(1);
}

const REKAPAN_SHEET = 'REKAPAN QUALITY';
const USER_SHEET = 'USER';

// === Setup Google Sheets API ===
let serviceAccount;
try {
  // Handle both direct JSON and base64 encoded
  let keyData = GOOGLE_SERVICE_ACCOUNT_KEY;
  
  // Check if it's base64 encoded
  if (!keyData.startsWith('{')) {
    try {
      keyData = Buffer.from(keyData, 'base64').toString('utf-8');
    } catch (e) {
      console.log('Not base64 encoded, using as is');
    }
  }
  
  serviceAccount = JSON.parse(keyData);
  console.log('Google Service Account parsed successfully');
} catch (e) {
  console.error('ERROR parsing GOOGLE_SERVICE_ACCOUNT_KEY:', e.message);
  console.error('First 100 chars of key:', GOOGLE_SERVICE_ACCOUNT_KEY.substring(0, 100));
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// === Setup Telegram Bot dengan webhook untuk Railway ===
let bot;
const PORT = process.env.PORT || 3000;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = process.env.USE_WEBHOOK === 'true' || !!RAILWAY_STATIC_URL;

if (USE_WEBHOOK && RAILWAY_STATIC_URL) {
  // Webhook mode untuk Railway
  const express = require('express');
  const app = express();
  app.use(express.json());
  
  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/bot${TOKEN}`;
  
  bot.setWebHook(webhookUrl).then(() => {
    console.log(`Webhook set to: ${webhookUrl}`);
  }).catch(err => {
    console.error('Failed to set webhook:', err);
  });
  
  app.post(`/bot${TOKEN}`, (req, res) => {
    bot.processUpdate(req.body);
    res.sendStatus(200);
  });
  
  app.get('/', (req, res) => {
    res.send('Bot is running!');
  });
  
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
} else {
  // Polling mode untuk development
  bot = new TelegramBot(TOKEN, { polling: true });
  console.log('Bot running in polling mode');
}

// === Helper: Ambil data dari sheet dengan error handling ===
async function getSheetData(sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    return res.data.values || [];
  } catch (error) {
    console.error(`Error getting sheet data from ${sheetName}:`, error.message);
    throw error;
  }
}

// === Helper: Tambah data ke sheet dengan error handling ===
async function appendSheetData(sheetName, values) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
  } catch (error) {
    console.error(`Error appending data to ${sheetName}:`, error.message);
    throw error;
  }
}

// === Helper: Update range sheet data ===
async function updateSheetData(sheetName, range, values) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!${range}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values },
    });
  } catch (error) {
    console.error(`Error updating sheet data:`, error.message);
    throw error;
  }
}

// === Helper: Kirim pesan Telegram dengan retry logic dan reply ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  const maxRetries = 3;
  
  async function sendWithRetry(message, retries = 0) {
    try {
      return await bot.sendMessage(chatId, message, { parse_mode: 'HTML', ...options });
    } catch (error) {
      if (retries < maxRetries) {
        console.log(`Retry ${retries + 1} sending message to ${chatId}`);
        await new Promise(resolve => setTimeout(resolve, 1000 * (retries + 1)));
        return sendWithRetry(message, retries + 1);
      }
      throw error;
    }
  }
  
  if (text.length <= maxLength) {
    return sendWithRetry(text);
  } else {
    // Split by line, try not to break in the middle of a line
    const lines = text.split('\n');
    let chunk = '';
    let promises = [];
    for (let i = 0; i < lines.length; i++) {
      if ((chunk + lines[i] + '\n').length > maxLength) {
        promises.push(sendWithRetry(chunk));
        chunk = '';
      }
      chunk += lines[i] + '\n';
    }
    if (chunk.trim()) promises.push(sendWithRetry(chunk));
    return Promise.all(promises);
  }
}

// === Helper: Kirim file CSV ===
async function sendCSVFile(chatId, csvContent, filename, options = {}) {
  try {
    const tempDir = '/tmp';
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const filePath = path.join(tempDir, filename);
    fs.writeFileSync(filePath, csvContent, 'utf8');
    
    await bot.sendDocument(chatId, filePath, {
      caption: `📊 File CSV berhasil digenerate!\nFilename: ${filename}`,
      ...options
    });
    
    // Cleanup file setelah dikirim
    fs.unlinkSync(filePath);
  } catch (error) {
    console.error('Error sending CSV file:', error);
    throw error;
  }
}

// === Helper: Cek user aktif dengan error handling ===
async function getUserData(username) {
  try {
    const data = await getSheetData(USER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const userSheetUsername = (data[i][1] || '').replace('@', '').toLowerCase();
      const inputUsername = (username || '').replace('@', '').toLowerCase();
      const userStatus = (data[i][3] || '').toUpperCase();
      if (userSheetUsername === inputUsername && userStatus === 'AKTIF') {
        return data[i];
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting user data:', error);
    return null;
  }
}

// === Helper: Cek admin ===
async function isAdmin(username) {
  const user = await getUserData(username);
  return user && (user[2] || '').toUpperCase() === 'ADMIN';
}

// === Helper: Get today's date string ===
function getTodayDateString() {
  const today = new Date();
  return today.toLocaleDateString('id-ID', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
    timeZone: 'Asia/Jakarta'
  });
}

// === Helper: Parse tanggal dari string Indonesia ke Date object ===
function parseIndonesianDate(dateStr) {
  const months = {
    'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
    'mei': '05', 'juni': '06', 'juli': '07', 'agustus': '08',
    'september': '09', 'oktober': '10', 'november': '11', 'desember': '12'
  };
  
  const parts = dateStr.toLowerCase().split(' ');
  if (parts.length >= 4) {
    const day = parts[1].padStart(2, '0');
    const month = months[parts[2]];
    const year = parts[3];
    if (month) {
      return new Date(`${year}-${month}-${day}`);
    }
  }
  return null;
}

// === Helper: Filter data berdasarkan periode ===
function filterDataByPeriod(data, period, customDate = null) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let startDate, endDate;
  
  if (customDate) {
    // Parse custom date format (dd/mm/yyyy atau dd-mm-yyyy)
    const datePattern = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/;
    const match = customDate.match(datePattern);
    if (match) {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1; // Month is 0-indexed
      const year = parseInt(match[3]);
      const targetDate = new Date(year, month, day);
      
      if (period === 'daily') {
        startDate = new Date(targetDate);
        endDate = new Date(targetDate);
        endDate.setHours(23, 59, 59, 999);
      } else if (period === 'weekly') {
        // Get week start (Monday)
        const dayOfWeek = targetDate.getDay();
        const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        startDate = new Date(targetDate);
        startDate.setDate(targetDate.getDate() + mondayOffset);
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 6);
        endDate.setHours(23, 59, 59, 999);
      } else if (period === 'monthly') {
        startDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
        endDate = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0);
        endDate.setHours(23, 59, 59, 999);
      }
    }
  } else {
    // Default periods (current)
    switch (period) {
      case 'daily':
        startDate = new Date(today);
        endDate = new Date(today);
        endDate.setHours(23, 59, 59, 999);
        break;
      case 'weekly':
        const dayOfWeek = today.getDay();
        const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        startDate = new Date(today);
        startDate.setDate(today.getDate() + mondayOffset);
        endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 6);
        endDate.setHours(23, 59, 59, 999);
        break;
      case 'monthly':
        startDate = new Date(today.getFullYear(), today.getMonth(), 1);
        endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        endDate.setHours(23, 59, 59, 999);
        break;
      default:
        return data.slice(1); // Return all data except header
    }
  }
  
  const filtered = [];
  for (let i = 1; i < data.length; i++) {
    const dateStr = data[i][0];
    if (dateStr) {
      const rowDate = parseIndonesianDate(dateStr);
      if (rowDate && rowDate >= startDate && rowDate <= endDate) {
        filtered.push(data[i]);
      }
    }
  }
  
  return filtered;
}

// === Helper: Generate CSV content ===
function generateCSV(data, headers) {
  let csv = headers.join(',') + '\n';
  
  data.forEach(row => {
    const csvRow = row.map(cell => {
      const cellStr = (cell || '').toString();
      // Escape double quotes and wrap in quotes if contains comma or quotes
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    });
    csv += csvRow.join(',') + '\n';
  });
  
  return csv;
}

// === Handler pesan masuk dengan error handling lengkap ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const messageId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const chatType = msg.chat.type;
  
  // Log untuk debugging
  console.log(`Message received - Chat: ${chatId}, User: @${username}, Type: ${chatType}, Text: ${text.substring(0, 50)}`);
  
  try {
    // === Hanya proses /aktivasi di group, command lain diabaikan ===
    if ((chatType === 'group' || chatType === 'supergroup') && !/^\/aktivasi\b/i.test(text)) {
      return;
    }
    
    // === /exportcari: Export detail aktivasi user ke CSV ===
    if (/^\/exportcari\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const userTeknisi = (user[1] || username).replace('@', '').toLowerCase();
      const userActivations = [];
      
      // Headers untuk CSV
      const headers = ['TANGGAL', 'AO', 'WORKORDER', 'SERVICE_NO', 'CUSTOMER_NAME', 'OWNER', 'WORKZONE', 'SN_ONT', 'NIK_ONT', 'STB_ID', 'NIK_STB', 'TEKNISI'];
      
      // Filter data untuk user ini
      for (let i = 1; i < data.length; i++) {
        const teknisiData = (data[i][11] || '').replace('@', '').toLowerCase();
        if (teknisiData === userTeknisi) {
          userActivations.push(data[i]);
        }
      }
      
      if (userActivations.length === 0) {
        return sendTelegram(chatId, '❌ Tidak ada data aktivasi untuk diekspor.', { reply_to_message_id: messageId });
      }
      
      // Generate CSV
      const csvContent = generateCSV(userActivations, headers);
      const filename = `aktivasi_${userTeknisi}_${new Date().toISOString().split('T')[0]}.csv`;
      
      await sendCSVFile(chatId, csvContent, filename, { reply_to_message_id: messageId });
    }
    
    // === /ps: Laporan harian detail dengan support tanggal custom ===
    else if (/^\/ps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /ps hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      // Parse parameter tanggal jika ada
      const args = text.split(' ').slice(1);
      const customDate = args.length > 0 ? args[0] : null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      const filteredData = customDate ? 
        filterDataByPeriod(data, 'daily', customDate) : 
        filterDataByPeriod(data, 'daily');
      
      let total = filteredData.length;
      let teknisiMap = {}, workzoneMap = {}, ownerMap = {};
      
      filteredData.forEach(row => {
        const teknisi = (row[11] || '-').toUpperCase();
        const workzone = (row[6] || '-').toUpperCase();
        const owner = (row[5] || '-').toUpperCase();
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        ownerMap[owner] = (ownerMap[owner] || 0) + 1;
      });
      
      const dateLabel = customDate ? `Tanggal: ${customDate}` : `Tanggal: ${getTodayDateString()}`;
      let msg = `📊 <b>LAPORAN AKTIVASI HARIAN</b>\n${dateLabel}\nTotal Aktivasi: ${total} SSL\n\n`;
      
      if (total === 0) {
        msg += '⚠️ Belum ada data aktivasi untuk periode ini.\n\n';
      } else {
        msg += `METRICS PERIODE INI:\n- Teknisi Aktif: ${Object.keys(teknisiMap).length}\n- Workzone Tercover: ${Object.keys(workzoneMap).length}\n- Owner: ${Object.keys(ownerMap).length}\n\n`;
        
        msg += 'PERFORMA TEKNISI:\n';
        Object.entries(teknisiMap).sort((a,b)=>b[1]-a[1]).forEach(([t,c],i)=>{
          msg+=`${i+1}. ${t}: ${c} SSL\n`;
        });
        
        msg += '\nPERFORMA WORKZONE:\n';
        Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).forEach(([w,c],i)=>{
          msg+=`${i+1}. ${w}: ${c} SSL\n`;
        });
        
        msg += '\nPERFORMA OWNER:\n';
        Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c],i)=>{
          msg+=`${i+1}. ${o}: ${c} SSL\n`;
        });
      }
      
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /weekly: Laporan mingguan ===
    else if (/^\/weekly\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /weekly hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const args = text.split(' ').slice(1);
      const customDate = args.length > 0 ? args[0] : null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      const filteredData = filterDataByPeriod(data, 'weekly', customDate);
      
      let total = filteredData.length;
      let teknisiMap = {}, workzoneMap = {}, ownerMap = {};
      
      filteredData.forEach(row => {
        const teknisi = (row[11] || '-').toUpperCase();
        const workzone = (row[6] || '-').toUpperCase();
        const owner = (row[5] || '-').toUpperCase();
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        ownerMap[owner] = (ownerMap[owner] || 0) + 1;
      });
      
      const periodLabel = customDate ? `Minggu dari: ${customDate}` : 'Minggu ini';
      let msg = `📈 <b>LAPORAN AKTIVASI MINGGUAN</b>\n${periodLabel}\nTotal Aktivasi: ${total} SSL\n\n`;
      
      if (total === 0) {
        msg += '⚠️ Belum ada data aktivasi untuk periode ini.\n\n';
      } else {
        msg += `METRICS MINGGUAN:\n- Teknisi Aktif: ${Object.keys(teknisiMap).length}\n- Workzone Tercover: ${Object.keys(workzoneMap).length}\n- Owner: ${Object.keys(ownerMap).length}\n\n`;
        
        msg += 'TOP 10 TEKNISI MINGGU INI:\n';
        Object.entries(teknisiMap).sort((a,b)=>b[1]-a[1]).slice(0,10).forEach(([t,c],i)=>{
          const medal = i < 3 ? ['🥇', '🥈', '🥉'][i] : `${i+1}.`;
          msg+=`${medal} ${t}: ${c} SSL\n`;
        });
        
        msg += '\nWORKZONE TERBAIK:\n';
        Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).slice(0,5).forEach(([w,c],i)=>{
          msg+=`${i+1}. ${w}: ${c} SSL\n`;
        });
      }
      
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /monthly: Laporan bulanan ===
    else if (/^\/monthly\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /monthly hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const args = text.split(' ').slice(1);
      const customDate = args.length > 0 ? args[0] : null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      const filteredData = filterDataByPeriod(data, 'monthly', customDate);
      
      let total = filteredData.length;
      let teknisiMap = {}, workzoneMap = {}, ownerMap = {};
      
      filteredData.forEach(row => {
        const teknisi = (row[11] || '-').toUpperCase();
        const workzone = (row[6] || '-').toUpperCase();
        const owner = (row[5] || '-').toUpperCase();
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        ownerMap[owner] = (ownerMap[owner] || 0) + 1;
      });
      
      const periodLabel = customDate ? `Bulan dari: ${customDate}` : 'Bulan ini';
      let msg = `📅 <b>LAPORAN AKTIVASI BULANAN</b>\n${periodLabel}\nTotal Aktivasi: ${total} SSL\n\n`;
      
      if (total === 0) {
        msg += '⚠️ Belum ada data aktivasi untuk periode ini.\n\n';
      } else {
        msg += `METRICS BULANAN:\n- Teknisi Aktif: ${Object.keys(teknisiMap).length}\n- Workzone Tercover: ${Object.keys(workzoneMap).length}\n- Owner: ${Object.keys(ownerMap).length}\n- Rata-rata per hari: ${(total / 30).toFixed(1)} SSL\n\n`;
        
        msg += 'TOP 15 TEKNISI BULAN INI:\n';
        Object.entries(teknisiMap).sort((a,b)=>b[1]-a[1]).slice(0,15).forEach(([t,c],i)=>{
          const medal = i < 3 ? ['🥇', '🥈', '🥉'][i] : `${i+1}.`;
          msg+=`${medal} ${t}: ${c} SSL\n`;
        });
        
        msg += '\nWORKZONE TERBAIK:\n';
        Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).slice(0,8).forEach(([w,c],i)=>{
          msg+=`${i+1}. ${w}: ${c} SSL\n`;
        });
      }
      
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /topteknisi: Ranking teknisi terbaik ===
    else if (/^\/topteknisi\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /topteknisi hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const args = text.split(' ').slice(1);
      const period = args[0] || 'all'; // all, daily, weekly, monthly
      const customDate = args[1] || null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      let filteredData;
      
      switch (period.toLowerCase()) {
        case 'daily':
          filteredData = filterDataByPeriod(data, 'daily', customDate);
          break;
        case 'weekly':
          filteredData = filterDataByPeriod(data, 'weekly', customDate);
          break;
        case 'monthly':
          filteredData = filterDataByPeriod(data, 'monthly', customDate);
          break;
        default:
          filteredData = data.slice(1); // All data
      }
      
      let teknisiMap = {};
      filteredData.forEach(row => {
        const teknisi = (row[11] || '-').toUpperCase();
        if (teknisi !== '-') {
          teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        }
      });
      
      const sortedTeknisi = Object.entries(teknisiMap).sort((a,b) => b[1] - a[1]);
      const periodLabel = {
        daily: customDate ? `Harian (${customDate})` : 'Hari ini',
        weekly: customDate ? `Mingguan (${customDate})` : 'Minggu ini',
        monthly: customDate ? `Bulanan (${customDate})` : 'Bulan ini',
        all: 'Keseluruhan'
      };
      
      let msg = `🏆 <b>RANKING TEKNISI TERBAIK</b>\nPeriode: ${periodLabel[period.toLowerCase()] || 'Keseluruhan'}\n\n`;
      
      if (sortedTeknisi.length === 0) {
        msg += '⚠️ Belum ada data teknisi untuk periode ini.\n';
      } else {
        msg += `Total Teknisi Aktif: ${sortedTeknisi.length}\n\n`;
        msg += '🏅 <b>TOP 20 TEKNISI:</b>\n';
        
        sortedTeknisi.slice(0, 20).forEach(([teknisi, count], index) => {
          let icon = '';
          if (index === 0) icon = '🥇';
          else if (index === 1) icon = '🥈';
          else if (index === 2) icon = '🥉';
          else icon = `${index + 1}.`;
          
          msg += `${icon} ${teknisi}: <b>${count} SSL</b>\n`;
        });
        
        if (sortedTeknisi.length > 20) {
          msg += `\n... dan ${sortedTeknisi.length - 20} teknisi lainnya`;
        }
      }
      
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /allps: breakdown owner, sektor, top teknisi ===
    else if (/^\/allps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /allps hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      let total = Math.max(0, data.length - 1);
      let ownerMap = {}, sektorMap = {}, teknisiMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const owner = (data[i][5] || '-').toUpperCase();
        const sektor = (data[i][6] || '-').toUpperCase();
        const teknisi = (data[i][11] || '-').toUpperCase();
        ownerMap[owner] = (ownerMap[owner] || 0) + 1;
        sektorMap[sektor] = (sektorMap[sektor] || 0) + 1;
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
      }
      
      let msg = '📊 <b>RINGKASAN AKTIVASI TOTAL</b>\n';
      msg += `TOTAL KESELURUHAN: ${total} SSL\n\nBERDASARKAN OWNER:\n`;
      Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c])=>{
        msg+=`- ${o}: ${c}\n`;
      });
      msg += '\nBERDASARKAN SEKTOR/WORKZONE:\n';
      Object.entries(sektorMap).sort((a,b)=>b[1]-a[1]).forEach(([s,c])=>{
        msg+=`- ${s}: ${c}\n`;
      });
      
      let teknisiArr = Object.entries(teknisiMap).map(([name,count])=>({name,count}));
      teknisiArr.sort((a,b)=>b.count-a.count);
      msg += '\nTOP TEKNISI:\n';
      teknisiArr.slice(0,5).forEach((t,i)=>{
        msg+=`${i+1}. ${t.name}: ${t.count}\n`;
      });
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /cari: menampilkan total dari user tersebut (FIXED) ===
    else if (/^\/cari\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const userTeknisi = (user[1] || username).replace('@', '').toLowerCase();
      let count = 0;
      let ownerMap = {}, workzoneMap = {};
      
      // Cari semua data dari teknisi yang sesuai
      for (let i = 1; i < data.length; i++) {
        const teknisiData = (data[i][11] || '').replace('@', '').toLowerCase();
        if (teknisiData === userTeknisi) {
          count++;
          const owner = (data[i][5] || '-').toUpperCase();
          const workzone = (data[i][6] || '-').toUpperCase();
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>STATISTIK ANDA</b>\n👤 Teknisi: ${user[1] || username}\n📈 Total Aktivasi: ${count} SSL\n\n`;
      
      if (count === 0) {
        msg += '⚠️ Belum ada data aktivasi yang tercatat untuk Anda.\n';
      } else {
        msg += 'DETAIL PER OWNER:\n';
        Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c])=>{
          msg+=`- ${o}: ${c}\n`;
        });
        msg += '\nDETAIL PER WORKZONE:\n';
        Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).forEach(([s,c])=>{
          msg+=`- ${s}: ${c}\n`;
        });
        
        msg += '\n💾 <i>Tip: Gunakan /exportcari untuk download data lengkap dalam format CSV</i>';
      }
      
      msg += `\nUpdated: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /@username: menampilkan total dari username tersebut (FIXED FORMAT) ===
    else if (/^\/[A-Za-z0-9_]+$/.test(text) && !text.match(/^\/cari|^\/ps|^\/allps|^\/clean|^\/clear|^\/help|^\/start|^\/aktivasi|^\/exportcari|^\/weekly|^\/monthly|^\/topteknisi/i)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command ini hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const targetUsername = text.substring(1).toLowerCase(); // Remove / and convert to lowercase
      const data = await getSheetData(REKAPAN_SHEET);
      let count = 0;
      let ownerMap = {}, workzoneMap = {};
      
      // Cari data berdasarkan username (dengan atau tanpa @)
      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][11] || '').replace('@', '').toLowerCase();
        if (teknisi === targetUsername) {
          count++;
          const owner = (data[i][5] || '-').toUpperCase();
          const workzone = (data[i][6] || '-').toUpperCase();
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>STATISTIK TEKNISI</b>\n👤 Username: ${text}\n📈 Total Aktivasi: ${count} SSL\n\n`;
      
      if (count === 0) {
        msg += '⚠️ Belum ada data aktivasi yang tercatat untuk teknisi ini.\n';
      } else {
        msg += 'DETAIL PER OWNER:\n';
        Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c])=>{
          msg+=`- ${o}: ${c}\n`;
        });
        msg += '\nDETAIL PER WORKZONE:\n';
        Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).forEach(([s,c])=>{
          msg+=`- ${s}: ${c}\n`;
        });
      }
      
      msg += `\nUpdated: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /clear: untuk menghapus duplikat di sheet berdasarkan AO ===
    else if (/^\/clear\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /clear hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      if (data.length <= 1) {
        return sendTelegram(chatId, '✅ Sheet sudah bersih, tidak ada data duplikat.', { reply_to_message_id: messageId });
      }
      
      const seen = new Set();
      const uniqueData = [data[0]]; // Keep header
      let duplicateCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const ao = (data[i][1] || '').toUpperCase().trim(); // AO is in column index 1
        
        if (!seen.has(ao) && ao) {
          seen.add(ao);
          uniqueData.push(data[i]);
        } else {
          duplicateCount++;
        }
      }
      
      if (duplicateCount === 0) {
        return sendTelegram(chatId, '✅ Sheet sudah bersih, tidak ada data duplikat.', { reply_to_message_id: messageId });
      }
      
      // Update sheet with clean data
      await updateSheetData(REKAPAN_SHEET, `A1:L${uniqueData.length}`, uniqueData);
      return sendTelegram(chatId, `✅ Berhasil menghapus ${duplicateCount} data duplikat berdasarkan AO. Sheet telah dibersihkan.`, { reply_to_message_id: messageId });
    }
    
    // === /aktivasi: parsing multi-format yang diperbaiki, cek duplikat berdasarkan AO, simpan ===
    else if (/^\/aktivasi\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const inputText = text.replace(/^\/aktivasi\s*/i, '').trim();
      if (!inputText) {
        return sendTelegram(chatId, 'Silakan kirim data aktivasi setelah /aktivasi.', { reply_to_message_id: messageId });
      }
      
      // Use the extracted parser module for parsing multi-format messages
      const parsed = parseAktivasi(inputText, user, username);
      
      // Validasi minimal AO harus ada
      let missing = [];
      if (!parsed.ao) missing.push('AO');
      if (missing.length > 0) {
        return sendTelegram(chatId, `❌ Data tidak lengkap. Field berikut wajib diisi: ${missing.join(', ')}`, { reply_to_message_id: messageId });
      }
      
      // === Cek duplikat: AO sudah ada di sheet ===
      const data = await getSheetData(REKAPAN_SHEET);
      let isDuplicate = false;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][1] || '').toUpperCase().trim() === parsed.ao.toUpperCase().trim()) {
          isDuplicate = true;
          break;
        }
      }
      if (isDuplicate) {
        return sendTelegram(chatId, '❌ Data duplikat. AO sudah pernah diinput.', { reply_to_message_id: messageId });
      }
      
      // Susun data sesuai urutan kolom sheet
      const tanggal = getTodayDateString();
      
      const row = [
        tanggal,               // TANGGAL
        parsed.ao,             // AO
        parsed.workorder,      // WORKORDER
        parsed.serviceNo,      // SERVICE NO
        parsed.customerName,   // CUSTOMER NAME
        parsed.owner,          // OWNER
        parsed.workzone,       // WORKZONE
        parsed.snOnt,          // SN ONT
        parsed.nikOnt,         // NIK ONT
        parsed.stbId,          // STB ID
        parsed.nikStb,         // NIK STB
        parsed.teknisi         // TEKNISI
      ];
      
      await appendSheetData(REKAPAN_SHEET, row);
      
      // Tampilkan konfirmasi dengan data yang berhasil diparse
      let confirmMsg = '✅ Data berhasil disimpan ke sheet, GASPOLLL 🚀🚀!\n\n';
      confirmMsg += '<b>Lanjut GROUP FULFILLMENT dan PT1</b>\n';
      
      return sendTelegram(chatId, confirmMsg, { reply_to_message_id: messageId });
    }
    
    // === /help: Command list yang diperbaiki dan lebih detail ===
    else if (/^\/help\b/i.test(text) || /^\/start\b/i.test(text)) {
      let helpMsg = '🤖 <b>Bot Rekapan Quality - Panduan Lengkap</b>\n\n';
      
      helpMsg += '📝 <b>COMMANDS UNTUK USER:</b>\n';
      helpMsg += '• <code>/aktivasi [data]</code> - Input data aktivasi\n';
      helpMsg += '• <code>/cari</code> - Lihat statistik total aktivasi Anda\n';
      helpMsg += '• <code>/exportcari</code> - Download data aktivasi Anda dalam format CSV\n';
      helpMsg += '• <code>/help</code> - Tampilkan bantuan ini\n\n';
      
      helpMsg += '📊 <b>FORMAT INPUT AKTIVASI:</b>\n';
      helpMsg += 'Bot mendukung 3 format input:\n';
      helpMsg += '1. <b>Auto-detect BGES/WMS:</b> Copy paste langsung dari sistem\n';
      helpMsg += '2. <b>Auto-detect TSEL:</b> Copy paste langsung dari sistem\n';
      helpMsg += '3. <b>Format Manual:</b>\n';
      helpMsg += '   AO : SC123456\n';
      helpMsg += '   SERVICE NO : 12345678901\n';
      helpMsg += '   CUSTOMER NAME : JOHN DOE\n';
      helpMsg += '   OWNER : BGES\n';
      helpMsg += '   WORKZONE : MEDAN\n';
      helpMsg += '   SN ONT : ZTEG12345678\n';
      helpMsg += '   NIK ONT : 987654321\n\n';
      
      if (await isAdmin(username)) {
        helpMsg += '👑 <b>ADMIN COMMANDS:</b>\n';
        helpMsg += '• <code>/ps [tanggal]</code> - Laporan harian\n';
        helpMsg += '   Contoh: /ps atau /ps 01/09/2025\n';
        helpMsg += '• <code>/weekly [tanggal]</code> - Laporan mingguan\n';
        helpMsg += '   Contoh: /weekly atau /weekly 01/09/2025\n';
        helpMsg += '• <code>/monthly [tanggal]</code> - Laporan bulanan\n';
        helpMsg += '   Contoh: /monthly atau /monthly 01/09/2025\n';
        helpMsg += '• <code>/topteknisi [periode] [tanggal]</code> - Ranking teknisi\n';
        helpMsg += '   Periode: all, daily, weekly, monthly\n';
        helpMsg += '   Contoh: /topteknisi monthly 01/09/2025\n';
        helpMsg += '• <code>/allps</code> - Ringkasan total keseluruhan\n';
        helpMsg += '• <code>/[username]</code> - Statistik teknisi tertentu\n';
        helpMsg += '   Contoh: /HKS_HENDRA_16951456\n';
        helpMsg += '• <code>/clear</code> - Hapus data duplikat dari sheet\n\n';
      }
      
      helpMsg += '💡 <b>TIPS PENGGUNAAN:</b>\n';
      helpMsg += '• Field wajib: AO, SERVICE NO, CUSTOMER NAME, OWNER, WORKZONE, SN ONT, NIK ONT\n';
      helpMsg += '• Bot otomatis mendeteksi format BGES, WMS, dan TSEL\n';
      helpMsg += '• Gunakan format tanggal: DD/MM/YYYY atau DD-MM-YYYY\n';
      helpMsg += '• Data duplikat (berdasarkan AO) akan ditolak sistem\n';
      helpMsg += '• Export CSV tersedia untuk backup data personal\n\n';
      
      helpMsg += '🚀 <b>Bot siap membantu aktivasi Anda!</b>\n';
      helpMsg += '📅 Generated: ' + new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'}) + ' WIB';
      
      return sendTelegram(chatId, helpMsg, { reply_to_message_id: messageId });
    }
    
    // Default response for unknown commands
    else if (text.startsWith('/')) {
      return sendTelegram(chatId, '❓ Command tidak dikenali. Ketik /help untuk melihat daftar command yang tersedia untuk Anda.', { reply_to_message_id: messageId });
    }
    
  } catch (err) {
    console.error('Error processing message:', err);
    return sendTelegram(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi nanti.', { reply_to_message_id: messageId });
  }
});

// Error handling untuk uncaught exceptions
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

console.log('Bot Telegram Rekapan started successfully!');
console.log('Mode:', USE_WEBHOOK ? 'Webhook' : 'Polling');
if (USE_WEBHOOK) {
  console.log('Listening on port:', PORT);
}

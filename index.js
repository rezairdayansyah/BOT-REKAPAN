require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

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

// === Helper: Clean customer name by removing WORKZONE suffix ===
function cleanCustomerName(name) {
  if (!name) return '';
  
  // Remove "WORKZONE" and any text after it (case insensitive)
  // This will handle patterns like "RZKLDESSY YUNANDA WORKZONE" -> "RZKLDESSY YUNANDA"
  const cleaned = name.replace(/\s+WORKZONE.*$/i, '').trim();
  
  return cleaned;
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

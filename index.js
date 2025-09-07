require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

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
  today.setHours(today.getHours() + 7); // Adjust for WIB timezone
  return today.toLocaleDateString('id-ID', { 
    weekday: 'long', 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric',
    timeZone: 'Asia/Jakarta'
  });
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
    
    // === /ps: Laporan harian detail (hanya data hari ini) ===
    if (/^\/ps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /ps hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const todayStr = getTodayDateString();
      
      let total = 0;
      let teknisiMap = {}, workzoneMap = {}, ownerMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const tgl = (data[i][0] || '').trim();
        if (tgl === todayStr) {
          total++;
          const teknisi = (data[i][11] || '-').toUpperCase();
          const workzone = (data[i][6] || '-').toUpperCase();
          const owner = (data[i][5] || '-').toUpperCase();
          teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>LAPORAN AKTIVASI HARIAN</b>\nTanggal: ${todayStr}\nTotal Aktivasi: ${total} SSL\n\n`;
      msg += `METRICS:\n- Teknisi Aktif: ${Object.keys(teknisiMap).length}\n- Workzone Tercover: ${Object.keys(workzoneMap).length}\n- Owner: ${Object.keys(ownerMap).length}\n\n`;
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
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /allps: breakdown owner, sektor, top teknisi ===
    if (/^\/allps\b/i.test(text)) {
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
    
    // === /cari: menampilkan total dari user tersebut ===
    if (/^\/cari\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const userTeknisi = user[1] || username;
      let count = 0;
      let ownerMap = {}, workzoneMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][11] || '').toLowerCase();
        if (teknisi === userTeknisi.toLowerCase()) {
          count++;
          const owner = (data[i][5] || '-').toUpperCase();
          const workzone = (data[i][6] || '-').toUpperCase();
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>STATISTIK ANDA</b>\n👤 Teknisi: ${userTeknisi}\n📈 Total Aktivasi: ${count} SSL\n\nDETAIL PER OWNER:\n`;
      Object.entries(ownerMap).forEach(([o,c])=>{
        msg+=`- ${o}: ${c}\n`;
      });
      msg += '\nDETAIL PER WORKZONE:\n';
      Object.entries(workzoneMap).forEach(([s,c])=>{
        msg+=`- ${s}: ${c}\n`;
      });
      msg += `\nUpdated: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /username: menampilkan total dari username tersebut (khusus admin) ===
    if (/^\/\w+$/.test(text) && text.startsWith('/@')) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command ini hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const targetUsername = text.substring(2); // Remove /@
      const data = await getSheetData(REKAPAN_SHEET);
      let count = 0;
      let ownerMap = {}, workzoneMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][11] || '').toLowerCase().replace('@', '');
        if (teknisi === targetUsername.toLowerCase()) {
          count++;
          const owner = (data[i][5] || '-').toUpperCase();
          const workzone = (data[i][6] || '-').toUpperCase();
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>STATISTIK TEKNISI</b>\n👤 Username: @${targetUsername}\n📈 Total Aktivasi: ${count} SSL\n\nDETAIL PER OWNER:\n`;
      Object.entries(ownerMap).forEach(([o,c])=>{
        msg+=`- ${o}: ${c}\n`;
      });
      msg += '\nDETAIL PER WORKZONE:\n';
      Object.entries(workzoneMap).forEach(([s,c])=>{
        msg+=`- ${s}: ${c}\n`;
      });
      msg += `\nUpdated: ${new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'})} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /clean: untuk menghapus duplikat di sheet ===
    if (/^\/clean\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /clean hanya untuk admin.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      if (data.length <= 1) {
        return sendTelegram(chatId, '✅ Sheet sudah bersih, tidak ada data duplikat.', { reply_to_message_id: messageId });
      }
      
      const seen = new Set();
      const uniqueData = [data[0]]; // Keep header
      let duplicateCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const snOnt = (data[i][7] || '').toUpperCase();
        const nikOnt = (data[i][8] || '').toUpperCase();
        const key = `${snOnt}-${nikOnt}`;
        
        if (!seen.has(key) && snOnt && nikOnt) {
          seen.add(key);
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
      return sendTelegram(chatId, `✅ Berhasil menghapus ${duplicateCount} data duplikat. Sheet telah dibersihkan.`, { reply_to_message_id: messageId });
    }
    
    // === /aktivasi: parsing multi-format, cek duplikat, simpan ===
    if (/^\/aktivasi\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const inputText = text.replace(/^\/aktivasi\s*/i, '').trim();
      if (!inputText) {
        return sendTelegram(chatId, 'Silakan kirim data aktivasi setelah /aktivasi.', { reply_to_message_id: messageId });
      }
      
      // === Parsing multi-format yang diperbaiki ===
      function parseAktivasi(text, userRow) {
        const lines = text.split('\n').map(l=>l.trim()).filter(l=>l);
        const upper = text.toUpperCase();
        let ao='', workorder='', serviceNo='', customerName='', owner='', workzone='', snOnt='', nikOnt='', stbId='', nikStb='', teknisi='';
        teknisi = userRow[1] || username;
        
        // Helper untuk mencari nilai dengan berbagai pola
        function findValue(patterns) {
          for (const pattern of patterns) {
            for (const line of lines) {
              const match = line.match(pattern);
              if (match && match[1]) {
                return match[1].trim();
              }
            }
          }
          return '';
        }
        
        // === BGES ===
        if (upper.includes('BGES')) {
          owner = 'BGES';
          
          // AO/Workorder - mencari pola SC diikuti angka
          ao = findValue([
            /SC(\d+)/i,
            /AO[:\s]+([A-Z0-9]+)/i,
            /WORKORDER[:\s]+([A-Z0-9-]+)/i
          ]);
          workorder = ao; // Untuk BGES, AO dan workorder biasanya sama
          
          // Service No - mencari angka panjang atau pola khusus
          serviceNo = findValue([
            /SERVICE\s*NO[:\s]+(\d+)/i,
            /(\d{10,15})\s+null\s+MIA/i,
            /(\d{10,15})/
          ]);
          
          // Customer Name - berbagai pola
          customerName = findValue([
            /CUSTOMER\s*NAME[:\s]+(.+)/i,
            /null\s+\d{8}\s+([A-Z\s]+?)(?:\s+[A-Za-z]+,|\s+Tijue,|\s+\d{10,15}|\s+null)/i,
            /(WMS\s+[A-Z\s]+)/i,
            /(BGES\s+[A-Z\s]+)/i
          ]);
          
          // Workzone
          workzone = findValue([
            /WORKZONE[:\s]+([A-Z0-9]+)/i,
            /AO\|INTERNET\s+([A-Z]{3})/i,
            /([A-Z]{3})\s*$/i
          ]);
          
          // SN ONT - berbagai brand
          snOnt = findValue([
            /SN\s*ONT[:\s]+([A-Z0-9]+)/i,
            /(ZTEG[A-Z0-9]+)/i,
            /(HWTC[A-Z0-9]+)/i,
            /(HUAW[A-Z0-9]+)/i,
            /(FHTT[A-Z0-9]+)/i,
            /(FIBR[A-Z0-9]+)/i
          ]);
          
          nikOnt = findValue([/NIK\s*ONT[:\s]+(\d+)/i]);
          stbId = findValue([/STB\s*ID[:\s]+([A-Z0-9]+)/i]);
          nikStb = findValue([/NIK\s*STB[:\s]+(\d+)/i]);
        }
        // === WMS ===
        else if (upper.includes('WMS')) {
          owner = 'WMS';
          
          // AO/Workorder
          ao = findValue([
            /SC(\d+)/i,
            /AO[:\s]+([A-Z0-9]+)/i,
            /WORKORDER[:\s]+([A-Z0-9-]+)/i
          ]);
          workorder = ao;
          
          // Service No
          serviceNo = findValue([
            /SERVICE\s*NO[:\s]+(\d+)/i,
            /(\d{10,15})/
          ]);
          
          // Customer Name - khusus untuk WMS
          customerName = findValue([
            /CUSTOMER\s*NAME[:\s]+(.+)/i,
            /(WMS\s+[A-Z\s]+)/i
          ]);
          
          // Workzone
          workzone = findValue([
            /WORKZONE[:\s]+([A-Z0-9]+)/i,
            /([A-Z]{3})\s*$/i
          ]);
          
          // SN ONT
          snOnt = findValue([
            /SN\s*ONT[:\s]+([A-Z0-9]+)/i,
            /(ZTEG[A-Z0-9]+)/i,
            /(HWTC[A-Z0-9]+)/i,
            /(HUAW[A-Z0-9]+)/i,
            /(FHTT[A-Z0-9]+)/i,
            /(FIBR[A-Z0-9]+)/i
          ]);
          
          nikOnt = findValue([/NIK\s*ONT[:\s]+(\d+)/i]);
          stbId = findValue([/STB\s*ID[:\s]+([A-Z0-9]+)/i]);
          nikStb = findValue([/NIK\s*STB[:\s]+(\d+)/i]);
        }
        // === TSEL ===
        else if (upper.includes('TSEL')) {
          owner = 'TSEL';
          
          ao = findValue([
            /AO[:\s]+([A-Z0-9]+)/i,
            /SC(\d+)/i
          ]);
          workorder = findValue([/WORKORDER[:\s]+([A-Z0-9-]+)/i]) || ao;
          serviceNo = findValue([/SERVICE\s*NO[:\s]+(\d+)/i]);
          customerName = findValue([/CUSTOMER\s*NAME[:\s]+(.+)/i]);
          workzone = findValue([/WORKZONE[:\s]+([A-Z0-9]+)/i]);
          snOnt = findValue([
            /SN\s*ONT[:\s]+([A-Z0-9]+)/i,
            /(ZTEG[A-Z0-9]+)/i,
            /(HWTC[A-Z0-9]+)/i,
            /(HUAW[A-Z0-9]+)/i,
            /(FHTT[A-Z0-9]+)/i,
            /(FIBR[A-Z0-9]+)/i
          ]);
          nikOnt = findValue([/NIK\s*ONT[:\s]+(\d+)/i]);
          stbId = findValue([/STB\s*ID[:\s]+([A-Z0-9]+)/i]);
          nikStb = findValue([/NIK\s*STB[:\s]+(\d+)/i]);
        }
        // === fallback: label/manual/regex ===
        else {
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
      
      const parsed = parseAktivasi(inputText, user);
      
      // Validasi minimal SN ONT dan NIK ONT harus ada
      let missing = [];
      if (!parsed.snOnt) missing.push('SN ONT');
      if (!parsed.nikOnt) missing.push('NIK ONT');
      if (missing.length > 0) {
        return sendTelegram(chatId, `❌ Data tidak lengkap. Field berikut wajib diisi: ${missing.join(', ')}`, { reply_to_message_id: messageId });
      }
      
      // === Cek duplikat: SN ONT dan NIK ONT sudah ada di sheet ===
      const data = await getSheetData(REKAPAN_SHEET);
      let isDuplicate = false;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][7] || '').toUpperCase() === parsed.snOnt.toUpperCase() && 
            (data[i][8] || '').toUpperCase() === parsed.nikOnt.toUpperCase()) {
          isDuplicate = true;
          break;
        }
      }
      if (isDuplicate) {
        return sendTelegram(chatId, '❌ Data duplikat. SN ONT dan NIK ONT sudah pernah diinput.', { reply_to_message_id: messageId });
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
      return sendTelegram(chatId, '✅ Data berhasil disimpan ke sheet, GASPOLLL 🚀🚀!', { reply_to_message_id: messageId });
    }
    
    // === /help: Command list ===
    if (/^\/help\b/i.test(text) || /^\/start\b/i.test(text)) {
      let helpMsg = '🤖 <b>Bot Rekapan Quality</b>\n\n';
      helpMsg += '<b>Commands User:</b>\n';
      helpMsg += '/aktivasi - Input data aktivasi\n';
      helpMsg += '/cari - Lihat total aktivasi Anda\n';
      helpMsg += '/help - Tampilkan bantuan ini\n';
      
      if (await isAdmin(username)) {
        helpMsg += '\n<b>Admin Commands:</b>\n';
        helpMsg += '/ps - Laporan harian\n';
        helpMsg += '/allps - Ringkasan total\n';
        helpMsg += '/@username - Statistik berdasarkan username\n';
        helpMsg += '/clean - Hapus data duplikat\n';
      }
      
      return sendTelegram(chatId, helpMsg, { reply_to_message_id: messageId });
    }
    
    // Default response for unknown commands
    if (text.startsWith('/')) {
      return sendTelegram(chatId, '❓ Command tidak dikenali. Ketik /help untuk melihat daftar command.', { reply_to_message_id: messageId });
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

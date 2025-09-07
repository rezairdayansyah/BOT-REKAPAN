require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === Konfigurasi dari environment variables ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

const REKAPAN_SHEET = 'REKAPAN QUALITY';
const USER_SHEET = 'USER';

// === Setup Google Sheets API ===
console.log('GOOGLE_SERVICE_ACCOUNT_KEY:', typeof GOOGLE_SERVICE_ACCOUNT_KEY, GOOGLE_SERVICE_ACCOUNT_KEY ? 'OK' : 'NOT FOUND');
if (!GOOGLE_SERVICE_ACCOUNT_KEY) {
  throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set!');
}
let serviceAccount;
try {
  serviceAccount = JSON.parse(GOOGLE_SERVICE_ACCOUNT_KEY);
} catch (e) {
  console.error('ERROR parsing GOOGLE_SERVICE_ACCOUNT_KEY:', e.message);
  throw e;
}
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// === Setup Telegram Bot ===
const bot = new TelegramBot(TOKEN, { polling: true });

// === Helper: Ambil data dari sheet ===
async function getSheetData(sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: sheetName,
  });
  return res.data.values || [];
}

// === Helper: Tambah data ke sheet ===
async function appendSheetData(sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

// === Helper: Kirim pesan Telegram ===
function sendTelegram(chatId, text, options = {}) {
  return bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
}

// === Helper: Cek user aktif ===
async function getUserData(username) {
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
}

// === Helper: Cek admin ===
async function isAdmin(username) {
  const user = await getUserData(username);
  return user && (user[2] || '').toUpperCase() === 'ADMIN';
}

// === Handler pesan masuk ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';

  try {
    // === /ps ===
    if (/^\/ps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /ps hanya untuk admin.');
      }
      const data = await getSheetData(REKAPAN_SHEET);
      const today = new Date().toLocaleDateString('en-CA'); // yyyy-mm-dd
      let total = 0, sukses = 0, gagal = 0;
      for (let i = 1; i < data.length; i++) {
        const tgl = (data[i][0] || '').split(' ')[0];
        if (tgl === today) {
          total++;
          if ((data[i][6] || '').toUpperCase() === 'SUKSES') sukses++;
          else gagal++;
        }
      }
      return sendTelegram(chatId, `📊 <b>Laporan Hari Ini (${today})</b>\nTotal: ${total}\n✅ Sukses: ${sukses}\n❌ Gagal: ${gagal}`);
    }

    // === /cari <SN> ===
    if (/^\/cari\b/i.test(text)) {
      const sn = text.split(' ')[1];
      if (!sn) return sendTelegram(chatId, 'Format: /cari <SN>');
      const data = await getSheetData(REKAPAN_SHEET);
      let found = [];
      for (let i = 1; i < data.length; i++) {
        if ((data[i][2] || '').toUpperCase() === sn.toUpperCase()) {
          found.push(data[i]);
        }
      }
      if (found.length === 0) return sendTelegram(chatId, `SN <b>${sn}</b> tidak ditemukan.`);
      let msgText = `📄 <b>Data SN ${sn}:</b>\n`;
      found.forEach(row => {
        msgText += `Tgl: ${row[0]}, User: ${row[1]}, Status: ${row[6]}\n`;
      });
      return sendTelegram(chatId, msgText);
    }

    // === /allps ===
    if (/^\/allps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /allps hanya untuk admin.');
      }
      const data = await getSheetData(REKAPAN_SHEET);
      let total = data.length - 1;
      let sukses = 0, gagal = 0;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][6] || '').toUpperCase() === 'SUKSES') sukses++;
        else gagal++;
      }
      return sendTelegram(chatId, `📊 <b>Rekap Semua Data</b>\nTotal: ${total}\n✅ Sukses: ${sukses}\n❌ Gagal: ${gagal}`);
    }

    // === /nik <NIK> ===
    if (/^\/nik\b/i.test(text)) {
      const nik = text.split(' ')[1];
      if (!nik) return sendTelegram(chatId, 'Format: /nik <NIK>');
      const data = await getSheetData(REKAPAN_SHEET);
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][3] || '').toUpperCase() === nik.toUpperCase()) count++;
      }
      return sendTelegram(chatId, `NIK <b>${nik}</b> ditemukan pada <b>${count}</b> data.`);
    }

    // === /aktivasi (hanya isi kolom yang Anda minta) ===
    if (/^\/aktivasi\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.');

      // Ambil semua teks setelah "/aktivasi"
      const inputText = text.replace(/^\/aktivasi\s*/i, '').trim();
      if (!inputText) return sendTelegram(chatId, 'Silakan kirim data aktivasi setelah /aktivasi.');

      // Parsing field dari input multi-baris
      const lines = inputText.split('\n').map(l => l.trim()).filter(l => l);

      // Helper untuk cari value dari baris berlabel
      function getValue(label) {
        const line = lines.find(l => l.toUpperCase().startsWith(label.toUpperCase() + ' :'));
        return line ? line.split(':').slice(1).join(':').trim() : '';
      }

      // Helper untuk cari value dari baris tanpa label (untuk BGES/WMS)
      function findByPattern(pattern) {
        const regex = new RegExp(pattern, 'i');
        const line = lines.find(l => regex.test(l));
        return line ? (line.match(regex)[1] || '').trim() : '';
      }

      // === Parsing field penting ===
      function getField(label, fallbackPattern) {
        let val = getValue(label);
        if (!val && fallbackPattern) val = findByPattern(fallbackPattern);
        return val || '';
      }

      // OWNER
      let owner = getValue('OWNER');
      if (!owner) {
        const ownerLine = lines.find(l => /OWNER\s+\w+/i.test(l));
        if (ownerLine) {
          owner = ownerLine.split('OWNER')[1].trim();
        } else {
          if (inputText.toUpperCase().includes('BGES')) owner = 'BGES';
          else if (inputText.toUpperCase().includes('WMS')) owner = 'WMS';
          else if (inputText.toUpperCase().includes('TSEL')) owner = 'TSEL';
        }
      }

      const ao = getField('AO', /AO[ :|]+([A-Z0-9]+)/);
      const workorder = getField('WORKORDER', /WORKORDER[ :|]+([A-Z0-9]+)/);
      const serviceNo = getField('SERVICE NO', /SERVICE NO[ :|]+([0-9]+)/);
      const customerName = getField('CUSTOMER NAME', /CUSTOMER NAME[ :|]+(.+)/);
      const workzone = getField('WORKZONE', /WORKZONE[ :|]+([A-Z0-9]+)/);
      const snOnt = getField('SN ONT', /SN ONT[ :|]+([A-Z0-9]+)/);
      const nikOnt = getField('NIK ONT', /NIK ONT[ :|]+([0-9]+)/);
      const stbId = getField('STB ID', /STB ID[ :|]+([A-Z0-9]+)/);
      const nikStb = getField('NIK STB', /NIK STB[ :|]+([0-9]+)/);
      const teknisi = getField('TEKNISI', /TEKNISI[ :|]+(.+)/);

      // Validasi minimal SN ONT dan NIK ONT harus ada
      if (!snOnt || !nikOnt) {
        return sendTelegram(chatId, '❌ Data tidak lengkap. Minimal harus ada SN ONT dan NIK ONT.');
      }

      // Susun data sesuai urutan kolom sheet Anda
      const now = new Date();
      const tanggal = now.toLocaleString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      const row = [
        tanggal,     // TANGGAL
        ao,          // AO
        workorder,   // WORKORDER
        serviceNo,   // SERVICE NO
        customerName,// CUSTOMER NAME
        owner,       // OWNER
        workzone,    // WORKZONE
        snOnt,       // SN ONT
        nikOnt,      // NIK ONT
        stbId,       // STB ID
        nikStb,      // NIK STB
        teknisi      // TEKNISI
      ];

      // Simpan ke sheet
      await appendSheetData(REKAPAN_SHEET, row);

      // Balasan singkat
      return sendTelegram(chatId, '✅ Data berhasil disimpan ke sheet, GASPOLLL 🚀🚀!');
    }
    // === Default: Help ===
    return sendTelegram(chatId, `🤖 Bot aktif. Command:\n/ps\n/cari <SN>\n/allps\n/nik <NIK>\n/aktivasi <SN> <NIK> <KETERANGAN>\n/myid`);
  } catch (err) {
    console.error(err);
    return sendTelegram(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi.');
  }
});


console.log('Bot Telegram Rekapan aktif!');

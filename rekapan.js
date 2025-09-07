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

// === Helper: Kirim pesan Telegram (otomatis split jika >4000 char) ===
function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  if (text.length <= maxLength) {
    return bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
  } else {
    // Split by line, try not to break in the middle of a line
    const lines = text.split('\n');
    let chunk = '';
    let promises = [];
    for (let i = 0; i < lines.length; i++) {
      if ((chunk + lines[i] + '\n').length > maxLength) {
        promises.push(bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options }));
        chunk = '';
      }
      chunk += lines[i] + '\n';
    }
    if (chunk.trim()) promises.push(bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options }));
    return Promise.all(promises);
  }
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


// === Handler pesan masuk dengan fitur lengkap ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const chatType = msg.chat.type;

  try {
    // === Hanya proses /aktivasi di group, command lain diabaikan ===
    if ((chatType === 'group' || chatType === 'supergroup') && !/^\/aktivasi\b/i.test(text)) {
      return;
    }

    // === /ps: Laporan harian detail ===
    if (/^\/ps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /ps hanya untuk admin.');
      }
      const data = await getSheetData(REKAPAN_SHEET);
      const today = new Date();
      const todayStr = today.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      let total = 0;
      let teknisiMap = {}, workzoneMap = {}, ownerMap = {};
      let details = [];
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
          details.push({ teknisi, workzone, owner });
        }
      }
      let msg = `📊 <b>LAPORAN AKTIVASI HARIAN</b>\nTanggal: ${todayStr}\nTotal Aktivasi: ${total} SSL\n\n`;
      msg += `METRICS:\n- Teknisi Aktif: ${Object.keys(teknisiMap).length}\n- Workzone Tercover: ${Object.keys(workzoneMap).length}\n- Owner: ${Object.keys(ownerMap).length}\n\n`;
      msg += 'PERFORMA TEKNISI:\n';
      Object.entries(teknisiMap).sort((a,b)=>b[1]-a[1]).forEach(([t,c],i)=>{msg+=`${i+1}. ${t}: ${c} SSL\n`;});
      msg += '\nPERFORMA WORKZONE:\n';
      Object.entries(workzoneMap).sort((a,b)=>b[1]-a[1]).forEach(([w,c],i)=>{msg+=`${i+1}. ${w}: ${c} SSL\n`;});
      msg += '\nPERFORMA OWNER:\n';
      Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c],i)=>{msg+=`${i+1}. ${o}: ${c} SSL\n`;});
      msg += `\nDATA SOURCE: REKAPAN_QUALITY\nGENERATED: ${today.toLocaleString('id-ID')} WIB`;
      return sendTelegram(chatId, msg);
    }

    // === /allps: breakdown owner, sektor, top teknisi ===
    if (/^\/allps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Akses ditolak. Command /allps hanya untuk admin.');
      }
      const data = await getSheetData(REKAPAN_SHEET);
      let total = data.length - 1;
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
      Object.entries(ownerMap).sort((a,b)=>b[1]-a[1]).forEach(([o,c])=>{msg+=`- ${o}: ${c}\n`;});
      msg += '\nBERDASARKAN SEKTOR/WORKZONE:\n';
      Object.entries(sektorMap).sort((a,b)=>b[1]-a[1]).forEach(([s,c])=>{msg+=`- ${s}: ${c}\n`;});
      let teknisiArr = Object.entries(teknisiMap).map(([name,count])=>({name,count}));
      teknisiArr.sort((a,b)=>b.count-a.count);
      msg += '\nTOP TEKNISI:\n';
      teknisiArr.slice(0,5).forEach((t,i)=>{msg+=`${i+1}. ${t.name}: ${t.count}\n`;});
      return sendTelegram(chatId, msg);
    }

    // === /nik <NIK>: statistik berdasarkan NIK ===
    if (/^\/nik\b/i.test(text)) {
      const nik = text.split(' ')[1];
      if (!nik) return sendTelegram(chatId, 'Format: /nik <NIK>');
      const data = await getSheetData(REKAPAN_SHEET);
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][8] || '').toUpperCase() === nik.toUpperCase()) count++;
      }
      return sendTelegram(chatId, `NIK <b>${nik}</b> ditemukan pada <b>${count}</b> data.`);
    }

    // === /cari <SN>: cari data berdasarkan SN ONT ===
    if (/^\/cari\b/i.test(text)) {
      const sn = text.split(' ')[1];
      if (!sn) return sendTelegram(chatId, 'Format: /cari <SN>');
      const data = await getSheetData(REKAPAN_SHEET);
      let found = [];
      for (let i = 1; i < data.length; i++) {
        if ((data[i][7] || '').toUpperCase() === sn.toUpperCase()) {
          found.push(data[i]);
        }
      }
      if (found.length === 0) return sendTelegram(chatId, `SN <b>${sn}</b> tidak ditemukan.`);
      let msgText = `📄 <b>Data SN ${sn}:</b>\n`;
      found.forEach(row => {
        msgText += `Tgl: ${row[0]}, User: ${row[1]}, Status: ${row[6] || '-'}\n`;
      });
      return sendTelegram(chatId, msgText);
    }

    // === /stat <teknisi>: statistik teknisi per nama/username ===
    if (/^\/stat\b/i.test(text)) {
      const param = text.split(' ')[1];
      if (!param) return sendTelegram(chatId, 'Format: /stat <nama_teknisi>');
      const data = await getSheetData(REKAPAN_SHEET);
      let total = 0, ownerMap = {}, sektorMap = {}, detailArr = [];
      for (let i = 1; i < data.length; i++) {
        const teknisi = (data[i][11] || '').toLowerCase();
        if (teknisi.includes(param.toLowerCase())) {
          total++;
          const owner = (data[i][5] || '-').toUpperCase();
          const sektor = (data[i][6] || '-').toUpperCase();
          ownerMap[owner] = (ownerMap[owner] || 0) + 1;
          sektorMap[sektor] = (sektorMap[sektor] || 0) + 1;
          detailArr.push(data[i]);
        }
      }
      let msg = `📊 STATISTIK TEKNISI\n👤 Teknisi: ${param}\n📈 Total Aktivasi: ${total} SSL\n\nDETAIL PER OWNER:\n`;
      Object.entries(ownerMap).forEach(([o,c])=>{msg+=`- ${o}: ${c}\n`;});
      msg += '\nDETAIL PER SEKTOR:\n';
      Object.entries(sektorMap).forEach(([s,c])=>{msg+=`- ${s}: ${c}\n`;});
      msg += '\nStatus: AKTIF\n';
      msg += `Updated: ${(new Date()).toLocaleString('id-ID')} WIB`;
      return sendTelegram(chatId, msg);
    }

    // === /aktivasi: parsing multi-format, cek duplikat, simpan ===
    if (/^\/aktivasi\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.');
      const inputText = text.replace(/^\/aktivasi\s*/i, '').trim();
      if (!inputText) return sendTelegram(chatId, 'Silakan kirim data aktivasi setelah /aktivasi.');

      // === Parsing multi-format (BGES, WMS, TSEL, default) ===
      function parseAktivasi(text, username) {
        const lines = text.split('\n').map(l=>l.trim()).filter(l=>l);
        const upper = text.toUpperCase();
        let ao='', workorder='', serviceNo='', customerName='', owner='', workzone='', snOnt='', nikOnt='', stbId='', nikStb='', teknisi='';
        teknisi = user[1] || username;
        // BGES
        if (upper.includes('BGES')) {
          owner = 'BGES';
          ao = lines.find(l=>/SC\d+/i.test(l))?.match(/SC(\d+)/i)?.[1] || '';
          serviceNo = lines.find(l=>/SERVICE NO/i.test(l))?.split(':')[1]?.trim() || '';
          snOnt = lines.find(l=>/SN ONT/i.test(l))?.split(':')[1]?.trim() || '';
          nikOnt = lines.find(l=>/NIK ONT/i.test(l))?.split(':')[1]?.trim() || '';
          customerName = lines.find(l=>/CUSTOMER NAME/i.test(l))?.split(':')[1]?.trim() || '';
          workzone = lines.find(l=>/WORKZONE/i.test(l))?.split(':')[1]?.trim() || '';
        } else if (upper.includes('WMS')) {
          owner = 'WMS';
          ao = lines.find(l=>/AO/i.test(l))?.split(':')[1]?.trim() || '';
          serviceNo = lines.find(l=>/SERVICE NO/i.test(l))?.split(':')[1]?.trim() || '';
          snOnt = lines.find(l=>/SN ONT/i.test(l))?.split(':')[1]?.trim() || '';
          nikOnt = lines.find(l=>/NIK ONT/i.test(l))?.split(':')[1]?.trim() || '';
          customerName = lines.find(l=>/CUSTOMER NAME/i.test(l))?.split(':')[1]?.trim() || '';
          workzone = lines.find(l=>/WORKZONE/i.test(l))?.split(':')[1]?.trim() || '';
        } else if (upper.includes('TSEL')) {
          owner = 'TSEL';
          ao = lines.find(l=>/AO/i.test(l))?.split(':')[1]?.trim() || '';
          serviceNo = lines.find(l=>/SERVICE NO/i.test(l))?.split(':')[1]?.trim() || '';
          snOnt = lines.find(l=>/SN ONT/i.test(l))?.split(':')[1]?.trim() || '';
          nikOnt = lines.find(l=>/NIK ONT/i.test(l))?.split(':')[1]?.trim() || '';
          customerName = lines.find(l=>/CUSTOMER NAME/i.test(l))?.split(':')[1]?.trim() || '';
          workzone = lines.find(l=>/WORKZONE/i.test(l))?.split(':')[1]?.trim() || '';
        } else {
          // fallback: ambil label
          function getValue(label) {
            const line = lines.find(l => l.toUpperCase().startsWith(label.toUpperCase() + ' :'));
            return line ? line.split(':').slice(1).join(':').trim() : '';
          }
          ao = getValue('AO');
          workorder = getValue('WORKORDER');
          serviceNo = getValue('SERVICE NO');
          customerName = getValue('CUSTOMER NAME');
          owner = getValue('OWNER');
          workzone = getValue('WORKZONE');
          snOnt = getValue('SN ONT');
          nikOnt = getValue('NIK ONT');
          stbId = getValue('STB ID');
          nikStb = getValue('NIK STB');
        }
        return { ao, workorder, serviceNo, customerName, owner, workzone, snOnt, nikOnt, stbId, nikStb, teknisi };
      }

      const parsed = parseAktivasi(inputText, username);
      // Validasi minimal SN ONT dan NIK ONT harus ada
      if (!parsed.snOnt || !parsed.nikOnt) {
        return sendTelegram(chatId, '❌ Data tidak lengkap. Minimal harus ada SN ONT dan NIK ONT.');
      }

      // === Cek duplikat: SN ONT dan NIK ONT sudah ada di sheet ===
      const data = await getSheetData(REKAPAN_SHEET);
      let isDuplicate = false;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][7] || '').toUpperCase() === parsed.snOnt.toUpperCase() && (data[i][8] || '').toUpperCase() === parsed.nikOnt.toUpperCase()) {
          isDuplicate = true;
          break;
        }
      }
      if (isDuplicate) {
        return sendTelegram(chatId, '❌ Data duplikat. SN ONT dan NIK ONT sudah pernah diinput.');
      }

      // Susun data sesuai urutan kolom sheet
      const now = new Date();
      const tanggal = now.toLocaleString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      const row = [
        tanggal,     // TANGGAL
        parsed.ao,          // AO
        parsed.workorder,   // WORKORDER
        parsed.serviceNo,   // SERVICE NO
        parsed.customerName,// CUSTOMER NAME
        parsed.owner,       // OWNER
        parsed.workzone,    // WORKZONE
        parsed.snOnt,       // SN ONT
        parsed.nikOnt,      // NIK ONT
        parsed.stbId,       // STB ID
        parsed.nikStb,      // NIK STB
        parsed.teknisi      // TEKNISI
      ];
      await appendSheetData(REKAPAN_SHEET, row);
      return sendTelegram(chatId, '✅ Data berhasil disimpan ke sheet, GASPOLLL 🚀🚀!');
    }

    // === /testparsing: fungsi testing parsing (opsional) ===
    if (/^\/testparsing\b/i.test(text)) {
      let testMsg = `/aktivasi\nOWNER : BGES\nAO : SC123456\nSERVICE NO : 9876543210\nCUSTOMER NAME : PT TEST\nWORKZONE : ZONE1\nSN ONT : ZTEGDA140D99\nNIK ONT : 12345678`;
      let parsed = (()=>{
        const lines = testMsg.split('\n').map(l=>l.trim()).filter(l=>l);
        let ao = lines.find(l=>/SC\d+/i.test(l))?.match(/SC(\d+)/i)?.[1] || '';
        let snOnt = lines.find(l=>/SN ONT/i.test(l))?.split(':')[1]?.trim() || '';
        let nikOnt = lines.find(l=>/NIK ONT/i.test(l))?.split(':')[1]?.trim() || '';
        return { ao, snOnt, nikOnt };
      })();
      return sendTelegram(chatId, `Test parsing:\nAO: ${parsed.ao}\nSN ONT: ${parsed.snOnt}\nNIK ONT: ${parsed.nikOnt}`);
    }

    // === Default: Help ===
    return sendTelegram(chatId, `🤖 Bot aktif. Command:\n/ps\n/cari <SN>\n/allps\n/nik <NIK>\n/aktivasi <SN> <NIK> <KETERANGAN>\n/stat <teknisi>\n/testparsing`);
  } catch (err) {
    console.error(err);
    return sendTelegram(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi.');
  }
});


console.log('Bot Telegram Rekapan aktif!');

require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === ENV VARIABLES ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GROUP_CHAT_ID = process.env.GROUP_CHAT_ID;
let GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

if (!TOKEN) { console.error('❌ TELEGRAM_TOKEN not set'); process.exit(1); }
if (!SHEET_ID) { console.error('❌ SHEET_ID not set'); process.exit(1); }
if (!GOOGLE_SERVICE_ACCOUNT_JSON) { console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON not set'); process.exit(1); }
if (!GROUP_CHAT_ID) { console.warn('⚠️ GROUP_CHAT_ID not set - TTR alerts disabled'); }

// === PARSE GOOGLE SERVICE ACCOUNT ===
let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  if (!keyData.startsWith('{')) {
    try { keyData = Buffer.from(keyData, 'base64').toString('utf-8'); } catch (e) { }
  }
  serviceAccount = JSON.parse(keyData);
  console.log('✅ Google Service Account parsed');
} catch (e) {
  console.error('❌ Failed to parse JSON:', e.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// === CONSTANTS ===
const ASSURANCE_SHEET = 'PROGRES ASSURANCE';
const ORDER_ASSURANCE_SHEET = 'ORDER ASSURANCE';
const MASTER_SHEET = 'MASTER';

const TTR_TABLE = {
  'HVC_DIAMOND': 3,
  'FFG': 3,
  'BGES HSI': 4,
  'DATIN K2': 3.6,
  'DATIN K3': 7.2,
  'HVC_PLATINUM': 6,
  'HVC_GOLD': 12,
  'REGULER': 36,
};

const BULAN_ID = {
  'januari': 1, 'februari': 2, 'maret': 3, 'april': 4,
  'mei': 5, 'juni': 6, 'juli': 7, 'agustus': 8,
  'september': 9, 'oktober': 10, 'november': 11, 'desember': 12,
};

// === HELPER: Auto-detect ORDER ASSURANCE column indices from header ===
function getOrderColumns(data) {
  const defaults = { incident: 1, teknisi: 2, ttrCustomer: 3, workzone: 4, customerType: 5, status: 9, hasilUkur: -1 };
  if (!data || data.length < 1) return defaults;

  const header = data[0] || [];
  const cols = { ...defaults };

  for (let c = 0; c < header.length; c++) {
    const h = (header[c] || '').toUpperCase().trim();
    if (h === 'INCIDENT' || h === 'NO INCIDENT') cols.incident = c;
    else if (h === 'TEKNISI') cols.teknisi = c;
    else if (h.includes('TTR') && h.includes('CUSTOMER')) cols.ttrCustomer = c;
    else if (h === 'WORKZONE' && c < 10) cols.workzone = c;
    else if (h === 'CUSTOMER TYPE' || h === 'CUSTOMER_TYPE') {
      if (c < 10) cols.customerType = c;
    }
    else if (h === 'STATUS') cols.status = c;
    else if (h === 'HASIL UKUR' || h === 'HASIL_UKUR') cols.hasilUkur = c;
  }

  console.log(`📍 ORDER columns: INC=${cols.incident}, TEKNISI=${cols.teknisi}, TTR=${cols.ttrCustomer}, WZ=${cols.workzone}, CUST_TYPE=${cols.customerType}, STATUS=${cols.status}, HASIL_UKUR=${cols.hasilUkur}`);
  return cols;
}

// === CACHING ===
const cache = {
  masterData: null, masterDataTime: 0,
  assuranceData: null, assuranceDataTime: 0,
  orderAssuranceData: null, orderAssuranceDataTime: 0,
  cacheExpiry: 5 * 60 * 1000,
};

// === STATE ===
const alertState = { warned: new Set(), expired: new Set() };
const userChatIds = {};

// === HELPER: Get sheet data with caching ===
async function getSheetData(sheetName, useCache = true) {
  try {
    if (useCache) {
      if (sheetName === MASTER_SHEET && cache.masterData && Date.now() - cache.masterDataTime < cache.cacheExpiry) {
        return cache.masterData;
      }
      if (sheetName === ASSURANCE_SHEET && cache.assuranceData && Date.now() - cache.assuranceDataTime < cache.cacheExpiry) {
        return cache.assuranceData;
      }
      if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData && Date.now() - cache.orderAssuranceDataTime < cache.cacheExpiry) {
        return cache.orderAssuranceData;
      }
    }

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    const data = res.data.values || [];

    if (sheetName === MASTER_SHEET) { cache.masterData = data; cache.masterDataTime = Date.now(); }
    else if (sheetName === ASSURANCE_SHEET) { cache.assuranceData = data; cache.assuranceDataTime = Date.now(); }
    else if (sheetName === ORDER_ASSURANCE_SHEET) { cache.orderAssuranceData = data; cache.orderAssuranceDataTime = Date.now(); }

    return data;
  } catch (error) {
    console.error(`Error reading ${sheetName}:`, error.message);
    if (sheetName === MASTER_SHEET && cache.masterData) return cache.masterData;
    if (sheetName === ASSURANCE_SHEET && cache.assuranceData) return cache.assuranceData;
    if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData) return cache.orderAssuranceData;
    throw error;
  }
}

// === HELPER: Append to sheet ===
async function appendSheetData(sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

// === HELPER: Update a single cell ===
async function updateSheetCell(sheetName, cell, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${sheetName}'!${cell}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [[value]] },
  });
}

// === HELPER: Batch update multiple cells (preserves formulas in skipped columns) ===
async function batchUpdateCells(sheetName, updates) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: {
      valueInputOption: 'USER_ENTERED',
      data: updates.map(u => ({
        range: `'${sheetName}'!${u.range}`,
        values: [[u.value]],
      })),
    },
  });
}

// === HELPER: Find next row with empty incident column ===
function findNextEmptyRow(data, incidentColIdx) {
  for (let i = 1; i < data.length; i++) {
    if (!(data[i][incidentColIdx] || '').trim()) {
      return i + 1; // 1-based row number
    }
  }
  return data.length + 1; // Append at end if no empty row found
}

// === HELPER: Send Telegram (with chunking) ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  try {
    if (text.length <= maxLength) {
      return await bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
    }
    const lines = text.split('\n');
    let chunk = '';
    for (const line of lines) {
      if ((chunk + line + '\n').length > maxLength) {
        await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
        chunk = '';
      }
      chunk += line + '\n';
    }
    if (chunk.trim()) {
      await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
    }
  } catch (error) {
    console.error('Error sending message:', error.message);
  }
}

// === HELPER: Timeout wrapper ===
function withTimeout(promise, ms = 10000) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout - Google API too slow')), ms)),
  ]);
}

// === HELPER: Get user role from MASTER ===
async function getUserRole(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      if (sheetUser === inputUser && status === 'AKTIF') return role;
    }
    return null;
  } catch (error) {
    console.error('Error getting user role:', error.message);
    return null;
  }
}

// === HELPER: Check authorization ===
async function checkAuthorization(username, requiredRoles = []) {
  try {
    const userRole = await withTimeout(getUserRole(username), 8000);
    if (!userRole) return { authorized: false, role: null, message: '❌ Anda tidak terdaftar di sistem.' };
    if (requiredRoles.length > 0 && !requiredRoles.includes(userRole))
      return { authorized: false, role: userRole, message: `❌ Akses ditolak. Role ${userRole} tidak memiliki izin.` };
    return { authorized: true, role: userRole };
  } catch (error) {
    return { authorized: false, role: null, message: '❌ Terjadi kesalahan saat verifikasi.' };
  }
}

// === HELPER: Get active admins from MASTER ===
async function getActiveAdmins() {
  try {
    const data = await getSheetData(MASTER_SHEET);
    const admins = [];
    for (let i = 1; i < data.length; i++) {
      const uname = (data[i][8] || '').replace('@', '').trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      if (role === 'ADMIN' && status === 'AKTIF' && uname) admins.push(uname);
    }
    return admins;
  } catch (error) {
    console.error('Error getting admins:', error.message);
    return [];
  }
}

// === HELPER: Get workzone mappings from ORDER ASSURANCE (auto-detect columns) ===
function getWorkzoneMappings(data) {
  if (!data || data.length < 2) return [];

  // Auto-detect MAPPING TEAM and WORKZONE columns from header
  const header = data[0] || [];
  let teamColIdx = -1;
  let wzColIdx = -1;

  for (let c = 0; c < header.length; c++) {
    const h = (header[c] || '').toUpperCase().trim();
    if (h.includes('MAPPING') && h.includes('TEAM')) teamColIdx = c;
    // Only match WORKZONE columns after column 10 (to skip main WORKZONE at col E)
    if (c > 10 && h.includes('WORKZONE')) wzColIdx = c;
  }

  // Fallback to Q(16) and R(17) if not found
  if (teamColIdx === -1) teamColIdx = 16;
  if (wzColIdx === -1) wzColIdx = teamColIdx + 1;

  console.log(`📍 Mapping columns detected: MAPPING TEAM=col ${teamColIdx}, WORKZONE=col ${wzColIdx}`);

  const mappings = [];
  for (let i = 1; i < data.length; i++) {
    const team = (data[i][teamColIdx] || '').trim();
    const wz = (data[i][wzColIdx] || '').trim();
    if (team && wz) mappings.push({ team, workzone: wz });
  }

  console.log(`📍 Found ${mappings.length} workzone mappings`);

  // Remove duplicates
  const seen = new Set();
  return mappings.filter(m => {
    const key = `${m.team}|${m.workzone}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// === HELPER: Find mapping team for a workzone ===
function findMappingTeam(workzone, status, mappings) {
  if (!workzone && status !== 'GAMAS') return null;
  if (status === 'GAMAS') {
    const allMap = mappings.find(m => m.workzone.toUpperCase() === 'ALL');
    return allMap ? allMap.team : null;
  }
  for (const mapping of mappings) {
    if (mapping.workzone.toUpperCase() === 'ALL') continue;
    const zones = mapping.workzone.split(/[&,]/).map(z => z.trim().toUpperCase());
    if (zones.includes(workzone.toUpperCase())) return mapping.team;
  }
  return null;
}

// === HELPER: Parse TTR duration string to hours ===
function parseTTRHours(ttrStr) {
  if (!ttrStr || ttrStr === '-' || ttrStr === '') return null;
  const parts = ttrStr.toString().trim().split(':');
  if (parts.length >= 2) {
    const h = parseFloat(parts[0]) || 0;
    const m = parseFloat(parts[1]) || 0;
    const s = parts.length > 2 ? (parseFloat(parts[2]) || 0) : 0;
    return h + m / 60 + s / 3600;
  }
  const num = parseFloat(ttrStr);
  return isNaN(num) ? null : num;
}

// === HELPER: Format hours to HH:MM:SS ===
function formatHours(hours) {
  const h = Math.floor(hours);
  const m = Math.floor((hours - h) * 60);
  const s = Math.floor(((hours - h) * 60 - m) * 60);
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

// === HELPER: Parse Indonesian date ===
function parseIndonesianDate(dateStr) {
  if (!dateStr) return null;
  const cleaned = dateStr.replace(/^[^,]*,\s*/, '').trim();
  const parts = cleaned.split(/\s+/);
  if (parts.length < 3) return null;
  const day = parseInt(parts[0]);
  const month = BULAN_ID[parts[1].toLowerCase()];
  const year = parseInt(parts[2]);
  if (!day || !month || !year) return null;
  return { day, month, year };
}

// === HELPER: Get today in Jakarta timezone ===
function getTodayJakarta() {
  const now = new Date();
  return {
    day: parseInt(now.toLocaleDateString('id-ID', { day: 'numeric', timeZone: 'Asia/Jakarta' })),
    month: parseInt(now.toLocaleDateString('id-ID', { month: 'numeric', timeZone: 'Asia/Jakarta' })),
    year: parseInt(now.toLocaleDateString('id-ID', { year: 'numeric', timeZone: 'Asia/Jakarta' })),
  };
}

// === HELPER: Parse assurance input text ===
function parseAssurance(text, username) {
  let data = {
    incidentNo: '', closeDesc: '',
    dropcore: '', patchcord: '', soc: '', pslave: '',
    passive1_8: '', passive1_4: '', pigtail: '', adaptor: '',
    roset: '', rj45: '', lan: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  const incidentMatch = text.match(/INC[0-9]+/i);
  if (incidentMatch) data.incidentNo = incidentMatch[0].trim().toUpperCase();

  const closeMatch = text.match(/CLOSE\s*:\s*(.+?)(?=\n|MATERIAL|$)/i);
  if (closeMatch && closeMatch[1]) data.closeDesc = closeMatch[1].trim();

  const patterns = {
    dropcore: /DROPCORE\s*:\s*([0-9\.]+)/i, patchcord: /PATCHCORD\s*:\s*([0-9\.]+)/i,
    soc: /SOC\s*:\s*([0-9\.]+)/i, pslave: /PSLAVE\s*:\s*([0-9\.]+)/i,
    passive1_8: /PASSIVE\s*1\/8\s*:\s*([0-9\.]+)/i, passive1_4: /PASSIVE\s*1\/4\s*:\s*([0-9\.]+)/i,
    pigtail: /PIGTAIL\s*:\s*([0-9\.]+)/i, adaptor: /ADAPTOR\s*:\s*([0-9\.]+)/i,
    roset: /ROSET\s*:\s*([0-9\.]+)/i, rj45: /RJ\s*45\s*:\s*([0-9\.]+)/i,
    lan: /LAN\s*:\s*([0-9\.]+)/i,
  };
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) data[key] = match[1].trim();
  }
  return data;
}

// === HELPER: Parse tiket baru message ===
function parseTicketBaru(text) {
  if (!text.includes('Data Tiket Baru Diterima')) return null;

  const result = {};
  const patterns = {
    teknisi: /Teknisi:\s*(.+)/i,
    incident: /Incident:\s*(INC\d+)/i,
    ttrCustomer: /TTR Customer:\s*(.+)/i,
    reportedDate: /Reported Date:\s*(.+)/i,
    tiketGamas: /Tiket gamas:\s*(.*)/i,
    contactPhone: /Contact Phone:\s*(.+)/i,
    contactName: /Contact name:\s*(.+)/i,
    customerType: /Customer Type:\s*(.+)/i,
    serviceNo: /Service no:\s*(.+)/i,
    odp: /ODP:\s*(.+)/i,
    bookingDate: /Booking date:\s*(.+)/i,
    hasilUkur: /Hasil Ukur:\s*(.+)/i,
    guaranteStatus: /Guarante Status:\s*(.+)/i,
    symtom: /Symtom:\s*(.+)/i,
  };

  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    result[key] = match ? match[1].trim() : '';
  }

  return result;
}

// === HELPER: Extract workzone from ODP string ===
// ODP-SGI-FH/33 FH/D02/33.01 → SGI
function extractWorkzoneFromODP(odp) {
  if (!odp) return '';
  const match = odp.match(/ODP-([A-Z0-9]+)-/i);
  return match ? match[1].toUpperCase() : '';
}



// === BOT SETUP ===
const PORT = process.env.PORT || 3002;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = !!RAILWAY_STATIC_URL;
let bot;

if (USE_WEBHOOK) {
  const express = require('express');
  const app = express();
  app.use(express.json());
  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/assurance${TOKEN}`;
  bot.setWebHook(webhookUrl).then(() => console.log(`✅ Webhook set: ${webhookUrl}`)).catch(err => console.error('❌ Webhook error:', err.message));
  app.post(`/assurance${TOKEN}`, (req, res) => { bot.processUpdate(req.body); res.sendStatus(200); });
  app.get('/', (req, res) => res.send('Bot Assurance is running!'));
  app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
} else {
  bot = new TelegramBot(TOKEN, { polling: { interval: 300, autoStart: true, params: { timeout: 10, allowed_updates: ['message'] } } });
  console.log('✅ Bot running in polling mode');
  bot.on('polling_error', (error) => console.error(error.code === 'EFATAL' ? '❌ Polling fatal:' : '⚠️ Polling error:', error.message));
}

// ============================================================
// MONITORING FUNCTIONS (TTR Alerts + Auto-Fill Teknisi)
// ============================================================

async function autoFillTeknisi() {
  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const cols = getOrderColumns(data);
    const mappings = getWorkzoneMappings(data);
    let filled = 0;

    for (let i = 1; i < data.length; i++) {
      const teknisi = (data[i][cols.teknisi] || '').trim();
      const status = (data[i][cols.status] || '').toUpperCase().trim();
      const workzone = (data[i][cols.workzone] || '').trim();

      if (teknisi || !status || status === 'CLOSE') continue;

      const team = findMappingTeam(workzone, status, mappings);
      if (team) {
        const rowNum = i + 1;
        await updateSheetCell(ORDER_ASSURANCE_SHEET, `C${rowNum}`, team);
        filled++;
        console.log(`🔧 Auto-fill teknisi row ${rowNum}: ${team}`);
      }
    }

    if (filled > 0) {
      cache.orderAssuranceData = null; // Invalidate cache
      console.log(`✅ Auto-fill: ${filled} teknisi filled`);
    }
  } catch (error) {
    console.error('❌ Auto-fill error:', error.message);
  }
}

async function checkTTRAlerts() {
  if (!GROUP_CHAT_ID) return;

  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const cols = getOrderColumns(data);
    const mappings = getWorkzoneMappings(data);
    const admins = await getActiveAdmins();
    const adminTags = admins.map(a => `@${a}`).join(' ');

    let expiredList = [];
    let warningList = [];
    let hasNewExpired = false;
    let hasNewWarning = false;

    for (let i = 1; i < data.length; i++) {
      const incident = (data[i][cols.incident] || '').trim();
      const ttrStr = (data[i][cols.ttrCustomer] || '').trim();
      const workzone = (data[i][cols.workzone] || '').trim();
      const custType = (data[i][cols.customerType] || '').trim().toUpperCase();
      const status = (data[i][cols.status] || '').toUpperCase().trim();

      if (status !== 'OPEN' || !incident || !ttrStr) continue;

      const elapsed = parseTTRHours(ttrStr);
      if (elapsed === null) continue;

      let maxTTR = null;
      for (const [type, hours] of Object.entries(TTR_TABLE)) {
        if (type.toUpperCase() === custType) { maxTTR = hours; break; }
      }
      if (maxTTR === null) continue;

      const team = findMappingTeam(workzone, status, mappings) || (data[i][cols.teknisi] || '-').trim();
      const cleanTeam = team.replace(/@@/g, '@');

      // === EXPIRED ===
      if (elapsed >= maxTTR) {
        const overtime = elapsed - maxTTR;
        expiredList.push({ incident, ttrStr, custType, maxTTR, overtime, team: cleanTeam });
        if (!alertState.expired.has(incident)) {
          hasNewExpired = true;
          alertState.expired.add(incident);
          alertState.warned.delete(incident);
        }
      }
      // === WARNING (1 jam sebelum expired) ===
      else if (elapsed >= maxTTR - 1 && elapsed < maxTTR) {
        const sisa = maxTTR - elapsed;
        warningList.push({ incident, ttrStr, custType, maxTTR, sisa, team: cleanTeam });
        if (!alertState.warned.has(incident)) {
          hasNewWarning = true;
          alertState.warned.add(incident);
        }
      }
    }

    // Kirim alert EXPIRED jika ada ticket baru yang expired
    if (hasNewExpired && expiredList.length > 0) {
      expiredList.sort((a, b) => b.overtime - a.overtime);
      let msg = `🔴 EXPIRED: ${expiredList.length} tickets\n`;
      expiredList.forEach(e => {
        msg += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | OT: +${formatHours(e.overtime)}\n`;
        msg += `  👷 ${e.team}\n`;
      });
      if (adminTags) msg += `\ncc bg ${adminTags}`;
      await sendTelegram(GROUP_CHAT_ID, msg);
      console.log(`🔴 TTR EXPIRED alert: ${expiredList.length} tickets`);
    }

    // Kirim alert WARNING jika ada ticket baru yang mendekati expired
    if (hasNewWarning && warningList.length > 0) {
      warningList.sort((a, b) => a.sisa - b.sisa);
      let msg = `⚠️ MENDEKATI EXPIRED:\n\n`;
      warningList.forEach((e, idx) => {
        msg += `${idx + 1}. ${e.incident}\n`;
        msg += `   ${e.custType} (Max: ${e.maxTTR} Jam)\n`;
        msg += `   Elapsed: ${e.ttrStr} | Sisa: ${formatHours(e.sisa)}\n`;
        msg += `   Teknisi: ${e.team}\n\n`;
      });
      if (adminTags) msg += `cc bg ${adminTags}`;
      await sendTelegram(GROUP_CHAT_ID, msg);
      console.log(`⚠️ TTR WARNING alert: ${warningList.length} tickets`);

      // DM ke setiap teknisi yang tiketnya mendekati expired
      for (const entry of warningList) {
        // Extract usernames dari team (e.g., "@user1 & @user2")
        const usernames = entry.team.split(/\s*&\s*/).map(u => u.replace('@', '').trim().toLowerCase()).filter(u => u);
        for (const uname of usernames) {
          const dmChatId = userChatIds[uname];
          if (dmChatId) {
            let dmMsg = `⚠️ PERINGATAN TTR MENDEKATI EXPIRED!\n\n`;
            dmMsg += `Incident: ${entry.incident}\n`;
            dmMsg += `${entry.custType} (Max: ${entry.maxTTR} Jam)\n`;
            dmMsg += `Elapsed: ${entry.ttrStr} | Sisa: ${formatHours(entry.sisa)}\n`;
            dmMsg += `\nSegera selesaikan tiket ini!`;
            try {
              await sendTelegram(dmChatId, dmMsg);
              console.log(`📩 DM warning sent to @${uname}`);
            } catch (dmErr) {
              console.log(`⚠️ Could not DM @${uname}: ${dmErr.message}`);
            }
          }
        }
      }
    }

    // Cleanup: remove closed incidents from alert state
    const openIncidents = new Set();
    for (let i = 1; i < data.length; i++) {
      const st = (data[i][cols.status] || '').toUpperCase().trim();
      if (st === 'OPEN') openIncidents.add((data[i][cols.incident] || '').trim());
    }
    for (const inc of alertState.warned) { if (!openIncidents.has(inc)) alertState.warned.delete(inc); }
    for (const inc of alertState.expired) { if (!openIncidents.has(inc)) alertState.expired.delete(inc); }

  } catch (error) {
    console.error('❌ TTR check error:', error.message);
  }
}

// === AUTO-POST: Build sisa ticket report (reusable) ===
async function buildSisaTicketReport() {
  const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
  if (!data || data.length < 2) return null;

  const cols = getOrderColumns(data);
  const mappings = getWorkzoneMappings(data);

  const now = new Date();
  const dayNameID = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'][now.getDay()];
  const dateStr = now.toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' }).toUpperCase();
  const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jakarta' });

  let ticketsByTeam = {};
  let totalOpen = 0;
  let gamasTickets = {};
  let totalGamas = 0;

  for (let i = 1; i < data.length; i++) {
    const incident = (data[i][cols.incident] || '-').trim();
    const ttrCustomer = (data[i][cols.ttrCustomer] || '-').trim();
    const workzone = (data[i][cols.workzone] || '').trim();
    const custType = (data[i][cols.customerType] || '').trim();
    const hasilUkur = cols.hasilUkur >= 0 ? (data[i][cols.hasilUkur] || '').trim() : '';
    const status = (data[i][cols.status] || '').toUpperCase().trim();

    if (status === 'OPEN') {
      const team = findMappingTeam(workzone, status, mappings) || workzone || '-';
      const cleanTeam = team.replace(/@@/g, '@');
      if (!ticketsByTeam[cleanTeam]) ticketsByTeam[cleanTeam] = [];
      ticketsByTeam[cleanTeam].push({ incident, ttr: ttrCustomer, custType, hasilUkur });
      totalOpen++;
    } else if (status === 'GAMAS') {
      const team = findMappingTeam(workzone, 'GAMAS', mappings) || workzone || '-';
      const cleanTeam = team.replace(/@@/g, '@');
      if (!gamasTickets[cleanTeam]) gamasTickets[cleanTeam] = [];
      gamasTickets[cleanTeam].push({ incident, ttr: ttrCustomer, custType, hasilUkur });
      totalGamas++;
    }
  }

  const numEmoji = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟'];
  const sortedTeams = Object.keys(ticketsByTeam).sort();

  let response = `🔴 <b>SISA TICKET OPEN</b>\n📅 ${dayNameID}, ${dateStr} | ${timeStr}\n\n`;
  response += `🎫 <b>OPEN : ${totalOpen} tickets</b>\n\n`;

  if (sortedTeams.length === 0) {
    response += '<i>Tidak ada ticket yang masih OPEN</i>\n';
  } else {
    sortedTeams.forEach((teamName, idx) => {
      const tickets = ticketsByTeam[teamName].sort((a, b) => (parseTTRHours(a.ttr) || 0) - (parseTTRHours(b.ttr) || 0));
      const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
      response += `${num} <b>${teamName}</b> [${tickets.length}]\n`;
      tickets.forEach(t => {
        const hu = t.hasilUkur ? ` | ${t.hasilUkur}` : '';
        response += `   🔹 ${t.incident} | ${t.ttr} | ${t.custType}${hu}\n`;
      });
      response += '\n';
    });
  }

  if (totalGamas > 0) {
    response += `🎫 <b>GAMAS : ${totalGamas} tickets</b>\n\n`;
    const sortedGamas = Object.keys(gamasTickets).sort();
    sortedGamas.forEach((teamName, idx) => {
      const tickets = gamasTickets[teamName].sort((a, b) => (parseTTRHours(a.ttr) || 0) - (parseTTRHours(b.ttr) || 0));
      const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
      response += `${num} <b>${teamName}</b> [${tickets.length}]\n`;
      tickets.forEach(t => {
        const hu = t.hasilUkur ? ` | ${t.hasilUkur}` : '';
        response += `   🔹 ${t.incident} | ${t.ttr} | ${t.custType}${hu}\n`;
      });
      response += '\n';
    });
  }

  return response;
}

// === AUTO-POST: Kirim sisa ticket ke group setiap 1 jam ===
async function autoPostSisaTicket() {
  if (!GROUP_CHAT_ID) return;
  try {
    const report = await buildSisaTicketReport();
    if (report) {
      await sendTelegram(GROUP_CHAT_ID, report);
      console.log('📊 Auto-post sisa ticket ke group');
    }
  } catch (error) {
    console.error('❌ Auto-post sisa ticket error:', error.message);
  }
}

// ============================================================
// MESSAGE HANDLER
// ============================================================
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const msgId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const groupType = msg.chat.type;

  if (!text) return;

  // Store chat ID for potential DM (from private and group)
  if (username) {
    const unameLower = username.replace('@', '').toLowerCase();
    // Simpan user's Telegram ID untuk DM (msg.from.id selalu user ID, bukan group ID)
    if (msg.from.id) userChatIds[unameLower] = msg.from.id;
  }

  console.log(`📨 [${groupType}] [@${username}] ${text.substring(0, 60)}`);

  try {
    // ============================================================
    // AUTO-INPUT TIKET BARU ke ORDER ASSURANCE
    // Deteksi pesan "Data Tiket Baru Diterima" (non-command)
    // ============================================================
    if (text.includes('Data Tiket Baru Diterima')) {
      try {
        const tiket = parseTicketBaru(text);
        if (!tiket || !tiket.incident) {
          console.log('⚠️ Tiket baru detected but could not parse incident');
          return;
        }

        // Get ORDER ASSURANCE data
        const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
        const orderCols = getOrderColumns(orderData);

        // Cek duplikat incident
        for (let i = 1; i < orderData.length; i++) {
          const existingInc = (orderData[i][orderCols.incident] || '').trim().toUpperCase();
          if (existingInc === tiket.incident.toUpperCase()) {
            console.log(`⚠️ Duplicate incident ${tiket.incident} - skipping`);
            return sendTelegram(chatId, `⚠️ Incident <b>${tiket.incident}</b> sudah ada di ORDER ASSURANCE.`, { reply_to_message_id: msgId });
          }
        }

        // Extract workzone from ODP (e.g., ODP-SGI-FH/33 → SGI)
        const workzone = extractWorkzoneFromODP(tiket.odp);

        // Map teknisi using workzone → MAPPING TEAM
        const mappings = getWorkzoneMappings(orderData);
        const mappedTeknisi = findMappingTeam(workzone, 'OPEN', mappings) || tiket.teknisi;

        // Format tanggal saat ini
        const tanggal = new Date().toLocaleDateString('id-ID', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
        });

        // Format timestamp saat ini (hanya waktu HH:mm:ss)
        const now = new Date();
        const jam = String(now.toLocaleString('id-ID', { hour: '2-digit', timeZone: 'Asia/Jakarta', hour12: false })).padStart(2, '0');
        const menit = String(now.toLocaleString('id-ID', { minute: '2-digit', timeZone: 'Asia/Jakarta' })).padStart(2, '0');
        const detik = String(now.toLocaleString('id-ID', { second: '2-digit', timeZone: 'Asia/Jakarta' })).padStart(2, '0');
        const timestamp = `${jam}:${menit}:${detik}`;

        // Find next row with empty incident (preserves formulas in D and L)
        const nextRow = findNextEmptyRow(orderData, orderCols.incident);

        // Batch update specific cells — SKIP kolom D (rumus) dan L (rumus)
        const cellUpdates = [
          { range: `A${nextRow}`, value: tanggal },              // A - Tanggal
          { range: `B${nextRow}`, value: tiket.incident },       // B - Incident
          { range: `C${nextRow}`, value: mappedTeknisi },        // C - Teknisi (mapped)
          // D - SKIP (rumus TTR CUSTOMER)
          { range: `E${nextRow}`, value: workzone },             // E - Workzone
          { range: `F${nextRow}`, value: 'TSEL' },               // F - Customer Segment
          { range: `G${nextRow}`, value: tiket.customerType },   // G - Customer Type
          { range: `H${nextRow}`, value: tiket.serviceNo },      // H - Service No
          { range: `I${nextRow}`, value: tiket.odp },            // I - Device Name (ODP)
          { range: `J${nextRow}`, value: 'OPEN' },               // J - Status
          { range: `K${nextRow}`, value: tiket.ttrCustomer || timestamp }, // K - TTR Dashboard (durasi, fallback ke timestamp jika kosong)
          // L - SKIP (rumus NOW)
          { range: `M${nextRow}`, value: timestamp },            // M - Timestamp
        ];

        await withTimeout(batchUpdateCells(ORDER_ASSURANCE_SHEET, cellUpdates), 10000);
        cache.orderAssuranceData = null; // Invalidate cache

        console.log(`✅ Tiket baru recorded at row ${nextRow}: ${tiket.incident} | Teknisi: ${mappedTeknisi} | WZ: ${workzone}`);

        let confirmMsg = `✅ <b>Data Tiket Baru berhasil disimpan!</b> (Row ${nextRow})\n\n`;
        confirmMsg += `📋 <b>Incident:</b> ${tiket.incident}\n`;
        confirmMsg += `👷 <b>Teknisi:</b> ${mappedTeknisi}\n`;
        confirmMsg += `📍 <b>Workzone:</b> ${workzone}`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ Tiket Baru Error:', err.message);
        return sendTelegram(chatId, `❌ Error menyimpan tiket: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // Skip non-command messages after tiket baru check
    if (!text.startsWith('/')) return;
    // ============================================================
    // /INPUT - Input data assurance + auto-close di ORDER ASSURANCE
    // ============================================================
    if (/^\/INPUT\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/INPUT\s*/i, '').trim();
        if (!inputText) return sendTelegram(chatId, '❌ Silakan kirim data assurance setelah /INPUT.', { reply_to_message_id: msgId });

        const parsed = parseAssurance(inputText, username);
        const missing = ['incidentNo', 'closeDesc'].filter(f => !parsed[f]);
        if (missing.length > 0) return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });

        // Simpan ke PROGRES ASSURANCE (termasuk timestamp di kolom P)
        const inputTimestamp = new Date().toLocaleString('id-ID', {
          day: '2-digit', month: '2-digit', year: 'numeric',
          hour: '2-digit', minute: '2-digit', second: '2-digit',
          timeZone: 'Asia/Jakarta', hour12: false,
        });
        const row = [
          parsed.dateCreated, parsed.incidentNo,
          parsed.dropcore, parsed.patchcord, parsed.soc, parsed.pslave,
          parsed.passive1_8, parsed.passive1_4, parsed.pigtail, parsed.adaptor,
          parsed.roset, parsed.rj45, parsed.lan, parsed.closeDesc, parsed.teknisi,
          inputTimestamp, // P - Timestamp saat /INPUT
        ];
        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);
        cache.assuranceData = null; // Invalidate cache agar /rekap_hari langsung update

        // Auto-close di ORDER ASSURANCE + cek KAWAL TTR (COMPLY/NOT COMPLY)
        let orderClosed = false;
        let kawalTTR = '';
        try {
          const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
          const orderCols = getOrderColumns(orderData);
          const statusColLetter = String.fromCharCode(65 + orderCols.status);
          for (let i = 1; i < orderData.length; i++) {
            const incInOrder = (orderData[i][orderCols.incident] || '').trim().toUpperCase();
            if (incInOrder === parsed.incidentNo) {
              // Cek TTR untuk KAWAL TTR
              const ttrStr = (orderData[i][orderCols.ttrCustomer] || '').trim();
              const custType = (orderData[i][orderCols.customerType] || '').trim().toUpperCase();
              const elapsed = parseTTRHours(ttrStr);

              // Cari max TTR dari tabel berdasarkan customer type
              let maxTTR = null;
              for (const [type, hours] of Object.entries(TTR_TABLE)) {
                if (type.toUpperCase() === custType) { maxTTR = hours; break; }
              }

              // Tentukan COMPLY atau NOT COMPLY
              if (elapsed !== null && maxTTR !== null) {
                kawalTTR = elapsed <= maxTTR ? 'COMPLY' : 'NOT COMPLY';
              } else {
                kawalTTR = 'COMPLY'; // Default jika data tidak lengkap
              }

              // Update STATUS ke CLOSE dan KAWAL TTR ke kolom R
              await batchUpdateCells(ORDER_ASSURANCE_SHEET, [
                { range: `${statusColLetter}${i + 1}`, value: 'CLOSE' },
                { range: `R${i + 1}`, value: kawalTTR },
              ]);
              cache.orderAssuranceData = null;
              alertState.warned.delete(parsed.incidentNo);
              alertState.expired.delete(parsed.incidentNo);
              orderClosed = true;
              console.log(`✅ Auto-close ORDER: ${parsed.incidentNo} row ${i + 1} | KAWAL TTR: ${kawalTTR}`);
              break;
            }
          }
        } catch (closeErr) {
          console.error('⚠️ Auto-close error:', closeErr.message);
        }

        let confirmMsg = `✅ Data Assurance berhasil disimpan!\n\n`;
        confirmMsg += `<b>Incident:</b> ${parsed.incidentNo}\n`;
        confirmMsg += `<b>Close:</b> ${parsed.closeDesc}\n`;
        if (orderClosed) confirmMsg += `<b>Status ORDER:</b> ✅ Auto-CLOSE | <b>KAWAL TTR:</b> ${kawalTTR}\n`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /INPUT Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /sisa_ticket - Ticket OPEN grouped by mapping team
    // ============================================================
    else if (/^\/sisa_ticket\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const response = await withTimeout(buildSisaTicketReport(), 10000);
        if (!response) return sendTelegram(chatId, '<i>Tidak ada data</i>', { reply_to_message_id: msgId });

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /sisa_ticket Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /cek_ttr - Cek TTR warning & expired secara manual
    // ============================================================
    else if (/^\/cek_ttr\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET, false), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data</i>', { reply_to_message_id: msgId });

        const cols = getOrderColumns(data);
        const mappings = getWorkzoneMappings(data);
        let expiredList = [];
        let warningList = [];
        let safeList = [];

        for (let i = 1; i < data.length; i++) {
          const incident = (data[i][cols.incident] || '').trim();
          const ttrStr = (data[i][cols.ttrCustomer] || '').trim();
          const workzone = (data[i][cols.workzone] || '').trim();
          const custType = (data[i][cols.customerType] || '').trim().toUpperCase();
          const status = (data[i][cols.status] || '').toUpperCase().trim();

          if (status !== 'OPEN' || !incident || !ttrStr) continue;

          const elapsed = parseTTRHours(ttrStr);
          if (elapsed === null) continue;

          let maxTTR = null;
          for (const [type, hours] of Object.entries(TTR_TABLE)) {
            if (type.toUpperCase() === custType) { maxTTR = hours; break; }
          }
          if (maxTTR === null) continue;

          const team = findMappingTeam(workzone, status, mappings) || (data[i][cols.teknisi] || '-').trim();
          const cleanTeam = team.replace(/@@/g, '@');

          const entry = { incident, ttrStr, custType, maxTTR, elapsed, team: cleanTeam };

          if (elapsed >= maxTTR) {
            entry.overtime = elapsed - maxTTR;
            expiredList.push(entry);
          } else if (elapsed >= maxTTR - 1) {
            entry.sisa = maxTTR - elapsed;
            warningList.push(entry);
          } else {
            safeList.push(entry);
          }
        }

        // Sort by overtime/sisa descending
        expiredList.sort((a, b) => b.overtime - a.overtime);
        warningList.sort((a, b) => a.sisa - b.sisa);

        let response = `⏱ <b>STATUS TTR - SEMUA TICKET OPEN</b>\n\n`;

        // EXPIRED
        response += `🔴 <b>EXPIRED: ${expiredList.length} tickets</b>\n`;
        if (expiredList.length === 0) {
          response += '<i>Tidak ada ticket expired</i>\n\n';
        } else {
          expiredList.forEach(e => {
            response += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | OT: +${formatHours(e.overtime)}\n`;
            response += `  👷 ${e.team}\n`;
          });
          response += '\n';
        }

        // WARNING
        response += `⚠️ <b>MENDEKATI EXPIRED: ${warningList.length} tickets</b>\n`;
        if (warningList.length === 0) {
          response += '<i>Tidak ada ticket mendekati expired</i>\n\n';
        } else {
          warningList.forEach(e => {
            response += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | Sisa: ${formatHours(e.sisa)}\n`;
            response += `  👷 ${e.team}\n`;
          });
          response += '\n';
        }

        // SAFE
        response += `✅ <b>AMAN: ${safeList.length} tickets</b>`;

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /cek_ttr Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /material_used - Total material keseluruhan
    // ============================================================
    else if (/^\/material_used\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let materialMap = {
          'DROPCORE': 0, 'PATCHCORD': 0, 'SOC': 0, 'PSLAVE': 0,
          'PASSIVE 1/8': 0, 'PASSIVE 1/4': 0, 'PIGTAIL': 0, 'ADAPTOR': 0,
          'ROSET': 0, 'RJ 45': 0, 'LAN': 0,
        };
        const materialColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        const materialNames = Object.keys(materialMap);

        for (let i = 1; i < data.length; i++) {
          materialColumns.forEach((colIdx, idx) => {
            materialMap[materialNames[idx]] += parseInt((data[i][colIdx] || '0').trim()) || 0;
          });
        }

        const entries = Object.entries(materialMap).filter(([_, c]) => c > 0).sort((a, b) => b[1] - a[1]);
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📦 <b>PENGGUNAAN MATERIAL</b>\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada material yang dipakai</i>';
        } else {
          entries.forEach(([mat, count]) => { response += `📊 <b>${mat}</b> : ${count} unit\n`; });
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /material_used Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_hari - Rekap close per teknisi HARI INI
    // ============================================================
    else if (/^\/rekap_hari\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();
        let map = {};
        let incidents = [];

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.day === today.day && d.month === today.month && d.year === today.year) {
            map[teknisi] = (map[teknisi] || 0) + 1;
            if (incident) incidents.push(incident.toUpperCase());
          }
        }

        // Hitung COMPLY/NOT COMPLY dari ORDER ASSURANCE kolom R
        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase(); // Kolom R = index 17
          if (incidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const now = new Date();
        const todayStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
        const total = entries.reduce((sum, [_, c]) => sum + c, 0);

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - HARI INI</b>\n📅 ${todayStr}\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada data hari ini</i>';
        } else {
          entries.forEach(([tek, c], i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${tek}</b> : ${c} tickets\n`;
          });
          response += `\n📋 <b>Total: ${total} tickets</b>\n`;
          response += `✅ COMPLY: ${comply} | ❌ NOT COMPLY: ${notComply}`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rekap_hari Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_bulan - Rekap close per teknisi BULAN INI
    // ============================================================
    else if (/^\/rekap_bulan\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();
        let map = {};
        let incidents = [];

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.month === today.month && d.year === today.year) {
            map[teknisi] = (map[teknisi] || 0) + 1;
            if (incident) incidents.push(incident.toUpperCase());
          }
        }

        // Hitung COMPLY/NOT COMPLY
        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase();
          if (incidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const now = new Date();
        const bulanStr = now.toLocaleDateString('id-ID', { month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
        const total = entries.reduce((sum, [_, c]) => sum + c, 0);

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - BULAN INI</b>\n📅 ${bulanStr}\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada data bulan ini</i>';
        } else {
          entries.forEach(([tek, c], i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${tek}</b> : ${c} tickets\n`;
          });
          response += `\n📋 <b>Total: ${total} tickets</b>\n`;
          response += `✅ COMPLY: ${comply} | ❌ NOT COMPLY: ${notComply}`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rekap_bulan Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_tahun - Rekap close per teknisi TAHUN INI
    // ============================================================
    else if (/^\/rekap_tahun\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();

        // Group by month, then by teknisi
        const bulanNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        let monthData = {}; // { month: { teknisi: count } }
        let allIncidents = [];
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.year === today.year) {
            if (!monthData[d.month]) monthData[d.month] = {};
            monthData[d.month][teknisi] = (monthData[d.month][teknisi] || 0) + 1;
            grandTotal++;
            if (incident) allIncidents.push(incident.toUpperCase());
          }
        }

        // Hitung COMPLY/NOT COMPLY total
        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase();
          if (allIncidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - TAHUN ${today.year}</b>\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;

        const sortedMonths = Object.keys(monthData).map(Number).sort((a, b) => a - b);
        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data tahun ini</i>';
        } else {
          sortedMonths.forEach(month => {
            const entries = Object.entries(monthData[month]).sort((a, b) => b[1] - a[1]);
            const monthTotal = entries.reduce((sum, [_, c]) => sum + c, 0);
            response += `📅 <b>${bulanNames[month]} ${today.year}</b> [${monthTotal} tickets]\n`;
            entries.forEach(([tek, c], i) => {
              const icon = i < 3 ? medal[i] : '🔸';
              response += `  ${icon} ${tek} : ${c}\n`;
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tickets</b>\n`;
          response += `✅ COMPLY: ${comply} | ❌ NOT COMPLY: ${notComply}`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rekap_tahun Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /chatid - Get chat ID (untuk setup GROUP_CHAT_ID)
    // ============================================================
    else if (/^\/chatid\b/i.test(text)) {
      return sendTelegram(chatId, `📍 <b>Chat ID:</b> <code>${chatId}</code>\n<b>Type:</b> ${msg.chat.type}`, { reply_to_message_id: msgId });
    }

    // ============================================================
    // /help or /start
    // ============================================================
    else if (/^\/(help|start)\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const helpMsg = `🤖 <b>Bot Assurance</b>

<b>📝 INPUT COMMAND:</b>
/INPUT - Input data assurance (auto-close ORDER)

<b>📊 MONITORING (ADMIN):</b>
/sisa_ticket - Ticket yang masih OPEN
/cek_ttr - Cek TTR warning & expired
/material_used - Total material yang dipakai
/rekap_hari - Rekap close teknisi HARI INI
/rekap_bulan - Rekap close teknisi BULAN INI
/rekap_tahun - Rekap close teknisi TAHUN INI

<b>📋 FORMAT /INPUT:</b>
/INPUT INC47052822
CLOSE: deskripsi perbaikan
DROPCORE: 0
PATCHCORD: 0
SOC: 0
PSLAVE: 2
PASSIVE 1/8: 0
PASSIVE 1/4: 0
PIGTAIL: 0
ADAPTOR: 0
ROSET: 0
RJ 45: 0
LAN: 0

<b>⚙️ FITUR OTOMATIS:</b>
• Auto-fill teknisi berdasarkan workzone
• TTR monitoring & alert ke group
• Auto-close status saat /INPUT`;

        return sendTelegram(chatId, helpMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /help Error:', err.message);
        return sendTelegram(chatId, '❌ Terjadi kesalahan.', { reply_to_message_id: msgId });
      }
    }

  } catch (err) {
    console.error('Error:', err.message);
    sendTelegram(chatId, '❌ Terjadi kesalahan sistem.', { reply_to_message_id: msgId });
  }
});

// === ERROR HANDLER ===
process.on('unhandledRejection', (reason) => console.error('Error:', reason));

// === STARTUP ===
console.log('\n🚀 Bot Assurance started!');
console.log(`Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);
console.log('═'.repeat(50));
console.log('✅ Auto-Cache Enabled (5 min expiry)');
console.log('✅ TTR Monitoring Enabled (5 min interval)');
console.log('✅ Auto-Post Sisa Ticket Enabled (1 jam interval)');
console.log('✅ Auto-Fill Teknisi Enabled');
console.log('✅ Timeout Protection Enabled');
console.log('═'.repeat(50));

// Start monitoring after 10 seconds delay
setTimeout(() => {
  console.log('🔄 Starting initial monitoring check...');
  autoFillTeknisi();
  checkTTRAlerts();

  // TTR check + Auto-fill setiap 5 menit
  setInterval(() => {
    autoFillTeknisi();
    checkTTRAlerts();
  }, 5 * 60 * 1000);

  // Auto-post sisa ticket ke group setiap 1 jam
  setInterval(() => {
    autoPostSisaTicket();
  }, 60 * 60 * 1000);
}, 10000);

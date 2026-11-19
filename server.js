'use strict';

const express = require('express');
const http = require('http');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { Server } = require('socket.io');
const ExcelJS = require('exceljs');
const JSZip = require('jszip'); // ExcelJS'nin zaten kullandığı ZIP kütüphanesi

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
  // WebSocket ve polling desteği: Render ve yerel ağ için aynı yapı.
  transports: ['websocket', 'polling'],
  pingInterval: 25000,
  pingTimeout: 20000,
});

const PORT = Number(process.env.PORT) || 3000;
const QUESTION_DURATION_SECONDS_DEFAULT = 20;
const RESULT_DURATION_MS = 5000;
const MAX_QUESTIONS_PER_GAME = 50;
const MAX_PLAYERS_PER_ROOM = 100;

// Fenova etkinlik kayıtları doğrudan kullanıcının masaüstündeki Excel dosyasına yazılır.
// Sunucu Windows bilgisayarda çalışmalıdır. Excel dosyası sunucu çalışırken açık tutulmamalıdır.
const EXCEL_FILE_PATH = 'C:/Users/PC/OneDrive/Desktop/Fenova-Sunucu-Kayit-Sablonu.xlsx';
const EXCEL_LOG_SHEET_NAME = 'Bağlantı Kayıtları';
const EXCEL_SUMMARY_SHEET_NAME = 'Özet';

// Bazı XLSX şablonları ana Excel XML etiketlerini <x:workbook>, <x:sheets>
// biçiminde üretir. ExcelJS bu dosyalarda "sheets" hatası verebildiğinden,
// yalnızca XML ad alanı ön ekini aynı XLSX dosyası içinde normalleştiriyoruz.
// Excel dosyasının konumu ve sayfaları değişmez; kalıcı ek dosya oluşturulmaz.
async function normalizeExcelXmlNamespaces(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel dosyası bulunamadı: ${filePath}`);
  }

  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const workbookEntry = zip.file('xl/workbook.xml');
  if (!workbookEntry) {
    throw new Error('Excel dosyasında xl/workbook.xml bulunamadı. Dosya geçerli bir .xlsx olmayabilir.');
  }

  const workbookXml = await workbookEntry.async('string');
  if (!workbookXml.includes('<x:workbook')) return false;

  const mainNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  let changed = false;

  for (const [entryName, entry] of Object.entries(zip.files)) {
    if (entry.dir || !/\.(xml|rels)$/i.test(entryName)) continue;

    const xml = await entry.async('string');
    const normalizedXml = xml
      .replace(/xmlns:x="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g, `xmlns="${mainNamespace}"`)
      .replace(/(<\/?\s*)x:/g, '$1');

    if (normalizedXml !== xml) {
      zip.file(entryName, normalizedXml);
      changed = true;
    }
  }

  if (!changed) return false;

  const fixedBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });
  const temporaryPath = `${filePath}.fenova-tmp`;

  try {
    fs.writeFileSync(temporaryPath, fixedBuffer);
    fs.renameSync(temporaryPath, filePath);
  } finally {
    if (fs.existsSync(temporaryPath)) fs.unlinkSync(temporaryPath);
  }

  console.log('Fenova: Excel dosyasının XML biçimi otomatik olarak düzeltildi.');
  return true;
}

let excelWorkbook = null;
let excelLogSheet = null;
let excelSummarySheet = null;
let excelWriteQueue = Promise.resolve();
const excelStats = {
  total: 0,
  logins: 0,
  logouts: 0,
  usernames: new Set(),
};

function updateExcelSummary() {
  if (!excelSummarySheet) return;
  excelSummarySheet.getCell('B4').value = excelStats.total;
  excelSummarySheet.getCell('B5').value = excelStats.logins;
  excelSummarySheet.getCell('B6').value = excelStats.logouts;
  excelSummarySheet.getCell('B7').value = excelStats.usernames.size;
  excelSummarySheet.getCell('A10').value =
    'Not: Kayıtlar bu dosyaya sunucu çalışırken otomatik eklenir. Excel dosyasını görüntülemek için sunucuyu durdurup dosyayı açın.';
}

async function initializeExcelLogger() {
  await normalizeExcelXmlNamespaces(EXCEL_FILE_PATH);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE_PATH);

  const logSheet = workbook.getWorksheet(EXCEL_LOG_SHEET_NAME);
  const summarySheet = workbook.getWorksheet(EXCEL_SUMMARY_SHEET_NAME);
  if (!logSheet) {
    throw new Error(`Excel içinde "${EXCEL_LOG_SHEET_NAME}" sayfası bulunamadı.`);
  }
  if (!summarySheet) {
    throw new Error(`Excel içinde "${EXCEL_SUMMARY_SHEET_NAME}" sayfası bulunamadı.`);
  }

  // Şablondaki 4. satır başlıklardır; mevcut kayıtlar korunur ve 5. satırdan devam edilir.
  const headers = ['Tarih-Saat', 'Olay', 'Görünen Ad', 'Kullanıcı Adı', 'Socket ID', 'Oda Kodu', 'Oyun Modu', 'Not'];
  const headerRow = logSheet.getRow(4);
  headers.forEach((header, index) => {
    if (!headerRow.getCell(index + 1).value) headerRow.getCell(index + 1).value = header;
  });

  // Excel'de önceden bulunan kayıtları sayarak özet bölümünü doğru başlat.
  for (let rowNumber = 5; rowNumber <= logSheet.rowCount; rowNumber += 1) {
    const row = logSheet.getRow(rowNumber);
    const eventName = String(row.getCell(2).value || '').trim();
    if (!eventName) continue;
    excelStats.total += 1;
    if (eventName === 'ogrenci_giris') excelStats.logins += 1;
    if (eventName === 'ogrenci_cikis') excelStats.logouts += 1;
    const username = String(row.getCell(4).value || '').trim();
    if (username) excelStats.usernames.add(username);
  }

  excelWorkbook = workbook;
  excelLogSheet = logSheet;
  excelSummarySheet = summarySheet;
  updateExcelSummary();
  return true;
}

const excelReady = initializeExcelLogger().catch((error) => {
  console.error('Fenova Excel kayıt sistemi başlatılamadı:', error.stack || error.message);
  console.error('Kontrol et: Excel dosyası belirtilen konumda olmalı ve sunucu aynı Windows bilgisayarda çalışmalı.');
  return false;
});

function recordActivity(eventName, options = {}) {
  const user = options.user || null;
  const now = new Date();
  const record = {
    id: crypto.randomUUID(),
    timestamp: now.toISOString(),
    localTime: new Intl.DateTimeFormat('tr-TR', {
      dateStyle: 'short', timeStyle: 'medium', timeZone: 'Europe/Istanbul'
    }).format(now),
    timeZone: 'Europe/Istanbul',
    event: String(eventName || 'unknown').slice(0, 60),
    userId: user?.id || options.userId || null,
    username: user?.username ? `@${user.username}` : (options.username || null),
    displayName: user?.displayName || options.displayName || null,
    socketId: options.socketId || null,
    roomCode: options.roomCode || null,
    mode: options.mode || null,
    details: options.details && typeof options.details === 'object' ? options.details : {},
  };

  // Eşzamanlı olayların aynı Excel dosyasına çakışmadan sırayla yazılmasını sağlar.
  excelWriteQueue = excelWriteQueue.then(async () => {
    const initialized = await excelReady;
    if (!initialized || !excelWorkbook || !excelLogSheet) return;

    const detailText = JSON.stringify(record.details || {});
    const note = detailText.length > 2500 ? `${detailText.slice(0, 2497)}...` : detailText;
    const row = excelLogSheet.addRow([
      `${record.localTime} (Europe/Istanbul)`,
      record.event,
      record.displayName || '',
      record.username || '',
      record.socketId || '',
      record.roomCode || '',
      record.mode || '',
      note,
    ]);
    row.height = 22;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = { vertical: 'middle', wrapText: true };
    });

    excelStats.total += 1;
    if (record.event === 'ogrenci_giris') excelStats.logins += 1;
    if (record.event === 'ogrenci_cikis') excelStats.logouts += 1;
    if (record.username) excelStats.usernames.add(record.username);
    updateExcelSummary();

    await excelWorkbook.xlsx.writeFile(EXCEL_FILE_PATH);
  }).catch((error) => {
    console.error('Fenova Excel kaydı yazılamadı:', error.message);
    console.error('Excel dosyası açıksa kapat; OneDrive eşitlemesi bitince yeni bir etkinlikte tekrar denenecektir.');
  });
}

const questionsPath = path.join(__dirname, 'questions.json');
if (!fs.existsSync(questionsPath)) {
  throw new Error('questions.json bulunamadı. Dosyayı server.js ile aynı klasöre koy.');
}

const questionFile = JSON.parse(fs.readFileSync(questionsPath, 'utf8'));
if (!questionFile || !Array.isArray(questionFile.units)) {
  throw new Error('questions.json biçimi hatalı: units dizisi bulunamadı.');
}

function createUnitMap(unitList) {
  const result = new Map();
  if (!Array.isArray(unitList)) return result;
  for (const unit of unitList) {
    const id = String(unit?.id ?? '');
    if (!id || !unit?.title || !Array.isArray(unit.questions)) continue;
    const validQuestions = unit.questions.filter((q) =>
      q && typeof q.soru === 'string' &&
      Array.isArray(q.secenekler) && q.secenekler.length === 4 &&
      Number.isInteger(q.dogruCevap) && q.dogruCevap >= 0 && q.dogruCevap <= 3
    );
    if (validQuestions.length) result.set(id, { id, title: String(unit.title), questions: validQuestions });
  }
  return result;
}

const units = createUnitMap(questionFile.units);
if (units.size === 0) throw new Error('questions.json içinde geçerli ünite/soru bulunamadı.');

// Hızlı Oyun her sınıf için ayrı soru havuzu kullanır.
// Şu anki questions.json yalnızca 7. sınıf üniteleri içerdiğinden üst seviye "units" alanı 7. sınıfa bağlanır.
// İleride questionFile.grades dizisine { grade: 6, units: [...] } benzeri sınıf paketleri eklenebilir.
const quickQuestionPools = new Map();
if (Array.isArray(questionFile.grades)) {
  for (const gradePack of questionFile.grades) {
    const grade = Number(gradePack?.grade ?? gradePack?.sinif);
    const gradeUnits = createUnitMap(gradePack?.units);
    if ([3, 4, 5, 6, 7, 8].includes(grade) && gradeUnits.size) quickQuestionPools.set(grade, gradeUnits);
  }
} else if (questionFile.grades && typeof questionFile.grades === 'object') {
  for (const [gradeKey, gradePack] of Object.entries(questionFile.grades)) {
    const grade = Number(gradeKey);
    const gradeUnits = createUnitMap(Array.isArray(gradePack) ? gradePack : gradePack?.units);
    if ([3, 4, 5, 6, 7, 8].includes(grade) && gradeUnits.size) quickQuestionPools.set(grade, gradeUnits);
  }
}
if (!quickQuestionPools.has(7) && units.size) quickQuestionPools.set(7, units);

// Öğrenci hesapları: kullanıcı bilgileri public klasörünün dışında saklanır.
app.use(express.json({ limit: '32kb' }));

const usersPath = path.join(__dirname, 'users.json');
const freeAvatarIds = new Set([
  'rocket', 'fox', 'panda', 'owl', 'robot', 'cat',
  'dino', 'frog', 'octopus', 'tiger', 'atom', 'planet'
]);
const allowedAvatarIds = new Set([
  ...freeAvatarIds,
  'dragon', 'unicorn', 'shark', 'phoenix', 'alien', 'capybara',
  'gold-robot', 'crystal-fox', 'galaxy-cat', 'neon-panda'
]);
const sessions = new Map();
let users = [];
function createFriendCode() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code;
  do {
    code = 'FN-' + Array.from({ length: 6 }, () => alphabet[Math.floor(Math.random() * alphabet.length)]).join('');
  } while (users.some((user) => user.friendCode === code));
  return code;
}

function normalizeUserRecord(user) {
  let changed = false;
  const storedGrade = Number(user.grade);
  if (![3, 4, 5, 6, 7, 8].includes(storedGrade)) {
    if (user.grade !== null) { user.grade = null; changed = true; }
  } else if (typeof user.grade !== 'number') {
    user.grade = storedGrade; changed = true;
  }
  if (!Object.prototype.hasOwnProperty.call(user, 'welcomePackClaimedAt')) { user.welcomePackClaimedAt = null; changed = true; }
  if (!user.dailyRewards || typeof user.dailyRewards !== 'object') {
    user.dailyRewards = { lastGiftDate: null, lastSpinDate: null, giftsClaimed: 0, wheelsSpun: 0 };
    changed = true;
  } else {
    for (const [key, fallback] of Object.entries({ lastGiftDate: null, lastSpinDate: null, giftsClaimed: 0, wheelsSpun: 0 })) {
      if (!Object.prototype.hasOwnProperty.call(user.dailyRewards, key)) { user.dailyRewards[key] = fallback; changed = true; }
    }
    if (!Number.isFinite(user.dailyRewards.giftsClaimed)) { user.dailyRewards.giftsClaimed = 0; changed = true; }
    if (!Number.isFinite(user.dailyRewards.wheelsSpun)) { user.dailyRewards.wheelsSpun = 0; changed = true; }
  }
  if (!user.friendCode) { user.friendCode = createFriendCode(); changed = true; }
  if (!Array.isArray(user.friends)) { user.friends = []; changed = true; }
  if (!Array.isArray(user.friendRequests)) { user.friendRequests = []; changed = true; }
  if (!user.wallet || typeof user.wallet !== 'object') { user.wallet = { kua: 500, diamonds: 20 }; changed = true; }
  if (!Number.isFinite(user.wallet.kua)) { user.wallet.kua = 500; changed = true; }
  if (!Number.isFinite(user.wallet.diamonds)) { user.wallet.diamonds = 20; changed = true; }
  if (!Array.isArray(user.ownedItems)) { user.ownedItems = ['rocket']; changed = true; }
  if (!user.ownedItems.includes('rocket')) { user.ownedItems.push('rocket'); changed = true; }
  if (!user.equippedItems || typeof user.equippedItems !== 'object') { user.equippedItems = { frame: 'frame-basic', theme: 'theme-default' }; changed = true; }
  if (!Array.isArray(user.badges)) { user.badges = ['welcome']; changed = true; }
  if (!user.badges.includes('welcome')) { user.badges.push('welcome'); changed = true; }
  if (!user.stats || typeof user.stats !== 'object') { user.stats = { gamesPlayed: 0, gamesWon: 0, correctAnswers: 0, perfectGames: 0 }; changed = true; }
  for (const key of ['gamesPlayed', 'gamesWon', 'correctAnswers', 'perfectGames']) {
    if (!Number.isFinite(user.stats[key])) { user.stats[key] = 0; changed = true; }
  }
  if (!allowedAvatarIds.has(user.avatarId)) { user.avatarId = 'rocket'; changed = true; }
  return changed;
}

try {
  if (fs.existsSync(usersPath)) {
    const loaded = JSON.parse(fs.readFileSync(usersPath, 'utf8'));
    users = Array.isArray(loaded) ? loaded : [];
  } else {
    fs.writeFileSync(usersPath, '[]\n', 'utf8');
  }
  let migrated = false;
  for (const user of users) migrated = normalizeUserRecord(user) || migrated;
  if (migrated) {
    const temp = `${usersPath}.tmp`;
    fs.writeFileSync(temp, JSON.stringify(users, null, 2), 'utf8');
    fs.renameSync(temp, usersPath);
  }
} catch (error) {
  throw new Error(`users.json okunamadı: ${error.message}`);
}

function saveUsers() {
  const tempPath = `${usersPath}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify(users, null, 2), 'utf8');
  fs.renameSync(tempPath, usersPath);
}

function cleanDisplayName(value) {
  if (typeof value !== 'string') return '';
  return value.trim().replace(/\s+/g, ' ').slice(0, 24);
}

function normalizeUsername(value) {
  if (typeof value !== 'string') return '';
  return value.trim().toLowerCase();
}

function validUsername(value) {
  return /^[a-zA-Z0-9_.-]{3,20}$/.test(value);
}

function cleanAvatarId(value) {
  return typeof value === 'string' && allowedAvatarIds.has(value) ? value : 'rocket';
}

function passwordDigest(password, salt) {
  return crypto.scryptSync(password, salt, 64).toString('hex');
}

function publicUser(user) {
  return {
    id: user.id,
    username: user.username,
    displayName: user.displayName,
    grade: Number.isInteger(user.grade) ? user.grade : null,
    avatarId: user.avatarId,
    friendCode: user.friendCode,
    wallet: { kua: user.wallet?.kua || 0, diamonds: user.wallet?.diamonds || 0 },
    badges: Array.isArray(user.badges) ? user.badges : [],
    ownedItems: Array.isArray(user.ownedItems) ? user.ownedItems : [],
    equippedItems: user.equippedItems || { frame: 'frame-basic', theme: 'theme-default' },
    stats: user.stats || { gamesPlayed: 0, gamesWon: 0, correctAnswers: 0, perfectGames: 0 },
    friendsCount: Array.isArray(user.friends) ? user.friends.length : 0,
    createdAt: user.createdAt,
    welcomePackClaimedAt: user.welcomePackClaimedAt || null,
    dailyRewards: user.dailyRewards || { lastGiftDate: null, lastSpinDate: null, giftsClaimed: 0, wheelsSpun: 0 },
  };
}

function createSession(user) {
  const token = crypto.randomBytes(32).toString('hex');
  sessions.set(token, user.id);
  return { token, user: publicUser(user) };
}

function requireAuth(req, res, next) {
  const header = req.get('authorization') || '';
  const token = header.startsWith('Bearer ') ? header.slice(7).trim() : '';
  const userId = token ? sessions.get(token) : null;
  const user = userId ? users.find((item) => item.id === userId) : null;
  if (!user) return res.status(401).json({ error: 'Oturum bulunamadı. Lütfen tekrar giriş yap.' });
  req.authToken = token;
  req.user = user;
  next();
}

// Sosyal özellikler için Socket.IO bağlantısı giriş yapan hesaba bağlanır.
io.use((socket, next) => {
  const token = String(socket.handshake.auth?.token || '');
  const userId = sessions.get(token);
  const user = userId ? users.find((item) => item.id === userId) : null;
  if (!user) return next(new Error('Oturum doğrulanamadı. Lütfen tekrar giriş yap.'));
  socket.data.userId = user.id;
  next();
});

app.post('/api/auth/register', (req, res) => {
  const displayName = cleanDisplayName(req.body?.displayName);
  const username = normalizeUsername(req.body?.username);
  const password = req.body?.password;
  const grade = Number(req.body?.grade);
  const requestedAvatar = cleanAvatarId(req.body?.avatarId);
  const avatarId = freeAvatarIds.has(requestedAvatar) ? requestedAvatar : 'rocket';

  if (displayName.length < 2) return res.status(400).json({ error: 'Görünen ad en az 2 karakter olmalı.' });
  if (![3, 4, 5, 6, 7, 8].includes(grade)) return res.status(400).json({ error: 'Lütfen 3–8 arasında sınıfını seç.' });
  if (!validUsername(username)) return res.status(400).json({ error: 'Kullanıcı adı 3–20 karakter olmalı; İngilizce harf, rakam, nokta, tire veya alt çizgi kullan.' });
  if (typeof password !== 'string' || password.length < 8 || password.length > 72) {
    return res.status(400).json({ error: 'Şifre en az 8, en fazla 72 karakter olmalı.' });
  }
  if (users.some((item) => item.username === username)) {
    return res.status(409).json({ error: 'Bu kullanıcı adı zaten alınmış. Başka bir kullanıcı adı dene.' });
  }

  const salt = crypto.randomBytes(16).toString('hex');
  const user = {
    id: crypto.randomUUID(),
    username,
    displayName,
    grade,
    avatarId,
    passwordSalt: salt,
    passwordHash: passwordDigest(password, salt),
    friendCode: createFriendCode(),
    friends: [],
    friendRequests: [],
    wallet: { kua: 500, diamonds: 20 },
    ownedItems: ['rocket'],
    equippedItems: { frame: 'frame-basic', theme: 'theme-default' },
    badges: ['welcome'],
    welcomePackClaimedAt: null,
    dailyRewards: { lastGiftDate: null, lastSpinDate: null, giftsClaimed: 0, wheelsSpun: 0 },
    stats: { gamesPlayed: 0, gamesWon: 0, correctAnswers: 0, perfectGames: 0 },
    createdAt: new Date().toISOString(),
  };
  users.push(user);
  try {
    saveUsers();
  } catch (error) {
    users = users.filter((item) => item.id !== user.id);
    console.error('Kullanıcı kaydedilemedi:', error);
    return res.status(500).json({ error: 'Hesap kaydedilemedi. Sunucunun klasör yazma iznini kontrol et.' });
  }
  console.log(`Yeni öğrenci hesabı oluşturuldu: @${username}`);
  recordActivity('ogrenci_kayit', { user, details: { avatarId } });
  res.status(201).json(createSession(user));
});

app.post('/api/auth/login', (req, res) => {
  const username = normalizeUsername(req.body?.username);
  const password = req.body?.password;
  if (!username || typeof password !== 'string') return res.status(400).json({ error: 'Kullanıcı adı ve şifre gerekli.' });
  const user = users.find((item) => item.username === username);
  if (!user) return res.status(401).json({ error: 'Kullanıcı adı veya şifre hatalı.' });
  const attempted = Buffer.from(passwordDigest(password, user.passwordSalt), 'hex');
  const stored = Buffer.from(user.passwordHash, 'hex');
  if (attempted.length !== stored.length || !crypto.timingSafeEqual(attempted, stored)) {
    return res.status(401).json({ error: 'Kullanıcı adı veya şifre hatalı.' });
  }
  console.log(`Öğrenci giriş yaptı: @${username}`);
  recordActivity('ogrenci_giris', { user, details: { method: 'password' } });
  res.json(createSession(user));
});

app.get('/api/auth/me', requireAuth, (req, res) => {
  res.json({ user: publicUser(req.user) });
});

app.put('/api/auth/profile', requireAuth, (req, res) => {
  const displayName = cleanDisplayName(req.body?.displayName);
  if (displayName.length < 2) return res.status(400).json({ error: 'Görünen ad en az 2 karakter olmalı.' });
  const requestedAvatar = cleanAvatarId(req.body?.avatarId);
  const canUseAvatar = freeAvatarIds.has(requestedAvatar) || req.user.ownedItems.includes(`avatar-${requestedAvatar}`);
  if (!canUseAvatar) return res.status(403).json({ error: 'Bu karakter mağazadan alınmalı.' });
  const grade = Number(req.body?.grade ?? req.user.grade);
  if (![3, 4, 5, 6, 7, 8].includes(grade)) return res.status(400).json({ error: 'Lütfen 3–8 arasında sınıfını seç.' });
  const oldDisplayName = req.user.displayName;
  const oldAvatarId = req.user.avatarId;
  const oldGrade = req.user.grade;
  req.user.displayName = displayName;
  req.user.avatarId = requestedAvatar;
  req.user.grade = grade;
  try {
    saveUsers();
  } catch (error) {
    console.error('Profil kaydedilemedi:', error);
    return res.status(500).json({ error: 'Profil kaydedilemedi.' });
  }
  recordActivity('profil_guncellendi', { user: req.user, details: { oldDisplayName, displayName, oldAvatarId, avatarId: requestedAvatar, oldGrade, grade } });
  res.json({ user: publicUser(req.user) });
});

app.post('/api/auth/logout', requireAuth, (req, res) => {
  recordActivity('ogrenci_cikis', { user: req.user, details: { method: 'manual' } });
  sessions.delete(req.authToken);
  res.json({ ok: true });
});

// Arkadaşlar, gruplar ve sohbetler ayrı bir JSON dosyasında tutulur.
const socialPath = path.join(__dirname, 'social.json');
let socialStore = { groups: [], messages: [] };
try {
  if (fs.existsSync(socialPath)) {
    const loaded = JSON.parse(fs.readFileSync(socialPath, 'utf8'));
    if (loaded && typeof loaded === 'object') socialStore = {
      groups: Array.isArray(loaded.groups) ? loaded.groups : [],
      messages: Array.isArray(loaded.messages) ? loaded.messages : [],
    };
  } else {
    fs.writeFileSync(socialPath, JSON.stringify(socialStore, null, 2), 'utf8');
  }
} catch (error) {
  throw new Error(`social.json okunamadı: ${error.message}`);
}

function saveSocial() {
  const temp = `${socialPath}.tmp`;
  fs.writeFileSync(temp, JSON.stringify(socialStore, null, 2), 'utf8');
  fs.renameSync(temp, socialPath);
}

function findUser(id) { return users.find((user) => user.id === id); }
function userRoom(id) { return `user:${id}`; }
function cleanText(value, max = 300) {
  if (typeof value !== 'string') return '';
  return value.replace(/[<>]/g, '').replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, max);
}
function areFriends(a, b) { return Array.isArray(a.friends) && a.friends.includes(b.id) && Array.isArray(b.friends) && b.friends.includes(a.id); }
function addBadge(user, badgeId) {
  if (!user.badges.includes(badgeId)) user.badges.push(badgeId);
}
function makeUniqueGroupCode() {
  let code;
  do { code = 'FN-' + crypto.randomBytes(3).toString('hex').toUpperCase(); }
  while (socialStore.groups.some((group) => group.code === code));
  return code;
}

const badgeCatalog = [
  { id: 'welcome', name: 'Fenova Kaşifi', icon: '🧭', description: 'Fenova ailesine katıldın.' },
  { id: 'first_game', name: 'İlk Yarışmam', icon: '🚀', description: 'İlk yarışmanı tamamladın.' },
  { id: 'quiz_10', name: 'Bilim Çırağı', icon: '🧪', description: '10 doğru cevap verdin.' },
  { id: 'quiz_50', name: 'Bilim Ustası', icon: '🔬', description: '50 doğru cevap verdin.' },
  { id: 'first_win', name: 'Şampiyon', icon: '🏆', description: 'Bir yarışmada birinci oldun.' },
  { id: 'perfect_game', name: 'Kusursuz Zihin', icon: '💎', description: 'Bir yarışmada tüm soruları doğru bildin.' },
  { id: 'friend_1', name: 'Yeni Dost', icon: '🤝', description: 'İlk arkadaşını ekledin.' },
  { id: 'friend_5', name: 'Sosyal Kaşif', icon: '🌟', description: '5 arkadaş edindin.' },
  { id: 'group_creator', name: 'Takım Kaptanı', icon: '🛡️', description: 'İlk grubunu oluşturdun.' },
  { id: 'collector_3', name: 'Koleksiyoncu', icon: '🎒', description: 'Mağazadan 3 eşya aldın.' },
];

const shopItems = [
  { id: 'avatar-dragon', name: 'Ateş Ejderhası', emoji: '🐉', kind: 'avatar', currency: 'kua', price: 650, rarity: 'Nadir', description: 'Alev gibi hızlı, bilim gibi güçlü.' },
  { id: 'avatar-unicorn', name: 'Kristal Unicorn', emoji: '🦄', kind: 'avatar', currency: 'kua', price: 700, rarity: 'Nadir', description: 'Renkli ve ışıltılı bir yol arkadaşı.' },
  { id: 'avatar-shark', name: 'Turbo Köpekbalığı', emoji: '🦈', kind: 'avatar', currency: 'kua', price: 800, rarity: 'Epik', description: 'Okyanusun en hızlı araştırmacısı.' },
  { id: 'avatar-phoenix', name: 'Anka Kuşu', emoji: '🔥', kind: 'avatar', currency: 'diamonds', price: 18, rarity: 'Efsanevi', description: 'Her yarışmada yeniden doğar.' },
  { id: 'avatar-alien', name: 'Galaksi Uzaylısı', emoji: '👽', kind: 'avatar', currency: 'kua', price: 900, rarity: 'Epik', description: 'Başka gezegenlerden gelen zeka.' },
  { id: 'avatar-capybara', name: 'Sakin Kapibara', emoji: '🦫', kind: 'avatar', currency: 'kua', price: 550, rarity: 'Nadir', description: 'Sakin kal, doğru cevabı bul.' },
  { id: 'avatar-gold-robot', name: 'Altın Robot', emoji: '🤖', kind: 'avatar', currency: 'diamonds', price: 25, rarity: 'Efsanevi', description: 'Geleceğin özel seri karakteri.' },
  { id: 'frame-galaxy', name: 'Galaksi Çerçevesi', emoji: '🌌', kind: 'frame', currency: 'kua', price: 400, rarity: 'Nadir', description: 'Profiline yıldızlı bir görünüm ekler.' },
  { id: 'frame-crown', name: 'Şampiyon Çerçevesi', emoji: '👑', kind: 'frame', currency: 'diamonds', price: 12, rarity: 'Epik', description: 'Başarılarını parlat.' },
  { id: 'theme-neon', name: 'Neon Laboratuvar', emoji: '🧬', kind: 'theme', currency: 'kua', price: 500, rarity: 'Nadir', description: 'Ekranın için neon laboratuvar teması.' },
];

// Mağaza kodları sunucu tarafında SHA-256 özetiyle tutulur. Kodun kendisi dosyada saklanmaz.
const storeCodesPath = path.join(__dirname, 'store-codes.json');
let storeCodes = { codes: [] };
try {
  if (fs.existsSync(storeCodesPath)) {
    const loadedCodes = JSON.parse(fs.readFileSync(storeCodesPath, 'utf8'));
    if (loadedCodes && Array.isArray(loadedCodes.codes)) storeCodes = loadedCodes;
  } else {
    fs.writeFileSync(storeCodesPath, JSON.stringify(storeCodes, null, 2) + '\n', 'utf8');
  }
} catch (error) {
  throw new Error(`store-codes.json okunamadı: ${error.message}`);
}

function saveStoreCodes() {
  const temp = `${storeCodesPath}.tmp`;
  fs.writeFileSync(temp, JSON.stringify(storeCodes, null, 2), 'utf8');
  fs.renameSync(temp, storeCodesPath);
}

function normalizeStoreCode(value) {
  if (typeof value !== 'string') return '';
  return value.toUpperCase().replace(/[^A-Z0-9-]/g, '').slice(0, 40);
}

function hashStoreCode(code) {
  return crypto.createHash('sha256').update(code).digest('hex');
}

const lootBoxes = [
  {
    id: 'science-mini', name: 'Bilim Paketi', emoji: '📦', currency: 'kua', price: 250,
    rarity: 'Başlangıç', description: 'Küçük ama sürprizlerle dolu bir paket.',
    pool: [
      { type: 'kua', amount: 100, weight: 35, name: '100 Kua', emoji: '🪙', rarity: 'Sıradan' },
      { type: 'kua', amount: 200, weight: 30, name: '200 Kua', emoji: '🪙', rarity: 'Sıradan' },
      { type: 'diamonds', amount: 1, weight: 18, name: '1 Elmas', emoji: '💎', rarity: 'Nadir' },
      { type: 'item', itemId: 'frame-galaxy', weight: 10, name: 'Galaksi Çerçevesi', emoji: '🌌', rarity: 'Epik' },
      { type: 'item', itemId: 'theme-neon', weight: 7, name: 'Neon Laboratuvar', emoji: '🧬', rarity: 'Epik' },
    ],
  },
  {
    id: 'science-chest', name: 'Bilim Sandığı', emoji: '🧰', currency: 'kua', price: 650,
    rarity: 'Epik', description: 'Daha değerli ödüller ve özel eşyalar içerir.',
    pool: [
      { type: 'kua', amount: 250, weight: 25, name: '250 Kua', emoji: '🪙', rarity: 'Sıradan' },
      { type: 'kua', amount: 500, weight: 20, name: '500 Kua', emoji: '🪙', rarity: 'Nadir' },
      { type: 'diamonds', amount: 2, weight: 20, name: '2 Elmas', emoji: '💎', rarity: 'Nadir' },
      { type: 'item', itemId: 'avatar-dragon', weight: 12, name: 'Ateş Ejderhası', emoji: '🐉', rarity: 'Epik' },
      { type: 'item', itemId: 'avatar-unicorn', weight: 10, name: 'Kristal Unicorn', emoji: '🦄', rarity: 'Epik' },
      { type: 'item', itemId: 'frame-crown', weight: 6, name: 'Şampiyon Çerçevesi', emoji: '👑', rarity: 'Efsanevi' },
      { type: 'diamonds', amount: 5, weight: 7, name: '5 Elmas', emoji: '💎', rarity: 'Efsanevi' },
    ],
  },
  {
    id: 'galaxy-capsule', name: 'Galaksi Kapsülü', emoji: '🌠', currency: 'diamonds', price: 12,
    rarity: 'Efsanevi', description: 'Nadir karakterler ve parlak ödüller için özel kapsül.',
    pool: [
      { type: 'diamonds', amount: 3, weight: 25, name: '3 Elmas', emoji: '💎', rarity: 'Nadir' },
      { type: 'kua', amount: 900, weight: 20, name: '900 Kua', emoji: '🪙', rarity: 'Epik' },
      { type: 'item', itemId: 'avatar-shark', weight: 15, name: 'Turbo Köpekbalığı', emoji: '🦈', rarity: 'Epik' },
      { type: 'item', itemId: 'avatar-phoenix', weight: 14, name: 'Anka Kuşu', emoji: '🔥', rarity: 'Efsanevi' },
      { type: 'item', itemId: 'avatar-gold-robot', weight: 12, name: 'Altın Robot', emoji: '🤖', rarity: 'Efsanevi' },
      { type: 'item', itemId: 'frame-crown', weight: 14, name: 'Şampiyon Çerçevesi', emoji: '👑', rarity: 'Efsanevi' },
    ],
  },
];

function chooseWeightedReward(pool) {
  const total = pool.reduce((sum, reward) => sum + reward.weight, 0);
  let roll = crypto.randomInt(total);
  for (const reward of pool) {
    roll -= reward.weight;
    if (roll < 0) return { ...reward };
  }
  return { ...pool[pool.length - 1] };
}

function grantStoreReward(user, reward) {
  if (reward.type === 'kua') {
    user.wallet.kua += reward.amount;
    return { ...reward, details: `${reward.amount} Kua hesabına eklendi.` };
  }
  if (reward.type === 'diamonds') {
    user.wallet.diamonds += reward.amount;
    return { ...reward, details: `${reward.amount} Elmas hesabına eklendi.` };
  }
  if (reward.type === 'item') {
    const item = shopItems.find((entry) => entry.id === reward.itemId);
    if (!item) throw new Error('Ödül kataloğunda bulunamadı.');
    if (user.ownedItems.includes(item.id)) {
      user.wallet.kua += 200;
      return { type: 'kua', amount: 200, name: 'Tekrar Ödülü', emoji: '🪙', rarity: 'Nadir', details: `${item.name} zaten sende olduğu için yerine 200 Kua verildi.` };
    }
    user.ownedItems.push(item.id);
    return { ...reward, name: item.name, emoji: item.emoji, details: `${item.name} koleksiyonuna eklendi.` };
  }
  throw new Error('Ödül türü geçersiz.');
}


// Her hesap için yalnızca bir kez verilebilen ücretsiz ilk giriş paketi.
const welcomePackPool = [
  { type: 'kua', amount: 250, weight: 38, name: '250 Kua', emoji: '🪙', rarity: 'Sıradan' },
  { type: 'kua', amount: 400, weight: 24, name: '400 Kua', emoji: '🪙', rarity: 'Nadir' },
  { type: 'diamonds', amount: 2, weight: 18, name: '2 Elmas', emoji: '💎', rarity: 'Nadir' },
  { type: 'item', itemId: 'frame-galaxy', weight: 12, name: 'Galaksi Çerçevesi', emoji: '🌌', rarity: 'Epik' },
  { type: 'item', itemId: 'theme-neon', weight: 8, name: 'Neon Laboratuvar', emoji: '🧬', rarity: 'Epik' },
];

app.get('/api/shop/welcome-pack/status', requireAuth, (req, res) => {
  res.json({ available: !req.user.welcomePackClaimedAt, claimedAt: req.user.welcomePackClaimedAt || null });
});

app.post('/api/shop/welcome-pack/claim', requireAuth, (req, res) => {
  if (req.user.welcomePackClaimedAt) {
    return res.status(409).json({ error: 'İlk giriş paketin daha önce açılmış.' });
  }

  const previousWallet = { ...req.user.wallet };
  const previousItems = [...req.user.ownedItems];
  const selected = chooseWeightedReward(welcomePackPool);
  let reward;
  try {
    reward = grantStoreReward(req.user, selected);
    req.user.welcomePackClaimedAt = new Date().toISOString();
    saveUsers();
  } catch (error) {
    req.user.wallet = previousWallet;
    req.user.ownedItems = previousItems;
    req.user.welcomePackClaimedAt = null;
    console.error('İlk giriş paketi kaydedilemedi:', error);
    return res.status(500).json({ error: 'Hediye kaydedilemedi. Lütfen tekrar dene.' });
  }

  recordActivity('ilk_giris_paketi_acildi', { user: req.user, details: { rewardName: reward?.name || null, rewardType: reward?.type || null, rarity: reward?.rarity || null } });
  const user = publicUser(req.user);
  io.to(userRoom(req.user.id)).emit('profile:update', user);
  res.json({ ok: true, reward, user });
});

// Günlük hediyeler ve çark: tarih sunucu tarafında Türkiye saatine göre hesaplanır.
function getIstanbulDateKey(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Europe/Istanbul', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

const DAILY_GIFT_KUA = 100;
const DAILY_WHEEL_REWARDS = [
  { id: 'kua-50', type: 'kua', amount: 50, weight: 25, name: '50 Kua', emoji: '🪙', rarity: 'Sıradan', sectorIndex: 0 },
  { id: 'kua-100', type: 'kua', amount: 100, weight: 22, name: '100 Kua', emoji: '🪙', rarity: 'Sıradan', sectorIndex: 1 },
  { id: 'kua-150', type: 'kua', amount: 150, weight: 18, name: '150 Kua', emoji: '🪙', rarity: 'Nadir', sectorIndex: 2 },
  { id: 'kua-250', type: 'kua', amount: 250, weight: 12, name: '250 Kua', emoji: '🪙', rarity: 'Nadir', sectorIndex: 3 },
  { id: 'diamond-1', type: 'diamonds', amount: 1, weight: 12, name: '1 Elmas', emoji: '💎', rarity: 'Nadir', sectorIndex: 4 },
  { id: 'diamond-2', type: 'diamonds', amount: 2, weight: 7, name: '2 Elmas', emoji: '💎', rarity: 'Epik', sectorIndex: 5 },
  { id: 'diamond-3', type: 'diamonds', amount: 3, weight: 3, name: '3 Elmas', emoji: '💎', rarity: 'Efsanevi', sectorIndex: 6 },
  { id: 'kua-500', type: 'kua', amount: 500, weight: 1, name: '500 Kua', emoji: '🪙', rarity: 'Efsanevi', sectorIndex: 7 },
];

function dailyRewardStatus(user) {
  const today = getIstanbulDateKey();
  const state = user.dailyRewards || {};
  return {
    today,
    dailyGift: {
      available: state.lastGiftDate !== today,
      amount: DAILY_GIFT_KUA,
      claimedToday: state.lastGiftDate === today,
      lastClaimDate: state.lastGiftDate || null,
    },
    wheel: {
      available: state.lastSpinDate !== today,
      spinsRemaining: state.lastSpinDate === today ? 0 : 1,
      spunToday: state.lastSpinDate === today,
      lastSpinDate: state.lastSpinDate || null,
      rewards: DAILY_WHEEL_REWARDS.map(({ id, type, amount, name, emoji, rarity, sectorIndex }) => ({ id, type, amount, name, emoji, rarity, sectorIndex })),
    },
    resetText: 'Her gün Türkiye saatiyle 00.00’da yenilenir.',
  };
}

app.get('/api/rewards/status', requireAuth, (req, res) => {
  res.json({ status: dailyRewardStatus(req.user), user: publicUser(req.user) });
});

app.post('/api/rewards/daily-gift/claim', requireAuth, (req, res) => {
  const user = req.user;
  const today = getIstanbulDateKey();
  if (user.dailyRewards.lastGiftDate === today) {
    return res.status(409).json({ error: 'Günlük hediyeni bugün zaten aldın. Yarın tekrar gel!' });
  }

  const previousWallet = { ...user.wallet };
  const previousState = { ...user.dailyRewards };
  try {
    user.wallet.kua += DAILY_GIFT_KUA;
    user.dailyRewards.lastGiftDate = today;
    user.dailyRewards.giftsClaimed += 1;
    saveUsers();
  } catch (error) {
    user.wallet = previousWallet;
    user.dailyRewards = previousState;
    console.error('Günlük hediye kaydedilemedi:', error);
    return res.status(500).json({ error: 'Günlük hediye kaydedilemedi. Tekrar dene.' });
  }

  recordActivity('gunluk_hediye_alindi', { user, details: { rewardType: 'kua', amount: DAILY_GIFT_KUA, date: today } });
  const updatedUser = publicUser(user);
  io.to(userRoom(user.id)).emit('profile:update', updatedUser);
  res.json({ ok: true, amount: DAILY_GIFT_KUA, message: `${DAILY_GIFT_KUA} Kua hesabına eklendi!`, status: dailyRewardStatus(user), user: updatedUser });
});

app.post('/api/rewards/wheel/spin', requireAuth, (req, res) => {
  const user = req.user;
  const today = getIstanbulDateKey();
  if (user.dailyRewards.lastSpinDate === today) {
    return res.status(409).json({ error: 'Bugünkü çark hakkını kullandın. Yarın tekrar çevirebilirsin!' });
  }

  const previousWallet = { ...user.wallet };
  const previousState = { ...user.dailyRewards };
  let reward;
  try {
    const selected = chooseWeightedReward(DAILY_WHEEL_REWARDS);
    reward = grantStoreReward(user, selected);
    user.dailyRewards.lastSpinDate = today;
    user.dailyRewards.wheelsSpun += 1;
    saveUsers();
  } catch (error) {
    user.wallet = previousWallet;
    user.dailyRewards = previousState;
    console.error('Çark ödülü kaydedilemedi:', error);
    return res.status(500).json({ error: 'Çark ödülü kaydedilemedi. Tekrar dene.' });
  }

  recordActivity('gunluk_cark_cevrildi', { user, details: { rewardName: reward?.name || null, rewardType: reward?.type || null, amount: reward?.amount || 0, rarity: reward?.rarity || null, sectorIndex: reward?.sectorIndex ?? null, date: today } });
  const updatedUser = publicUser(user);
  io.to(userRoom(user.id)).emit('profile:update', updatedUser);
  res.json({ ok: true, reward, sectorIndex: reward.sectorIndex, status: dailyRewardStatus(user), user: updatedUser });
});

app.get('/api/social/summary', requireAuth, (req, res) => {
  const user = req.user;
  const friends = user.friends.map(findUser).filter(Boolean).map(publicUser);
  const incomingRequests = user.friendRequests
    .map((request) => ({ ...request, from: findUser(request.fromId) }))
    .filter((request) => request.from)
    .map((request) => ({ id: request.id, createdAt: request.createdAt, from: publicUser(request.from) }));
  const groups = socialStore.groups.filter((group) => group.members.includes(user.id)).map((group) => ({
    id: group.id, name: group.name, code: group.code, memberCount: group.members.length, createdBy: group.createdBy,
  }));
  res.json({ user: publicUser(user), friends, incomingRequests, groups });
});

app.post('/api/social/friends/request', requireAuth, (req, res) => {
  const code = String(req.body?.friendCode || '').trim().toUpperCase();
  const target = users.find((user) => user.friendCode === code);
  if (!target) return res.status(404).json({ error: 'Bu arkadaş koduyla bir öğrenci bulunamadı.' });
  if (target.id === req.user.id) return res.status(400).json({ error: 'Kendi kodunu kullanamazsın.' });
  if (areFriends(req.user, target)) return res.status(409).json({ error: 'Bu öğrenci zaten arkadaşın.' });
  if (target.friendRequests.some((item) => item.fromId === req.user.id)) return res.status(409).json({ error: 'Arkadaşlık isteğin zaten gönderilmiş.' });
  const reverse = req.user.friendRequests.find((item) => item.fromId === target.id);
  if (reverse) {
    req.user.friendRequests = req.user.friendRequests.filter((item) => item.id !== reverse.id);
    target.friends.push(req.user.id); req.user.friends.push(target.id);
    addBadge(req.user, 'friend_1'); addBadge(target, 'friend_1');
    if (req.user.friends.length >= 5) addBadge(req.user, 'friend_5');
    if (target.friends.length >= 5) addBadge(target, 'friend_5');
    saveUsers();
    io.to(userRoom(target.id)).emit('social:update'); io.to(userRoom(req.user.id)).emit('social:update');
    recordActivity('arkadas_olundu', { user: req.user, details: { otherUserId: target.id, otherUsername: `@${target.username}` } });
    return res.json({ ok: true, accepted: true, message: 'Karşılıklı istekler eşleşti. Artık arkadaşsınız.' });
  }
  const request = { id: crypto.randomUUID(), fromId: req.user.id, createdAt: new Date().toISOString() };
  target.friendRequests.push(request);
  saveUsers();
  io.to(userRoom(target.id)).emit('social:update');
  recordActivity('arkadaslik_istegi_gonderildi', { user: req.user, details: { targetUserId: target.id, targetUsername: `@${target.username}` } });
  res.status(201).json({ ok: true, message: `${target.displayName} kişisine arkadaşlık isteği gönderildi.` });
});

app.post('/api/social/friends/respond', requireAuth, (req, res) => {
  const requestId = String(req.body?.requestId || '');
  const accept = Boolean(req.body?.accept);
  const index = req.user.friendRequests.findIndex((item) => item.id === requestId);
  if (index < 0) return res.status(404).json({ error: 'Arkadaşlık isteği bulunamadı.' });
  const [request] = req.user.friendRequests.splice(index, 1);
  const sender = findUser(request.fromId);
  if (accept && sender && !areFriends(req.user, sender)) {
    req.user.friends.push(sender.id); sender.friends.push(req.user.id);
    addBadge(req.user, 'friend_1'); addBadge(sender, 'friend_1');
    if (req.user.friends.length >= 5) addBadge(req.user, 'friend_5');
    if (sender.friends.length >= 5) addBadge(sender, 'friend_5');
  }
  saveUsers();
  if (sender) io.to(userRoom(sender.id)).emit('social:update');
  io.to(userRoom(req.user.id)).emit('social:update');
  recordActivity(accept ? 'arkadaslik_istegi_kabul' : 'arkadaslik_istegi_reddedildi', { user: req.user, details: { senderUserId: sender?.id || request.fromId, senderUsername: sender ? `@${sender.username}` : null } });
  res.json({ ok: true, accepted: Boolean(accept) });
});

app.post('/api/social/groups', requireAuth, (req, res) => {
  const name = cleanText(req.body?.name, 32);
  if (name.length < 3) return res.status(400).json({ error: 'Grup adı en az 3 karakter olmalı.' });
  const owned = socialStore.groups.filter((group) => group.createdBy === req.user.id).length;
  if (owned >= 10) return res.status(400).json({ error: 'En fazla 10 grup oluşturabilirsin.' });
  const group = { id: crypto.randomUUID(), name, code: makeUniqueGroupCode(), createdBy: req.user.id, members: [req.user.id], createdAt: new Date().toISOString() };
  socialStore.groups.push(group); saveSocial(); addBadge(req.user, 'group_creator'); saveUsers();
  recordActivity('grup_olusturuldu', { user: req.user, details: { groupId: group.id, groupName: group.name, groupCode: group.code } });
  res.status(201).json({ group: { id: group.id, name: group.name, code: group.code, memberCount: 1, createdBy: group.createdBy } });
});

app.post('/api/social/groups/join', requireAuth, (req, res) => {
  const code = String(req.body?.code || '').trim().toUpperCase();
  const group = socialStore.groups.find((item) => item.code === code);
  if (!group) return res.status(404).json({ error: 'Grup kodu bulunamadı.' });
  if (!group.members.includes(req.user.id)) {
    if (group.members.length >= 50) return res.status(400).json({ error: 'Bu grup dolu.' });
    group.members.push(req.user.id); saveSocial();
    for (const memberId of group.members) io.to(userRoom(memberId)).emit('social:update');
    recordActivity('gruba_katildi', { user: req.user, details: { groupId: group.id, groupName: group.name, groupCode: group.code } });
  }
  res.json({ ok: true, group: { id: group.id, name: group.name, code: group.code, memberCount: group.members.length, createdBy: group.createdBy } });
});

app.get('/api/social/groups/:groupId/messages', requireAuth, (req, res) => {
  const group = socialStore.groups.find((item) => item.id === req.params.groupId);
  if (!group || !group.members.includes(req.user.id)) return res.status(403).json({ error: 'Bu grubun üyesi değilsin.' });
  const messages = socialStore.messages.filter((item) => item.type === 'group' && item.groupId === group.id).slice(-100).map((message) => ({
    id: message.id, text: message.text, createdAt: message.createdAt,
    from: publicUser(findUser(message.fromId) || { id: '', username: 'silinmis', displayName: 'Silinmiş hesap', avatarId: 'rocket', wallet: {}, badges: [], ownedItems: [], stats: {} }),
  }));
  res.json({ messages });
});

app.post('/api/social/groups/:groupId/messages', requireAuth, (req, res) => {
  const group = socialStore.groups.find((item) => item.id === req.params.groupId);
  if (!group || !group.members.includes(req.user.id)) return res.status(403).json({ error: 'Bu grubun üyesi değilsin.' });
  const text = cleanText(req.body?.text, 300);
  if (!text) return res.status(400).json({ error: 'Mesaj boş olamaz.' });
  if (/https?:\/\/|www\./i.test(text)) return res.status(400).json({ error: 'Güvenlik için bağlantı paylaşımı kapalı.' });
  const message = { id: crypto.randomUUID(), type: 'group', groupId: group.id, fromId: req.user.id, text, createdAt: new Date().toISOString() };
  socialStore.messages.push(message); if (socialStore.messages.length > 10000) socialStore.messages = socialStore.messages.slice(-10000); saveSocial();
  recordActivity('grup_mesaji_gonderildi', { user: req.user, details: { groupId: group.id, groupName: group.name, messageLength: text.length } });
  const payload = { id: message.id, groupId: group.id, text, createdAt: message.createdAt, from: publicUser(req.user) };
  for (const memberId of group.members) io.to(userRoom(memberId)).emit('group:new', payload);
  res.status(201).json({ message: payload });
});

app.get('/api/social/messages/:friendId', requireAuth, (req, res) => {
  const friend = findUser(req.params.friendId);
  if (!friend || !areFriends(req.user, friend)) return res.status(403).json({ error: 'Mesajlaşmak için önce arkadaş olmalısınız.' });
  const messages = socialStore.messages.filter((item) => item.type === 'dm' && ((item.fromId === req.user.id && item.toId === friend.id) || (item.fromId === friend.id && item.toId === req.user.id))).slice(-100).map((message) => ({
    id: message.id, text: message.text, createdAt: message.createdAt, fromId: message.fromId,
  }));
  res.json({ messages, friend: publicUser(friend) });
});

app.post('/api/social/messages/:friendId', requireAuth, (req, res) => {
  const friend = findUser(req.params.friendId);
  if (!friend || !areFriends(req.user, friend)) return res.status(403).json({ error: 'Mesajlaşmak için önce arkadaş olmalısınız.' });
  const text = cleanText(req.body?.text, 300);
  if (!text) return res.status(400).json({ error: 'Mesaj boş olamaz.' });
  if (/https?:\/\/|www\./i.test(text)) return res.status(400).json({ error: 'Güvenlik için bağlantı paylaşımı kapalı.' });
  const message = { id: crypto.randomUUID(), type: 'dm', fromId: req.user.id, toId: friend.id, text, createdAt: new Date().toISOString() };
  socialStore.messages.push(message); if (socialStore.messages.length > 10000) socialStore.messages = socialStore.messages.slice(-10000); saveSocial();
  recordActivity('ozel_mesaj_gonderildi', { user: req.user, details: { recipientUserId: friend.id, recipientUsername: `@${friend.username}`, messageLength: text.length } });
  const payload = { id: message.id, text, createdAt: message.createdAt, fromId: req.user.id, from: publicUser(req.user) };
  io.to(userRoom(friend.id)).emit('dm:new', payload);
  io.to(userRoom(req.user.id)).emit('dm:new', payload);
  res.status(201).json({ message: payload });
});

app.get('/api/shop/items', requireAuth, (req, res) => {
  res.json({ items: shopItems.map((item) => ({ ...item, owned: req.user.ownedItems.includes(item.id), equipped: item.kind === 'avatar' ? req.user.avatarId === item.id.replace('avatar-', '') : req.user.equippedItems?.[item.kind] === item.id })), wallet: req.user.wallet });
});

app.post('/api/shop/buy', requireAuth, (req, res) => {
  const item = shopItems.find((entry) => entry.id === String(req.body?.itemId || ''));
  if (!item) return res.status(404).json({ error: 'Mağaza ürünü bulunamadı.' });
  if (req.user.ownedItems.includes(item.id)) return res.status(409).json({ error: 'Bu ürün zaten koleksiyonunda.' });
  if ((req.user.wallet?.[item.currency] || 0) < item.price) return res.status(400).json({ error: item.currency === 'kua' ? 'Yeterli Kua paran yok.' : 'Yeterli elmasın yok.' });
  req.user.wallet[item.currency] -= item.price;
  req.user.ownedItems.push(item.id);
  if (item.kind === 'avatar') req.user.avatarId = item.id.replace('avatar-', '');
  else req.user.equippedItems[item.kind] = item.id;
  if (req.user.ownedItems.filter((id) => id !== 'rocket').length >= 3) addBadge(req.user, 'collector_3');
  saveUsers();
  recordActivity('magaza_urun_satinalindi', { user: req.user, details: { itemId: item.id, itemName: item.name, currency: item.currency, price: item.price } });
  const user = publicUser(req.user);
  io.to(userRoom(req.user.id)).emit('profile:update', user);
  res.json({ ok: true, user, message: `${item.name} koleksiyonuna eklendi!` });
});

app.post('/api/shop/equip', requireAuth, (req, res) => {
  const item = shopItems.find((entry) => entry.id === String(req.body?.itemId || ''));
  if (!item || !req.user.ownedItems.includes(item.id)) return res.status(403).json({ error: 'Bu ürüne sahip değilsin.' });
  if (item.kind === 'avatar') req.user.avatarId = item.id.replace('avatar-', '');
  else req.user.equippedItems[item.kind] = item.id;
  saveUsers();
  recordActivity('magaza_esya_donatildi', { user: req.user, details: { itemId: item.id, itemName: item.name, kind: item.kind } });
  const user = publicUser(req.user);
  io.to(userRoom(req.user.id)).emit('profile:update', user);
  res.json({ ok: true, user });
});

// Kutu kataloğu: olasılıklar istemciye gönderilmez, seçim yalnızca sunucuda yapılır.
app.get('/api/shop/boxes', requireAuth, (req, res) => {
  res.json({
    boxes: lootBoxes.map(({ pool, ...box }) => {
      // `pool` yukarıda box nesnesinden ayrıştırıldı; bu nedenle box.pool değil pool kullanılmalı.
      const rewards = Array.isArray(pool) ? pool : [];
      const totalWeight = rewards.reduce((sum, item) => {
        const weight = Number(item?.weight) || 0;
        return sum + Math.max(0, weight);
      }, 0);

      return {
        ...box,
        odds: rewards.map((reward) => ({
          name: String(reward?.name || 'Ödül'),
          rarity: String(reward?.rarity || 'common'),
          chance: totalWeight > 0
            ? Math.round(((Math.max(0, Number(reward?.weight) || 0) / totalWeight) * 1000)) / 10
            : 0,
        })),
        affordable: (req.user.wallet?.[box.currency] || 0) >= box.price,
      };
    }),
    wallet: req.user.wallet,
  });
});

app.post('/api/shop/boxes/open', requireAuth, (req, res) => {
  const box = lootBoxes.find((entry) => entry.id === String(req.body?.boxId || ''));
  if (!box) return res.status(404).json({ error: 'Paket bulunamadı.' });
  if ((req.user.wallet?.[box.currency] || 0) < box.price) {
    return res.status(400).json({ error: box.currency === 'kua' ? 'Bu paket için yeterli Kua yok.' : 'Bu paket için yeterli Elmas yok.' });
  }

  req.user.wallet[box.currency] -= box.price;
  let selected = chooseWeightedReward(box.pool);
  let reward;
  try {
    reward = grantStoreReward(req.user, selected);
    req.user.stats.boxesOpened = (req.user.stats.boxesOpened || 0) + 1;
    saveUsers();
  } catch (error) {
    req.user.wallet[box.currency] += box.price;
    return res.status(500).json({ error: 'Paket ödülü kaydedilemedi. Bakiyen düşürülmedi.' });
  }

  recordActivity('surpriz_paket_acildi', { user: req.user, details: { boxId: box.id, boxName: box.name, rewardName: reward?.name || null, rewardType: reward?.type || null, rarity: reward?.rarity || null, currency: box.currency, price: box.price } });
  const user = publicUser(req.user);
  io.to(userRoom(req.user.id)).emit('profile:update', user);
  res.json({ ok: true, box: { id: box.id, name: box.name }, reward, user });
});

app.post('/api/shop/codes/redeem', requireAuth, (req, res) => {
  const code = normalizeStoreCode(req.body?.code);
  if (code.length < 8) return res.status(400).json({ error: 'Geçerli bir mağaza kodu gir.' });
  const entry = storeCodes.codes.find((item) => item.hash === hashStoreCode(code));
  if (!entry) return res.status(404).json({ error: 'Bu mağaza kodu bulunamadı.' });
  if (entry.usedAt) return res.status(409).json({ error: 'Bu kod daha önce kullanılmış.' });

  const reward = entry.reward || {};
  const granted = [];
  if (Number.isInteger(reward.kua) && reward.kua > 0) {
    req.user.wallet.kua += Math.min(reward.kua, 100000);
    granted.push(`${Math.min(reward.kua, 100000)} Kua`);
  }
  if (Number.isInteger(reward.diamonds) && reward.diamonds > 0) {
    req.user.wallet.diamonds += Math.min(reward.diamonds, 1000);
    granted.push(`${Math.min(reward.diamonds, 1000)} Elmas`);
  }
  if (typeof reward.itemId === 'string') {
    const item = shopItems.find((candidate) => candidate.id === reward.itemId);
    if (item && !req.user.ownedItems.includes(item.id)) {
      req.user.ownedItems.push(item.id);
      granted.push(item.name);
    } else if (item) {
      req.user.wallet.kua += 200;
      granted.push('200 Kua (eşya zaten sende)');
    }
  }
  if (granted.length === 0) return res.status(400).json({ error: 'Bu kodun geçerli bir ödülü yok.' });

  entry.usedAt = new Date().toISOString();
  entry.usedBy = req.user.id;
  try {
    saveUsers();
    saveStoreCodes();
  } catch (error) {
    console.error('Mağaza kodu kaydedilemedi:', error);
    return res.status(500).json({ error: 'Kod kullanımı kaydedilemedi. Lütfen tekrar dene.' });
  }
  recordActivity('magaza_kodu_kullanildi', { user: req.user, details: { rewards: granted } });
  const user = publicUser(req.user);
  io.to(userRoom(req.user.id)).emit('profile:update', user);
  res.json({ ok: true, user, message: `Kod kabul edildi! Ödüllerin: ${granted.join(', ')}.` });
});

app.use(express.static(path.join(__dirname, 'public'), {
  etag: true,
  lastModified: true,
  maxAge: process.env.NODE_ENV === 'production' ? '7d' : 0,
  setHeaders(res, filePath) {
    if (filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-cache');
    } else if (/\.(?:css|js|webp|mp3|jpg|png)$/i.test(filePath)) {
      res.setHeader('Cache-Control', 'public, max-age=604800, stale-while-revalidate=86400');
    }
  },
}));

app.get('/healthz', (_req, res) => {
  res.status(200).json({ status: 'online', project: 'Fenova', socketIO: 'active' });
});

// Doğru cevap anahtarı API ile tarayıcıya gönderilmez.
app.get('/api/units', (_req, res) => {
  res.json([...units.values()].map((unit) => ({
    id: unit.id,
    title: unit.title,
    questionCount: unit.questions.length,
  })));
});

// Ligler ayrı bir dosyada saklanır; sahibi lig adını, sezonunu ve açıklamasını düzenleyebilir.
const leaguesPath = path.join(__dirname, 'leagues.json');
let leagues = [];
try {
  if (fs.existsSync(leaguesPath)) {
    const loaded = JSON.parse(fs.readFileSync(leaguesPath, 'utf8'));
    leagues = Array.isArray(loaded) ? loaded : [];
  } else {
    fs.writeFileSync(leaguesPath, '[]\n', 'utf8');
  }
  let migrated = false;
  for (const league of leagues) {
    if (!Array.isArray(league.members)) { league.members = []; migrated = true; }
    if (!league.scores || typeof league.scores !== 'object') { league.scores = {}; migrated = true; }
    if (!league.joinCode) { league.joinCode = crypto.randomBytes(3).toString('hex').toUpperCase(); migrated = true; }
    if (!Number.isFinite(league.maxMembers)) { league.maxMembers = 60; migrated = true; }
  }
  if (migrated) fs.writeFileSync(leaguesPath, JSON.stringify(leagues, null, 2), 'utf8');
} catch (error) {
  throw new Error(`leagues.json okunamadı: ${error.message}`);
}

function saveLeagues() {
  const temp = `${leaguesPath}.tmp`;
  fs.writeFileSync(temp, JSON.stringify(leagues, null, 2), 'utf8');
  fs.renameSync(temp, leaguesPath);
}

function cleanLeagueName(value, max = 42) {
  return typeof value === 'string' ? value.replace(/[<>]/g, '').replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, max) : '';
}

function makeLeagueCode() {
  let code;
  do { code = 'LIG-' + crypto.randomBytes(3).toString('hex').toUpperCase(); }
  while (leagues.some((league) => league.joinCode === code));
  return code;
}

function leagueStandings(league) {
  return league.members.map((userId) => {
    const user = findUser(userId);
    const score = league.scores[userId] || { points: 0, games: 0, wins: 0 };
    return user ? {
      userId,
      displayName: user.displayName,
      avatarId: user.avatarId,
      points: Number(score.points) || 0,
      games: Number(score.games) || 0,
      wins: Number(score.wins) || 0,
    } : null;
  }).filter(Boolean).sort((a, b) => b.points - a.points || b.wins - a.wins || a.displayName.localeCompare(b.displayName, 'tr'));
}

function publicLeague(league, userId) {
  const owner = findUser(league.ownerId);
  return {
    id: league.id,
    name: league.name,
    description: league.description,
    season: league.season,
    unitId: league.unitId,
    unitTitle: league.unitId === 'all' ? 'Tüm üniteler' : (units.get(String(league.unitId))?.title || 'Fen Bilimleri'),
    joinCode: league.joinCode,
    ownerName: owner?.displayName || 'Fenova öğrencisi',
    ownerId: league.ownerId,
    memberCount: league.members.length,
    maxMembers: league.maxMembers,
    isMember: league.members.includes(userId),
    isOwner: league.ownerId === userId,
    status: league.status || 'active',
    standings: leagueStandings(league),
    createdAt: league.createdAt,
  };
}

app.get('/api/leagues', requireAuth, (_req, res) => {
  res.json({ leagues: leagues.filter((league) => league.status !== 'archived').map((league) => publicLeague(league, _req.user.id)) });
});

app.post('/api/leagues', requireAuth, (req, res) => {
  const name = cleanLeagueName(req.body?.name, 36);
  const description = cleanLeagueName(req.body?.description, 150);
  const season = cleanLeagueName(req.body?.season || '2026 Sezonu', 32);
  const unitId = String(req.body?.unitId || 'all');
  const maxMembers = Math.min(100, Math.max(6, Math.floor(Number(req.body?.maxMembers) || 60)));
  if (name.length < 3) return res.status(400).json({ error: 'Lig adı en az 3 karakter olmalı.' });
  if (unitId !== 'all' && !units.has(unitId)) return res.status(400).json({ error: 'Geçerli bir ünite seç.' });
  const ownedCount = leagues.filter((league) => league.ownerId === req.user.id && league.status !== 'archived').length;
  if (ownedCount >= 5) return res.status(400).json({ error: 'En fazla 5 lig oluşturabilirsin.' });
  const league = {
    id: crypto.randomUUID(), name, description, season, unitId,
    joinCode: makeLeagueCode(), ownerId: req.user.id,
    members: [req.user.id], scores: {}, maxMembers,
    status: 'active', createdAt: new Date().toISOString(),
  };
  leagues.push(league);
  try { saveLeagues(); } catch (error) { leagues = leagues.filter((item) => item.id !== league.id); return res.status(500).json({ error: 'Lig kaydedilemedi.' }); }
  recordActivity('lig_olusturuldu', { user: req.user, details: { leagueId: league.id, leagueName: league.name, season: league.season, joinCode: league.joinCode } });
  res.status(201).json({ league: publicLeague(league, req.user.id) });
});

app.post('/api/leagues/join', requireAuth, (req, res) => {
  const code = cleanLeagueName(req.body?.code, 20).toUpperCase();
  const league = leagues.find((item) => item.joinCode === code && item.status !== 'archived');
  if (!league) return res.status(404).json({ error: 'Lig kodu bulunamadı.' });
  if (!league.members.includes(req.user.id)) {
    if (league.members.length >= league.maxMembers) return res.status(409).json({ error: 'Bu lig dolu.' });
    league.members.push(req.user.id);
    saveLeagues();
    recordActivity('lige_katildi', { user: req.user, details: { leagueId: league.id, leagueName: league.name, joinCode: league.joinCode } });
  }
  res.json({ league: publicLeague(league, req.user.id) });
});

app.put('/api/leagues/:id', requireAuth, (req, res) => {
  const league = leagues.find((item) => item.id === req.params.id);
  if (!league) return res.status(404).json({ error: 'Lig bulunamadı.' });
  if (league.ownerId !== req.user.id) return res.status(403).json({ error: 'Bu ligi yalnızca oluşturan öğrenci düzenleyebilir.' });
  const name = cleanLeagueName(req.body?.name, 36);
  const description = cleanLeagueName(req.body?.description, 150);
  const season = cleanLeagueName(req.body?.season, 32);
  if (name.length < 3 || !season) return res.status(400).json({ error: 'Lig adı ve sezon bilgisi gerekli.' });
  league.name = name;
  league.description = description;
  league.season = season;
  saveLeagues();
  recordActivity('lig_duzenlendi', { user: req.user, details: { leagueId: league.id, leagueName: league.name, season: league.season } });
  res.json({ league: publicLeague(league, req.user.id) });
});

app.post('/api/leagues/:id/join', requireAuth, (req, res) => {
  const league = leagues.find((item) => item.id === req.params.id && item.status !== 'archived');
  if (!league) return res.status(404).json({ error: 'Lig bulunamadı.' });
  if (!league.members.includes(req.user.id)) {
    if (league.members.length >= league.maxMembers) return res.status(409).json({ error: 'Bu lig dolu.' });
    league.members.push(req.user.id);
    saveLeagues();
    recordActivity('lige_katildi', { user: req.user, details: { leagueId: league.id, leagueName: league.name, joinCode: league.joinCode } });
  }
  res.json({ league: publicLeague(league, req.user.id) });
});

const rooms = new Map();
const footballMatches = new Map();
const quickQueues = new Map();              // grade -> socketId[]
const quickOpenRoomsByGrade = new Map();    // grade -> roomCode
const quickUserStatus = new Map();          // userId -> { socketId, grade, state, roomCode? }
const QUICK_MIN_PLAYERS = 4;
const QUICK_MAX_PLAYERS = 8;
const QUICK_COUNTDOWN_SECONDS = 20;
const QUICK_QUESTION_COUNT = 10;

function createRoomCode() {
  let code;
  do {
    code = String(Math.floor(1000 + Math.random() * 9000));
  } while (rooms.has(code) || footballMatches.has(code));
  return code;
}

function cleanName(value) {
  if (typeof value !== 'string') return '';
  return value.trim().replace(/\s+/g, ' ').slice(0, 20);
}

function shuffledCopy(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function makePlayerList(room) {
  return [...room.players.values()]
    .map((player) => ({
      id: player.id,
      isim: player.name,
      avatarId: player.avatarId || 'rocket',
      puan: player.score,
      kurucu: !room.isQuickMatch && player.id === room.hostId,
    }))
    .sort((a, b) => b.puan - a.puan || a.isim.localeCompare(b.isim, 'tr'));
}

function quickRoomPayload(room) {
  const remaining = room.countdownEndsAt
    ? Math.max(0, Math.ceil((room.countdownEndsAt - Date.now()) / 1000))
    : 0;
  return {
    code: room.code,
    grade: room.grade,
    state: room.state,
    playerCount: room.players.size,
    minPlayers: QUICK_MIN_PLAYERS,
    maxPlayers: QUICK_MAX_PLAYERS,
    countdown: remaining,
    players: makePlayerList(room),
  };
}

function publishRoomState(room) {
  io.to(room.code).emit('siralama_guncelle', makePlayerList(room));
  if (room.isQuickMatch) {
    for (const player of room.players.values()) io.to(player.id).emit('oda_yetki', false);
    io.to(room.code).emit('quick:roomUpdate', quickRoomPayload(room));
    return;
  }
  if (room.hostId) io.to(room.hostId).emit('oda_yetki', true);
  for (const player of room.players.values()) io.to(player.id).emit('oda_yetki', false);
}

function clearRoomTimers(room) {
  if (room.questionTimer) clearTimeout(room.questionTimer);
  if (room.resultTimer) clearTimeout(room.resultTimer);
  if (room.quickCountdownInterval) clearInterval(room.quickCountdownInterval);
  room.questionTimer = null;
  room.resultTimer = null;
  room.quickCountdownInterval = null;
}

function deleteRoom(code) {
  const room = rooms.get(code);
  if (!room) return;
  clearRoomTimers(room);
  if (room.isQuickMatch) {
    if (quickOpenRoomsByGrade.get(room.grade) === code) quickOpenRoomsByGrade.delete(room.grade);
    for (const player of room.players.values()) {
      const status = quickUserStatus.get(player.userId);
      if (status?.socketId === player.id) quickUserStatus.delete(player.userId);
    }
  }
  rooms.delete(code);
  console.log(`Oda kapatıldı: ${code}`);
}

function getQuickQueue(grade) {
  if (!quickQueues.has(grade)) quickQueues.set(grade, []);
  return quickQueues.get(grade);
}

function quickUserForSocket(socketId) {
  const socket = io.sockets.sockets.get(socketId);
  if (!socket) return null;
  const user = findUser(socket.data.userId);
  return user ? { socket, user } : null;
}

function sendQuickWaiting(socket, grade) {
  const queue = getQuickQueue(grade);
  const index = queue.indexOf(socket.id);
  socket.emit('quick:waiting', {
    grade,
    queuePosition: index >= 0 ? index + 1 : queue.length + 1,
    waitingCount: queue.length,
    minPlayers: QUICK_MIN_PLAYERS,
    maxPlayers: QUICK_MAX_PLAYERS,
    message: `${grade}. sınıftan oyuncular aranıyor. ${QUICK_MIN_PLAYERS} kişi olduğunda oda kurulacak.`,
  });
}

function attachQuickPlayer(socket, room) {
  const user = findUser(socket.data.userId);
  if (!user || !Number.isInteger(user.grade) || user.grade !== room.grade) return false;
  if (room.players.has(socket.id) || room.players.size >= QUICK_MAX_PLAYERS) return false;

  room.players.set(socket.id, {
    id: socket.id,
    userId: user.id,
    name: cleanName(user.displayName) || 'Oyuncu',
    avatarId: user.avatarId || 'rocket',
    score: 0,
    correctThisGame: 0,
    speedBonusTotal: 0,
  });
  quickUserStatus.set(user.id, { socketId: socket.id, grade: room.grade, state: 'lobby', roomCode: room.code });
  socket.join(room.code);
  socket.emit('quick:matched', quickRoomPayload(room));
  recordActivity('hizli_oyun_eslesti', {
    user,
    socketId: socket.id,
    roomCode: room.code,
    mode: 'hizli_oyun',
    details: { grade: room.grade, playerCount: room.players.size },
  });
  return true;
}

function requeueQuickLobby(room) {
  if (!room || !room.isQuickMatch || room.state !== 'lobby') return;
  clearRoomTimers(room);
  if (quickOpenRoomsByGrade.get(room.grade) === room.code) quickOpenRoomsByGrade.delete(room.grade);
  rooms.delete(room.code);

  for (const player of room.players.values()) {
    const socket = io.sockets.sockets.get(player.id);
    if (!socket) {
      const status = quickUserStatus.get(player.userId);
      if (status?.socketId === player.id) quickUserStatus.delete(player.userId);
      continue;
    }
    socket.leave(room.code);
    quickUserStatus.set(player.userId, { socketId: player.id, grade: room.grade, state: 'queue' });
    const queue = getQuickQueue(room.grade);
    if (!queue.includes(player.id)) queue.push(player.id);
    sendQuickWaiting(socket, room.grade);
  }
  processQuickQueue(room.grade);
}

function startQuickCountdown(room) {
  if (!room || room.quickCountdownInterval || room.state !== 'lobby') return;
  room.countdownEndsAt = Date.now() + QUICK_COUNTDOWN_SECONDS * 1000;
  io.to(room.code).emit('quick:countdown', quickRoomPayload(room));
  room.quickCountdownInterval = setInterval(() => {
    if (!rooms.has(room.code) || room.state !== 'lobby') {
      clearInterval(room.quickCountdownInterval);
      room.quickCountdownInterval = null;
      return;
    }
    const payload = quickRoomPayload(room);
    io.to(room.code).emit('quick:countdown', payload);
    if (payload.countdown <= 0) {
      clearInterval(room.quickCountdownInterval);
      room.quickCountdownInterval = null;
      if (room.players.size >= QUICK_MIN_PLAYERS) beginQuickGame(room);
      else requeueQuickLobby(room);
    }
  }, 1000);
}

function beginQuickGame(room) {
  if (!room || room.state !== 'lobby' || room.players.size < QUICK_MIN_PLAYERS) return false;
  clearRoomTimers(room);
  if (quickOpenRoomsByGrade.get(room.grade) === room.code) quickOpenRoomsByGrade.delete(room.grade);
  room.countdownEndsAt = null;
  room.state = 'starting';
  for (const player of room.players.values()) {
    quickUserStatus.set(player.userId, { socketId: player.id, grade: room.grade, state: 'playing', roomCode: room.code });
  }
  io.to(room.code).emit('quick:started', {
    code: room.code,
    grade: room.grade,
    playerCount: room.players.size,
    questionCount: room.questions.length,
  });
  recordActivity('hizli_oyun_basladi', {
    roomCode: room.code,
    mode: 'hizli_oyun',
    details: { grade: room.grade, playerCount: room.players.size, questionCount: room.questions.length },
  });
  sendNextQuestion(room);
  return true;
}

function processQuickQueue(grade) {
  const queue = getQuickQueue(grade);
  // Sadece hâlâ bağlı ve bu kuyruğa ait öğrencileri tut.
  for (let i = queue.length - 1; i >= 0; i -= 1) {
    const found = quickUserForSocket(queue[i]);
    if (!found || Number(found.user.grade) !== grade) {
      const socketId = queue[i];
      queue.splice(i, 1);
      if (found?.user) {
        const status = quickUserStatus.get(found.user.id);
        if (status?.socketId === socketId) quickUserStatus.delete(found.user.id);
      }
    }
  }

  const existingCode = quickOpenRoomsByGrade.get(grade);
  const openRoom = existingCode ? rooms.get(existingCode) : null;
  if (openRoom && openRoom.state === 'lobby' && openRoom.players.size < QUICK_MAX_PLAYERS) {
    while (queue.length && openRoom.players.size < QUICK_MAX_PLAYERS) {
      const socketId = queue.shift();
      const found = quickUserForSocket(socketId);
      if (found) attachQuickPlayer(found.socket, openRoom);
    }
    publishRoomState(openRoom);
    if (openRoom.players.size >= QUICK_MAX_PLAYERS) beginQuickGame(openRoom);
    return;
  }

  if (openRoom && openRoom.state !== 'lobby') quickOpenRoomsByGrade.delete(grade);
  if (queue.length < QUICK_MIN_PLAYERS) {
    queue.forEach((socketId) => {
      const found = quickUserForSocket(socketId);
      if (found) sendQuickWaiting(found.socket, grade);
    });
    return;
  }

  const selected = queue.splice(0, Math.min(QUICK_MAX_PLAYERS, queue.length));
  const code = createRoomCode();
  const gradeUnits = quickQuestionPools.get(grade);
  const allQuestions = gradeUnits
    ? [...gradeUnits.values()].flatMap((unit) => unit.questions.map((question) => ({ ...question, _unitTitle: unit.title })))
    : [];
  if (allQuestions.length < 4) {
    selected.forEach((socketId) => {
      const found = quickUserForSocket(socketId);
      if (found) {
        quickUserStatus.delete(found.user.id);
        found.socket.emit('quick:error', `${grade}. sınıf için henüz yeterli soru eklenmemiş.`);
      }
    });
    return;
  }

  const chosenQuestions = shuffledCopy(allQuestions).slice(0, Math.min(QUICK_QUESTION_COUNT, allQuestions.length));
  const room = {
    code,
    unitId: 'quick',
    unitTitle: `${grade}. Sınıf • Karma Fen Bilimleri`,
    leagueId: null,
    questions: chosenQuestions,
    durationSeconds: QUESTION_DURATION_SECONDS_DEFAULT,
    hostId: null,
    players: new Map(),
    answers: new Map(),
    questionIndex: -1,
    state: 'lobby',
    questionFinished: false,
    questionTimer: null,
    resultTimer: null,
    isQuickMatch: true,
    mode: 'hizli_oyun',
    grade,
    quickCountdownInterval: null,
    countdownEndsAt: null,
  };
  rooms.set(code, room);

  for (const socketId of selected) {
    const found = quickUserForSocket(socketId);
    if (found) attachQuickPlayer(found.socket, room);
  }
  room.hostId = room.players.keys().next().value || null;
  quickOpenRoomsByGrade.set(grade, code);
  publishRoomState(room);

  if (room.players.size >= QUICK_MAX_PLAYERS) beginQuickGame(room);
  else startQuickCountdown(room);
  processQuickQueue(grade);
}

function joinQuickQueue(socket) {
  const user = findUser(socket.data.userId);
  if (!user) return socket.emit('quick:error', 'Öğrenci hesabı doğrulanamadı.');
  const grade = Number(user.grade);
  if (![3, 4, 5, 6, 7, 8].includes(grade)) {
    return socket.emit('quick:error', 'Hızlı Oyun için önce profilinden sınıfını seçmelisin.');
  }
  const questionPool = quickQuestionPools.get(grade);
  const poolCount = questionPool ? [...questionPool.values()].reduce((sum, unit) => sum + unit.questions.length, 0) : 0;
  if (poolCount < 4) {
    return socket.emit('quick:error', `${grade}. sınıfın soru havuzu henüz hazır değil. Şu an elimizde 7. sınıf soru havuzu var; diğer sınıfların soruları eklendiğinde onlar da oynayabilecek.`);
  }

  const existingStatus = quickUserStatus.get(user.id);
  if (existingStatus) {
    const oldSocket = io.sockets.sockets.get(existingStatus.socketId);
    if (oldSocket && oldSocket.id !== socket.id) {
      return socket.emit('quick:error', 'Bu öğrenci hesabı zaten bir hızlı oyun eşleşmesinde. Önce diğer oturumu kapat.');
    }
    if (oldSocket && oldSocket.id === socket.id) {
      if (existingStatus.state === 'queue') sendQuickWaiting(socket, grade);
      else if (existingStatus.roomCode) {
        const room = rooms.get(existingStatus.roomCode);
        if (room) socket.emit('quick:matched', quickRoomPayload(room));
      }
      return;
    }
    quickUserStatus.delete(user.id);
  }

  leaveOtherRooms(socket);
  const openCode = quickOpenRoomsByGrade.get(grade);
  const openRoom = openCode ? rooms.get(openCode) : null;
  if (openRoom && openRoom.state === 'lobby' && openRoom.players.size < QUICK_MAX_PLAYERS) {
    attachQuickPlayer(socket, openRoom);
    publishRoomState(openRoom);
    if (openRoom.players.size >= QUICK_MAX_PLAYERS) beginQuickGame(openRoom);
    return;
  }

  const queue = getQuickQueue(grade);
  if (!queue.includes(socket.id)) queue.push(socket.id);
  quickUserStatus.set(user.id, { socketId: socket.id, grade, state: 'queue' });
  sendQuickWaiting(socket, grade);
  processQuickQueue(grade);
}

function removeQuickQueueSocket(socketId, userId = null) {
  let changedGrade = null;
  for (const [grade, queue] of quickQueues.entries()) {
    const index = queue.indexOf(socketId);
    if (index >= 0) {
      queue.splice(index, 1);
      changedGrade = grade;
    }
  }
  if (userId) {
    const status = quickUserStatus.get(userId);
    if (status?.socketId === socketId && status.state === 'queue') quickUserStatus.delete(userId);
  }
  if (changedGrade !== null) processQuickQueue(changedGrade);
}

function cancelQuickForSocket(socket) {
  const user = findUser(socket.data.userId);
  const status = user ? quickUserStatus.get(user.id) : null;
  removeQuickQueueSocket(socket.id, socket.data.userId);
  if (status?.roomCode) {
    const room = rooms.get(status.roomCode);
    if (room?.isQuickMatch) {
      removeSocketFromRoom(socket.id, room.code);
      socket.leave(room.code);
    }
  }
  if (user) {
    const after = quickUserStatus.get(user.id);
    if (after?.socketId === socket.id) quickUserStatus.delete(user.id);
  }
  socket.emit('quick:cancelled', { message: 'Hızlı oyun aramasından ayrıldın.' });
}

function applyQuizLeaguePoints(room, ranking) {
  if (!room.leagueId) return;
  const league = leagues.find((item) => item.id === room.leagueId && item.status !== 'archived');
  if (!league) return;
  ranking.forEach((rankedPlayer, index) => {
    const player = room.players.get(rankedPlayer.id);
    const user = player?.userId ? findUser(player.userId) : null;
    if (!user || !league.members.includes(user.id)) return;
    const previous = league.scores[user.id] || { points: 0, games: 0, wins: 0 };
    const placementPoints = index === 0 ? 5 : index === 1 ? 3 : 1;
    const correctPoints = Math.max(0, Number(player.correctThisGame) || 0) * 2;
    previous.points += placementPoints + correctPoints;
    previous.games += 1;
    if (index === 0) previous.wins += 1;
    league.scores[user.id] = previous;
    if (previous.wins >= 3) addBadge(user, 'league_champion');
  });
  saveLeagues();
}

function awardGameRewards(room) {
  if (room.rewardsApplied) return;
  room.rewardsApplied = true;
  const ranking = makePlayerList(room);
  const winnerId = ranking[0]?.id;
  recordActivity('yarisma_tamamlandi', { roomCode: room.code, mode: room.mode || 'bilgi_yarismasi', details: { unitId: room.unitId, unitTitle: room.unitTitle, questionCount: room.questions.length, ranking: ranking.slice(0, 10).map((p, i) => ({ rank: i + 1, name: p.isim, score: p.puan })) } });
  applyQuizLeaguePoints(room, ranking);
  for (const player of room.players.values()) {
    const user = player.userId ? findUser(player.userId) : null;
    if (!user) continue;
    user.stats.gamesPlayed += 1;
    user.wallet.kua += 100;
    user.wallet.diamonds += 1;
    if (player.id === winnerId) { user.stats.gamesWon += 1; user.wallet.kua += 100; addBadge(user, 'first_win'); }
    if (player.correctThisGame >= room.questions.length && room.questions.length > 0) { user.stats.perfectGames += 1; addBadge(user, 'perfect_game'); }
    if (user.stats.gamesPlayed >= 1) addBadge(user, 'first_game');
    if (user.stats.correctAnswers >= 10) addBadge(user, 'quiz_10');
    if (user.stats.correctAnswers >= 50) addBadge(user, 'quiz_50');
    io.to(userRoom(user.id)).emit('profile:update', publicUser(user));
  }
  saveUsers();
}

function finishGame(room) {
  clearRoomTimers(room);
  awardGameRewards(room);
  room.state = 'finished';
  if (room.isQuickMatch) {
    for (const player of room.players.values()) {
      const status = quickUserStatus.get(player.userId);
      if (status?.socketId === player.id) quickUserStatus.delete(player.userId);
    }
  }
  io.to(room.code).emit('oyun_bitti', {
    siralama: makePlayerList(room),
    toplamSoru: room.questions.length,
  });
  room.resultTimer = setTimeout(() => {
    const current = rooms.get(room.code);
    if (current && current.state === 'finished') deleteRoom(room.code);
  }, 10 * 60 * 1000);
}

function sendNextQuestion(room) {
  clearRoomTimers(room);
  room.questionIndex += 1;

  if (room.questionIndex >= room.questions.length) {
    finishGame(room);
    return;
  }

  room.state = 'playing';
  room.answers.clear();
  room.questionFinished = false;
  room.questionStartedAt = Date.now();

  const question = room.questions[room.questionIndex];
  // Cevap anahtarı sadece soru bitince gönderilir.
  const clientQuestion = {
    soru: question.soru,
    gorsel: question.gorsel || '',
    secenekler: question.secenekler,
  };

  io.to(room.code).emit('sorulari_yukle', {
    index: room.questionIndex,
    toplam: room.questions.length,
    unite: room.unitTitle,
    soru: clientQuestion,
    sure: room.durationSeconds,
  });

  room.questionTimer = setTimeout(
    () => finishQuestion(room.code),
    room.durationSeconds * 1000
  );
}

function finishQuestion(code) {
  const room = rooms.get(code);
  if (!room || room.state !== 'playing' || room.questionFinished) return;

  room.questionFinished = true;
  room.state = 'result';
  if (room.questionTimer) clearTimeout(room.questionTimer);
  room.questionTimer = null;

  const question = room.questions[room.questionIndex];
  for (const player of room.players.values()) {
    const answerData = room.answers.get(player.id);
    const answered = Boolean(answerData);
    const selected = answerData?.answerIndex;
    const correct = answered && selected === question.dogruCevap;
    let bonus = 0;
    let points = 0;
    if (correct) {
      player.correctThisGame = (player.correctThisGame || 0) + 1;
      const durationMs = room.durationSeconds * 1000;
      const elapsedMs = Math.max(0, (answerData.answeredAt || Date.now()) - (room.questionStartedAt || Date.now()));
      const remainingRatio = Math.max(0, Math.min(1, (durationMs - elapsedMs) / durationMs));
      bonus = Math.round(remainingRatio * 50);
      points = 100 + bonus;
      player.score += points;
      player.speedBonusTotal = (player.speedBonusTotal || 0) + bonus;
      const user = player.userId ? findUser(player.userId) : null;
      if (user) {
        user.stats.correctAnswers += 1;
        if (user.stats.correctAnswers >= 10) addBadge(user, 'quiz_10');
        if (user.stats.correctAnswers >= 50) addBadge(user, 'quiz_50');
        if ((player.speedBonusTotal || 0) >= 100) addBadge(user, 'speed_100');
      }
    }
    if (player.userId) {
      const answerUser = findUser(player.userId);
      recordActivity('yarisma_cevabi_sonuc', { user: answerUser, socketId: player.id, roomCode: room.code, mode: room.mode || 'bilgi_yarismasi', details: { questionNumber: room.questionIndex + 1, answered, correct, points, speedBonus: bonus } });
    }
    io.to(player.id).emit('cevap_sonucu', {
      cevaplandi: answered,
      dogruMu: correct,
      dogruIndex: question.dogruCevap,
      puan: points,
      bonus,
      temelPuan: correct ? 100 : 0,
    });
  }
  saveUsers();

  io.to(room.code).emit('soru_sonu', {
    dogruIndex: question.dogruCevap,
    dogruSecenek: question.secenekler[question.dogruCevap],
    cevaplayan: room.answers.size,
    oyuncuSayisi: room.players.size,
    siralama: makePlayerList(room),
  });

  publishRoomState(room);
  room.resultTimer = setTimeout(() => {
    const current = rooms.get(code);
    if (current && current.state === 'result') sendNextQuestion(current);
  }, RESULT_DURATION_MS);
}

function removeSocketFromRoom(socketId, code) {
  const room = rooms.get(code);
  if (!room) return;

  if (room.isQuickMatch) {
    const removedPlayer = room.players.get(socketId);
    if (!removedPlayer) return;
    room.players.delete(socketId);
    room.answers.delete(socketId);
    const status = quickUserStatus.get(removedPlayer.userId);
    if (status?.socketId === socketId) quickUserStatus.delete(removedPlayer.userId);

    if (room.state === 'lobby') {
      if (room.hostId === socketId) room.hostId = room.players.keys().next().value || null;
      if (room.players.size < QUICK_MIN_PLAYERS) {
        requeueQuickLobby(room);
        return;
      }
      publishRoomState(room);
      return;
    }

    if (room.state === 'playing' || room.state === 'result') {
      if (room.players.size === 0) {
        deleteRoom(code);
        return;
      }
      publishRoomState(room);
      if (room.state === 'playing' && room.players.size > 0 && room.answers.size >= room.players.size) finishQuestion(code);
      return;
    }
    publishRoomState(room);
    return;
  }

  const wasHost = room.hostId === socketId;
  const wasPlayer = room.players.has(socketId);
  if (!wasHost && !wasPlayer) return;

  room.players.delete(socketId);
  room.answers.delete(socketId);

  if (wasHost) {
    const nextHostId = room.players.keys().next().value;
    if (!nextHostId) {
      deleteRoom(code);
      return;
    }
    // Odayı devralan oyuncu artık sunucu olur; oyuncu puan tablosundan çıkar.
    room.players.delete(nextHostId);
    room.answers.delete(nextHostId);
    room.hostId = nextHostId;
    io.to(nextHostId).emit('sunucu_devralindi', 'Oda yöneticisi ayrıldı. Yarışmayı artık sen yönetiyorsun.');
  }

  publishRoomState(room);
  if (room.state === 'playing' && room.players.size > 0 && room.answers.size >= room.players.size) {
    finishQuestion(code);
  }
}


function leaveOtherRooms(socket, keepCode = null) {
  removeQuickQueueSocket(socket.id, socket.data.userId);
  for (const code of [...socket.rooms]) {
    if (code === socket.id || code === keepCode) continue;
    if (rooms.has(code)) removeSocketFromRoom(socket.id, code);
    const footballMatch = footballMatches.get(code);
    if (footballMatch) {
      if (footballMatch.hostId === socket.id) {
        clearFootballTimers(footballMatch);
        footballMatches.delete(code);
        io.to(code).emit('football:match-finished', { code, winner: 'cancelled', goals: { ...footballMatch.goals }, text: 'Maç yöneticisi ayrıldığı için maç kapatıldı.' });
      } else {
        for (const team of FOOTBALL_TEAMS) {
          const index = footballMatch.teams[team].findIndex((player) => player.id === socket.id);
          if (index >= 0) footballMatch.teams[team].splice(index, 1);
        }
        if (footballMatch.state === 'lobby') publishFootballLobby(footballMatch);
      }
    }
    socket.leave(code);
  }
}

const FOOTBALL_TEAMS = ['blue', 'orange'];
const FOOTBALL_TEAM_NAMES = { blue: 'Mavi Takım', orange: 'Turuncu Takım' };

function footballLobbyState(match) {
  const teamView = {};
  for (const team of FOOTBALL_TEAMS) {
    teamView[team] = match.teams[team].map((player) => ({
      id: player.id, name: player.name, avatarId: player.avatarId,
    }));
  }
  return {
    code: match.code,
    state: match.state,
    hostName: match.hostName,
    hostId: match.hostId,
    isHost: false,
    unitTitle: match.unitTitle,
    leagueName: match.leagueName || '',
    teams: teamView,
    maxPerTeam: 3,
    canStart: match.teams.blue.length === 3 && match.teams.orange.length === 3,
    goals: { ...match.goals },
    possessions: { ...match.possessions },
  };
}

function publishFootballLobby(match) {
  const state = footballLobbyState(match);
  io.to(match.code).emit('football:lobby', state);
  if (match.hostId) io.to(match.hostId).emit('football:lobby-host', { ...state, isHost: true });
}

function clearFootballTimers(match) {
  if (match.duel?.questionTimer) clearTimeout(match.duel.questionTimer);
  if (match.duel?.nextTimer) clearTimeout(match.duel.nextTimer);
  if (match.nextPossessionTimer) clearTimeout(match.nextPossessionTimer);
}

function footballParticipants(match) {
  return FOOTBALL_TEAMS.flatMap((team) => match.teams[team].map((player) => ({ ...player, team })));
}

function finishFootballMatch(match, forcedWinnerTeam = null) {
  if (!match || match.state === 'finished') return;
  clearFootballTimers(match);
  match.state = 'finished';
  const winner = forcedWinnerTeam || (match.goals.blue > match.goals.orange ? 'blue' : match.goals.orange > match.goals.blue ? 'orange' : 'draw');
  recordActivity('futbol_mac_tamamlandi', { roomCode: match.code, mode: 'futbol', details: { winner, goals: { ...match.goals }, unitTitle: match.unitTitle, bluePlayers: match.teams.blue.map((p) => p.name), orangePlayers: match.teams.orange.map((p) => p.name) } });
  if (match.leagueId) {
    const league = leagues.find((item) => item.id === match.leagueId);
    if (league) {
      for (const player of footballParticipants(match)) {
        if (!league.members.includes(player.userId)) continue;
        const stats = league.scores[player.userId] || { points: 0, games: 0, wins: 0 };
        stats.points += winner === 'draw' ? 7 : winner === player.team ? 10 : 4;
        stats.games += 1;
        if (winner === player.team) stats.wins += 1;
        league.scores[player.userId] = stats;
        const user = findUser(player.userId);
        if (user && stats.wins >= 3) addBadge(user, 'league_champion');
      }
      saveLeagues();
    }
  }
  for (const player of footballParticipants(match)) {
    const user = findUser(player.userId);
    if (!user) continue;
    user.stats.gamesPlayed += 1;
    if (match.goalScorers.includes(player.id)) addBadge(user, 'football_goal');
    if (winner === player.team) {
      user.stats.gamesWon += 1;
      user.wallet.kua += 100;
      addBadge(user, 'first_win');
    } else {
      user.wallet.kua += 50;
    }
    addBadge(user, 'first_game');
    io.to(userRoom(user.id)).emit('profile:update', publicUser(user));
  }
  saveUsers();
  const payload = {
    code: match.code,
    winner,
    goals: { ...match.goals },
    teamNames: FOOTBALL_TEAM_NAMES,
    teams: {
      blue: match.teams.blue.map((p) => ({ name: p.name, avatarId: p.avatarId })),
      orange: match.teams.orange.map((p) => ({ name: p.name, avatarId: p.avatarId })),
    },
    leagueName: match.leagueName || '',
  };
  io.to(match.code).emit('football:match-finished', payload);
  match.cleanupTimer = setTimeout(() => footballMatches.delete(match.code), 10 * 60 * 1000);
}

function endFootballPossession(match, outcomeText) {
  if (match.state !== 'playing') return;
  const previousAttack = match.attackTeam;
  match.possessions[previousAttack] += 1;
  io.to(match.code).emit('football:possession-ended', {
    text: outcomeText,
    goals: { ...match.goals },
    possessions: { ...match.possessions },
    nextTeam: FOOTBALL_TEAM_NAMES[previousAttack === 'blue' ? 'orange' : 'blue'],
  });
  if (match.possessions.blue >= 3 && match.possessions.orange >= 3) {
    finishFootballMatch(match);
    return;
  }
  const nextTeam = previousAttack === 'blue' ? 'orange' : 'blue';
  match.nextPossessionTimer = setTimeout(() => {
    if (match.state !== 'playing') return;
    if (match.possessions[nextTeam] >= 3) {
      const other = nextTeam === 'blue' ? 'orange' : 'blue';
      if (match.possessions[other] >= 3) return finishFootballMatch(match);
      match.attackTeam = other;
    } else {
      match.attackTeam = nextTeam;
    }
    startFootballPhase(match, 'dribble');
  }, 2800);
}

function resolveFootballPhase(match) {
  const duel = match.duel;
  if (!duel || match.state !== 'playing') return;
  clearFootballTimers(match);
  const attacker = duel.score.attacker;
  const defender = duel.score.defender;
  const attackerWins = attacker.correct > defender.correct || (attacker.correct === defender.correct && attacker.speed > defender.speed);
  const phaseLabel = { dribble: 'Top sürme', pass: 'Pas', shot: 'Şut / Kaleci' }[duel.phase];
  let text;
  let nextPhase = null;
  if (duel.phase === 'shot') {
    if (attackerWins) {
      match.goals[match.attackTeam] += 1;
      match.goalScorers.push(duel.attackerId);
      text = 'GOOOL! Şut düellosunu hücum oyuncusu kazandı.';
    } else {
      text = 'KALECİ KURTARDI! Gol olmadı.';
    }
  } else if (attackerWins) {
    text = duel.phase === 'dribble' ? 'Çalım başarılı! Hücum oyuncusu rakibini geçti.' : 'Pas başarılı! Hücum devam ediyor.';
    nextPhase = duel.phase === 'dribble' ? 'pass' : 'shot';
  } else {
    text = duel.phase === 'dribble' ? 'Savunma topu kaptı!' : 'Rakip pası kesti!';
  }
  io.to(match.code).emit('football:phase-result', {
    phase: duel.phase,
    phaseLabel,
    attackerWins,
    attackerCorrect: attacker.correct,
    defenderCorrect: defender.correct,
    attackerSpeed: attacker.speed,
    defenderSpeed: defender.speed,
    text,
    goals: { ...match.goals },
  });
  match.duel = null;
  if (nextPhase) {
    match.nextPossessionTimer = setTimeout(() => startFootballPhase(match, nextPhase), 2600);
  } else {
    endFootballPossession(match, text);
  }
}

function finishFootballQuestion(match) {
  const duel = match.duel;
  if (!duel || duel.questionResolved || match.state !== 'playing') return;
  duel.questionResolved = true;
  if (duel.questionTimer) clearTimeout(duel.questionTimer);
  const question = duel.questions[duel.questionIndex];
  const results = {};
  for (const [role, playerId] of [['attacker', duel.attackerId], ['defender', duel.defenderId]]) {
    const answer = duel.currentAnswers.get(playerId);
    const correct = Boolean(answer && answer.answerIndex === question.dogruCevap);
    let speed = 0;
    if (correct) {
      const elapsed = Math.max(0, (answer.answeredAt || Date.now()) - duel.questionStartedAt);
      speed = 10 + Math.round(Math.max(0, 1 - elapsed / (duel.durationSeconds * 1000)) * 20);
      duel.score[role].correct += 1;
      duel.score[role].speed += speed;
    }
    results[role] = { answered: Boolean(answer), correct, speed };
    io.to(playerId).emit('football:answer-result', {
      answered: Boolean(answer), correct, correctIndex: question.dogruCevap, speed,
      text: correct ? `Doğru! +${speed} hız puanı` : answer ? 'Yanlış cevap.' : 'Süre doldu.',
    });
  }
  io.to(match.code).emit('football:question-result', {
    index: duel.questionIndex,
    correctIndex: question.dogruCevap,
    results,
    score: { attacker: { ...duel.score.attacker }, defender: { ...duel.score.defender } },
  });
  duel.questionResolved = false;
  duel.questionIndex += 1;
  duel.nextTimer = setTimeout(() => {
    if (match.state !== 'playing' || match.duel !== duel) return;
    if (duel.questionIndex >= duel.questions.length) resolveFootballPhase(match);
    else sendFootballQuestion(match);
  }, 1800);
}

function sendFootballQuestion(match) {
  const duel = match.duel;
  if (!duel || match.state !== 'playing') return;
  duel.currentAnswers.clear();
  duel.questionResolved = false;
  duel.questionStartedAt = Date.now();
  const question = duel.questions[duel.questionIndex];
  io.to(match.code).emit('football:question', {
    phase: duel.phase,
    phaseLabel: { dribble: '1. Rakibi geç', pass: '2. Pası tamamla', shot: '3. Kaleciyle karşı karşıya' }[duel.phase],
    index: duel.questionIndex + 1,
    total: duel.questions.length,
    attackerId: duel.attackerId,
    defenderId: duel.defenderId,
    attackerName: duel.attackerName,
    defenderName: duel.defenderName,
    unitTitle: match.unitTitle,
    duration: duel.durationSeconds,
    soru: { soru: question.soru, secenekler: question.secenekler, gorsel: question.gorsel || '' },
    score: { attacker: { ...duel.score.attacker }, defender: { ...duel.score.defender } },
    goals: { ...match.goals },
  });
  duel.questionTimer = setTimeout(() => finishFootballQuestion(match), duel.durationSeconds * 1000);
}

function startFootballPhase(match, phase) {
  if (match.state !== 'playing') return;
  const attackTeam = match.attackTeam;
  const defenseTeam = attackTeam === 'blue' ? 'orange' : 'blue';
  const index = phase === 'dribble' ? 0 : phase === 'pass' ? 1 : 2;
  const attacker = match.teams[attackTeam][index];
  const defender = match.teams[defenseTeam][index];
  const unit = units.get(match.unitId);
  if (!attacker || !defender || !unit || unit.questions.length === 0) {
    endFootballPossession(match, 'Takım oyuncusu eksik olduğu için hücum sona erdi.');
    return;
  }
  match.duel = {
    phase,
    attackerId: attacker.id,
    defenderId: defender.id,
    attackerName: attacker.name,
    defenderName: defender.name,
    questions: shuffledCopy(unit.questions).slice(0, Math.min(3, unit.questions.length)),
    questionIndex: 0,
    durationSeconds: 12,
    currentAnswers: new Map(),
    questionStartedAt: 0,
    questionTimer: null,
    nextTimer: null,
    questionResolved: false,
    score: { attacker: { correct: 0, speed: 0 }, defender: { correct: 0, speed: 0 } },
  };
  io.to(match.code).emit('football:phase', {
    phase,
    phaseLabel: { dribble: 'TOP SÜRME DÜELLOSU', pass: 'PAS DÜELLOSU', shot: 'ŞUT / KALECİ DÜELLOSU' }[phase],
    attackTeam,
    attackTeamName: FOOTBALL_TEAM_NAMES[attackTeam],
    defenseTeamName: FOOTBALL_TEAM_NAMES[defenseTeam],
    attackerName: attacker.name,
    defenderName: defender.name,
    goals: { ...match.goals },
    possessions: { ...match.possessions },
  });
  sendFootballQuestion(match);
}

function startFootballMatch(match) {
  if (match.teams.blue.length !== 3 || match.teams.orange.length !== 3) return false;
  match.state = 'playing';
  match.attackTeam = Math.random() < 0.5 ? 'blue' : 'orange';
  match.possessions = { blue: 0, orange: 0 };
  io.to(match.code).emit('football:started', { code: match.code, attackTeam: match.attackTeam, goals: match.goals });
  startFootballPhase(match, 'dribble');
  return true;
}

io.on('connection', (socket) => {
  const authenticatedUser = findUser(socket.data.userId);
  socket.join(userRoom(socket.data.userId));
  console.log('Giriş yapmış kullanıcı bağlandı:', authenticatedUser ? `@${authenticatedUser.username}` : socket.id);
  recordActivity('socket_baglandi', { user: authenticatedUser, userId: socket.data.userId, socketId: socket.id });

  socket.on('quick:join', () => joinQuickQueue(socket));
  socket.on('quick:cancel', () => cancelQuickForSocket(socket));

  socket.on('football:create', (data = {}) => {
    const unitId = String(data.unitId || '1');
    const unit = units.get(unitId);
    if (!unit || unit.questions.length < 3) return socket.emit('football:error', 'Futbol maçı için seçilen ünitede en az 3 soru bulunmalı.');
    const leagueId = String(data.leagueId || '');
    const league = leagueId ? leagues.find((item) => item.id === leagueId && item.status !== 'archived') : null;
    if (leagueId && (!league || !league.members.includes(socket.data.userId))) return socket.emit('football:error', 'Lig puanlı maç oluşturmak için önce lige katılmalısın.');
    leaveOtherRooms(socket);
    const account = findUser(socket.data.userId);
    const code = createRoomCode();
    const match = {
      code, hostId: socket.id, hostUserId: socket.data.userId,
      hostName: account?.displayName || 'Maç yöneticisi', unitId,
      unitTitle: unit.title, leagueId: league?.id || null, leagueName: league?.name || '',
      teams: { blue: [], orange: [] }, state: 'lobby',
      goals: { blue: 0, orange: 0 }, possessions: { blue: 0, orange: 0 },
      attackTeam: 'blue', duel: null, goalScorers: [], nextPossessionTimer: null, cleanupTimer: null,
    };
    footballMatches.set(code, match);
    socket.join(code);
    socket.emit('football:created', { code });
    publishFootballLobby(match);
    console.log(`Futbol maçı oluşturuldu: ${code} | ${unit.title}`);
    recordActivity('futbol_mac_olusturuldu', { user: account, socketId: socket.id, roomCode: code, mode: 'futbol', details: { unitId, unitTitle: unit.title, leagueId: league?.id || null } });
  });

  socket.on('football:join', (data = {}) => {
    const code = String(data.code || '').trim();
    const team = String(data.team || 'blue');
    const match = footballMatches.get(code);
    const account = findUser(socket.data.userId);
    if (!match) return socket.emit('football:error', 'Futbol maç kodu bulunamadı.');
    if (!FOOTBALL_TEAMS.includes(team)) return socket.emit('football:error', 'Mavi veya turuncu takımı seç.');
    if (match.state !== 'lobby') return socket.emit('football:error', 'Bu maç başlamış.');
    if (match.hostId === socket.id) return socket.emit('football:error', 'Maç yöneticisi oyuncu listesinde yer almaz.');
    if (match.leagueId) {
      const league = leagues.find((item) => item.id === match.leagueId);
      if (!league || !league.members.includes(socket.data.userId)) return socket.emit('football:error', 'Bu maç lig puanlı. Önce ilgili lige katıl.');
    }
    const already = footballParticipants(match).find((player) => player.id === socket.id);
    if (already) return socket.emit('football:error', 'Zaten bu maçın bir takımındasın.');
    if (match.teams[team].length >= 3) return socket.emit('football:error', 'Seçtiğin takım dolu. Diğer takıma katılabilirsin.');
    leaveOtherRooms(socket, code);
    match.teams[team].push({
      id: socket.id, userId: socket.data.userId,
      name: account?.displayName || 'Oyuncu', avatarId: account?.avatarId || 'rocket',
    });
    socket.join(code);
    socket.emit('football:joined', { code, team, teamName: FOOTBALL_TEAM_NAMES[team] });
    publishFootballLobby(match);
    console.log(`${account?.displayName || 'Oyuncu'} ${code} kodlu maçta ${FOOTBALL_TEAM_NAMES[team]} takımına katıldı.`);
    recordActivity('futbol_macina_katildi', { user: account, socketId: socket.id, roomCode: code, mode: 'futbol', details: { team, teamName: FOOTBALL_TEAM_NAMES[team] } });
  });

  socket.on('football:leave', (codeInput) => {
    const code = String(codeInput || '').trim();
    const match = footballMatches.get(code);
    if (!match) return;
    if (match.hostId === socket.id) {
      clearFootballTimers(match);
      footballMatches.delete(code);
      io.to(code).emit('football:match-finished', { code, winner: 'cancelled', goals: { ...match.goals }, text: 'Maç yöneticisi maçı kapattı.' });
      return;
    }
    let removedTeam = null;
    for (const team of FOOTBALL_TEAMS) {
      const index = match.teams[team].findIndex((player) => player.id === socket.id);
      if (index >= 0) { match.teams[team].splice(index, 1); removedTeam = team; }
    }
    socket.leave(code);
    if (removedTeam && match.state === 'playing') finishFootballMatch(match, removedTeam === 'blue' ? 'orange' : 'blue');
    else if (removedTeam) publishFootballLobby(match);
  });

  socket.on('football:start', (codeInput) => {
    const code = String(codeInput || '').trim();
    const match = footballMatches.get(code);
    if (!match) return socket.emit('football:error', 'Maç bulunamadı.');
    if (match.hostId !== socket.id) return socket.emit('football:error', 'Maçı yalnızca oluşturan kişi başlatabilir.');
    if (match.state !== 'lobby') return socket.emit('football:error', 'Maç zaten başlamış.');
    if (!startFootballMatch(match)) return socket.emit('football:error', 'Maçı başlatmak için iki takımın da 3 oyuncuya ulaşması gerekiyor.');
    recordActivity('futbol_mac_basladi', { user: findUser(socket.data.userId), socketId: socket.id, roomCode: code, mode: 'futbol', details: { bluePlayers: match.teams.blue.map((p) => p.name), orangePlayers: match.teams.orange.map((p) => p.name), unitTitle: match.unitTitle } });
  });

  socket.on('football:answer', (data = {}) => {
    const code = String(data.code || '').trim();
    const match = footballMatches.get(code);
    const duel = match?.duel;
    const answerIndex = Number(data.answerIndex);
    if (!match || match.state !== 'playing' || !duel) return;
    if (socket.id !== duel.attackerId && socket.id !== duel.defenderId) return;
    if (!Number.isInteger(answerIndex) || answerIndex < 0 || answerIndex > 3) return;
    if (duel.currentAnswers.has(socket.id)) return;
    duel.currentAnswers.set(socket.id, { answerIndex, answeredAt: Date.now() });
    const answerPlayer = footballParticipants(match).find((p) => p.id === socket.id);
    recordActivity('futbol_duello_cevabi', { user: answerPlayer ? findUser(answerPlayer.userId) : null, socketId: socket.id, roomCode: code, mode: 'futbol', details: { phase: duel.phase, questionNumber: duel.questionIndex + 1 } });
    io.to(code).emit('football:answer-progress', { answered: duel.currentAnswers.size, total: 2 });
    if (duel.currentAnswers.size >= 2) finishFootballQuestion(match);
  });

  socket.on('oda_olustur', (data = {}) => {
    const unitId = String(data.testNo || '1');
    const unit = units.get(unitId);
    if (!unit || unit.questions.length === 0) {
      socket.emit('hata', 'Seçilen ünitede henüz soru bulunmuyor.');
      return;
    }
    const leagueId = String(data.leagueId || '');
    const league = leagueId ? leagues.find((item) => item.id === leagueId && item.status !== 'archived') : null;
    if (leagueId && (!league || !league.members.includes(socket.data.userId))) {
      socket.emit('hata', 'Lig puanlı yarışma oluşturmak için önce ilgili lige katılmalısın.');
      return;
    }

    leaveOtherRooms(socket);

    const requestedCount = Number(data.soruSayisi || 10);
    const count = requestedCount <= 0
      ? Math.min(unit.questions.length, MAX_QUESTIONS_PER_GAME)
      : Math.min(Math.max(1, Math.floor(requestedCount)), unit.questions.length, MAX_QUESTIONS_PER_GAME);
    const duration = [10, 15, 20, 30, 45, 60].includes(Number(data.sure))
      ? Number(data.sure)
      : QUESTION_DURATION_SECONDS_DEFAULT;

    const code = createRoomCode();
    const room = {
      code,
      unitId,
      unitTitle: unit.title,
      leagueId: league?.id || null,
      questions: shuffledCopy(unit.questions).slice(0, count),
      durationSeconds: duration,
      hostId: socket.id,
      players: new Map(),
      answers: new Map(),
      questionIndex: -1,
      state: 'lobby',
      questionFinished: false,
      questionTimer: null,
      resultTimer: null,
    };

    rooms.set(code, room);
    socket.join(code);
    socket.emit('oda_olusturuldu', code);
    publishRoomState(room);
    console.log(`Oda oluşturuldu: ${code} | ${unit.title} | ${count} soru`);
    recordActivity('yarisma_odasi_olusturuldu', { user: findUser(socket.data.userId), socketId: socket.id, roomCode: code, mode: 'bilgi_yarismasi', details: { unitId, unitTitle: unit.title, questionCount: count, durationSeconds: duration, leagueId: league?.id || null } });
  });

  socket.on('oda_katil', (data = {}) => {
    const code = String(data.kod || '').trim();
    const account = findUser(socket.data.userId);
    const name = cleanName(account?.displayName || data.isim);
    const room = rooms.get(code);

    if (!name) return socket.emit('hata', 'Önce adını veya rumuzunu yazmalısın.');
    if (!room) return socket.emit('hata', 'Böyle bir oda bulunamadı. Oda kodunu kontrol et.');
    if (room.hostId === socket.id) return socket.emit('hata', 'Odayı sen kurdun; sunucu olarak zaten odadasın.');
    if (room.state !== 'lobby') return socket.emit('hata', 'Bu yarışma başlamış. Yeni bir oda beklemelisin.');
    if (room.players.size >= MAX_PLAYERS_PER_ROOM) return socket.emit('hata', 'Bu oda dolu.');
    if (room.leagueId) {
      const league = leagues.find((item) => item.id === room.leagueId);
      if (!league || !league.members.includes(socket.data.userId)) return socket.emit('hata', 'Bu yarışma lig puanı kazandırıyor. Önce ligler bölümünden bu lige katıl.');
    }

    leaveOtherRooms(socket, code);

    room.players.set(socket.id, { id: socket.id, userId: socket.data.userId, name, avatarId: account?.avatarId || cleanAvatarId(data.avatarId), score: 0, correctThisGame: 0 });
    socket.join(code);
    socket.emit('katilim_basarili', code);
    publishRoomState(room);
    console.log(`${name}, ${code} kodlu odaya katıldı.`);
    recordActivity('yarisma_odasi_katilim', { user: account, socketId: socket.id, roomCode: code, mode: 'bilgi_yarismasi', details: { playerName: name, unitTitle: room.unitTitle } });
  });

  socket.on('yarismani_baslat', (codeInput) => {
    const code = String(codeInput || '').trim();
    const room = rooms.get(code);
    if (!room) return socket.emit('hata', 'Oda bulunamadı.');
    if (room.hostId !== socket.id) return socket.emit('hata', 'Yarışmayı yalnızca oda yöneticisi başlatabilir.');
    if (room.state !== 'lobby') return socket.emit('hata', 'Yarışma zaten başlatılmış.');
    if (room.players.size === 0) return socket.emit('hata', 'Başlatmadan önce en az bir öğrenci katılmalı.');
    recordActivity('yarisma_basladi', { user: findUser(socket.data.userId), socketId: socket.id, roomCode: code, mode: 'bilgi_yarismasi', details: { playerCount: room.players.size, unitTitle: room.unitTitle } });
    sendNextQuestion(room);
  });

  socket.on('cevap_ver', (data = {}) => {
    const code = String(data.kod || '').trim();
    const answerIndex = Number(data.cevapIndex);
    const room = rooms.get(code);

    if (!room || room.state !== 'playing') return;
    if (!room.players.has(socket.id)) return;
    if (!Number.isInteger(answerIndex) || answerIndex < 0 || answerIndex > 3) return;
    if (room.answers.has(socket.id)) return;

    room.answers.set(socket.id, { answerIndex, answeredAt: Date.now() });

    io.to(code).emit('cevap_durumu', {
      cevaplayan: room.answers.size,
      oyuncuSayisi: room.players.size,
    });
    publishRoomState(room);

    if (room.answers.size >= room.players.size) finishQuestion(code);
  });

  socket.on('disconnect', () => {
    const disconnectedUser = findUser(socket.data.userId);
    recordActivity('socket_baglantisi_kesildi', { user: disconnectedUser, userId: socket.data.userId, socketId: socket.id });
    removeQuickQueueSocket(socket.id, socket.data.userId);
    for (const code of [...rooms.keys()]) removeSocketFromRoom(socket.id, code);
    for (const match of [...footballMatches.values()]) {
      if (match.hostId === socket.id) {
        clearFootballTimers(match);
        io.to(match.code).emit('football:match-finished', { code: match.code, winner: 'cancelled', goals: { ...match.goals }, text: 'Maç yöneticisi ayrıldığı için maç kapatıldı.' });
        footballMatches.delete(match.code);
        continue;
      }
      let removedTeam = null;
      for (const team of FOOTBALL_TEAMS) {
        const index = match.teams[team].findIndex((player) => player.id === socket.id);
        if (index >= 0) { match.teams[team].splice(index, 1); removedTeam = team; }
      }
      if (removedTeam && match.state === 'playing') {
        const winningTeam = removedTeam === 'blue' ? 'orange' : 'blue';
        finishFootballMatch(match, winningTeam);
      } else if (removedTeam) {
        publishFootballLobby(match);
      }
    }
    console.log('Kullanıcı ayrıldı:', socket.id);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  recordActivity('sunucu_basladi', { mode: 'system', details: { port: PORT, units: units.size } });
  const totalQuestions = [...units.values()].reduce((sum, unit) => sum + unit.questions.length, 0);
  console.log('====================================');
  console.log('FENOVA KAHOOT SUNUCUSU ÇALIŞIYOR');
  console.log(`Port: ${PORT}`);
  console.log(`Ünite: ${units.size}`);
  console.log(`Soru: ${totalQuestions}`);
  console.log('Socket.IO: Aktif');
  console.log('Öğrenci kayıt/giriş: Aktif');
  console.log('Arkadaşlık, gruplar, mesajlaşma ve mağaza: Aktif');
  console.log('Lig yönetimi ve lig puanları: Aktif');
  console.log('Futbol bilgi düellosu (3v3): Aktif');
  console.log('Doğru cevap hız bonusu: 0–50 puan');
  console.log('Günlük hediye: günde +100 Kua');
  console.log('Günlük çark: hesap başına günde 1 ücretsiz dönüş');
  console.log('Yerel adres: http://localhost:' + PORT);
  excelReady.then((ready) => {
    console.log(ready ? ('Excel kayıt sistemi: AKTİF — ' + EXCEL_FILE_PATH) : 'Excel kayıt sistemi: KAPALI — yukarıdaki hata mesajını kontrol et.');
  });
  console.log('====================================');
});

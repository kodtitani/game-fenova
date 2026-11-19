#!/usr/bin/env node
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const file = path.join(__dirname, 'store-codes.json');
const args = process.argv.slice(2);
function getArg(name, fallback = '0') { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : fallback; }
const kua = Number(getArg('--kua', '0'));
const diamonds = Number(getArg('--diamonds', '0'));
const itemId = getArg('--item', '');
const allowedItems = new Set(['avatar-dragon','avatar-unicorn','avatar-shark','avatar-phoenix','avatar-alien','avatar-capybara','avatar-gold-robot','frame-galaxy','frame-crown','theme-neon']);
if (!Number.isInteger(kua) || kua < 0 || kua > 100000 || !Number.isInteger(diamonds) || diamonds < 0 || diamonds > 1000 || (itemId && !allowedItems.has(itemId)) || (kua === 0 && diamonds === 0 && !itemId)) {
  console.log('Kullanım: node create-store-code.js --kua 500 --diamonds 5 [--item avatar-dragon]');
  console.log('Kua: 0-100000, Elmas: 0-1000. En az bir ödül belirt.');
  process.exit(1);
}
let data = { codes: [] };
if (fs.existsSync(file)) {
  data = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (!data || !Array.isArray(data.codes)) throw new Error('store-codes.json biçimi hatalı.');
}
const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
function part(n) { return Array.from({length:n}, () => alphabet[crypto.randomInt(alphabet.length)]).join(''); }
let code;
let hash;
do { code = `FNV-${part(4)}-${part(6)}`; hash = crypto.createHash('sha256').update(code).digest('hex'); }
while (data.codes.some((entry) => entry.hash === hash));
const reward = {};
if (kua) reward.kua = kua;
if (diamonds) reward.diamonds = diamonds;
if (itemId) reward.itemId = itemId;
data.codes.push({ hash, reward, createdAt: new Date().toISOString(), usedAt: null, usedBy: null });
const temp = `${file}.tmp`;
fs.writeFileSync(temp, JSON.stringify(data, null, 2) + '\n', 'utf8');
fs.renameSync(temp, file);
console.log('\nYeni tek kullanımlık Fenova mağaza kodu:');
console.log(code);
console.log('Ödül:', JSON.stringify(reward));
console.log('\nKod yalnızca bir öğrenci tarafından kullanılabilir. Kodu güvenli şekilde paylaş.');
console.log('Render gibi bir sunucuya aktaracaksan store-codes.json dosyasını da proje dağıtımına ekle.');

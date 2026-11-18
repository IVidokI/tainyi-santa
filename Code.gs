// Таємний Санта ПСХ — бекенд (Google Apps Script, прив'язаний до Google-таблиці).
//
// Аркуші:
//   «Учасники» — колонка A: ПІБ (з рядка 2, рядок 1 — заголовок). Джерело правди про учасників.
//   «Пари»     — створюється автоматично при першому зверненні (див. ensurePairs_):
//                A Санта | B Кому дарує | C Токен | D–H Факти 1–5 | I Спроби | J Статус | K Здогадки | L Оновлено
//
// Розподіл — один випадковий цикл по всіх учасниках (A→B→C→…→A): ніхто не витягне себе,
// кожен має рівно одного Санту і дарує рівно одній людині.
//
// Усі запити — POST з JSON у тілі (Content-Type: text/plain, щоб не було CORS preflight).

var SHEET_PEOPLE_ = 'Учасники';
var SHEET_PAIRS_ = 'Пари';
var MAX_ATTEMPTS_ = 3;
var FACT_MAX_LEN_ = 300;
var COL_ = { GIVER: 0, RECIPIENT: 1, TOKEN: 2, FACTS: 3, ATTEMPTS: 8, STATUS: 9, GUESSES: 10, UPDATED: 11 };

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'list':      return out_({ status: 'ok', names: listNames_() });
      case 'login':     return out_(login_(data));
      case 'saveFacts': return out_(saveFacts_(data));
      case 'open':      return out_(open_(data));
      case 'guess':     return out_(guess_(data));
      default:          return out_({ status: 'error', message: 'Невідома дія' });
    }
  } catch (err) {
    return out_({ status: 'error', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_({ status: 'ok', message: 'Таємний Санта ПСХ' });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- допоміжне ----------

function norm_(s) {
  return String(s || '').replace(/[ʼ’`']/g, "'").replace(/ё/gi, 'е').replace(/\s+/g, ' ').trim().toLowerCase();
}

function listNames_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEOPLE_);
  if (!sh) throw new Error('Немає аркуша «' + SHEET_PEOPLE_ + '»');
  var seen = {}, names = [];
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var n = String(r[0] || '').replace(/\s+/g, ' ').trim();
    if (n && !seen[norm_(n)]) { seen[norm_(n)] = true; names.push(n); }
  });
  return names.sort(function (a, b) { return a.localeCompare(b, 'uk'); });
}

function shuffle_(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

function newToken_() {
  var chars = 'abcdefghijkmnpqrstuvwxyz23456789', t = '';
  for (var i = 0; i < 12; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

// Створює аркуш «Пари» один раз. Якщо аркуш уже є — нічого не змінює
// (щоб перезавантаження/повторний вхід не міняло розподіл).
function ensurePairs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PAIRS_);
  if (sh) return sh;
  var names = listNames_();
  if (names.length < 3) throw new Error('Потрібно щонайменше 3 учасники в аркуші «' + SHEET_PEOPLE_ + '»');
  shuffle_(names);
  sh = ss.insertSheet(SHEET_PAIRS_);
  var rows = [['Санта', 'Кому дарує', 'Токен', 'Факт 1', 'Факт 2', 'Факт 3', 'Підказка 1', 'Підказка 2', 'Спроби', 'Статус', 'Здогадки', 'Оновлено']];
  var used = {};
  for (var i = 0; i < names.length; i++) {
    var tok; do { tok = newToken_(); } while (used[tok]); used[tok] = true;
    rows.push([names[i], names[(i + 1) % names.length], tok, '', '', '', '', '', 0, 'чекає', '', '']);
  }
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return sh;
}

// Повертає {sh, rows (усі значення), idx (0-based індекс у rows)} для Санти з таким ПІБ.
function findByGiver_(name) {
  var sh = ensurePairs_();
  var rows = sh.getDataRange().getValues();
  var key = norm_(name);
  for (var i = 1; i < rows.length; i++) if (norm_(rows[i][COL_.GIVER]) === key) return { sh: sh, rows: rows, idx: i };
  return null;
}

function findByToken_(token) {
  var sh = ensurePairs_();
  var rows = sh.getDataRange().getValues();
  var t = String(token || '');
  if (!t) return null;
  for (var i = 1; i < rows.length; i++) if (String(rows[i][COL_.TOKEN]) === t) return { sh: sh, rows: rows, idx: i };
  return null;
}

function factsOf_(row) {
  return [0, 1, 2, 3, 4].map(function (k) { return String(row[COL_.FACTS + k] || ''); });
}

// ---------- дії ----------

// Санта вводить своє ПІБ → дізнається, кому дарує (повторний вхід показує того ж).
function login_(d) {
  var f = findByGiver_(d.name);
  if (!f) return { status: 'error', message: 'Не знайшли такого ПІБ у списку учасників. Оберіть себе зі списку.' };
  var row = f.rows[f.idx];
  return {
    status: 'ok',
    giver: row[COL_.GIVER],
    recipient: row[COL_.RECIPIENT],
    token: row[COL_.TOKEN],
    facts: factsOf_(row),
    locked: Number(row[COL_.ATTEMPTS]) > 0,
    state: String(row[COL_.STATUS])
  };
}

function saveFacts_(d) {
  var f = findByGiver_(d.name);
  if (!f) return { status: 'error', message: 'Учасника не знайдено' };
  var row = f.rows[f.idx];
  if (Number(row[COL_.ATTEMPTS]) > 0) return { status: 'error', message: 'Одержувач уже почав відгадувати — змінити факти не можна.' };
  var facts = (d.facts || []).map(function (s) { return String(s || '').replace(/\s+/g, ' ').trim(); });
  if (facts.length !== 5 || facts.some(function (s) { return !s; })) return { status: 'error', message: 'Потрібно заповнити всі 5 фактів' };
  if (facts.some(function (s) { return s.length > FACT_MAX_LEN_; })) return { status: 'error', message: 'Факт занадто довгий (максимум ' + FACT_MAX_LEN_ + ' символів)' };
  var r = f.idx + 1;
  f.sh.getRange(r, COL_.FACTS + 1, 1, 5).setValues([facts]);
  f.sh.getRange(r, COL_.STATUS + 1).setValue('грає');
  f.sh.getRange(r, COL_.UPDATED + 1).setValue(new Date());
  return { status: 'ok', token: row[COL_.TOKEN] };
}

// Одержувач сканує QR (токен) → бачить 3 факти + відкриті підказки.
function open_(d) {
  var f = findByToken_(d.token);
  if (!f) return { status: 'error', message: 'Цей QR-код не розпізнано.' };
  return view_(f.rows[f.idx]);
}

function view_(row) {
  var facts = factsOf_(row);
  var attempts = Number(row[COL_.ATTEMPTS]) || 0;
  var status = String(row[COL_.STATUS]);
  if (!facts[0]) return { status: 'ok', state: 'empty' };
  var shown = facts.slice(0, 3);
  if (attempts >= 1) shown.push(facts[3]);
  if (attempts >= 2) shown.push(facts[4]);
  var guesses = String(row[COL_.GUESSES] || '').split(' | ').filter(String);
  var res = { status: 'ok', state: status === 'вгадано' ? 'won' : (status === 'не вгадано' ? 'lost' : 'play'), facts: shown, attempts: attempts, max: MAX_ATTEMPTS_, guesses: guesses };
  if (res.state === 'won') res.giver = row[COL_.GIVER];
  return res;
}

function guess_(d) {
  var f = findByToken_(d.token);
  if (!f) return { status: 'error', message: 'Цей QR-код не розпізнано.' };
  var row = f.rows[f.idx];
  var attempts = Number(row[COL_.ATTEMPTS]) || 0;
  var status = String(row[COL_.STATUS]);
  if (!row[COL_.FACTS]) return { status: 'error', message: 'Санта ще не написав факти.' };
  if (status === 'вгадано' || status === 'не вгадано' || attempts >= MAX_ATTEMPTS_) return view_(row);
  var name = String(d.name || '').trim();
  if (!name) return { status: 'error', message: 'Оберіть ім\'я' };
  var guesses = String(row[COL_.GUESSES] || '').split(' | ').filter(String);
  if (guesses.some(function (g) { return norm_(g) === norm_(name); })) return { status: 'error', message: 'Цю людину ви вже називали.' };

  attempts++;
  guesses.push(name);
  var correct = norm_(name) === norm_(row[COL_.GIVER]);
  var newStatus = correct ? 'вгадано' : (attempts >= MAX_ATTEMPTS_ ? 'не вгадано' : 'грає');
  var r = f.idx + 1;
  f.sh.getRange(r, COL_.ATTEMPTS + 1).setValue(attempts);
  f.sh.getRange(r, COL_.STATUS + 1).setValue(newStatus);
  f.sh.getRange(r, COL_.GUESSES + 1).setValue(guesses.join(' | '));
  f.sh.getRange(r, COL_.UPDATED + 1).setValue(new Date());
  row[COL_.ATTEMPTS] = attempts; row[COL_.STATUS] = newStatus; row[COL_.GUESSES] = guesses.join(' | ');
  var v = view_(row);
  v.correct = correct;
  return v;
}

// ---------- для адміністратора (▶ Run у редакторі) ----------

// Повністю перегенерувати розподіл (видаляє аркуш «Пари» разом із фактами!).
function resetPairs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PAIRS_);
  if (sh) ss.deleteSheet(sh);
  ensurePairs_();
}

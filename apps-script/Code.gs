/**
 * Органайзер — сервер синхронизации на Google Apps Script.
 *
 * Данные хранятся в файле organizer-data.json в папке «Органайзер»
 * на вашем Google Диске. Скрипт принимает изменения с устройств,
 * сводит их (побеждает более позднее изменение каждой записи)
 * и возвращает устройству всё, чего у него ещё нет.
 *
 * «Входящие»: файлы organizer-inbox-*.json на вашем Диске (их может
 * создавать Claude по вашей просьбе) при следующей синхронизации
 * вносятся в данные и отправляются в корзину.
 *
 * Скриншоты: команда на iPhone распознаёт текст на последнем скриншоте
 * и присылает его сюда (action: "ocr"); скрипт находит на нём названия
 * ваших счетов и обновляет суммы.
 *
 * Таймшит: записи «Работа» из приложения переводятся на английский
 * (LanguageApp) и вписываются в таблицу недели в папке «Органайзер/Таймшиты»
 * по образцу таблицы Timesheet_template; рядом сохраняется .xlsx.
 *
 * Первый запуск: выберите функцию setup и нажмите «Выполнить».
 * В журнале появится ключ доступа — его нужно вставить в приложение.
 */

const FOLDER_NAME = 'Органайзер';
const FILE_NAME = 'organizer-data.json';
const INBOX_PREFIX = 'organizer-inbox';
const TOMBSTONE_TTL_DAYS = 180; // сколько дней помнить об удалённых записях
const KINDS = ['group', 'item', 'account', 'settings', 'task', 'goal', 'tx', 'log', 'project'];

// --- Таймшит ---
const TS_FOLDER = 'Таймшиты';               // папка внутри «Органайзер» для рабочих Google-таблиц
const TS_EXPORT_FOLDER = 'Отчеты для Эрика'; // сюда кладётся готовый .xlsx (папка синхронизируется с компьютером)
const TS_TEMPLATE = 'Timesheet_template';   // таблица-образец (Google Таблица)
const TS_USER = 'Sergey Korobkov';          // как в названии файла
const TS_SLOTS = { 1: '9:00-13:00', 2: '13:30-17:30' };
const TS_ROWS = { 1: 7, 2: 8 };             // строки утра и дня первого проекта, текст в столбце C
const TS_MAX_PROJECTS = 6;                  // проекты дня: строки 7–8, 9–10 … 17–18
const DAYS_EN = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const MONTHS_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const ID_RE = /^[A-Za-z0-9_\-.~:@+]{1,200}$/;

/** Однократная настройка: создаёт ключ доступа и файл данных. Повторный запуск ключ не меняет. */
function setup() {
  // Google может выдать разрешения не все (галочки на экране согласия) — просим недостающие
  try { ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL); } catch (err) {}
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('SECRET');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('SECRET', key);
  }
  const file = getDataFile_();
  Logger.log('Файл данных: ' + file.getUrl());
  Logger.log('Владелец: ' + Session.getEffectiveUser().getEmail());
  Logger.log('КЛЮЧ ДОСТУПА (вставьте в приложение): ' + key);
  // проверка доступа к переводчику и внешним запросам (чтобы Google спросил разрешения сразу)
  Logger.log('Проверка перевода: ' + LanguageApp.translate('Проверка связи', 'ru', 'en'));
  try { UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); } catch (err) {}
  // проверка доступа к таблицам (без этого разрешения таймшит не собирается)
  const tpl = findSheet_(null, TS_TEMPLATE);
  if (!tpl) {
    Logger.log('Таблица-образец «' + TS_TEMPLATE + '» не найдена — таймшит собрать не получится');
  } else {
    try {
      SpreadsheetApp.openById(tpl.getId());
      Logger.log('Доступ к Google Таблицам: есть');
    } catch (err) {
      Logger.log('НЕТ ДОСТУПА К GOOGLE ТАБЛИЦАМ: ' + (err && err.message || err));
      Logger.log('Запустите setup ещё раз и на экране Google отметьте ВСЕ галочки (или «Выбрать все»).');
      return;
    }
  }
  // пересобираем недели, которые раньше не получились (например, не хватало разрешений)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) { Logger.log('Сейчас идёт синхронизация — таймшит пересоберётся при следующей'); return; }
  try {
    const db = readDb_(file);
    const failed = failedWeeks_(db, 0);
    if (!failed.length) { Logger.log('Таймшиты в порядке'); return; }
    failed.forEach(function (monday) {
      try {
        const info = rebuildWeek_(db, monday);
        Logger.log('Таймшит недели ' + monday + ': ' + (info.error ? 'ошибка — ' + info.error : info.name));
      } catch (err) {
        setTsRecord_(db, monday, { error: String(err && err.message || err) });
        Logger.log('Таймшит недели ' + monday + ': ошибка — ' + (err && err.message || err));
      }
    });
    writeDb_(file, db);
  } finally {
    lock.releaseLock();
  }
}

/** Недели, таймшит которых в прошлый раз не собрался (ошибка старше minAgeMs). */
function failedWeeks_(db, minAgeMs) {
  const now = Date.now();
  return Object.keys(db.records).filter(function (id) {
    const r = db.records[id];
    return r && r.kind === 'tsfile' && r.error && r.week && !r.del && now - (Number(r.u) || 0) >= minAgeMs;
  }).map(function (id) { return db.records[id].week; });
}

/** Пересобрать таймшит текущей недели вручную (из редактора скрипта). */
function rebuildCurrentWeek() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) { Logger.log('Занято, попробуйте ещё раз'); return; }
  try {
    const file = getDataFile_();
    const db = readDb_(file);
    const now = new Date();
    const ymd = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const info = rebuildWeek_(db, mondayYmd_(ymd));
    writeDb_(file, db);
    Logger.log(JSON.stringify(info));
  } finally {
    lock.releaseLock();
  }
}

/** Сменить ключ доступа (если ссылка с ключом попала к посторонним). */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('SECRET');
  setup();
}

/** Разобрать «входящие» вручную (из редактора скрипта), не дожидаясь синхронизации. */
function processInboxNow() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) { Logger.log('Занято, попробуйте ещё раз'); return; }
  try {
    const file = getDataFile_();
    const db = readDb_(file);
    const n = applyInbox_(db);
    if (n) writeDb_(file, db);
    Logger.log('Внесено записей из входящих: ' + n);
  } finally {
    lock.releaseLock();
  }
}

/** Открытие адреса в браузере — простая проверка, что скрипт работает. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.key) return json_({ ok: true, service: 'life-organizer', hint: 'Скрипт работает. Вставьте этот адрес в настройки приложения.' });
  return handle_({ action: 'ping', key: p.key });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  return handle_(req);
}

function handle_(req) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) return json_({ ok: false, error: 'not_setup' });
  if (!req || req.key !== secret) return json_({ ok: false, error: 'auth' });
  if (req.action === 'ping') return json_({ ok: true, service: 'life-organizer' });
  if (req.action === 'ocr') return handleOcr_(req);
  if (req.action !== 'sync') return json_({ ok: false, error: 'bad_action' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'busy' });
  try {
    const file = getDataFile_();
    const db = readDb_(file);
    let changed = false;

    const weeks = {}; // недели таймшита, которые нужно пересобрать

    // 1. Входящие файлы
    try {
      if (applyInbox_(db, weeks) > 0) changed = true;
    } catch (err) {
      // входящие не должны ломать синхронизацию
    }

    // 2. Изменения с устройства
    const since = Number(req.since) || 0;
    const forced = {};
    const changes = Array.isArray(req.changes) ? req.changes : [];
    changes.forEach(function (r) {
      if (!r || typeof r.id !== 'string' || !r.id || r.id.length > 200) return;
      const ex = db.records[r.id];
      if (!ex || (Number(r.u) || 0) > (Number(ex.u) || 0)) {
        if (r.kind === 'log') prepareLog_(r, ex, weeks);
        db.rev += 1;
        r.rv = db.rev;
        db.records[r.id] = r;
        changed = true;
      } else {
        // На сервере версия новее: вернём её устройству
        forced[r.id] = true;
      }
    });

    // 3. Забываем давно удалённые записи
    const cutoff = Date.now() - TOMBSTONE_TTL_DAYS * 86400000;
    Object.keys(db.records).forEach(function (id) {
      const r = db.records[id];
      if (r && r.del && (Number(r.u) || 0) < cutoff) {
        delete db.records[id];
        changed = true;
      }
    });

    // 4. Таймшит: пересобираем затронутые недели и те, что раньше не получились (не чаще раза в минуту)
    failedWeeks_(db, 60000).forEach(function (monday) { weeks[monday] = true; });
    Object.keys(weeks).forEach(function (monday) {
      try { rebuildWeek_(db, monday); } catch (err) { setTsRecord_(db, monday, { error: String(err && err.message || err) }); }
      changed = true;
    });

    if (changed) writeDb_(file, db);

    const out = [];
    Object.keys(db.records).forEach(function (id) {
      const r = db.records[id];
      if ((Number(r.rv) || 0) > since || forced[id]) out.push(r);
    });
    return json_({ ok: true, rev: db.rev, records: out });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Вносит файлы organizer-inbox-*.json, принадлежащие владельцу скрипта.
 * Формат файла:
 *   {
 *     "records": [ { "id": "...", "kind": "item", ... } ],      // создать или заменить запись целиком
 *     "patches": [ { "id": "...", "set": { "amount": 700000 } } ], // поменять отдельные поля
 *     "deletes": [ "id1", "id2" ]                                  // удалить
 *   }
 * Возвращает число внесённых изменений.
 */
function applyInbox_(db, weeks) {
  weeks = weeks || {};
  const me = Session.getEffectiveUser().getEmail();
  const it = DriveApp.searchFiles("title contains '" + INBOX_PREFIX + "' and trashed = false");
  let applied = 0;
  const files = [];
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return a.getDateCreated().getTime() - b.getDateCreated().getTime(); });

  files.forEach(function (f) {
    if (f.getName().indexOf(INBOX_PREFIX) !== 0) return;
    let owner = '';
    try { owner = f.getOwner() ? f.getOwner().getEmail() : ''; } catch (err) { owner = ''; }
    if (!me || owner !== me) return; // чужие файлы не трогаем

    let box;
    try { box = JSON.parse(f.getBlob().getDataAsString('UTF-8')); } catch (err) { box = null; }
    if (box && typeof box === 'object') {
      const now = Date.now();
      const stamp = function (id) {
        const ex = db.records[id];
        return Math.max(now, ex ? (Number(ex.u) || 0) + 1 : 0);
      };
      (Array.isArray(box.records) ? box.records : []).forEach(function (r) {
        if (!r || typeof r.id !== 'string' || !ID_RE.test(r.id) || KINDS.indexOf(r.kind) < 0) return;
        const rec = JSON.parse(JSON.stringify(r));
        rec.u = stamp(rec.id);
        if (rec.kind === 'log') prepareLog_(rec, db.records[rec.id], weeks);
        db.rev += 1;
        rec.rv = db.rev;
        db.records[rec.id] = rec;
        applied++;
      });
      (Array.isArray(box.patches) ? box.patches : []).forEach(function (p) {
        if (!p || typeof p.id !== 'string' || !p.set || typeof p.set !== 'object') return;
        const ex = db.records[p.id];
        if (!ex || ex.del) return;
        const rec = JSON.parse(JSON.stringify(ex));
        Object.keys(p.set).forEach(function (k) {
          if (k === 'id' || k === 'u' || k === 'rv') return;
          rec[k] = p.set[k];
        });
        rec.u = stamp(rec.id);
        if (rec.kind === 'log') prepareLog_(rec, ex, weeks);
        db.rev += 1;
        rec.rv = db.rev;
        db.records[rec.id] = rec;
        applied++;
      });
      (Array.isArray(box.deletes) ? box.deletes : []).forEach(function (id) {
        const ex = db.records[id];
        if (!ex || ex.del) return;
        db.rev += 1;
        if (ex.kind === 'log' && ex.date) weeks[mondayYmd_(ex.date)] = true;
        db.records[id] = { id: id, kind: ex.kind, del: true, date: ex.date, u: stamp(id), rv: db.rev };
        applied++;
      });
    }
    f.setTrashed(true); // разобран (или повреждён) — убираем в корзину, восстановить можно оттуда
  });
  return applied;
}

/* ================= Суммы со скриншота ================= */

function handleOcr_(req) {
  const text = String(req.text || '');
  if (!text.trim()) return json_({ ok: false, error: 'empty', message: 'На скриншоте не найден текст' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'busy', message: 'Сервер занят, попробуйте ещё раз' });
  try {
    const file = getDataFile_();
    const db = readDb_(file);
    const res = applyOcr_(db, text);
    if (res.changed.length) writeDb_(file, db);
    let message;
    if (!res.found.length) message = 'Счета на скриншоте не найдены. Названия счетов в органайзере должны совпадать с названиями на скриншоте.';
    else if (!res.changed.length) message = 'Суммы не изменились (' + res.found.length + ' ' + plural_(res.found.length, ['счёт', 'счёта', 'счетов']) + ')';
    else message = 'Обновлено: ' + res.changed.map(function (c) { return c.name + ' ' + fmt_(c.amount); }).join('; ');
    return json_({ ok: true, message: message, changed: res.changed, found: res.found.length });
  } finally {
    lock.releaseLock();
  }
}

function normText_(s) {
  return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9%]/g, '');
}

/* «85 000,00 руб», «2 700,000 €», «-1 234,5», «1,234.56» → число */
function parseOcrAmount_(line) {
  const m = String(line).replace(/[−–]/g, '-').match(/^\s*(-?\s*\d[\d\s\u00a0\u202f.,']*)/);
  if (!m) return null;
  let s = m[1].replace(/[\s\u00a0\u202f']/g, '').replace(/[.,]+$/, '');
  const lastComma = s.lastIndexOf(','), lastDot = s.lastIndexOf('.');
  if (lastComma >= 0 && lastDot >= 0) {
    // оба разделителя: десятичный — последний
    if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (lastComma >= 0) {
    s = s.replace(/,(?=.*,)/g, '').replace(',', '.');
  } else if (lastDot >= 0 && (s.match(/\./g) || []).length > 1) {
    s = s.replace(/\./g, '');
  }
  const v = parseFloat(s);
  return isFinite(v) ? Math.round(v * 100) / 100 : null;
}

function applyOcr_(db, text) {
  const lines = String(text).split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  const accs = Object.keys(db.records).map(function (id) { return db.records[id]; })
    .filter(function (r) { return r && r.kind === 'account' && !r.del && r.name; });
  const found = [], changed = [], seen = {};

  lines.forEach(function (line, i) {
    const nl = normText_(line.replace(/[.…]+\s*$/, ''));
    if (!nl) return;
    let best = null, bestScore = 0, bestLen = 0, rest = '';
    accs.forEach(function (a) {
      const na = normText_(a.name);
      if (!na) return;
      let score = 0;
      if (nl === na) score = 3;
      else if (nl.indexOf(na) === 0) score = 2;               // в строке после названия идёт сумма
      else if (nl.length >= 6 && na.indexOf(nl) === 0) score = 1; // название обрезано на скриншоте
      if (score > bestScore || (score === bestScore && score && na.length > bestLen)) { best = a; bestScore = score; bestLen = na.length; }
    });
    if (!best || seen[best.id]) return;

    // сумма: в той же строке после названия или в одной из двух следующих строк
    let amount = null;
    if (bestScore === 2) {
      const afterName = line.length > best.name.length ? line.slice(best.name.length) : '';
      amount = parseOcrAmount_(afterName.replace(/^[^0-9\-−]+/, ''));
    }
    for (let k = 1; amount === null && k <= 2 && i + k < lines.length; k++) {
      amount = parseOcrAmount_(lines[i + k]);
    }
    if (amount === null) return;
    seen[best.id] = true;
    found.push(best.id);
    if (Number(best.amount) !== amount) {
      const rec = JSON.parse(JSON.stringify(best));
      rec.amount = amount;
      rec.at = Date.now();
      rec.u = Math.max(Date.now(), (Number(best.u) || 0) + 1);
      db.rev += 1;
      rec.rv = db.rev;
      db.records[rec.id] = rec;
      changed.push({ name: rec.name, amount: amount, cur: rec.cur || 'RUB' });
    }
  });
  return { found: found, changed: changed };
}

function plural_(n, f) {
  const a = Math.abs(n) % 100, b = a % 10;
  if (a > 10 && a < 20) return f[2];
  if (b > 1 && b < 5) return f[1];
  if (b === 1) return f[0];
  return f[2];
}

function fmt_(n) {
  const s = (Math.round(n * 100) / 100).toString().split('.');
  return s[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (s[1] ? ',' + s[1] : '');
}

/* ================= Таймшит ================= */

/** Перевод записи и отметка недели (старой и новой, если дату поменяли). */
function prepareLog_(r, ex, weeks) {
  if (ex && ex.kind === 'log' && ex.date) weeks[mondayYmd_(ex.date)] = true;
  if (r.del) { if (r.date) weeks[mondayYmd_(r.date)] = true; return; }
  if (!r.date) return;
  weeks[mondayYmd_(r.date)] = true;
  const text = String(r.text || '').trim();
  if (!text) { r.en = ''; r.enSrc = ''; return; }
  if (r.enSrc === text && r.en) return;               // уже переведено
  if (ex && ex.enSrc === text && ex.en) { r.en = ex.en; r.enSrc = text; return; }
  if (/[а-яё]/i.test(text)) {
    try { r.en = LanguageApp.translate(text, 'ru', 'en'); } catch (err) { r.en = ''; }
  } else {
    r.en = text;                                       // уже по-английски
  }
  r.enSrc = text;
}

function parseYmd_(ymd) {
  const p = String(ymd).split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
}
function ymdOf_(d) {
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function mondayYmd_(ymd) {
  const d = parseYmd_(ymd);
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
  return ymdOf_(d);
}

function tsFolder_() {
  const base = getDataFile_().getParents();
  const parent = base.hasNext() ? base.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName(TS_FOLDER);
  return it.hasNext() ? it.next() : parent.createFolder(TS_FOLDER);
}

/** Папка для .xlsx: «Отчеты для Эрика», если она есть на Диске, иначе — папка таймшитов. */
function exportFolder_() {
  const it = DriveApp.getFoldersByName(TS_EXPORT_FOLDER);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return tsFolder_();
}

function findSheet_(folder, name) {
  const it = folder ? folder.getFilesByName(name) : DriveApp.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && f.getMimeType() === MimeType.GOOGLE_SHEETS) return f;
  }
  return null;
}

function setTsRecord_(db, monday, data) {
  const id = 'ts-' + monday;
  const ex = db.records[id] || {};
  const rec = Object.assign({}, ex, data, { id: id, kind: 'tsfile', week: monday, updated: new Date().toISOString() });
  if (!data.error) delete rec.error;
  rec.u = Math.max(Date.now(), (Number(ex.u) || 0) + 1);
  db.rev += 1;
  rec.rv = db.rev;
  db.records[id] = rec;
  return rec;
}

/**
 * Пересобирает таблицу недели (понедельник monday) по всем записям «Работа»:
 * B2 — номер недели в месяце, B3 — дата, C7/C8 — утро/день с переводом.
 * Затем сохраняет копию в .xlsx рядом.
 */
function rebuildWeek_(db, monday) {
  const mon = parseYmd_(monday);
  const wk = Math.ceil(mon.getUTCDate() / 7);
  const name = 'Timesheet_wk' + wk + '_' + MONTHS_EN[mon.getUTCMonth()] + '_' + TS_USER;
  const folder = tsFolder_();

  let file = findSheet_(folder, name);
  let fresh = false;
  if (!file) {
    const tpl = findSheet_(folder, TS_TEMPLATE) || findSheet_(null, TS_TEMPLATE);
    if (!tpl) throw new Error('Не найдена таблица-образец «' + TS_TEMPLATE + '»');
    file = tpl.makeCopy(name, folder);
    fresh = true;
  }
  const ss = SpreadsheetApp.openById(file.getId());

  const logs = Object.keys(db.records).map(function (id) { return db.records[id]; })
    .filter(function (r) { return r && r.kind === 'log' && !r.del && r.date; });

  ss.getSheets().forEach(function (sh) {
    const dayIdx = DAYS_EN.indexOf(sh.getName().trim());
    if (dayIdx < 0) return;
    const d = new Date(mon.getTime());
    d.setUTCDate(d.getUTCDate() + dayIdx);
    const ymd = ymdOf_(d);
    sh.getRange('B2').setValue(wk);
    sh.getRange('B3').setNumberFormat('@').setValue(('0' + d.getUTCDate()).slice(-2) + '.' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '.' + d.getUTCFullYear());
    const day = logs.filter(function (r) { return r.date === ymd; })
      .sort(function (a, b) { return (Number(a.createdAt) || 0) - (Number(b.createdAt) || 0); });

    // проекты дня в порядке первой записи; без проекта — строки 7/8 с номером и названием из образца
    const projects = [], index = {};
    day.forEach(function (r) {
      const key = (r.pcode || r.pname) ? (String(r.pcode || '') + '|' + String(r.pname || '')) : '';
      if (!(key in index)) { index[key] = projects.length; projects.push({ key: key, code: r.pcode || '', name: r.pname || '', items: [] }); }
      projects[index[key]].items.push(r);
    });
    if (!projects.length) projects.push({ key: '', items: [] });

    // дополнительные строки проектов (9–18) очищаем и заполняем заново
    sh.getRange(TS_ROWS[1] + 2, 1, (TS_MAX_PROJECTS - 1) * 2, 3).clearContent();
    projects.slice(0, TS_MAX_PROJECTS).forEach(function (p, k) {
      [1, 2].forEach(function (slot) {
        const row = TS_ROWS[slot] + 2 * k;
        const texts = p.items.filter(function (r) { return Number(r.slot) === slot; })
          .map(function (r) { return String(r.en || r.text || '').trim(); }).filter(Boolean);
        if (p.key) {
          sh.getRange(row, 1).setValue(p.code);
          sh.getRange(row, 2).setValue(p.name);
        }
        sh.getRange(row, 3).setValue(TS_SLOTS[slot] + (texts.length ? ' ' + texts.join('; ') : ''));
      });
    });
    if (fresh) sh.getRange(TS_ROWS[1], 4, 2, 7).clearContent(); // часы D:J утра и дня
  });
  SpreadsheetApp.flush();

  // копия в .xlsx (старую с тем же именем — в корзину)
  const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(name + '.xlsx');
  const out = exportFolder_();
  // файл с таким же названием: если неделю собираем впервые — это ваш исходный файл,
  // его не удаляем, а переименовываем; иначе это наша прошлая версия — в корзину
  const ours = !!(db.records['ts-' + monday] && db.records['ts-' + monday].xlsxUrl);
  const old = out.getFilesByName(name + '.xlsx');
  while (old.hasNext()) {
    const f = old.next();
    if (ours) f.setTrashed(true);
    else f.setName(name + ' (до органайзера).xlsx');
  }
  const xlsx = out.createFile(blob);

  return setTsRecord_(db, monday, { name: name, sheetUrl: ss.getUrl(), xlsxUrl: xlsx.getUrl(), folderUrl: out.getUrl() });
}

function readDb_(file) {
  let db;
  try {
    db = JSON.parse(file.getBlob().getDataAsString('UTF-8') || '{}');
  } catch (err) {
    db = {};
  }
  db.rev = Number(db.rev) || 0;
  db.records = db.records || {};
  return db;
}

function writeDb_(file, db) {
  db.updated = new Date().toISOString();
  file.setContent(JSON.stringify(db));
}

function getDataFile_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FILE_ID');
  if (id) {
    try {
      const f = DriveApp.getFileById(id);
      if (!f.isTrashed()) return f;
    } catch (err) {
      // файл удалён или недоступен — создадим заново
    }
  }
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  const files = folder.getFilesByName(FILE_NAME);
  const file = files.hasNext()
    ? files.next()
    : folder.createFile(FILE_NAME, JSON.stringify({ rev: 0, records: {} }), 'application/json');
  props.setProperty('FILE_ID', file.getId());
  return file;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

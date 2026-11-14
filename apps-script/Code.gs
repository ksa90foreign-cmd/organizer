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
 * Первый запуск: выберите функцию setup и нажмите «Выполнить».
 * В журнале появится ключ доступа — его нужно вставить в приложение.
 */

const FOLDER_NAME = 'Органайзер';
const FILE_NAME = 'organizer-data.json';
const INBOX_PREFIX = 'organizer-inbox';
const TOMBSTONE_TTL_DAYS = 180; // сколько дней помнить об удалённых записях
const KINDS = ['group', 'item', 'account', 'settings', 'task', 'goal', 'tx'];
const ID_RE = /^[A-Za-z0-9_\-.~:@+]{1,200}$/;

/** Однократная настройка: создаёт ключ доступа и файл данных. Повторный запуск ключ не меняет. */
function setup() {
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
  if (req.action !== 'sync') return json_({ ok: false, error: 'bad_action' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'busy' });
  try {
    const file = getDataFile_();
    const db = readDb_(file);
    let changed = false;

    // 1. Входящие файлы
    try {
      if (applyInbox_(db) > 0) changed = true;
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
function applyInbox_(db) {
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
        db.rev += 1;
        rec.rv = db.rev;
        db.records[rec.id] = rec;
        applied++;
      });
      (Array.isArray(box.deletes) ? box.deletes : []).forEach(function (id) {
        const ex = db.records[id];
        if (!ex || ex.del) return;
        db.rev += 1;
        db.records[id] = { id: id, kind: ex.kind, del: true, u: stamp(id), rv: db.rev };
        applied++;
      });
    }
    f.setTrashed(true); // разобран (или повреждён) — убираем в корзину, восстановить можно оттуда
  });
  return applied;
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

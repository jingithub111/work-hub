/**
 * Bound Work Hub web app. The public page never contains the owner key.
 * Run initializeWorkHub in the bound editor before updating the deployment.
 */
const LINKS_SHEET = 'Work Hub Links';
const SETTINGS_SHEET = 'Work Hub Settings';
const VISITS_SHEET = 'Work Hub Visits';
const LINKS_HEADER = ['ID', 'Name', 'URL', 'Note', 'Category', 'Business', 'Color', 'Sort order'];
const OWNER_KEY = 'workhub.owner';
const BOOK_ID = 'workhub.spreadsheet';
const SHARE_PREFIX = 'workhub.share.';
const APP_URL = 'https://jingithub111.github.io/work-hub/';
const BOARD_CACHE = 'workhub.board.v3';
const VISIT_CACHE = 'workhub.visits.v3';
const CACHE_SECONDS = 60;

function initializeWorkHub() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Open the Apps Script editor from the Work Hub spreadsheet.');
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty(OWNER_KEY) || newKey_();
  props.setProperties({[BOOK_ID]: book.getId(), [OWNER_KEY]: key});
  console.log('ลิงก์เจ้าของ (เก็บเป็นส่วนตัว) ' + APP_URL + '#owner=' + key);
}

function doGet() {
  const props = PropertiesService.getScriptProperties();
  const ready = Boolean(props.getProperty(OWNER_KEY) && props.getProperty(BOOK_ID));
  return json_({ok: true, version: 3, ready});
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();
  const owner = props.getProperty(OWNER_KEY);
  const bookId = props.getProperty(BOOK_ID);
  const key = params.key || '';
  if (!owner || !bookId || !/^[a-f0-9]{64}$/.test(key)) return json_({ok: false, error: 'access'});
  const isOwner = key === owner;
  const sharedScope = () => parse_(props.getProperty(SHARE_PREFIX + key), null);
  let scope = isOwner ? null : sharedScope();
  const action = params.action;
  const allowed = isOwner ? ['read', 'visit', 'save', 'shareCreate', 'shareRevoke'] : ['read', 'visit', 'add', 'edit'];
  if ((!isOwner && !scope) || !allowed.includes(action)) return json_({ok: false, error: 'access'});

  const cache = CacheService.getScriptCache();
  let book;
  const openBook = () => {
    if (!book) book = SpreadsheetApp.openById(bookId);
    return book;
  };
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    if (action === 'read') {
      const snapshot = cache.getAll([BOARD_CACHE, VISIT_CACHE]);
      const cachedBoard = parse_(snapshot[BOARD_CACHE], null);
      const cachedVisits = parse_(snapshot[VISIT_CACHE], null);
      if (cachedBoard && cachedVisits) {
        if (scope && !hasScope_(cachedBoard, scope)) return json_({ok: false, error: 'access'});
        return json_(view_(cachedBoard, scope, props, cachedVisits));
      }
    }
    acquired = lock.tryLock(10000);
    if (!acquired) return json_({ok: false, error: 'busy'});
    scope = isOwner ? null : sharedScope();
    if (!isOwner && !scope) return json_({ok: false, error: 'access'});

    if (action === 'read' || action === 'visit') {
      let data = parse_(cache.get(BOARD_CACHE), null);
      if (!data) {
        data = readBoard_(openBook()).data;
        cacheJson_(cache, BOARD_CACHE, data);
      }
      if (scope && !hasScope_(data, scope)) return json_({ok: false, error: 'access'});
      if (action === 'read') return json_(view_(data, scope, props, visitCounts_(cache, openBook)));
      const link = data.links.find(item => item.id === params.id && (!scope || inScope_(item, scope)));
      if (!link || !/^[a-f0-9]{32}$/.test(params.visitor || '')) {
        return json_({ok: false, error: 'access'});
      }
      return json_({ok: true, count: recordVisit_(cache, openBook, link.id, params.visitor)});
    }

    const board = readBoard_(openBook());
    if (scope && !hasScope_(board.data, scope)) return json_({ok: false, error: 'access'});
    if (isOwner && params.action === 'save') {
      if (params.revision !== revision_(board.data)) return json_({ok: false, error: 'conflict'});
      const next = JSON.parse(params.data);
      if (!Array.isArray(next.bizs) || !next.bizs.length || !Array.isArray(next.links) ||
          !next.links.every(link => next.bizs.includes(link.biz))) {
        return json_({ok: false, error: 'invalid'});
      }
      // Write the table before pruning shares. The public read also checks that the business exists.
      cache.remove(BOARD_CACHE);
      writeBoard_(book, next);
      SpreadsheetApp.flush();
      pruneShares_(props, next);
      const updated = readBoard_(book).data;
      cacheJson_(cache, BOARD_CACHE, updated);
      return json_(view_(updated, null, props, visitCounts_(cache, openBook)));
    }

    if (isOwner && (params.action === 'shareCreate' || params.action === 'shareRevoke')) {
      const target = {business: params.business, category: params.category === undefined ? null : params.category};
      if (!hasScope_(board.data, target)) return json_({ok: false, error: 'invalid'});
      const share = shares_(props).find(item => item.business === target.business && item.category === target.category);
      if (params.action === 'shareCreate' && !share) {
        props.setProperty(SHARE_PREFIX + newKey_(), JSON.stringify(target));
      }
      if (params.action === 'shareRevoke' && share) {
        props.deleteProperty(SHARE_PREFIX + share.key);
      }
      return json_({ok: true, shares: shares_(props)});
    }

    if (!isOwner && (action === 'add' || action === 'edit')) {
      if (action === 'edit' && params.revision !== revision_(project_(board.data, scope))) {
        return json_({ok: false, error: 'conflict'});
      }
      const edit = editFields_(JSON.parse(params.data));
      if (!edit) return json_({ok: false, error: 'invalid'});
      if (scope.category !== null && scope.category !== edit.cat) return json_({ok: false, error: 'access'});
      const entry = board.entries.find(item => item.link.id === params.id);

      if (action === 'add') {
        if (!/^[a-f0-9]{32}$/.test(params.id || '')) return json_({ok: false, error: 'invalid'});
        if (entry) {
          if (!inScope_(entry.link, scope)) return json_({ok: false, error: 'access'});
          return json_(view_(board.data, scope, props, visitCounts_(cache, openBook)));
        }
        const link = {
          id: params.id, name: edit.name, url: edit.url, note: edit.note, cat: edit.cat,
          biz: scope.business, color: edit.color
        };
        cache.remove(BOARD_CACHE);
        board.data.links.push(link);
        if (board.hasBusinessColumn) {
          board.sheet.appendRow(linkRow_(link, board.sheet.getLastRow()));
        } else {
          writeBoard_(book, board.data);
        }
      } else {
        if (!entry || !inScope_(entry.link, scope)) return json_({ok: false, error: 'access'});
        cache.remove(BOARD_CACHE);
        Object.assign(entry.link, edit);
        if (entry.row && board.hasBusinessColumn) {
          board.sheet.getRange(entry.row, 1, 1, LINKS_HEADER.length).setValues([linkRow_(entry.link, entry.order)]);
        } else {
          writeBoard_(book, board.data);
        }
      }
      SpreadsheetApp.flush();
      const updated = board.hasBusinessColumn ? board.data : readBoard_(book).data;
      if (action === 'edit') pruneShares_(props, updated);
      cacheJson_(cache, BOARD_CACHE, updated);
      return json_(view_(updated, scope, props, visitCounts_(cache, openBook)));
    }
    return json_({ok: false, error: 'access'});
  } catch (err) {
    console.error(err && err.name ? err.name : 'Work Hub request failed');
    return json_({ok: false, error: 'unavailable'});
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function newKey_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
}

function shares_(props) {
  const values = props.getProperties();
  return Object.keys(values).filter(key => key.startsWith(SHARE_PREFIX))
    .map(key => Object.assign({key: key.slice(SHARE_PREFIX.length)}, JSON.parse(values[key])));
}

function inScope_(link, scope) {
  return link.biz === scope.business && (scope.category === null || (link.cat || '').trim() === scope.category);
}

function hasScope_(data, scope) {
  return data.bizs.includes(scope.business) && (scope.category === null || data.links.some(link => inScope_(link, scope)));
}

function pruneShares_(props, data) {
  shares_(props).filter(share => !hasScope_(data, share))
    .forEach(share => props.deleteProperty(SHARE_PREFIX + share.key));
}

function project_(data, scope) {
  return {
    title: scope.business,
    subtitle: scope.category === null ? 'Work Hub' : 'หมวด ' + (scope.category || 'ยังไม่จัดหมวด'),
    bizs: [scope.business],
    fixedCategory: scope.category,
    links: data.links.filter(link => inScope_(link, scope))
  };
}

function view_(data, scope, props, counts) {
  const visible = scope ? project_(data, scope) : data;
  const views = Object.fromEntries(visible.links.map(link => [link.id, Number(counts[link.id]) || 0]));
  const result = {ok: true, data: visible, revision: revision_(visible), views};
  if (!scope) result.shares = shares_(props);
  return result;
}

function visitRows_(sheet) {
  const count = sheet ? sheet.getLastRow() - 1 : 0;
  return count > 0 ? sheet.getRange(2, 1, count, 2).getValues() : [];
}

function cacheJson_(cache, key, value) {
  try {
    cache.put(key, JSON.stringify(value), CACHE_SECONDS);
  } catch (err) {
    // Large boards can exceed the cache limit. Reads still use the sheet.
  }
}

function visitCounts_(cache, openBook) {
  const stored = parse_(cache.get(VISIT_CACHE), null);
  if (stored) return Object.assign(Object.create(null), stored);
  const counts = Object.create(null);
  for (const [id, digest] of visitRows_(openBook().getSheetByName(VISITS_SHEET))) {
    if (id && digest) counts[id] = (counts[id] || 0) + 1;
  }
  cacheJson_(cache, VISIT_CACHE, counts);
  return counts;
}

function recordVisit_(cache, openBook, id, visitor) {
  const book = openBook();
  const counts = visitCounts_(cache, openBook);
  let sheet = book.getSheetByName(VISITS_SHEET);
  if (!sheet) {
    sheet = book.insertSheet(VISITS_SHEET);
    sheet.appendRow(['Link ID', 'Visit digest']);
    sheet.setFrozenRows(1);
  }
  const digest = revision_([id, visitor]);
  const count = Number(counts[id]) || 0;
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount > 0 && sheet.getRange(2, 2, rowCount, 1).createTextFinder(digest).matchEntireCell(true).findNext()) {
    return count;
  }
  cache.remove(VISIT_CACHE);
  sheet.appendRow([sheetText_(id), digest]);
  SpreadsheetApp.flush();
  counts[id] = count + 1;
  cacheJson_(cache, VISIT_CACHE, counts);
  return count + 1;
}

function revision_(data) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(data))
  );
}

function readBoard_(book) {
  const sheet = book.getSheetByName(LINKS_SHEET);
  const settings = book.getSheetByName(SETTINGS_SHEET);
  const settingRows = settings ? settings.getDataRange().getValues().slice(1) : [];
  const stored = Object.fromEntries(settingRows.filter(row => row[0]).map(row => row.slice(0, 2)));
  const table = sheet ? sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), LINKS_HEADER.length).getValues() : [];
  const hasBusinessColumn = Boolean(table[0] && table[0][5] === 'Business');
  const legacySheet = book.getSheetByName('Work Hub Data');
  const legacy = !table[0] || table[0][0] !== 'ID';
  const backup = legacy ? parse_(legacySheet ? legacySheet.getRange('A1').getValue() : table[0]?.[0], {}) : {};
  const names = parse_(stored.bizs, backup.bizs || []);
  const bizs = Array.isArray(names) ? names.filter(Boolean).map(String) : [];
  if (!bizs.length) bizs.push('ธุรกิจของฉัน');
  const entries = legacy
    ? (backup.links || []).map(link => ({link}))
    : table.slice(1).map((row, index) => ({
      row: index + 2,
      order: row[hasBusinessColumn ? 7 : 6] || index + 1,
      link: {
        id: String(row[0]), name: String(row[1]), url: String(row[2]), note: String(row[3]),
        cat: String(row[4]), biz: hasBusinessColumn ? String(row[5]) : '',
        color: String(row[hasBusinessColumn ? 6 : 5] || '#4C7DF7')
      }
    })).filter(item => item.link.id);
  for (const {link} of entries) {
    link.biz = link.biz || bizs[0];
    if (!bizs.includes(link.biz)) bizs.push(link.biz);
  }
  const data = {
    title: String(stored.title || backup.title || 'Work Hub'),
    subtitle: String(stored.subtitle || backup.subtitle || ''),
    bizs,
    links: entries.map(item => item.link)
  };
  return {data, sheet, hasBusinessColumn, entries};
}

function parse_(raw, fallback) {
  try {
    return JSON.parse(raw) || fallback;
  } catch (err) {
    return fallback;
  }
}

function editFields_(fields) {
  if (!fields || typeof fields !== 'object') return null;
  const {name, url, note, cat, color} = fields;
  if (typeof name !== 'string' || !name.trim() || name.length > 60 ||
      typeof url !== 'string' || url.length > 4096 || !/^https?:\/\/[^/?#\s]+(?:[/?#]\S*)?$/i.test(url) ||
      typeof note !== 'string' || note.length > 80 || typeof cat !== 'string' || cat.length > 30 ||
      !/^#[a-f0-9]{6}$/i.test(color)) return null;
  return {name: name.trim(), url, note, cat, color};
}

function sheetText_(value) {
  const text = String(value ?? '');
  return /^[=']/.test(text) ? "'" + text : text;
}

function linkRow_(link, order) {
  return [link.id, link.name, link.url, link.note, link.cat, link.biz, link.color]
    .map(sheetText_).concat(order);
}

function writeBoard_(book, data) {
  const sheet = book.getSheetByName(LINKS_SHEET) || book.insertSheet(LINKS_SHEET);
  const rows = [LINKS_HEADER].concat(data.links.map((link, index) => linkRow_(link, index + 1)));
  const count = Math.max(sheet.getLastRow(), rows.length);
  if (count > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), count - sheet.getMaxRows());
  while (rows.length < count) rows.push(Array(LINKS_HEADER.length).fill(''));
  sheet.getRange(1, 1, count, LINKS_HEADER.length).setValues(rows);
  sheet.setFrozenRows(1);

  const settings = book.getSheetByName(SETTINGS_SHEET) || book.insertSheet(SETTINGS_SHEET);
  settings.getRange(1, 1, 5, 2).setValues([
    ['Setting', 'Value'],
    ['title', sheetText_(data.title)],
    ['subtitle', sheetText_(data.subtitle)],
    ['bizs', JSON.stringify(data.bizs)],
    ['updatedAt', new Date()]
  ]);
}

function json_(response) {
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

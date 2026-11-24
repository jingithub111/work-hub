/**
 * Bound Work Hub web app. The public page never contains the owner key.
 * Run initializeWorkHub in the bound editor before updating the deployment.
 */
const LINKS_SHEET = 'Work Hub Links';
const SETTINGS_SHEET = 'Work Hub Settings';
const LINKS_HEADER = ['ID', 'Name', 'URL', 'Note', 'Category', 'Business', 'Color', 'Sort order'];
const OWNER_KEY = 'workhub.owner';
const BOOK_ID = 'workhub.spreadsheet';
const SHARE_PREFIX = 'workhub.share.';
const APP_URL = 'https://jingithub111.github.io/work-hub/';

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
  return json_({ok: true, version: 2, ready});
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();
  const owner = props.getProperty(OWNER_KEY);
  const bookId = props.getProperty(BOOK_ID);
  const key = params.key || '';
  if (!owner || !bookId || !/^[a-f0-9]{64}$/.test(key)) return json_({ok: false, error: 'access'});
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    acquired = lock.tryLock(10000);
    if (!acquired) return json_({ok: false, error: 'busy'});
    const isOwner = key === owner;
    const scope = isOwner ? null : parse_(props.getProperty(SHARE_PREFIX + key), null);
    if (!isOwner && !scope) return json_({ok: false, error: 'access'});

    const book = SpreadsheetApp.openById(bookId);
    const board = readBoard_(book);
    if (scope && !hasScope_(board.data, scope)) return json_({ok: false, error: 'access'});

    if (params.action === 'read') {
      return json_(view_(board.data, scope, props));
    }

    if (isOwner && params.action === 'save') {
      if (params.revision !== revision_(board.data)) return json_({ok: false, error: 'conflict'});
      const next = JSON.parse(params.data);
      if (!Array.isArray(next.bizs) || !next.bizs.length || !Array.isArray(next.links) ||
          !next.links.every(link => next.bizs.includes(link.biz))) {
        return json_({ok: false, error: 'invalid'});
      }
      // Write the table before pruning shares. The public read also checks that the business exists.
      writeBoard_(book, next);
      SpreadsheetApp.flush();
      pruneShares_(props, next);
      return json_(view_(readBoard_(book).data, null, props));
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

    if (!isOwner && params.action === 'edit') {
      if (params.revision !== revision_(project_(board.data, scope))) {
        return json_({ok: false, error: 'conflict'});
      }
      const fields = JSON.parse(params.data);
      const entry = board.entries.find(item => item.link.id === params.id && inScope_(item.link, scope));
      if (!entry) return json_({ok: false, error: 'access'});
      const edit = editFields_(fields);
      if (!edit) return json_({ok: false, error: 'invalid'});
      if (scope.category !== null && scope.category !== edit.cat) return json_({ok: false, error: 'access'});

      if (entry.row && board.hasBusinessColumn) {
        const old = entry.link;
        board.sheet.getRange(entry.row, 1, 1, LINKS_HEADER.length)
          .setValues([linkRow_(Object.assign({}, old, edit), entry.order)]);
      } else {
        Object.assign(entry.link, edit);
        writeBoard_(book, board.data);
      }
      SpreadsheetApp.flush();
      const updated = readBoard_(book).data;
      pruneShares_(props, updated);
      return json_(view_(updated, scope, props));
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

function view_(data, scope, props) {
  const visible = scope ? project_(data, scope) : data;
  const result = {ok: true, data: visible, revision: revision_(visible)};
  if (!scope) result.shares = shares_(props);
  return result;
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

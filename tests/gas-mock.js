// Google Apps Script 서비스 모의 환경 — Code.gs를 Node에서 실행해 서버 로직을 검증합니다.
// (실제 Google 서버가 아니므로 권한 승인·메일 발송·트리거는 실기기 UAT에서 다시 확인)
const vm = require('vm'), fs = require('fs'), crypto = require('crypto');

function makeEnv({ owner = 'owner@test.kr', active = owner, books: shared } = {}) {
  const books = shared || {}, mails = [], props = {}, cache = {}, logs = [];
  const mkSheet = (name) => {
    const sh = { name, data: [], maxRows: 1000 };
    const lastRow = () => { for (let i = sh.data.length - 1; i >= 0; i--) if ((sh.data[i] || []).some(c => c !== '' && c != null)) return i + 1; return 0; };
    const lastCol = () => Math.max(0, ...sh.data.map(r => (r || []).length));
    const range = (r, c, nr, nc) => {
      if (r + nr - 1 > sh.maxRows) throw new Error('범위가 시트 크기를 넘음: ' + name);
      const R = {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const row = sh.data[r - 1 + i] || []; const v = row[c - 1 + j]; return v === undefined ? '' : v; })),
        setValues: (vals) => { vals.forEach((row, i) => { const rr = sh.data[r - 1 + i] = sh.data[r - 1 + i] || []; row.forEach((v, j) => rr[c - 1 + j] = v); }); return R; },
        clearContent: () => { for (let i = 0; i < nr; i++) { const rr = sh.data[r - 1 + i]; if (rr) for (let j = 0; j < nc; j++) rr[c - 1 + j] = ''; } return R; },
        setNumberFormat: () => R, setFontWeight: () => R,
      };
      return R;
    };
    return Object.assign(sh, {
      getName: () => name, getLastRow: lastRow, getMaxRows: () => sh.maxRows,
      insertRowsAfter: (_, n) => { sh.maxRows += n; },
      getRange: (r, c, nr = 1, nc = 1) => range(r, c, nr, nc),
      getDataRange: () => range(1, 1, Math.max(1, lastRow()), Math.max(1, lastCol())),
      setFrozenRows: () => {},
    });
  };
  const mkBook = () => { const sheets = {}; return {
    sheets, getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = mkSheet(n)),
    getSheets: () => Object.values(sheets), getSpreadsheetTimeZone: () => 'Asia/Seoul' }; };
  const pad = n => String(n).padStart(2, '0');
  const user = { email: active };
  const g = {
    console,
    SpreadsheetApp: { openById: id => books[id] || (books[id] = mkBook()) },
    Utilities: {
      formatDate: (d, tz, f) => { const k = new Date(d.getTime() + 9 * 3600e3);
        const ymd = `${k.getUTCFullYear()}-${pad(k.getUTCMonth() + 1)}-${pad(k.getUTCDate())}`;
        return f === 'yyyy-MM-dd' ? ymd : `${ymd} ${pad(k.getUTCHours())}:${pad(k.getUTCMinutes())}:${pad(k.getUTCSeconds())}`; },
      getUuid: () => crypto.randomUUID(),
    },
    Session: { getActiveUser: () => ({ getEmail: () => user.email }), getEffectiveUser: () => ({ getEmail: () => owner }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = String(v); }, remove: k => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; }, getKeys: () => Object.keys(props) }) },
    MailApp: { sendEmail: (to, subject, body) => mails.push({ to, subject, body }) },
    Logger: { log: m => logs.push(m) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => { const t = { timeBased: () => t, everyHours: () => t, everyDays: () => t, atHour: () => t, inTimezone: () => t, create: () => t }; return t; } },
    HtmlService: { createHtmlOutputFromFile: () => { const o = { setTitle: () => o, addMetaTag: () => o, setXFrameOptionsMode: () => o }; return o; }, XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' } },
    ContentService: { MimeType: { JSON: 'application/json' }, createTextOutput: t => { const o = { text: t, mime: null, setMimeType: m => { o.mime = m; return o; }, getContent: () => t }; return o; } },
  };
  const ctx = vm.createContext(g);
  return { ctx, books, mails, props, cache, logs, user,
    load(file) { vm.runInContext(fs.readFileSync(file, 'utf8'), ctx, { filename: file }); },
    run(code) { return vm.runInContext(code, ctx); },
    // 시트를 열 이름 기준 객체 배열로 읽기
    table(name) { const b = Object.values(books)[0]; const sh = b.sheets[name]; if (!sh) return null; const [h, ...rows] = sh.data; return rows.filter(r => r && r.some(c => c !== '')).map(r => Object.fromEntries(h.map((k, i) => [k, r[i] === undefined ? '' : r[i]]))); },
  };
}
module.exports = { makeEnv };

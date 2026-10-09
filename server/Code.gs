/**
 * 순장사관학교 Operations OS v4.1.1 — Code.gs (Apps Script 서버)
 * v4.1: 이메일 인증코드 로그인 · 명단 관리(교육생·강사·운영진) · 조편성 · 일괄 입력 · 보관/복원 · 버전 충돌 방지
 * v4.1.1 (배포 보완): 관리용 함수 외부 호출 차단 · 편집기에서 인수 없이 실행 가능 · 시트 연결 1회 재사용(속도) ·
 *                     기록 시각 보존 · 1000행 초과 대비 · 일일 요약 1일 1회 제한
 * v4.2 (원본 엑셀 대비 보강): 예산 수입/지출 구분(화면) · 집결시간 설정 · 회비 금액·교육생별 납부 · 휴강일 반영 날짜 계산 ·
 *                     대표기도 담당공동체 · 개강 준비 11개+메모 · 담당장로(Elder) 조회 역할
 *      ※ v4.1.1 시트를 쓰던 경우 upgradeToV42() 한 번 실행 → 새 열·설정·준비 항목이 추가됩니다 (기존 데이터 유지)
 * 대시보드 v4(운영매뉴얼형)의 데이터 모델과 1:1로 맞춘 서버 함수 모음.
 *
 * 최초 1회 실행 순서 (편집기 상단 함수 선택 → ▶ 실행. 모든 함수는 인수 없이 실행됩니다)
 *   1) setupV4()          v4 시트 17개 생성 + 기본값(개강 2026-10-18, 10주 일정, 예산 항목)
 *   2) migrateFromV3()    (선택) 같은 파일에 v3.0 시트가 있으면 명단·강사·과제·예산·설정 이관
 *   3) seedQASample()     (QA 전용) 가상 교육생 25명·주간 Care·과제 이력 입력
 *   4) v4_Access 시트에 실제 운영진 이메일 입력(Gmail 가능) → getSystemHealth() = READY 확인
 *      ※ 이미 v4를 쓰던 시트는 upgradeToV41() 한 번만 실행하면 열 추가·id 부여·강사 목록 생성이 끝납니다
 *   5) 아래 ALERT_EMAIL에 Red 알림 받을 이메일 입력(선택) → setAlertEmail() → installTriggers()
 *   6) 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) — 로그인은 이메일 인증코드로 처리하므로 시트는 공유하지 않습니다
 *
 * 보안: 웹앱은 '모든 사용자'에게 열려 있으므로, 이름이 _로 끝나지 않는 함수는 누구나 브라우저에서 호출할 수 있습니다.
 *       그래서 설치·관리용 함수는 adminOnly_()로 '스크립트 소유자가 편집기에서 실행할 때'만 동작하게 막았습니다.
 */

// ===== 1. 환경 =====
const ENVIRONMENT = 'QA';   // UAT 완료 후에만 'PROD'로 변경하고 새 버전으로 재배포
const ALERT_EMAIL = '';     // Red Care 알림을 추가로 받을 이메일 (비워두면 v4_Access의 팀장·담당교역자에게만 발송)
// Google Drive 「순장사관학교 Operations OS v4」 폴더에 생성한 v4 전용 스프레드시트 (2026-09-28)
const ENV = {
  PROD: { id: '여기에_PROD_스프레드시트_ID', label: 'PRODUCTION' },
  QA:   { id: '여기에_QA_스프레드시트_ID', label: 'QA TEST' }
};
const TZ = 'Asia/Seoul';
const OPEN_DATE_DEFAULT = '2026-10-18';
const APP_TITLE = '2026 순장학교 운영매뉴얼';

// ===== 2. 스키마 (시트명 → 헤더). 모든 셀은 일반 텍스트(@)로 저장 =====
const SCHEMA = {
  v4_Config:     ['key', 'value', 'note'],
  v4_Access:     ['email', 'name', 'role', 'table', 'active', 'team', 'ver', 'updated', 'by'],
  v4_Team:       ['team', 'members'],
  v4_Week:       ['no', 'date', 'topic', 'speaker', 'sp', 'ppt', 'thanks', 'prayer', 'p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'speakerId', 'community'],
  v4_Trainee:    ['name', 'table', 'status', 'updated', 'id', 'community', 'source', 'note', 'archived', 'ver', 'by', 'paid'],
  v4_Speaker:    ['id', 'name', 'title', 'org', 'topics', 'history', 'note', 'archived', 'ver', 'updated', 'by'],
  v4_WeeklyCare: ['ts', 'week', 'name', 'table', 'att', 'part', 'qt', 'sig', 'cat', 'act', 'note', 'by'],
  v4_Case:       ['name', 'stage', 'esc', 'closedW', 'ts', 'by'],
  v4_Assignment: ['ts', 'name', 'kind', 'week', 'delta', 'by'],
  v4_Official:   ['name', 'result', 'ts', 'by'],
  v4_Task:       ['key', 'done', 'ts', 'by'],
  v4_Milestone:  ['id', 'status', 'ts', 'by'],
  v4_Prep:       ['idx', 'item', 'done', 'memo'],
  v4_Budget:     ['item', 'plan', 'spend', 'status'],
  v4_DDay:       ['week', 'idx', 'done', 'issue', 'ts', 'by'],
  v4_Debrief:    ['week', 'keep', 'problem', 'care', 'next', 'ts', 'by'],
  Audit_Log:     ['ts', 'email', 'role', 'action', 'detail']
};

// ===== 3. 기준값 (ERP LITE v3.0 원문) =====
const TOPICS = ['오리엔테이션·개강예배 및 온누리교회 목회철학', '역동적인 순장', '순예배의 도전과 균형', '순예배의 인도와 실제',
  '순예배 인도를 위한 키워드', '선교하는 순', '순원 돌봄의 키워드', '성령충만한 순', '순장의 사명(수료예배)', '아웃팅'];
const PREPLAN = ['순장학교 프로그램 기획 (일정·강사·기도)', '예산 신청', '순장학교 모집 공지', '운영스텝 모집·단톡방 개설', '교재 제본 (인쇄소 확정)',
  '아이스쿨 순장학교 등록 (계정정보는 별도 보관)', '비품 확인 및 세탁', '강사 선정 및 일정 확인', '장로님 기도 일정 확인',
  '회비 결정 (금액·입금 안내)', '과제 확인 (수료 기준·체크표)'];
const BUDGET = ['교회예산', '회비', '교재비', '강사사례비', '수료선물비', '일회용품 및 기타', '아웃팅비용 지원', '예비비', '식사비', '간식비'];
const TEAMS = ['교육팀', '학사팀', '섬김팀', 'AV', '찬양'];
const TABLES = ['T1', 'T2', 'T3', 'T4', 'T5'];
const E = {
  role: ['TeamLeader', 'Secretary', 'TableLeader', 'Pastor', 'Elder'],
  att: ['Present', 'Late', 'Absent', 'Excused'], sig: ['Green', 'Yellow', 'Red'], qt: ['정상', '부분', '미완료', '-'],
  cat: ['', '출석', '관계', '가정', '신앙', '건강', '기타'], sp: ['미섭외', '섭외중', '확정', '변경'], ppt: ['미수령', '초안', 'Final'],
  thanks: ['미완료', '진행중', '완료'], ms: ['Not Started', 'In Progress', 'Done', 'Hold'], bst: ['미집행', '진행중', '집행완료'],
  status: ['Registered', 'Active', 'Applied', 'Hold', 'Withdraw', 'Completed'], stage: ['신규', '연락·면담 중', '종결'],
  official: ['미정', '수료', '보류'], kind: ['sil', 'gong', 'out', 'dok']
};
// 주간 사이클 (대시보드 CYCLE과 동일 id) — 권한·일일 요약에 사용
const CYCLE = [
  ['d7a', 'D-7', -7, -4, '주간 운영회의: 강사·목표·명단·Risk Top 3 확인', '팀장'], ['d7b', 'D-7', -7, -4, '강사 D-7 Reminder 발송', '교육팀'],
  ['d7c', 'D-7', -7, -4, '명단·결석자·Care 현황 확인', '학사팀'], ['d3a', 'D-3', -3, -2, '강사 PPT 수령', '교육팀'],
  ['d3b', 'D-3', -3, -2, '교재·출석부·과제체크·명찰 준비', '학사팀'], ['d3c', 'D-3', -3, -2, 'AV·찬양 자료·마이크 점검', 'AV'],
  ['d3d', 'D-3', -3, -2, '다과 수량 확정', '섬김팀'], ['d3e', 'D-3', -3, -2, '기능팀 Final 점검 공지', '총무'],
  ['d1a', 'D-1', -1, -1, '교육생 주차 안내 발송', '총무'], ['d1b', 'D-1', -1, -1, 'Staff D-1 안내', '총무'],
  ['d0a', 'D-Day', 0, 0, 'Run-of-Show 실행 · Debrief 기록', '전체'], ['p1a', 'D+1', 1, 1, '출석·과제·Care·예산 Update', '총무'],
  ['p1b', 'D+1', 1, 1, '강사 감사 인사', '교육팀'], ['p1c', 'D+1', 1, 1, '교육생 감사·적용과제 안내', '총무'],
  ['p2a', 'D+2', 2, 3, 'Yellow/Red·결석자 개별 연락', '테이블리더'], ['p2b', 'D+2', 2, 3, 'Red Care 교역자 연결 확인', '팀장']
].map(r => ({ id: r[0], st: r[1], off: r[2], end: r[3], t: r[4], o: r[5] }));
const MS = [
  ['m1', '기획 M-2', 'O', -61, -31, '기수 목표·기간·예산·조직 확정'], ['m2', '준비 M-1', 'O', -30, -14, '10주 주제·강사 섭외·일정 확정'],
  ['m3', '모집 W-4', 'O', -28, -14, '모집공지·추천·신청 접수'], ['m4', '등록 W-2', 'O', -14, -8, '참가자 90% 확정·조편성 초안'],
  ['m5', '개강준비 D-7', 'O', -7, -4, '강사·교재·AV·다과·Staff 준비율 점검'], ['m6', '개강준비 D-3', 'O', -3, -2, 'PPT·명찰·교재·안내판 Final'],
  ['m7', '개강준비 D-1', 'O', -1, -1, '교육생 Welcome / Staff 안내'], ['m8', '현장 D-Day', 'O', 0, 0, '개강 현장 운영'],
  ['m9', '정규운영', 'O', 0, 63, 'D-7~D+2 표준 사이클'], ['m10', '수료 D-21', 'G', -21, -1, '수료 예상자·미충족 요건 사전관리'],
  ['m11', '수료 D-Day', 'G', 0, 0, '수료·파송·공동체 Handover'], ['m12', '사후 D+30', 'G', 30, 30, '공동체 배치 여부 확인'],
  ['m13', '사후 D+60~90', 'G', 60, 90, '현장 정착·코칭 Follow-up']
].map(r => ({ id: r[0], name: r[1], b: r[2], s: r[3], e: r[4], t: r[5] }));

// 역할별 허용 작업 (서버 재검증 — 화면 숨김과 별개로 강제)
const OPS = {
  task: ['TeamLeader', 'Secretary', 'TableLeader'], prep: ['TeamLeader', 'Secretary'], recalc: ['TeamLeader', 'Secretary'],
  wk: ['TeamLeader', 'Secretary'], dd: ['TeamLeader', 'Secretary'], ddIssue: ['TeamLeader', 'Secretary'], db: ['TeamLeader', 'Secretary'],
  care: ['TeamLeader', 'TableLeader'], caseStage: ['TeamLeader', 'TableLeader', 'Pastor'], escalate: ['TeamLeader'],
  adj: ['TeamLeader', 'Secretary'],
  tSave: ['TeamLeader', 'Secretary', 'TableLeader'], tArchive: ['TeamLeader', 'Secretary'], tBulk: ['TeamLeader', 'Secretary'],
  spSave: ['TeamLeader', 'Secretary', 'Pastor'], spArchive: ['TeamLeader', 'Secretary', 'Pastor'], wkSpeaker: ['TeamLeader', 'Secretary', 'Pastor'],
  stSave: ['TeamLeader'],
  ms: ['TeamLeader', 'Secretary'], prepPlan: ['TeamLeader', 'Secretary'], bud: ['TeamLeader', 'Secretary'],
  official: ['TeamLeader', 'Pastor'], team: ['TeamLeader'], cfg: ['TeamLeader'], crit: ['TeamLeader']
};

// ===== 4. 공통 유틸 =====
// 한 번의 실행(요청) 안에서는 스프레드시트·시트 핸들을 재사용 — openById를 요청당 1회로 줄여 응답 속도 개선
let SS_ = null; const SH_ = {};
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(ENV[ENVIRONMENT].id)); }
function sheet_(name) { return SH_[name] || (SH_[name] = ss_().getSheetByName(name)); }
/** 설치·관리용 함수 보호: 스크립트 소유자가 편집기(또는 소유자 본인 브라우저)에서 실행할 때만 허용 */
function adminOnly_() {
  let a = '', e = '';
  try { a = (Session.getActiveUser().getEmail() || '').toLowerCase(); e = (Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (x) { }
  need_(a && a === e, '권한 없음: 이 함수는 스크립트 소유자가 Apps Script 편집기에서만 실행할 수 있습니다.');
}
/** 시트 행이 모자라면 미리 늘림 (getRange가 시트 크기를 넘으면 오류) */
function ensureRows_(sh, lastRow) { const m = sh.getMaxRows(); if (lastRow > m) sh.insertRowsAfter(m, Math.max(200, lastRow - m)); }
function ensureSheet_(name) {
  const ss = ss_(); let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  SH_[name] = sh;
  const h = SCHEMA[name];
  sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function read_(name) {
  const sh = sheet_(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name + ' — setupV4()를 먼저 실행하세요.');
  const h = SCHEMA[name], v = sh.getDataRange().getValues();
  return v.slice(1).filter(r => r.some(c => c !== '' && c !== null))
    .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? ymd_(r[i]) : (r[i] === null ? '' : r[i])); return o; });
}
function write_(name, rows) {
  const sh = sheet_(name) || ensureSheet_(name), h = SCHEMA[name], last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, h.length).clearContent();
  if (rows.length) {
    ensureRows_(sh, rows.length + 1);
    sh.getRange(2, 1, rows.length, h.length).setNumberFormat('@')
      .setValues(rows.map(o => h.map(k => (o[k] === undefined || o[k] === null) ? '' : String(o[k]))));
  }
}
/** 한 행 추가. appendRow 대신 텍스트 서식(@)으로 직접 기록 → '2026-10-18 15:30:00' 같은 시각이 날짜로 바뀌어 시간이 잘리는 문제 방지.
 *  호출부는 모두 withLock_ 안에서 실행되어 행 번호 경쟁이 없음 */
function append_(name, obj) {
  const sh = sheet_(name) || ensureSheet_(name), h = SCHEMA[name], r = sh.getLastRow() + 1;
  ensureRows_(sh, r);
  sh.getRange(r, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => (obj[k] === undefined || obj[k] === null) ? '' : String(obj[k]))]);
}
const isActive_ = a => bool_(a.active !== '' && a.active !== undefined ? a.active : true);
function upsert_(name, match, obj) {
  const rows = read_(name), i = rows.findIndex(match);
  if (i >= 0) rows[i] = Object.assign(rows[i], obj); else rows.push(obj);
  write_(name, rows);
}
const num_ = v => { const n = Number(v); return isNaN(n) ? 0 : n; };
const bool_ = v => v === true || String(v).toLowerCase() === 'true';
const ymd_ = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
const now_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const today_ = () => ymd_(new Date());
function addDays_(s, n) { const p = s.split('-').map(Number); return ymd_(new Date(Date.UTC(p[0], p[1] - 1, p[2] + n, 3))); }
function diffDays_(a, b) { const f = s => { const p = s.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / 864e5; }; return f(a) - f(b); }
function need_(cond, msg) { if (!cond) throw new Error(msg); }
function oneOf_(v, list, label) { need_(list.indexOf(v) >= 0, label + ' 값이 올바르지 않습니다: ' + v); return v; }
function clip_(s, n) { return String(s || '').replace(/[\r\n]+/g, ' ').slice(0, n); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(15000); try { return fn(); } finally { l.releaseLock(); } }

// ===== 5. 웹앱 진입점 =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// ===== 6. 세션·권한 (이메일 인증코드 로그인) =====
const TOKEN_DAYS = 14, CODE_MIN = 10;
function accessRow_(email) {
  email = String(email || '').trim().toLowerCase();
  return email ? read_('v4_Access').find(r => String(r.email).toLowerCase() === email && isActive_(r)) : null;
}
function ctxOf_(row, email, via) {
  if (!row) return { active: false, loginRequired: true, email: email || '', env: ENVIRONMENT, envLabel: ENV[ENVIRONMENT].label };
  return { active: true, email: String(row.email).toLowerCase(), name: row.name, role: row.role, table: row.role === 'TableLeader' ? row.table : 'ALL',
    env: ENVIRONMENT, envLabel: ENV[ENVIRONMENT].label, via: via };
}
function tokenEmail_(tok) {
  if (!tok || !/^[A-Za-z0-9-]{20,80}$/.test(String(tok))) return '';
  const raw = PropertiesService.getScriptProperties().getProperty('tok_' + tok);
  if (!raw) return '';
  const t = JSON.parse(raw);
  if (t.exp < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('tok_' + tok); return ''; }
  return t.email;
}
/** tok: 로그인 토큰. 토큰이 없으면 Google 로그인 이메일(교회 Workspace 도메인일 때만 반환됨)로 판정 */
function getSessionContext(tok) {
  const te = tokenEmail_(tok);
  if (te) return ctxOf_(accessRow_(te), te, 'code');
  let ge = ''; try { ge = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { }
  if (ge && accessRow_(ge)) return ctxOf_(accessRow_(ge), ge, 'google');
  return ctxOf_(null, ge);
}
function requireCtx_(tok) {
  const c = getSessionContext(tok);
  need_(c.active, '로그인이 필요합니다. (로그인이 만료되었거나 등록되지 않은 계정)');
  return c;
}
/** 1단계: 인증코드 요청. 등록 여부와 관계없이 같은 안내를 돌려줘 계정 존재를 노출하지 않음 */
function requestLoginCode(email) {
  email = String(email || '').trim().toLowerCase();
  need_(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email), '이메일 형식을 확인해 주세요.');
  const cache = CacheService.getScriptCache(), rl = 'rl_' + email, n = num_(cache.get(rl));
  need_(n < 5, '요청이 많습니다. 30분 뒤 다시 시도해 주세요.');
  cache.put(rl, String(n + 1), 1800);
  const row = accessRow_(email);
  if (row) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put('otp_' + email, JSON.stringify({ code: code, tries: 0 }), CODE_MIN * 60);
    MailApp.sendEmail(email, '[순장학교] 로그인 인증코드 ' + code,
      row.name + '님, 운영 대시보드 로그인 인증코드는 ' + code + ' 입니다.\n' + CODE_MIN + '분 안에 입력해 주세요.\n\n본인이 요청하지 않았다면 이 메일을 무시하세요.');
  }
  return { ok: true, message: '등록된 운영진 이메일이면 인증코드를 보냈습니다. 메일함(스팸함 포함)을 확인해 주세요.' };
}
/** 2단계: 코드 확인 → 토큰 발급(14일) */
function verifyLoginCode(email, code) {
  email = String(email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), key = 'otp_' + email, raw = cache.get(key);
  need_(raw, '인증코드가 만료되었습니다. 다시 요청해 주세요.');
  const o = JSON.parse(raw);
  if (String(code || '').trim() !== o.code) {
    o.tries++; if (o.tries >= 5) cache.remove(key); else cache.put(key, JSON.stringify(o), CODE_MIN * 60);
    throw new Error(o.tries >= 5 ? '5회 틀렸습니다. 인증코드를 다시 요청해 주세요.' : '인증코드가 맞지 않습니다. (' + o.tries + '/5)');
  }
  cache.remove(key);
  const row = accessRow_(email); need_(row, '사용 중지된 계정입니다.');
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(k => k.indexOf('tok_') === 0).forEach(k => { try { if (JSON.parse(props.getProperty(k)).exp < Date.now()) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } });
  const tok = Utilities.getUuid() + '-' + Utilities.getUuid().slice(0, 8);
  props.setProperty('tok_' + tok, JSON.stringify({ email: email, exp: Date.now() + TOKEN_DAYS * 864e5 }));
  const ctx = ctxOf_(row, email, 'code');
  withLock_(() => audit_(ctx, 'login', '이메일 인증 로그인'));
  return { token: tok, ctx: ctx, data: getState_(ctx) };
}
function logout(tok) { if (tok && /^[A-Za-z0-9-]{20,80}$/.test(String(tok))) PropertiesService.getScriptProperties().deleteProperty('tok_' + tok); return true; }
/** 운영진 1명의 모든 로그인 끊기 — 보통은 웹앱 「명단 관리 › 운영진」에서 '사용'을 끄면 자동 처리됩니다.
 *  편집기에서 직접 하려면 아래 REVOKE_EMAIL에 이메일을 넣고 revokeTokens() 실행 */
const REVOKE_EMAIL = '';
function revokeTokens(email) { adminOnly_(); const t = email || REVOKE_EMAIL; need_(t, 'REVOKE_EMAIL에 이메일을 입력한 뒤 실행하세요.'); return revokeTokens_(t); }
function revokeTokens_(email) {
  const props = PropertiesService.getScriptProperties(); let n = 0;
  props.getKeys().filter(k => k.indexOf('tok_') === 0).forEach(k => { try { if (JSON.parse(props.getProperty(k)).email === String(email).toLowerCase()) { props.deleteProperty(k); n++; } } catch (e) { } });
  return n + '개 로그인 해제';
}

// ===== 7. 조회 API =====
function bootstrap(tok) { const ctx = getSessionContext(tok); return { ctx: ctx, data: ctx.active ? getState_(ctx) : null }; }
function refresh(tok) { const ctx = requireCtx_(tok); return { data: getState_(ctx) }; }

function readConfig_() {
  const m = {}; read_('v4_Config').forEach(r => m[r.key] = r.value);
  return {
    name: m.name || '2026 순장학교', open: m.open || OPEN_DATE_DEFAULT, venue: m.venue || '비전홀', time: m.time || '15:30~17:30',
    crit: { sil: num_(m.crit_sil || 2), gong: num_(m.crit_gong || 2), out: num_(m.crit_out || 1), dok: num_(m.crit_dok || 1), abs: num_(m.crit_abs || 2) },
    excAbs: bool_(m.excAbs), gather: m.gather || '14:30', fee: num_(m.fee === undefined || m.fee === '' ? 50000 : m.fee), holidays: m.holidays || ''
  };
}
function weeks_() {
  return read_('v4_Week').map(w => ({
    no: num_(w.no), date: String(w.date), topic: w.topic, speaker: w.speaker, sp: w.sp || '미섭외', ppt: w.ppt || '미수령',
    thanks: w.thanks || '미완료', prayer: w.prayer, community: w.community || '', prep: [1, 2, 3, 4, 5, 6, 7].map(j => num_(w['p' + j])), speakerId: w.speakerId || '', touched: true
  })).sort((a, b) => a.no - b.no);
}
/** 개강일부터 매주 일요일, 휴강일(쉼표 구분)은 건너뛰고 10주 날짜 */
function weekDates_(open, holidays) {
  const hs = String(holidays || '').split(/[\s,]+/).filter(Boolean), out = []; let d = open, k = 0;
  while (out.length < 10 && k++ < 60) { if (hs.indexOf(d) < 0) out.push(d); d = addDays_(d, 7); }
  return out;
}
function elapsed_(weeks) { const t = today_(); return (weeks || weeks_()).filter(w => w.date <= t).length; }

function getState_(ctx) {
  const scope = t => ctx.table === 'ALL' || t === ctx.table;
  const role = ctx.role, mask = role === 'Secretary' || role === 'Elder';
  const trainees = read_('v4_Trainee'), tableOf = {};
  trainees.filter(t => !bool_(t.archived)).forEach(t => tableOf[t.name] = t.table);
  const access = read_('v4_Access');
  const inS = name => scope(tableOf[name]);
  const dd = {};
  read_('v4_DDay').forEach(r => { const w = num_(r.week); dd[w] = dd[w] || { it: {}, db: {} }; dd[w].it[num_(r.idx)] = { done: bool_(r.done), issue: r.issue }; });
  read_('v4_Debrief').forEach(r => { const w = num_(r.week); dd[w] = dd[w] || { it: {}, db: {} }; dd[w].db = { keep: r.keep, problem: r.problem, care: r.care, next: r.next }; });
  const map = (name, kf, vf) => { const o = {}; read_(name).forEach(r => o[r[kf]] = vf(r)); return o; };
  const leaders = {};
  access.filter(a => a.role === 'TableLeader' && isActive_(a)).forEach(a => leaders[a.table] = a.name);
  const extra = map('v4_Team', 'team', r => r.members);
  const audit = role === 'TeamLeader' ? read_('Audit_Log').slice(-100).reverse().map(a => ({ ts: a.ts, user: a.email, role: a.role, a: a.action, d: a.detail })) : [];
  return {
    cfg: readConfig_(),
    weeks: weeks_(),
    tasks: map('v4_Task', 'key', r => bool_(r.done)),
    ms: map('v4_Milestone', 'id', r => r.status),
    prep: map('v4_Prep', 'idx', r => bool_(r.done)),
    prepMemo: map('v4_Prep', 'idx', r => r.memo || ''),
    budget: role === 'TableLeader' ? [] : read_('v4_Budget').map(b => ({ b: b.item, plan: num_(b.plan), spend: num_(b.spend), st: b.status || '미집행' })),
    dd: dd,
    roster: trainees.filter(t => scope(t.table) && (!bool_(t.archived) || role === 'TeamLeader' || role === 'Secretary')).map(t => ({
      id: t.id, name: t.name, table: t.table, status: t.status, community: t.community, source: t.source, note: t.note, paid: bool_(t.paid),
      archived: bool_(t.archived), ver: num_(t.ver) || 1, updated: t.updated, by: t.by })),
    speakers: read_('v4_Speaker').map(s => ({ id: s.id, name: s.name, title: s.title, org: s.org, topics: s.topics, history: s.history, note: s.note,
      archived: bool_(s.archived), ver: num_(s.ver) || 1, updated: s.updated })),
    staff: access.map(a => ({ email: role === 'TeamLeader' ? a.email : '', key: role === 'TeamLeader' ? a.email : '', name: a.name, role: a.role, table: a.table,
      team: a.team, active: isActive_(a), ver: num_(a.ver) || 1 })),
    logs: read_('v4_WeeklyCare').filter(r => scope(r.table)).map(r => ({
      w: num_(r.week), name: r.name, table: r.table, att: r.att, part: r.part || '-', qt: r.qt || '-', sig: r.sig,
      cat: mask ? '' : r.cat, act: mask ? '' : r.act, by: r.by
    })),
    cases: (() => { const o = {}; read_('v4_Case').filter(c => inS(c.name)).forEach(c => o[c.name] = { stage: c.stage, esc: bool_(c.esc), closedW: num_(c.closedW) }); return o; })(),
    adj: {},
    alog: read_('v4_Assignment').filter(a => inS(a.name)).map(a => ({ name: a.name, k: a.kind, w: num_(a.week), d: num_(a.delta) })),
    official: (() => { const o = {}; read_('v4_Official').filter(x => inS(x.name)).forEach(x => o[x.name] = x.result); return o; })(),
    teams: extra,
    leaders: leaders,
    audit: audit,
    serverToday: today_(),
    logsVersion: now_()
  };
}

// ===== 8. 변경 API (단일 진입점 + 서버 권한 검증 + Audit) =====
function apply(op, p) {
  const ctx = requireCtx_(p && p._tok);
  need_(OPS[op], '알 수 없는 작업입니다: ' + op);
  need_(OPS[op].indexOf(ctx.role) >= 0, ctx.role + ' 역할은 이 작업(' + op + ')을 할 수 없습니다.');
  p = p || {};
  // 기준 주차: PROD는 서버 날짜만 신뢰. QA는 화면 기준일 시뮬레이션(p._asOf)을 허용
  p._asOf = ENVIRONMENT === 'QA' && p._asOf !== undefined ? Math.max(0, Math.min(10, num_(p._asOf))) : elapsed_();
  withLock_(() => { const detail = HANDLERS[op](p, ctx); audit_(ctx, op, detail); });
  return { data: getState_(ctx) };
}
function audit_(ctx, action, detail) { append_('Audit_Log', { ts: now_(), email: ctx.email, role: ctx.role, action: action, detail: clip_(detail, 200) }); }

function traineeOf_(name) { const t = read_('v4_Trainee').find(x => x.name === name && !bool_(x.archived)); need_(t, '명단에 없는 교육생입니다: ' + name); return t; }
function newId_(p) { return p + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function verCheck_(row, ver) { need_(num_(ver) === (num_(row.ver) || 1), 'CONFLICT: 다른 운영진이 먼저 수정했습니다. 최신 내용을 다시 불러왔으니 확인 후 다시 저장해 주세요.'); }
function normTable_(v) { const s = String(v || '').trim().toUpperCase().replace(/조$/, '').replace(/^T/, ''); return /^[1-5]$/.test(s) ? 'T' + s : ''; }
/** 교육생 이름이 바뀌거나 Table을 옮기면 주간 Care·Case·과제·공식판정 기록을 함께 이어 붙임 */
function relinkTrainee_(oldName, newName, newTable) {
  [['v4_WeeklyCare', true], ['v4_Case', false], ['v4_Assignment', false], ['v4_Official', false]].forEach(([sh, hasTable]) => {
    const rows = read_(sh); let ch = false;
    rows.forEach(r => { if (r.name === oldName) { r.name = newName; if (hasTable) r.table = newTable; ch = true; } });
    if (ch) write_(sh, rows);
  });
}
function scopeCheck_(ctx, table) { need_(ctx.table === 'ALL' || ctx.table === table, '자기 Table(' + ctx.table + ')만 처리할 수 있습니다.'); }
function latestSig_(name, upto) {
  const l = read_('v4_WeeklyCare').filter(r => r.name === name && num_(r.week) <= upto).sort((a, b) => num_(a.week) - num_(b.week));
  return l.length ? l[l.length - 1].sig : 'None';
}
function setWeekField_(no, field, value) {
  const rows = read_('v4_Week'), w = rows.find(r => num_(r.no) === no);
  need_(w, '주차가 없습니다: ' + no); w[field] = value; write_('v4_Week', rows);
}

const HANDLERS = {
  task(p, ctx) {
    const [no, id] = String(p.k).split('|'); const t = CYCLE.find(c => c.id === id);
    need_(t && num_(no) >= 1 && num_(no) <= 10, '사이클 항목이 올바르지 않습니다.');
    if (ctx.role === 'TableLeader') need_(t.o === '테이블리더', '테이블리더 담당 항목만 체크할 수 있습니다.');
    upsert_('v4_Task', r => r.key === p.k, { key: p.k, done: !!p.v, ts: now_(), by: ctx.email });
    return p.k + ' → ' + (p.v ? '완료' : '해제');
  },
  prep(p) { need_(p.j >= 0 && p.j < 7 && [0, 1, 2, 3].indexOf(num_(p.v)) >= 0, '준비물 값 오류'); setWeekField_(num_(p.i) + 1, 'p' + (num_(p.j) + 1), num_(p.v)); return (num_(p.i) + 1) + '주차 준비물' + (num_(p.j) + 1) + ' → ' + p.v; },
  recalc() { const c = readConfig_(), ds = weekDates_(c.open, c.holidays), rows = read_('v4_Week'); rows.forEach(r => r.date = ds[num_(r.no) - 1]); write_('v4_Week', rows); return '개강일 ' + c.open + ' 기준 재계산' + (c.holidays ? ' (휴강 ' + c.holidays + ' 제외)' : ''); },
  wk(p) {
    const f = oneOf_(p.f, ['date', 'speaker', 'sp', 'ppt', 'thanks', 'prayer', 'community'], '항목'); let v = clip_(p.v, f === 'community' ? 20 : 40);
    if (f === 'date') need_(/^\d{4}-\d{2}-\d{2}$/.test(v), '날짜 형식은 YYYY-MM-DD입니다.');
    if (f === 'sp') oneOf_(v, E.sp, '섭외상태'); if (f === 'ppt') oneOf_(v, E.ppt, 'PPT'); if (f === 'thanks') oneOf_(v, E.thanks, '감사');
    setWeekField_(num_(p.i) + 1, f, v); return (num_(p.i) + 1) + '주차 ' + f + ' → ' + v;
  },
  dd(p, ctx) { upsert_('v4_DDay', r => num_(r.week) === num_(p.no) && num_(r.idx) === num_(p.i), { week: num_(p.no), idx: num_(p.i), done: !!p.done, ts: now_(), by: ctx.email }); return p.no + '주차 D-Day #' + (num_(p.i) + 1) + ' → ' + (p.done ? '완료' : '해제'); },
  ddIssue(p, ctx) { upsert_('v4_DDay', r => num_(r.week) === num_(p.no) && num_(r.idx) === num_(p.i), { week: num_(p.no), idx: num_(p.i), issue: clip_(p.issue, 120), ts: now_(), by: ctx.email }); return p.no + '주차 이슈 기록'; },
  db(p, ctx) { oneOf_(p.k, ['keep', 'problem', 'care', 'next'], 'Debrief'); const o = { week: num_(p.no), ts: now_(), by: ctx.email }; o[p.k] = clip_(p.v, 500); upsert_('v4_Debrief', r => num_(r.week) === num_(p.no), o); return p.no + '주차 Debrief ' + p.k; },
  care(p, ctx) {
    const r = p.rec || {}, t = traineeOf_(r.name); scopeCheck_(ctx, t.table);
    const weeks = weeks_(), w = num_(r.w);
    need_(w >= 1 && w <= 10, '주차 오류');
    need_(w <= p._asOf || w <= elapsed_(weeks), w + '주차 교육일 이전에는 입력할 수 없습니다.');
    oneOf_(r.att, E.att, '출석'); oneOf_(r.sig, E.sig, 'Care Signal'); oneOf_(r.qt, E.qt, 'QT/과제'); oneOf_(r.cat || '', E.cat, 'Category');
    const rec = { ts: now_(), week: w, name: t.name, table: t.table, att: r.att, part: r.att === 'Absent' ? '-' : String(num_(r.part) || 3), qt: r.qt,
      sig: r.sig, cat: r.cat || '', act: clip_(r.act, 60), note: clip_(r.note, 80), by: ctx.name };
    upsert_('v4_WeeklyCare', x => x.name === t.name && num_(x.week) === w, rec);
    if (r.sig !== 'Green') {
      const c = read_('v4_Case').find(x => x.name === t.name);
      if (c && c.stage === '종결' && w > num_(c.closedW)) upsert_('v4_Case', x => x.name === t.name, { stage: '신규', esc: false, ts: now_(), by: ctx.email });
    }
    if (r.sig === 'Red') { try { notifyRed_(t.table, w); } catch (e) { /* 트리거가 재시도 */ } }
    return w + '주차 ' + t.table + ' ' + t.name + ' ' + r.att + '/' + r.sig;
  },
  caseStage(p, ctx) {
    const t = traineeOf_(p.name); scopeCheck_(ctx, t.table); oneOf_(p.stage, E.stage, '단계');
    const sig = latestSig_(p.name, p._asOf);
    if (ctx.role === 'TableLeader') need_(p.stage === '연락·면담 중' && sig === 'Yellow', '테이블리더는 Yellow를 연락 단계로만 옮길 수 있습니다.');
    if (ctx.role === 'Pastor') need_(p.stage === '종결' && sig === 'Red', '담당교역자는 Red Review 완료(종결)만 처리합니다.');
    const cur = read_('v4_Case').find(x => x.name === p.name) || { esc: false };
    upsert_('v4_Case', x => x.name === p.name, { name: p.name, stage: p.stage, esc: p.stage === '종결' ? false : bool_(cur.esc),
      closedW: p.stage === '종결' ? p._asOf : (cur.closedW || 0), ts: now_(), by: ctx.email });
    return p.name + ' → ' + p.stage;
  },
  escalate(p, ctx) {
    const t = traineeOf_(p.name); need_(latestSig_(p.name, p._asOf) === 'Red', 'Red 신호만 교역자 연결 대상입니다.');
    upsert_('v4_Case', x => x.name === p.name, { name: p.name, stage: '연락·면담 중', esc: true, closedW: 0, ts: now_(), by: ctx.email });
    try { notifyEscalation_(t.table); } catch (e) { }
    return p.name + ' 교역자 연결';
  },
  adj(p, ctx) {
    const t = traineeOf_(p.name); oneOf_(p.k, E.kind, '과제 종류'); const d = num_(p.delta); need_(d === 1 || d === -1, '증감은 ±1입니다.');
    const cur = read_('v4_Assignment').filter(a => a.name === p.name && a.kind === p.k).reduce((s, a) => s + num_(a.delta), 0);
    need_(cur + d >= 0, '실적은 0보다 작을 수 없습니다.');
    append_('v4_Assignment', { ts: now_(), name: t.name, kind: p.k, week: Math.max(0, num_(p.w)), delta: d, by: ctx.email });
    return t.name + ' ' + p.k + ' ' + (d > 0 ? '+1' : '−1');
  },
  tSave(p, ctx) {
    const r = p.rec || {}, rows = read_('v4_Trainee');
    const name = clip_(r.name, 20).trim(), table = normTable_(r.table);
    if (ctx.role === 'TableLeader') {
      const row = rows.find(x => x.id === r.id && !bool_(x.archived)); need_(row, '교육생을 찾을 수 없습니다.');
      scopeCheck_(ctx, row.table); verCheck_(row, p.ver);
      row.note = clip_(r.note, 80); row.ver = (num_(row.ver) || 1) + 1; row.updated = now_(); row.by = ctx.email;
      write_('v4_Trainee', rows); return row.name + ' 비고 수정';
    }
    need_(name, '이름을 입력하세요.'); need_(table, 'Table은 T1~T5 중 하나입니다.'); oneOf_(r.status, E.status, '등록상태');
    need_(!rows.some(x => x.name === name && !bool_(x.archived) && x.id !== r.id), '같은 이름의 교육생이 있습니다. 구분 표기(예: 김하늘B)를 붙여 주세요.');
    const fields = { name: name, table: table, status: r.status, community: clip_(r.community, 30), source: clip_(r.source, 20), note: clip_(r.note, 80), paid: r.paid === true || r.paid === 'on', updated: now_(), by: ctx.email };
    if (!r.id) { rows.push(Object.assign({ id: newId_('T'), archived: false, ver: 1 }, fields)); write_('v4_Trainee', rows); return '교육생 등록 ' + name + ' ' + table; }
    const row = rows.find(x => x.id === r.id); need_(row, '교육생을 찾을 수 없습니다.'); verCheck_(row, p.ver);
    const oldName = row.name, oldTable = row.table;
    Object.assign(row, fields, { ver: (num_(row.ver) || 1) + 1 });
    write_('v4_Trainee', rows);
    if (oldName !== name || oldTable !== table) relinkTrainee_(oldName, name, table);
    return '교육생 수정 ' + (oldName !== name ? oldName + '→' : '') + name + (oldTable !== table ? ' ' + oldTable + '→' + table : '');
  },
  tArchive(p, ctx) {
    const rows = read_('v4_Trainee'), row = rows.find(x => x.id === p.id); need_(row, '교육생을 찾을 수 없습니다.'); verCheck_(row, p.ver);
    if (!p.archived) need_(!rows.some(x => x.name === row.name && !bool_(x.archived) && x.id !== row.id), '같은 이름의 교육생이 이미 있어 복원할 수 없습니다.');
    row.archived = !!p.archived; row.ver = (num_(row.ver) || 1) + 1; row.updated = now_(); row.by = ctx.email;
    write_('v4_Trainee', rows); return row.name + (p.archived ? ' 보관' : ' 복원');
  },
  tBulk(p, ctx) {
    const rows = read_('v4_Trainee'), list = (p.rows || []).slice(0, 200); let add = 0, upd = 0, skip = 0;
    const seen = {};
    list.forEach(x => {
      const name = clip_(x.name, 20).trim(), table = normTable_(x.table);
      if (!name || !table || seen[name]) { skip++; return; } seen[name] = 1;
      const row = rows.find(r => r.name === name && !bool_(r.archived));
      if (row) {
        const oldTable = row.table;
        Object.assign(row, { table: table, community: x.community ? clip_(x.community, 30) : row.community, note: x.note ? clip_(x.note, 80) : row.note,
          ver: (num_(row.ver) || 1) + 1, updated: now_(), by: ctx.email });
        if (oldTable !== table) relinkTrainee_(name, name, table); upd++;
      } else {
        rows.push({ id: newId_('T'), name: name, table: table, status: E.status.indexOf(x.status) >= 0 ? x.status : 'Registered', community: clip_(x.community, 30),
          source: clip_(x.source, 20), note: clip_(x.note, 80), paid: false, archived: false, ver: 1, updated: now_(), by: ctx.email }); add++;
      }
    });
    write_('v4_Trainee', rows); return '일괄 입력 신규 ' + add + ' · 수정 ' + upd + ' · 제외 ' + skip;
  },
  spSave(p, ctx) {
    const r = p.rec || {}, rows = read_('v4_Speaker'), name = clip_(r.name, 30).trim();
    need_(name, '강사 이름을 입력하세요.');
    const fields = { name: name, title: clip_(r.title, 20), org: clip_(r.org, 30), topics: clip_(r.topics, 80), history: clip_(r.history, 200), note: clip_(r.note, 120), updated: now_(), by: ctx.email };
    let id = r.id;
    if (!id) { id = newId_('S'); rows.push(Object.assign({ id: id, archived: false, ver: 1 }, fields)); }
    else { const row = rows.find(x => x.id === id); need_(row, '강사를 찾을 수 없습니다.'); verCheck_(row, p.ver); Object.assign(row, fields, { ver: (num_(row.ver) || 1) + 1 }); }
    write_('v4_Speaker', rows);
    const wk = read_('v4_Week'); let ch = false;
    wk.forEach(w => { if (w.speakerId === id && w.speaker !== name) { w.speaker = name; ch = true; } });
    if (Array.isArray(p.weeks)) wk.forEach(w => {
      const on = p.weeks.indexOf(num_(w.no)) >= 0;
      if (on && w.speakerId !== id) { w.speakerId = id; w.speaker = name; if (w.sp === '미섭외') w.sp = '섭외중'; ch = true; }
      if (!on && w.speakerId === id) { w.speakerId = ''; w.speaker = ''; ch = true; }
    });
    if (ch) write_('v4_Week', wk);
    return (r.id ? '강사 수정 ' : '강사 등록 ') + name + (Array.isArray(p.weeks) ? ' (담당 ' + p.weeks.join(',') + '주차)' : '');
  },
  spArchive(p, ctx) {
    const rows = read_('v4_Speaker'), row = rows.find(x => x.id === p.id); need_(row, '강사를 찾을 수 없습니다.'); verCheck_(row, p.ver);
    row.archived = !!p.archived; row.ver = (num_(row.ver) || 1) + 1; row.updated = now_(); row.by = ctx.email;
    write_('v4_Speaker', rows); return row.name + (p.archived ? ' 강사 보관' : ' 강사 복원');
  },
  wkSpeaker(p) {
    const no = num_(p.i) + 1, sid = String(p.speakerId || '');
    const s = sid ? read_('v4_Speaker').find(x => x.id === sid && !bool_(x.archived)) : null;
    need_(!sid || s, '강사 목록에 없는 강사입니다.');
    const rows = read_('v4_Week'), w = rows.find(r => num_(r.no) === no); need_(w, '주차가 없습니다: ' + no);
    w.speakerId = sid; w.speaker = s ? s.name : ''; if (s && w.sp === '미섭외') w.sp = '섭외중';
    write_('v4_Week', rows); return no + '주차 강사 → ' + (s ? s.name : '미정');
  },
  stSave(p, ctx) {
    const r = p.rec || {}, rows = read_('v4_Access'), email = String(r.email || '').trim().toLowerCase();
    need_(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email), '이메일 형식을 확인해 주세요.');
    need_(clip_(r.name, 20).trim(), '이름을 입력하세요.'); oneOf_(r.role, E.role, '역할'); oneOf_(r.team || '', [''].concat(TEAMS), '팀');
    const table = r.role === 'TableLeader' ? normTable_(r.table) : 'ALL'; need_(table, '테이블리더는 담당 Table(T1~T5)이 필요합니다.');
    const orig = String(p.orig || '').toLowerCase();
    need_(!rows.some(x => String(x.email).toLowerCase() === email && String(x.email).toLowerCase() !== orig), '이미 등록된 이메일입니다.');
    const active = r.active !== false;
    const fields = { email: email, name: clip_(r.name, 20).trim(), role: r.role, table: table, team: r.team || '', active: active, updated: now_(), by: ctx.email };
    if (orig) {
      const row = rows.find(x => String(x.email).toLowerCase() === orig); need_(row, '운영진을 찾을 수 없습니다.'); verCheck_(row, p.ver);
      if (orig === ctx.email) need_(active && r.role === 'TeamLeader', '본인 계정의 팀장 권한이나 사용 여부는 바꿀 수 없습니다.');
      Object.assign(row, fields, { ver: (num_(row.ver) || 1) + 1 });
      if (orig !== email || !active) revokeTokens_(orig);
    } else rows.push(Object.assign({ ver: 1 }, fields));
    need_(rows.some(x => x.role === 'TeamLeader' && isActive_(x)), '사용 중인 팀장이 최소 1명 있어야 합니다.');
    write_('v4_Access', rows); return (orig ? '운영진 수정 ' : '운영진 등록 ') + fields.name + ' (' + r.role + (table !== 'ALL' ? ' ' + table : '') + (active ? '' : ', 사용중지') + ')';
  },
  ms(p, ctx) { need_(MS.some(m => m.id === p.id), '마일스톤 오류'); oneOf_(p.st, E.ms, '상태'); upsert_('v4_Milestone', r => r.id === p.id, { id: p.id, status: p.st, ts: now_(), by: ctx.email }); return p.id + ' → ' + p.st; },
  prepPlan(p) { const i = num_(p.i); need_(i >= 0 && i < PREPLAN.length, '항목 오류');
    if (p.memo !== undefined) { upsert_('v4_Prep', r => num_(r.idx) === i, { idx: i, item: PREPLAN[i], memo: clip_(p.memo, 60) }); return PREPLAN[i] + ' 메모'; }
    upsert_('v4_Prep', r => num_(r.idx) === i, { idx: i, item: PREPLAN[i], done: !!p.v }); return PREPLAN[i] + ' → ' + (p.v ? '완료' : '해제'); },
  bud(p) {
    const rows = read_('v4_Budget'), b = rows[num_(p.i)]; need_(b, '예산 항목 오류'); oneOf_(p.f, ['plan', 'spend', 'st'], '예산 필드');
    if (p.f === 'st') b.status = oneOf_(p.v, E.bst, '집행상태'); else b[p.f] = Math.max(0, num_(p.v));
    write_('v4_Budget', rows); return b.item + ' ' + p.f + ' → ' + p.v;
  },
  official(p, ctx) { const t = traineeOf_(p.name); oneOf_(p.v, E.official, '판정'); upsert_('v4_Official', x => x.name === t.name, { name: t.name, result: p.v, ts: now_(), by: ctx.email }); return t.name + ' 공식 판정 ' + p.v; },
  team(p) { oneOf_(p.k, TEAMS, '팀'); upsert_('v4_Team', r => r.team === p.k, { team: p.k, members: clip_(p.v, 60) }); return p.k + ' 배정'; },
  cfg(p) {
    const k = oneOf_(p.k, ['name', 'open', 'venue', 'time', 'excAbs', 'gather', 'fee', 'holidays'], '설정'); let v = clip_(p.v, k === 'holidays' ? 200 : 40);
    if (k === 'open') need_(/^\d{4}-\d{2}-\d{2}$/.test(v), '날짜 형식은 YYYY-MM-DD입니다.');
    if (k === 'excAbs') v = String(v === 'true' || v === true);
    if (k === 'gather') need_(/^\d{2}:\d{2}$/.test(v), '집결시간 형식은 14:30입니다.');
    if (k === 'fee') v = String(Math.max(0, num_(v)));
    if (k === 'holidays') v = String(v).split(/[\s,]+/).filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x)).slice(0, 10).join(', ');
    setConfig_(k, v); if (k === 'open' || k === 'holidays') HANDLERS.recalc();
    return '설정 ' + k + ' → ' + v;
  },
  crit(p) { oneOf_(p.k, ['sil', 'gong', 'out', 'dok', 'abs'], '수료 기준'); const v = Math.max(0, Math.min(10, num_(p.v))); setConfig_('crit_' + p.k, v); return '수료 기준 ' + p.k + ' → ' + v; }
};
function setConfig_(k, v) { upsert_('v4_Config', r => r.key === k, { key: k, value: v }); }

// ===== 9. 초기 구성 =====
/** v4 → v4.1: 새 열 헤더 반영, 교육생 id·버전 부여, 주차 강사 텍스트를 강사 목록으로 전환. 여러 번 실행해도 안전 */
function upgradeToV41() { adminOnly_(); return upgradeToV41_(); }
function upgradeToV41_() {
  Object.keys(SCHEMA).forEach(ensureSheet_);
  const tr = read_('v4_Trainee'); let nT = 0;
  tr.forEach(t => { if (!t.id) { t.id = newId_('T'); nT++; } if (!t.ver) t.ver = 1; if (t.archived === '') t.archived = false; });
  write_('v4_Trainee', tr);
  const acc = read_('v4_Access'); acc.forEach(a => { if (!a.ver) a.ver = 1; }); write_('v4_Access', acc);
  const sp = read_('v4_Speaker'), wk = read_('v4_Week'); let nS = 0;
  wk.forEach(w => {
    if (w.speaker && !w.speakerId) {
      let s = sp.find(x => x.name === w.speaker && !bool_(x.archived));
      if (!s) { s = { id: newId_('S'), name: w.speaker, title: '', org: '', topics: w.topic, history: '', note: '', archived: false, ver: 1, updated: now_(), by: 'upgrade' }; sp.push(s); nS++; }
      w.speakerId = s.id;
    }
  });
  write_('v4_Speaker', sp); write_('v4_Week', wk);
  return 'upgradeToV41 완료: 교육생 id ' + nT + '건 부여, 강사 ' + nS + '명 생성';
}
const CFG_V42 = [{ key: 'gather', value: '14:30', note: '스텝 집결시간 (원본 큐시트: 교육준비 14:30)' },
  { key: 'fee', value: 50000, note: '회비 (원) — 5만원 가정' }, { key: 'holidays', value: '', note: '휴강일 (일요일, 쉼표 구분)' }];
/** v4.1.1 → v4.2: 새 열 헤더·설정 3개·개강 준비 2개 추가. 여러 번 실행해도 안전 (기존 값 유지) */
function upgradeToV42() {
  adminOnly_(); upgradeToV41_();
  const cfg = read_('v4_Config'); let nC = 0;
  CFG_V42.forEach(c => { if (!cfg.some(r => r.key === c.key)) { cfg.push(c); nC++; } }); write_('v4_Config', cfg);
  const pr = read_('v4_Prep'); let nP = 0;
  PREPLAN.forEach((t, i) => { if (!pr.some(r => num_(r.idx) === i)) { pr.push({ idx: i, item: t, done: false, memo: '' }); nP++; } }); write_('v4_Prep', pr);
  return 'upgradeToV42 완료: 설정 ' + nC + '개, 개강 준비 ' + nP + '개 추가';
}
function setupV4(openDate) {
  adminOnly_();
  const open = /^\d{4}-\d{2}-\d{2}$/.test(String(openDate || '')) ? openDate : OPEN_DATE_DEFAULT;
  Object.keys(SCHEMA).forEach(ensureSheet_);
  if (!read_('v4_Config').length) write_('v4_Config', [
    { key: 'name', value: '2026 순장학교', note: '기수명' }, { key: 'open', value: open, note: '1주차 개강일 (일)' },
    { key: 'venue', value: '비전홀', note: '교육장' }, { key: 'time', value: '15:30~17:30', note: '교육시간' },
    { key: 'crit_sil', value: 2, note: '순장실습 필요횟수 (초안)' }, { key: 'crit_gong', value: 2, note: '순장공부 필요횟수 (초안)' },
    { key: 'crit_out', value: 1, note: '아웃리치 필요 (초안)' }, { key: 'crit_dok', value: 1, note: '독후감 필요 (초안)' },
    { key: 'crit_abs', value: 2, note: '최대 결석 허용 (공식 기준 확인 필요)' }, { key: 'excAbs', value: 'false', note: '사유결석 결석 산입 여부' }]
    .concat(CFG_V42));
  if (!read_('v4_Week').length) write_('v4_Week', TOPICS.map((t, i) => ({ no: i + 1, date: addDays_(open, 7 * i), topic: t, speaker: '', sp: '미섭외', ppt: '미수령',
    thanks: '미완료', prayer: i === 0 ? '담당장로' : '', community: '', p1: 0, p2: 0, p3: 0, p4: 0, p5: 0, p6: 0, p7: 0 })));
  if (!read_('v4_Budget').length) write_('v4_Budget', BUDGET.map(b => ({ item: b, plan: 0, spend: 0, status: '미집행' })));
  if (!read_('v4_Prep').length) write_('v4_Prep', PREPLAN.map((t, i) => ({ idx: i, item: t, done: false })));
  if (!read_('v4_Team').length) write_('v4_Team', TEAMS.map(t => ({ team: t, members: '' })));
  if (!read_('v4_Access').length) write_('v4_Access', [
    { email: 'leader@example.com', name: '팀장(샘플)', role: 'TeamLeader', table: 'ALL', active: true },
    { email: 'secretary@example.com', name: '총무(샘플)', role: 'Secretary', table: 'ALL', active: true },
    { email: 'pastor@example.com', name: '담당교역자(샘플)', role: 'Pastor', table: 'ALL', active: true }]
    .concat(TABLES.map(t => ({ email: t.toLowerCase() + '@example.com', name: t + ' 리더', role: 'TableLeader', table: t, active: true }))));
  upgradeToV41_();
  return 'setupV4 완료: 시트 ' + Object.keys(SCHEMA).length + '개, 개강 ' + open + ' · v4_Access의 example.com 이메일을 실제 운영진 이메일로 교체하세요.';
}

/** 같은 스프레드시트에 v3.0 시트(명단·주차운영·강사관리·과제체크·예산·설정·준비)가 있으면 이관. 날짜는 v3 값을 쓰지 않고 개강일 기준으로 재산출 */
function migrateFromV3() {
  adminOnly_();
  const ss = ss_(), log = [];
  const grid = n => { const s = ss.getSheetByName(n); return s ? s.getDataRange().getValues() : null; };
  const headerRow = (g, label) => g.findIndex(r => r.some(c => String(c).trim() === label));
  // 명단: 신청자·조·등록상태
  const m = grid('명단');
  if (m) {
    const hr = headerRow(m, '신청자'), h = m[hr].map(String), ci = k => h.indexOf(k);
    const rows = m.slice(hr + 1).filter(r => String(r[ci('신청자')]).trim()).map(r => {
      const g = String(r[ci('조')]).trim(); const tb = /^T?\d$/.test(g) ? 'T' + g.replace('T', '') : 'T1';
      const st = String(r[ci('등록상태')]).trim(); return { name: String(r[ci('신청자')]).trim(), table: TABLES.indexOf(tb) >= 0 ? tb : 'T1', status: E.status.indexOf(st) >= 0 ? st : 'Registered',
        community: ci('공동체') >= 0 ? String(r[ci('공동체')] || '').trim() : '', updated: now_(), archived: false, ver: 1, by: 'migrate' };
    });
    if (rows.length) { write_('v4_Trainee', rows); log.push('명단 ' + rows.length + '명'); }
  }
  // 주차: 주제·강사·대표기도 (주차운영) + 섭외·PPT·감사 (강사관리)
  const wk = grid('주차운영'), sp = grid('강사관리'), weeks = read_('v4_Week');
  if (wk) { const hr = headerRow(wk, '주제'); wk.slice(hr + 1).forEach(r => { const w = weeks.find(x => num_(x.no) === num_(r[0])); if (w) { if (r[2]) w.topic = String(r[2]); if (r[3]) w.speaker = String(r[3]); if (r[4]) w.prayer = String(r[4]); } }); log.push('주차운영'); }
  if (sp) { const hr = headerRow(sp, '섭외상태'); sp.slice(hr + 1).forEach(r => { const w = weeks.find(x => num_(x.no) === num_(r[0])); if (w) {
    if (r[3]) w.speaker = String(r[3]); if (E.sp.indexOf(String(r[5])) >= 0) w.sp = String(r[5]); if (E.ppt.indexOf(String(r[8])) >= 0) w.ppt = String(r[8]); if (E.thanks.indexOf(String(r[10])) >= 0) w.thanks = String(r[10]); } }); log.push('강사관리'); }
  write_('v4_Week', weeks);
  // 과제체크: 누적 실적을 week 0 이벤트로 이관 (결석은 주간 입력에서 새로 집계)
  const kc = grid('과제체크');
  if (kc) {
    const hr = headerRow(kc, '이름'), ev = [];
    const val = v => { const s = String(v).trim(); if (s === '완료' || s.toUpperCase() === 'O') return 1; return num_(s); };
    kc.slice(hr + 1).forEach(r => { const n = String(r[1]).trim(); if (!n || n === '0') return;
      [['sil', 2], ['gong', 3], ['out', 4], ['dok', 5]].forEach(([k, c]) => { const v = val(r[c]); if (v > 0) ev.push({ ts: now_(), name: n, kind: k, week: 0, delta: v, by: 'migrate' }); }); });
    if (ev.length) { write_('v4_Assignment', read_('v4_Assignment').concat(ev)); log.push('과제 ' + ev.length + '건'); }
  }
  // 예산
  const bd = grid('예산');
  if (bd) { const rows = read_('v4_Budget'); bd.forEach(r => { const b = rows.find(x => x.item === String(r[0]).trim()); if (b) { b.plan = num_(r[1]); b.spend = num_(r[2]); if (E.bst.indexOf(String(r[3])) >= 0) b.status = String(r[3]); } }); write_('v4_Budget', rows); log.push('예산'); }
  // 설정: 수료 기준
  const st = grid('설정');
  if (st) { const map = { '기수명': 'name', '교육장': 'venue', '교육시간': 'time', '순장실습 필요횟수': 'crit_sil', '순장공부 필요횟수': 'crit_gong', '아웃리치 필요': 'crit_out', '독후감 필요': 'crit_dok', '최대 결석 허용': 'crit_abs' };
    st.forEach(r => { const k = map[String(r[0]).trim()]; if (k && r[1] !== '') setConfig_(k, r[1]); }); log.push('설정'); }
  // 준비 체크
  const pr = grid('준비');
  if (pr) { const rows = read_('v4_Prep'); pr.forEach(r => { const i = PREPLAN.findIndex(t => t.indexOf(String(r[1]).trim().slice(0, 6)) === 0); if (i >= 0 && r[2] === true) rows[i].done = true; }); write_('v4_Prep', rows); log.push('준비'); }
  upgradeToV41_();
  return 'migrateFromV3 완료: ' + (log.join(', ') || '이관할 v3 시트 없음') + ' · 연락처·계좌 정보는 이관하지 않습니다.';
}

/** QA 전용: 대시보드 QA 샘플과 동일한 가상 데이터(교육생 25명, 10주 주간 Care, 과제 이력) */
function seedQASample() {
  adminOnly_();
  need_(ENVIRONMENT === 'QA', 'seedQASample은 QA 환경에서만 실행합니다.');
  const S = [['김하늘', 'T1'], ['이서준', 'T1'], ['박지우', 'T1'], ['최민재', 'T1'], ['정다은', 'T1'], ['윤시온', 'T2'], ['한예린', 'T2'], ['오지훈', 'T2'], ['강수아', 'T2'], ['문태호', 'T2'],
    ['임채원', 'T3'], ['조현우', 'T3'], ['백서연', 'T3'], ['신도윤', 'T3'], ['유나영', 'T3'], ['장민준', 'T4'], ['송예은', 'T4'], ['권도현', 'T4'], ['남지민', 'T4'], ['홍서진', 'T4'],
    ['서유진', 'T5'], ['배현준', 'T5'], ['노하린', 'T5'], ['심재윤', 'T5'], ['류지아', 'T5']];
  let seed = 20261018;
  const r = () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
  const care = [], asg = [];
  S.forEach(([name, table]) => {
    for (let w = 1; w <= 10; w++) {
      const x = r(); let att = x < .06 ? 'Absent' : x < .12 ? 'Late' : x < .15 ? 'Excused' : 'Present';
      let part = att === 'Absent' ? '' : String(3 + Math.floor(r() * 3) - (r() < .12 ? 1 : 0));
      const qt = r() < .7 ? '정상' : r() < .8 ? '부분' : '미완료';
      let sig = 'Green', cat = '', act = '';
      if (w === 10) att = 'Present';
      if (name === '오지훈' && [2, 3, 5].indexOf(w) >= 0) att = 'Absent';
      if (name === '남지민' && [3, 4].indexOf(w) >= 0) att = 'Absent';
      if (name === '강수아' && [4, 6, 7].indexOf(w) >= 0) att = 'Absent';
      if (att === 'Absent') { sig = ['오지훈', '남지민', '강수아'].indexOf(name) >= 0 ? 'Yellow' : 'Green'; cat = '출석'; act = '개별 연락 후 상황 확인'; part = ''; }
      if (name === '백서연') { if (w === 3 || w === 4) { sig = 'Red'; cat = '신앙'; act = '담당교역자 연결'; } else if (w === 5) { sig = 'Yellow'; cat = '신앙'; act = '교역자 면담 후 관찰'; } }
      if (name === '한예린' && w === 6) { sig = 'Yellow'; cat = '관계'; act = '개별 연락'; }
      care.push({ ts: now_(), week: w, name: name, table: table, att: att, part: part || '-', qt: att === 'Absent' ? '-' : qt, sig: sig, cat: cat, act: act, note: '', by: '샘플' });
    }
    const pick = (lo, hi) => lo + Math.floor(r() * (hi - lo + 1));
    const two = () => { const a = pick(3, 6), b = pick(6, 9); return r() < .16 ? [a] : [a, b]; };
    const comp = { sil: two(), gong: two(), out: r() < .14 ? [] : [pick(5, 8)], dok: r() < .1 ? [] : [r() < .25 ? 9 : 8] };
    if (name === '김하늘') comp.dok = []; if (name === '문태호') comp.gong = [5];
    Object.keys(comp).forEach(k => comp[k].forEach(w => asg.push({ ts: now_(), name: name, kind: k, week: w, delta: 1, by: '샘플' })));
  });
  write_('v4_Trainee', S.map(([n, t], i) => ({ name: n, table: t, status: 'Active', updated: now_(), archived: false, ver: 1, by: '샘플', paid: i % 5 !== 0 })));
  write_('v4_Speaker', []);
  write_('v4_WeeklyCare', care); write_('v4_Assignment', asg); write_('v4_Case', []); write_('v4_Official', []);
  const t = today_(), weeks = read_('v4_Week');
  weeks.forEach((w, i) => {
    if (i < 4) { w.sp = '확정'; w.speaker = '샘플강사' + (i + 1); } else if (i < 7) w.sp = '섭외중';
    if (w.date < t) { w.ppt = 'Final'; for (let j = 1; j <= 7; j++) w['p' + j] = (j === 7 && i > 0) ? 3 : 2; if (addDays_(w.date, 1) < t) w.thanks = '완료'; }   // 지난 주차는 준비 완료로 가정
  });
  write_('v4_Week', weeks);
  // 지난 사이클 할 일·마일스톤은 수행 완료로 가정 (화면 QA 샘플과 동일하게 '지연'이 과도하게 쌓이지 않도록)
  const tasks = [], lastPast = weeks.filter(w => w.date < t).pop();
  weeks.forEach(w => CYCLE.forEach(c => { if (addDays_(w.date, c.end) < t && !(c.id === 'p1a' && lastPast && w.no === lastPast.no)) tasks.push({ key: w.no + '|' + c.id, done: true, ts: now_(), by: '샘플' }); }));
  write_('v4_Task', tasks);
  const open = weeks[0].date, grad = weeks[8].date;
  write_('v4_Milestone', MS.filter(m => addDays_(m.b === 'O' ? open : grad, m.e) < t).map(m => ({ id: m.id, status: 'Done', ts: now_(), by: '샘플' })));
  upgradeToV41_();
  return 'seedQASample 완료: 교육생 ' + S.length + '명, 주간 Care ' + care.length + '건, 과제 ' + asg.length + '건 (가상 데이터)';
}

// ===== 10. 헬스체크 =====
function getSystemHealth() {
  adminOnly_();
  const ss = ss_(), names = ss.getSheets().map(s => s.getName()), checks = [];
  const add = (ok, item, hint) => checks.push({ ok: ok, item: item, hint: ok ? '' : hint });
  const missing = Object.keys(SCHEMA).filter(n => names.indexOf(n) < 0);
  add(!missing.length, 'v4.1 시트 ' + Object.keys(SCHEMA).length + '개', '누락: ' + missing.join(', ') + ' → setupV4() 또는 upgradeToV41() 실행');
  if (!missing.length) add(read_('v4_Trainee').every(t => t.id), '교육생 id 부여', 'upgradeToV41() 실행');
  if (!missing.length) {
    const acc = read_('v4_Access');
    add(acc.every(a => E.role.indexOf(a.role) >= 0), 'v4_Access 역할 값', 'role은 TeamLeader/Secretary/TableLeader/Pastor/Elder 중 하나');
    add(!acc.some(a => /example\.com$/i.test(a.email)), 'v4_Access 실제 계정', 'example.com 샘플 이메일이 남아 있습니다');
    add(acc.some(a => a.role === 'TeamLeader'), 'TeamLeader 계정', 'TeamLeader 최소 1명 필요');
    add(true, 'TableLeader (선택) ' + acc.filter(a => a.role === 'TableLeader' && isActive_(a)).length + '명', '');
    const cfg = readConfig_(), wk = weeks_();
    add(wk.length === 10 && wk[0].date === cfg.open, '1주차 = 개강일', '주차 Board 1주차 날짜와 설정 개강일 불일치');
    add(wk.every(w => new Date(w.date + 'T12:00:00+09:00').getDay() === 0), '교육일 모두 일요일', '일요일이 아닌 교육일이 있습니다');
    const tr = read_('v4_Trainee');
    add(ENVIRONMENT === 'QA' || !tr.some(t => /샘플/.test(t.name)) && !read_('v4_WeeklyCare').some(c => c.by === '샘플'), 'PROD 샘플 데이터 없음', 'PROD에 샘플 데이터가 있습니다');
  }
  add(ss.getSpreadsheetTimeZone() === TZ, '스프레드시트 시간대 Asia/Seoul', '파일 > 설정에서 시간대 변경');
  add(!missing.length && !!alertTo_(), 'Red 알림 수신자', 'ALERT_EMAIL 입력 후 setAlertEmail() 실행, 또는 v4_Access에 실제 팀장·담당교역자 이메일 입력');
  add(ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'checkRedAlerts'), '알림 트리거', 'installTriggers() 실행');
  const fail = checks.filter(c => !c.ok);
  const res = { environment: ENV[ENVIRONMENT].label, spreadsheetId: ENV[ENVIRONMENT].id, status: fail.length ? 'CHECK' : 'READY', checks: checks, checkedAt: now_() };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// ===== 11. 알림 (Red Care · 일일 요약) — 이름·상담 내용은 메일에 넣지 않음 =====
function setAlertEmail(email) {
  adminOnly_(); const t = String(email || ALERT_EMAIL || '').trim();
  need_(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(t), '맨 위 ALERT_EMAIL에 이메일을 입력한 뒤 실행하세요.');
  PropertiesService.getScriptProperties().setProperty('ALERT_EMAIL', t); return 'Red 알림 수신자: ' + t;
}
function alertTo_() {
  const p = PropertiesService.getScriptProperties().getProperty('ALERT_EMAIL') || ALERT_EMAIL;
  const acc = read_('v4_Access').filter(a => isActive_(a) && (a.role === 'Pastor' || a.role === 'TeamLeader')).map(a => a.email);
  return [p].concat(acc).filter((v, i, a) => v && !/example\.com$/i.test(v) && a.indexOf(v) === i).join(',');
}
function notifyRed_(table, week) {
  const props = PropertiesService.getScriptProperties(), sent = JSON.parse(props.getProperty('RED_SENT') || '[]');
  const key = ENVIRONMENT + '|' + table + '|' + week + '|' + today_();
  if (sent.indexOf(key) >= 0 || !alertTo_()) return;
  MailApp.sendEmail(alertTo_(), '[순장학교] 새 Red Care (' + table + ' · ' + week + '주차)',
    '새 Red Care가 입력되었습니다.\nTable: ' + table + ' / ' + week + '주차\n\n이름과 상세는 운영 대시보드 Care 탭에서 확인해 주세요.\n(민감정보 보호를 위해 메일에는 최소정보만 포함합니다)');
  sent.push(key); props.setProperty('RED_SENT', JSON.stringify(sent.slice(-300)));
}
function notifyEscalation_(table) {
  const to = read_('v4_Access').filter(a => a.role === 'Pastor' && isActive_(a) && !/example\.com$/i.test(a.email)).map(a => a.email).join(',');
  if (to) MailApp.sendEmail(to, '[순장학교] 교역자 연결 요청 (' + table + ')', table + ' 교육생 1명에 대해 교역자 연결이 요청되었습니다.\n대시보드 Care 탭에서 확인 후 Review 완료로 종결해 주세요.');
}
/** 1시간 트리거: 최신 신호가 Red인데 알림 기록이 없는 건 재발송 (입력 시 즉시 발송 실패 대비) */
function checkRedAlerts() {
  const rows = read_('v4_WeeklyCare'), latest = {}, closed = {};
  read_('v4_Case').forEach(c => { if (c.stage === '종결') closed[c.name] = num_(c.closedW); });
  rows.forEach(r => { if (!latest[r.name] || num_(r.week) >= num_(latest[r.name].week)) latest[r.name] = r; });
  // 교역자 Review 완료(종결)된 건은 그 뒤 새 Red가 입력되기 전까지 재알림하지 않음
  Object.values(latest).filter(r => r.sig === 'Red' && !(r.name in closed && num_(r.week) <= closed[r.name])).forEach(r => notifyRed_(r.table, num_(r.week)));
}
/** 매일 07시: 오늘 진행·지연 중인 사이클 할 일과 마일스톤을 팀장·총무에게 요약 */
function sendDailyDigest() {
  const t = today_(), props = PropertiesService.getScriptProperties();
  if (props.getProperty('DIGEST_SENT') === t) return '오늘 요약은 이미 발송했습니다.';   // 외부 호출로 메일 한도를 소모하지 않도록 1일 1회
  const weeks = weeks_(), cfg = readConfig_(), done = {};
  read_('v4_Task').forEach(r => done[r.key] = bool_(r.done));
  const lines = [];
  weeks.forEach(w => CYCLE.forEach(c => {
    const s = addDays_(w.date, c.off), e = addDays_(w.date, c.end), k = w.no + '|' + c.id;
    if (done[k] || s > t || diffDays_(t, e) > 7) return;
    if (c.id === 'd3a' && w.ppt === 'Final') return;
    lines.push((t > e ? '[지연] ' : '[오늘] ') + w.no + '주차 ' + c.st + ' · ' + c.t + ' (' + c.o + ', ~' + e.slice(5) + ')');
  }));
  const ms = {}; read_('v4_Milestone').forEach(r => ms[r.id] = r.status);
  MS.forEach(m => { const base = m.b === 'O' ? cfg.open : weeks[8].date, s = addDays_(base, m.s), e = addDays_(base, m.e);
    if (m.id !== 'm9' && s <= t && (ms[m.id] || 'Not Started') !== 'Done' && diffDays_(t, e) <= 14) lines.push((t > e ? '[지연] ' : '[진행] ') + m.name + ' · ' + m.t + ' (~' + e.slice(5) + ')'); });
  props.setProperty('DIGEST_SENT', t);
  if (!lines.length) return;
  const to = read_('v4_Access').filter(a => isActive_(a) && (a.role === 'TeamLeader' || a.role === 'Secretary') && !/example\.com$/i.test(a.email)).map(a => a.email).join(',');
  if (to) MailApp.sendEmail(to, '[순장학교] ' + t + ' 오늘 할 일 ' + lines.length + '건', lines.join('\n') + '\n\n문구 복사와 체크는 대시보드 「오늘 할 일」에서 합니다.');
}
function installTriggers() {
  adminOnly_();
  ScriptApp.getProjectTriggers().filter(t => ['checkRedAlerts', 'sendDailyDigest'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkRedAlerts').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  return '트리거 설치: checkRedAlerts(1시간), sendDailyDigest(매일 07시)';
}

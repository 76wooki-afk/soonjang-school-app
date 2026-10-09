// 1.5단계(원본 보강) 서버 검증 — 실행: node tests/stage15.test.js
const path = require('path'), { execSync } = require('child_process'), fs = require('fs'), os = require('os');
const { makeEnv } = require('./gas-mock');
const CODE = path.join(__dirname, '../server/Code.gs');
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.log('  ✗ ' + m); } };
const throws = (fn, re, m) => { try { fn(); ok(false, m + ' (오류가 나야 함)'); } catch (e) { ok(re.test(e.message), m + ' → ' + e.message); } };

function login(env, email) {
  env.run(`requestLoginCode(${JSON.stringify(email)})`);
  const code = env.mails[env.mails.length - 1].subject.match(/\d{6}/)[0];
  return env.run(`verifyLoginCode(${JSON.stringify(email)}, '${code}')`).token;
}
const apply = (env, tok, op, p) => env.run(`apply(${JSON.stringify(op)}, ${JSON.stringify({ ...p, _tok: tok })})`).data;

// ── A. 새 시트 설치 ─────────────────────────────
const env = makeEnv(); env.load(CODE);
env.run('setupV4()'); env.run('seedQASample()');
const cfg = Object.fromEntries(env.table('v4_Config').map(r => [r.key, r.value]));
ok(cfg.gather === '14:30' && String(cfg.fee) === '50000' && cfg.holidays === '', 'A1 설정 기본값(집결 14:30·회비 50000·휴강 없음)');
ok(env.table('v4_Prep').length === 11, 'A2 개강 준비 11개');
ok(env.books[Object.keys(env.books)[0]].sheets.v4_Week.data[0].includes('community'), 'A3 주차 시트에 community 열');
ok(env.table('v4_Trainee').filter(t => t.paid === 'true').length === 20, 'A4 QA 샘플 회비 납부 20/25');
const wk0 = env.table('v4_Week').map(w => w.date);
ok(wk0[0] === '2026-10-18' && wk0[8] === '2026-12-13' && wk0[9] === '2026-12-20', 'A5 개강 10/18 · 수료 12/13 · 아웃팅 12/20');

// ── B. 팀장 로그인 후 보강 기능 ───────────────────
const tl = login(env, 'leader@example.com');
let d = apply(env, tl, 'cfg', { k: 'gather', v: '14:00' });
ok(d.cfg.gather === '14:00', 'B1 집결시간 변경');
throws(() => apply(env, tl, 'cfg', { k: 'gather', v: '2pm' }), /14:30/, 'B2 집결시간 형식 오류 거부');
d = apply(env, tl, 'cfg', { k: 'fee', v: '40000' }); ok(d.cfg.fee === 40000, 'B3 회비 변경');
d = apply(env, tl, 'cfg', { k: 'holidays', v: '2026-11-29, 잘못된값' });
const ds = d.weeks.map(w => w.date);
ok(d.cfg.holidays === '2026-11-29', 'B4 휴강일 정리(잘못된 값 제거)');
ok(!ds.includes('2026-11-29') && ds[5] === '2026-11-22' && ds[6] === '2026-12-06' && ds[9] === '2026-12-27', 'B5 휴강 이후 주차 1주씩 연기 → ' + ds.slice(4).join(','));
d = apply(env, tl, 'cfg', { k: 'holidays', v: '' }); ok(d.weeks[9].date === '2026-12-20', 'B6 휴강 해제 시 원래 일정 복귀');
d = apply(env, tl, 'wk', { i: 0, f: 'community', v: '화평' }); ok(d.weeks[0].community === '화평', 'B7 대표기도 담당공동체 저장');
const t0 = d.roster.find(t => !t.paid);
d = apply(env, tl, 'tSave', { rec: { id: t0.id, name: t0.name, table: t0.table, status: t0.status, community: '', source: '', note: '', paid: true }, ver: t0.ver });
ok(d.roster.find(t => t.id === t0.id).paid === true, 'B8 교육생 회비 납부 저장');
d = apply(env, tl, 'prepPlan', { i: 10, memo: '수료 기준표 인쇄' }); d = apply(env, tl, 'prepPlan', { i: 10, v: true });
ok(d.prepMemo[10] === '수료 기준표 인쇄' && d.prep[10] === true, 'B9 준비 체크 메모·완료가 함께 유지');
d = apply(env, tl, 'stSave', { rec: { email: 'elder@example.com', name: '담당장로', role: 'Elder', team: '', active: true } });
ok(d.staff.some(s => s.role === 'Elder'), 'B10 담당장로 계정 등록');

// ── C. 담당장로 권한 ─────────────────────────────
const el = login(env, 'elder@example.com');
const st = env.run(`refresh(${JSON.stringify(el)})`).data;
ok(st.logs.length > 0 && st.logs.every(l => l.cat === '' && l.act === ''), 'C1 담당장로: Care 상세 마스킹');
ok(st.budget.length === 10, 'C2 담당장로: 예산 조회 가능');
throws(() => apply(env, el, 'wk', { i: 0, f: 'prayer', v: 'x' }), /Elder 역할은/, 'C3 담당장로: 주차 수정 거부');
throws(() => apply(env, el, 'cfg', { k: 'fee', v: '1' }), /Elder 역할은/, 'C4 담당장로: 설정 수정 거부');

// ── D. 보안: 관리 함수 외부 호출 차단 ───────────────
env.user.email = '';
throws(() => env.run('upgradeToV42()'), /권한 없음/, 'D1 upgradeToV42 외부 호출 차단');
throws(() => env.run('setupV4()'), /권한 없음/, 'D2 setupV4 외부 호출 차단');
env.user.email = 'owner@test.kr';
const h = env.run('getSystemHealth()');
ok(h.checks.find(c => /TableLeader/.test(c.item)).ok, 'D3 헬스체크: 테이블리더는 선택 항목');

// ── E. v4.1.1 시트 → upgradeToV42 (기존 데이터 유지) ─────
const old = path.join(os.tmpdir(), 'code-v411.gs');
fs.writeFileSync(old, execSync('git show 778b582:server/Code.gs', { cwd: path.join(__dirname, '..') }));
const e1 = makeEnv(); e1.load(old); e1.run('setupV4()'); e1.run('seedQASample()');
const before = e1.table('v4_Trainee').length, careBefore = e1.table('v4_WeeklyCare').length;
const e2 = makeEnv({ books: e1.books }); e2.load(CODE);
const msg = e2.run('upgradeToV42()'); const msg2 = e2.run('upgradeToV42()');
ok(/설정 3개, 개강 준비 2개/.test(msg) && /설정 0개, 개강 준비 0개/.test(msg2), 'E1 업그레이드 1회 적용·재실행 안전 → ' + msg);
ok(e2.table('v4_Trainee').length === before && e2.table('v4_WeeklyCare').length === careBefore, 'E2 기존 교육생·Care 데이터 유지');
const tok2 = login(e2, 'leader@example.com'), s2 = e2.run(`refresh(${JSON.stringify(tok2)})`).data;
ok(s2.cfg.gather === '14:30' && s2.cfg.fee === 50000 && s2.roster.every(t => t.paid === false) && s2.weeks.every(w => w.community === ''), 'E3 새 항목 기본값으로 읽힘');

console.log(`\n1.5단계 서버 검증: ${pass}/${pass + fail} 통과`);
process.exit(fail ? 1 : 0);

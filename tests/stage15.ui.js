// 1.5단계 화면 검증 (휴대폰 390px · 시뮬레이션 모드) — 실행: node tests/stage15.ui.js <앱주소> <스크린샷폴더>
const { chromium } = require('playwright');
const [,, url, shots = '.'] = process.argv;
let pass = 0, fail = 0; const ok = (c, m) => { if (c) pass++; else { fail++; console.log('  ✗ ' + m); } };
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true, locale: 'ko-KR' });
  const p = await ctx.newPage(); const errs = [];
  p.on('pageerror', e => errs.push(e.message));
  p.on('console', m => { if (m.type() === 'error' && !/fonts\.g|net::|ERR_/.test(m.text())) errs.push(m.text()); });
  p.on('dialog', d => d.accept());
  await p.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  await p.goto(url); await p.evaluate(() => localStorage.clear()); await p.reload();
  const role = async i => { await p.selectOption('#roleSel', String(i)); await p.waitForTimeout(60); };
  const tab = async t => { await p.click(`#tabBar button[data-tab="${t}"]`); await p.waitForTimeout(60); };
  const txt = sel => p.$eval(sel, e => e.innerText);

  // 1) 9개 계정 × 전 탭 렌더링
  const roles = await p.$$eval('#roleSel option', o => o.map(x => x.textContent)); let views = 0;
  for (let i = 0; i < roles.length; i++) { await role(i);
    for (const t of await p.$$eval('#tabBar button', bs => bs.map(x => x.dataset.tab))) { await tab(t); views++;
      if ((await txt('#v-' + t)).trim().length < 20) errs.push('빈 화면 ' + roles[i] + '/' + t); } }
  ok(roles.length === 9 && roles[8].includes('담당장로'), '계정 9개(담당장로 포함)');
  await role(8); const eTabs = await p.$$eval('#tabBar button', bs => bs.map(x => x.dataset.tab));
  ok(eTabs.join() === 'today,home,care,grad,cal,manual', '담당장로 탭 6개(명단 관리 제외) → ' + eTabs);
  await tab('cal'); ok(await p.$$eval('#budget input,#budget select,#prepList input', a => a.length) === 0 || await p.$eval('#prepList input', e => e.disabled), '담당장로: 예산·준비 조회 전용');

  // 2) 예산: 수입/지출 구분 · 회비 자동
  await role(1); await tab('cal');
  const bud = await txt('#budget');
  ok(/수입 합계/.test(bud) && /지출 합계/.test(bud) && /잔액/.test(bud), '예산: 수입 합계·지출 합계·잔액 표시');
  ok(/50,000원×25명 · 납부 20/.test(bud) && /1,250,000/.test(bud) && /1,000,000/.test(bud), '예산: 회비 자동 계산(5만원×25명, 납부 20명)');
  await p.fill('#budget input[data-f="plan"][data-i="0"]', '2000000'); await p.dispatchEvent('#budget input[data-f="plan"][data-i="0"]', 'change');
  await p.fill('#budget input[data-f="plan"][data-i="3"]', '500000'); await p.dispatchEvent('#budget input[data-f="plan"][data-i="3"]', 'change');
  await p.fill('#budget input[data-f="spend"][data-i="3"]', '200000'); await p.dispatchEvent('#budget input[data-f="spend"][data-i="3"]', 'change');
  const bud2 = await txt('#budget');
  ok(/수입 합계\s+3,250,000\s+1,000,000/.test(bud2) && /지출 합계\s+500,000\s+200,000\s+집행률 40%/.test(bud2) && /잔액.*2,750,000\s+800,000/s.test(bud2), '예산: 수입 3,250,000 · 지출 500,000 · 집행률 40% · 잔액 2,750,000');

  // 3) 개강 준비 11개 + 메모
  ok(await p.$$eval('#prepList label.pchk', a => a.length) === 11, '개강 준비 11개');
  await p.fill('#prepList input[data-act="prepMemo"][data-i="4"]', '인쇄소 견적 2곳'); await p.dispatchEvent('#prepList input[data-act="prepMemo"][data-i="4"]', 'change');
  await tab('today'); await tab('cal');
  ok(await p.$eval('#prepList input[data-act="prepMemo"][data-i="4"]', e => e.value) === '인쇄소 견적 2곳', '준비 메모 저장 유지');
  await p.locator('#budget').locator('xpath=ancestor::div[contains(@class,"box")][1]').screenshot({ path: shots + '/v42-budget.png' });

  // 4) 집결시간 → D-Day·Run-of-Show·문구
  await tab('dday');
  ok((await txt('#ddList')).includes('집결\n14:30'), 'D-Day: 시작 전 항목이 집결 14:30 기준');
  ok(/14:30~15:20\s+50분/.test(await txt('#ros')), 'Run-of-Show: 14:30~15:20 준비 50분');
  ok(!/13:30/.test(await txt('#v-dday')), 'D-Day 화면에 13:30 없음');
  await tab('manual'); await p.fill('#settings input[data-k="gather"]', '14:00'); await p.dispatchEvent('#settings input[data-k="gather"]', 'change');
  await tab('dday'); ok(/14:00~15:20\s+80분/.test(await txt('#ros')), '집결 14:00 변경 → 준비 80분');
  await tab('manual'); await p.click('#tplLib button[data-k="rm_d1s"]');
  ok((await p.$eval('#mText', e => e.value)).includes('집결시간은 14:00'), 'Staff D-1 문구: 집결 14:00');
  await p.click('[data-act="closeModal"]');

  // 5) 휴강일 → 주차 연기
  await p.fill('#settings input[data-k="holidays"]', '2026-11-29'); await p.dispatchEvent('#settings input[data-k="holidays"]', 'change');
  await tab('board'); const bt = await txt('#board');
  ok(/12\/6\(일\)/.test(bt) && /12\/27\(일\)/.test(bt) && !/11\/29/.test(bt), '휴강 11/29 → 7주차 12/6 · 10주차 12/27');
  ok(/휴강 2026-11-29/.test(await txt('#boardHead')), 'Board 상단에 휴강 표시');
  // 6) 대표기도 담당공동체
  ok(await p.$$eval('#board input[data-f="community"]', a => a.length) === 10 && await p.$('#commList'), '담당공동체 칸 10주 + 추천 목록');
  await p.fill('#board input[data-f="community"][data-i="0"]', '화평'); await p.dispatchEvent('#board input[data-f="community"][data-i="0"]', 'change');
  await tab('dday'); await p.click('#ddWeeks button[data-w="1"]');
  ok(/(화평 공동체)/.test(await txt('#ros')), 'Run-of-Show 대표기도에 공동체 표시');
  await tab('manual'); await p.fill('#settings input[data-k="holidays"]', ''); await p.dispatchEvent('#settings input[data-k="holidays"]', 'change');

  // 7) 회비 납부: 명단·KPI
  await tab('admin'); ok(/회비 납부 20명/.test(await txt('#admBody')), '명단: 회비 납부 20명 집계');
  await p.click('#admBody tr.pick[data-id="Q0"]'); await p.check('#admForm input[name="paid"]'); await p.click('[data-adm="saveT"]');
  ok(/회비 납부 21명/.test(await txt('#admBody')), '명단: 납부 체크 저장 → 21명');
  await tab('home'); const k = await txt('#kpis');
  ok(/회비 납부\s+21\/25명/.test(k) && /미납 4명/.test(k), 'KPI: 회비 납부 21/25 · 미납 4');
  await p.screenshot({ path: shots + '/v42-home-kpi.png' });
  await role(3); await tab('home'); ok(!/회비 납부/.test(await txt('#kpis')), '테이블리더 KPI에는 회비 미표시');

  const hs = await p.evaluate(() => document.documentElement.scrollWidth > innerWidth + 1); ok(!hs, '가로 스크롤 깨짐 없음');
  ok(!errs.length, '스크립트 오류 0건 ' + errs.slice(0, 3).join(' | '));
  console.log(`\n1.5단계 화면 검증: ${pass}/${pass + fail} 통과 (화면 ${views}개 렌더링)`);
  await b.close(); process.exit(fail ? 1 : 0);
})();

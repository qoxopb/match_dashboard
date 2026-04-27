const { newBackgroundPage } = require('./browser');
const { getSheetsApi } = require('../sheets');
const config = require('../config.json');
const { shouldAbort, setActivePage } = require('../abort');

const ADMIN_URL = 'https://tutor-admin.qanda.ai/admin/tutor-pairing';

function getMonthlyTabName(type) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  return type === 'new' ? `[신규] ${yy}.${mm}` : `[재매칭] ${yy}.${mm}`;
}

function colNumberToLetter(n) {
  let s = '';
  n++;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// --- 시간 비교 유틸 ---
// 어드민 __NEXT_DATA__의 weeklyAvailablePeriods → 내부 형식 변환
function parseAdminSchedule(weeklyAvailablePeriods) {
  const keyMap = { monday: 'mon', tuesday: 'tue', wednesday: 'wed', thursday: 'thu', friday: 'fri', saturday: 'sat', sunday: 'sun' };
  const schedule = {};
  for (const [fullDay, hours] of Object.entries(weeklyAvailablePeriods || {})) {
    const key = keyMap[fullDay];
    if (key && Array.isArray(hours) && hours.length > 0) {
      schedule[key] = hours.sort((a, b) => a - b);
    }
  }
  return schedule;
}

// product name 문자열에서 주 n회, n분 파싱
function parseLessonFromName(prodName) {
  const weeklyM = prodName.match(/주\s*(\d+)\s*회/);
  const minM = prodName.match(/(\d+)\s*분/);
  if (!weeklyM && !minM) return null;
  const weekly = weeklyM ? parseInt(weeklyM[1]) : 1;
  const minutes = minM ? parseInt(minM[1]) : 60;
  const hoursNeeded = Math.ceil(minutes / 60); // 60→1, 90→2, 120→2
  return { weekly, minutes, hoursNeeded };
}

// 학생 시간표 유효성 검증: 주 n회 x n시간 연속 블록이 가능한지
function validateStudentSchedule(schedule, weekly, hoursNeeded) {
  let validDays = 0;
  for (const [day, hours] of Object.entries(schedule)) {
    if (hasConsecutiveBlock(hours, hoursNeeded)) validDays++;
  }
  return validDays >= weekly;
}

// 연속 시간 블록 존재 여부
function hasConsecutiveBlock(hours, needed) {
  if (hours.length < needed) return false;
  let count = 1;
  for (let i = 1; i < hours.length; i++) {
    if (hours[i] === hours[i - 1] + 1) {
      count++;
      if (count >= needed) return true;
    } else {
      count = 1;
    }
  }
  return count >= needed;
}

// 선생님 시간과 학생 시간 매칭: 주 n회 이상 요일에서 n시간 연속 겹침
function matchSchedules(studentSchedule, tutorSchedule, weekly, hoursNeeded) {
  const dayMap = { '월요일': 'mon', '화요일': 'tue', '수요일': 'wed', '목요일': 'thu', '금요일': 'fri', '토요일': 'sat', '일요일': 'sun' };
  let matchedDays = 0;

  for (const [korDay, tutorHours] of Object.entries(tutorSchedule)) {
    const dayKey = dayMap[korDay];
    if (!dayKey || !studentSchedule[dayKey]) continue;

    const studentHours = studentSchedule[dayKey];
    // 겹치는 시간
    const overlap = studentHours.filter(h => tutorHours.includes(h)).sort((a, b) => a - b);
    if (hasConsecutiveBlock(overlap, hoursNeeded)) matchedDays++;
  }

  return matchedDays >= weekly;
}

// --- 메인 ---
async function runAimForType(type) {
  const label = type === 'new' ? '신규' : '재매칭';
  console.log(`\n[aim] === ${label} 매칭 제안 시작 ===`);
  const startedAt = Date.now();

  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  const monthlyTab = getMonthlyTabName(type);

  // 헤더 + 데이터
  const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${monthlyTab}'!A1:BZ1` });
  const header = (headerRes.data.values || [])[0] || [];
  const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${monthlyTab}'!A2:BZ` });
  const rows = dataRes.data.values || [];

  const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
  const matchIdColName = type === 'new' ? 'match_id' : 'match ID';
  const statusColName = type === 'new' ? '매칭상태' : 'status';
  const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdColName));
  const statusIdx = header.findIndex(h => norm(h) === norm(statusColName));
  // 재매칭: pairing ID 컬럼
  const pairingIdIdx = type === 'rematch' ? header.findIndex(h => norm(h) === norm('pairing ID')) : -1;

  if (matchIdIdx < 0 || statusIdx < 0) throw new Error('헤더에서 컬럼을 찾을 수 없습니다');

  const excludeNew = ['매칭완료', '동일튜터유지', '동일 튜터 유지', '환불'];
  const excludeRematch = ['재매칭 완료', '재매칭완료', '환불/소멸', '환불', '소멸'];
  const excludeList = (type === 'new' ? excludeNew : excludeRematch).map(norm);

  // 메모 컬럼 찾기 (확인필요 시 메모 기재용)
  const memoIdx = header.findIndex(h => norm(h) === '매칭담당자메모' || norm(h) === 'memo');

  // 메모 컬럼 디버그
  console.log(`[aim] 메모 컬럼 위치: ${memoIdx} (헤더: "${memoIdx >= 0 ? header[memoIdx] : 'NOT FOUND'}")`);

  // AIM 키워드 설정 로드
  const aimKeywords = config.aimKeywords || [];
  const normNoSpace = s => (s || '').replace(/\s+/g, '').toLowerCase();

  const targets = [];
  let keywordSkipped = 0;
  rows.forEach((row, i) => {
    const matchId = (row[matchIdIdx] || '').trim();
    const status = (row[statusIdx] || '').trim();
    if (matchId && !excludeList.includes(norm(status))) {
      const formData = {};
      header.forEach((h, ci) => { formData[h] = row[ci] || ''; });

      // 메모 컬럼에서 키워드 매칭 (exclude 우선)
      const memoText = memoIdx >= 0 ? (row[memoIdx] || '') : '';
      const memoNorm = normNoSpace(memoText);
      let keywordAction = null;
      // 1차: exclude 키워드 우선 체크
      for (const kw of aimKeywords) {
        if (kw.action === 'exclude' && memoNorm.includes(normNoSpace(kw.keyword))) {
          keywordAction = kw;
          break;
        }
      }
      // 2차: exclude 아닌 키워드
      if (!keywordAction) {
        for (const kw of aimKeywords) {
          if (kw.action !== 'exclude' && memoNorm.includes(normNoSpace(kw.keyword))) {
            keywordAction = kw;
            break;
          }
        }
      }

      if (keywordAction && keywordAction.action === 'exclude') {
        keywordSkipped++;
        return; // 제외
      }

      const pairingId = pairingIdIdx >= 0 ? (row[pairingIdIdx] || '').trim() : '';
      targets.push({ matchId, rowNum: i + 2, formData, keywordAction, pairingId });
    }
  });

  if (keywordSkipped > 0) console.log(`[aim] 키워드 제외: ${keywordSkipped}건`);
  if (aimKeywords.length > 0) console.log(`[aim] 키워드 설정: ${aimKeywords.map(k => `"${k.keyword}"→${k.action}`).join(', ')}`);
  console.log(`[aim] 대상 ${targets.length}건`);
  if (targets.length === 0) return;

  const statusColLetter = colNumberToLetter(statusIdx);
  const memoColLetter = memoIdx >= 0 ? colNumberToLetter(memoIdx) : null;
  const page = await newBackgroundPage();
  setActivePage(page);
  const counters = { aimed: 0, manual: 0, noData: 0, skipped: 0, needCheck: 0, error: 0 };

  try {
    for (const target of targets) {
      if (shouldAbort()) { console.log(`[aim] 중단됨`); break; }
      const tag = `[aim] ${target.matchId}`;
      try {
        const result = await processOne(page, target, sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, tag, type);
        counters[result] = (counters[result] || 0) + 1;
      } catch (e) {
        // navigation interrupted → 페이지 안정화 후 1회 재시도
        if (/interrupted.*navigation/i.test(e.message)) {
          console.log(`${tag} 네비게이션 충돌 → 대기 후 재시도`);
          try { await page.waitForLoadState('domcontentloaded', { timeout: 5000 }); } catch {}
          await page.waitForTimeout(1000);
          try {
            const result = await processOne(page, target, sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, tag, type);
            counters[result] = (counters[result] || 0) + 1;
          } catch (e2) {
            console.error(`${tag} 재시도 에러: ${e2.message}`);
            counters.error++;
          }
        } else {
          console.error(`${tag} 에러: ${e.message}`);
          counters.error++;
        }
      }
    }
  } finally {
    await page.close().catch(() => {});
  }

  const elapsed = Date.now() - startedAt;
  const mins = Math.floor(elapsed / 60000);
  const secs = Math.floor((elapsed % 60000) / 1000);
  console.log(`[aim] === ${label} 완료: AIM ${counters.aimed || 0}, 수동 ${counters.manual || 0}, No Data ${counters.noData || 0}, 확인필요 ${counters.needCheck || 0}, 스킵 ${counters.skipped || 0}, 에러 ${counters.error} | ${mins > 0 ? mins + '분 ' : ''}${secs}초 ===`);
}

// --- systemError 처리 헬퍼 ---
async function handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, reason) {
  console.log(`${tag} 시스템 에러 → 확인필요: ${reason}`);
  await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '확인 필요');
  if (memoColLetter) {
    const existingRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}` });
    const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
    if (!existing.includes(reason)) {
      const newMemo = existing ? `${existing}\n${reason}` : reason;
      await sheets.spreadsheets.values.update({ spreadsheetId, range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[newMemo]] } });
    }
  }
  return 'needCheck';
}

// --- 매치아이디 1건 처리 ---
async function processOne(page, target, sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, tag, type) {
  // 검색 페이지 (재매칭은 pairingId, 신규는 matchId)
  const searchId = type === 'rematch' && target.pairingId ? target.pairingId : target.matchId;
  const searchFilter = type === 'rematch' && target.pairingId ? 'filters_id' : 'filters_matchId';
  await page.goto(
    `${ADMIN_URL}?page=1&size=20&${searchFilter}=${encodeURIComponent(searchId)}`,
    { waitUntil: 'domcontentloaded', timeout: 30000 }
  );
  await page.waitForSelector('table tbody, .ant-table-tbody', { timeout: 8000 });

  // 어드민 status 확인
  const adminStatus = await getAdminStatus(page);
  console.log(`${tag} 어드민 상태: "${adminStatus}"`);

  if (/신청서\s*미?작성/i.test(adminStatus)) {
    console.log(`${tag} 신청서 미작성 → 스킵`);
    return 'skipped';
  }

  // 매칭완료 → 시트 업데이트 후 다음
  if (/매칭\s*완료/i.test(adminStatus)) {
    console.log(`${tag} 매칭완료 → 시트 반영`);
    const completeStatus = type === 'new' ? '매칭완료' : '재매칭 완료';
    await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, completeStatus);
    return 'skipped';
  }

  // 상세 페이지 이동
  const detailHref = await page.evaluate(() => {
    const link = document.querySelector('table tbody a[href*="/tutor-pairing/"]');
    return link ? link.getAttribute('href') : null;
  });
  if (!detailHref) { console.log(`${tag} 상세 링크 못 찾음`); return 'error'; }

  await page.goto(new URL(detailHref, 'https://tutor-admin.qanda.ai').href, { waitUntil: 'domcontentloaded', timeout: 30000 });

  // __NEXT_DATA__에서 학생 시간표 + product 정보 한 번에 읽기
  try { await page.waitForSelector('#__NEXT_DATA__', { state: 'attached', timeout: 5000 }); } catch {}
  const pageData = await page.evaluate(() => {
    const script = document.getElementById('__NEXT_DATA__');
    if (!script) return null;
    try {
      const json = JSON.parse(script.textContent);
      const pp = json?.props?.pageProps;
      return {
        weeklyAvailablePeriods: pp?.applicant?.weeklyAvailablePeriods || null,
        productName: pp?.product?.name || null,
        productMinutes: pp?.product?.minutes || null,
        tutorGenders: pp?.applicant?.tutorStyles?.tutorGenders || [],
        tutorUniversities: pp?.applicant?.tutorStyles?.tutorUniversities || [],
        rank: pp?.applicant?.tutorRank || pp?.pairing?.tutorRank || '',
      };
    } catch { return null; }
  });

  // AIM 실패 → tutorPairingStatus를 MATCHING으로 변경 후 재시도
  if (/aim\s*실패/i.test(adminStatus)) {
    console.log(`${tag} AIM 실패 → tutorPairingStatus MATCHING으로 변경`);
    const fixResult = await fixPairingStatus(page, tag);
    if (fixResult) {
      // 상세 페이지로 다시 이동
      const detailUrl = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
      await page.goto(detailUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
      await page.waitForTimeout(2000);
    } else {
      console.log(`${tag} tutorPairingStatus 변경 실패`);
    }
  }

  // AIM 진행중(AIM_ING)이면 중단 후 재검색
  if (/aim.?ing|aim\s*진행|aim\s*매칭\s*중/i.test(adminStatus)) {
    console.log(`${tag} AIM 진행중 → 매칭 중단`);
    try { await page.waitForFunction(() => document.body.innerText.includes('AI 매칭'), { timeout: 8000 }); } catch {}
    await page.evaluate(() => {
      const btn = [...document.querySelectorAll('button')].find(b => /매칭\s*중단/i.test(b.textContent));
      if (btn) btn.click();
    });
    await page.waitForTimeout(2000);
  }

  // "매칭 중"(AIM 아닌)이면 중단 없이 바로 AIM 검색

  // === 특정 튜터풀 모드 → AIM 건너뛰고 바로 튜터풀 검색 ===
  if (target.keywordAction && target.keywordAction.action === 'tutorPool' && target.keywordAction.tutorIds) {
    // product + 시간 정보
    let lessonInfo = null;
    if (pageData?.productName) {
      lessonInfo = parseLessonFromName(pageData.productName);
    }
    const weekly = lessonInfo?.weekly || 1;
    const hoursNeeded = lessonInfo?.hoursNeeded || 1;
    const studentSchedule = pageData?.weeklyAvailablePeriods
      ? parseAdminSchedule(pageData.weeklyAvailablePeriods) : {};
    const studentGenders = (pageData?.tutorGenders || []).map(g => g === 'MALE' ? '남' : g === 'FEMALE' ? '여' : g);

    const isExpert = (pageData?.rank || '').toUpperCase() === 'EXPERT';
    console.log(`${tag} 특정튜터풀 모드: 주${weekly}회 ${lessonInfo?.minutes || '?'}분 | 성별: [${studentGenders.join(',')}] | PRO: ${isExpert}`);
    console.log(`${tag} 튜터풀 ID: [${target.keywordAction.tutorIds}]`);

    // PRO면 전문 강사 여부 토글 ON
    if (isExpert) {
      const toggleResult = await page.evaluate(() => {
        const all = [...document.querySelectorAll('*')];
        for (const el of all) {
          if (el.children.length === 0 && /전문\s*강사\s*여부/i.test(el.textContent.trim())) {
            let parent = el.parentElement;
            for (let d = 0; d < 5 && parent; d++) {
              const toggle = parent.querySelector('.ant-switch, button[role="switch"], input[type="checkbox"]');
              if (toggle) {
                const isOn = toggle.classList.contains('ant-switch-checked') || toggle.getAttribute('aria-checked') === 'true';
                if (!isOn) { toggle.click(); return 'ON으로 변경'; }
                return '이미 ON';
              }
              parent = parent.parentElement;
            }
          }
        }
        return '토글 못 찾음';
      });
      console.log(`${tag} 전문강사 토글 (튜터풀): ${toggleResult}`);
      await page.waitForTimeout(300);
    }

    const poolResult = await tryTutorPoolMatch(page, tag, studentSchedule, weekly, hoursNeeded, target.keywordAction.tutorIds, studentGenders);
    if (poolResult === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
    if (poolResult === 'matched') {
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
      console.log(`${tag} 튜터풀 매칭 성공 → "매칭중"`);
      return 'manual';
    }
    console.log(`${tag} 튜터풀 매칭 실패`);
    return 'noData';
  }

  // === 교대/메디컬 포함 시 → AIM 건너뛰고 변형 수동매칭 우선 ===
  const uniTypes = (pageData?.tutorUniversities || []).map(u => u.toUpperCase());
  const hasGyodae = uniTypes.includes('GYODAE') || uniTypes.includes('교대');
  const hasMedical = uniTypes.includes('MEDICAL') || uniTypes.includes('메디컬');

  if (hasGyodae || hasMedical) {
    const uniLabel = hasGyodae ? '교대' : '메디컬';
    console.log(`${tag} ${uniLabel} 포함 → AIM 건너뛰고 변형 수동매칭 우선`);

    // product + 시간 정보 (여기서 미리 계산)
    let lessonInfo = null;
    if (pageData?.productName) {
      console.log(`${tag} product name: "${pageData.productName}"`);
      lessonInfo = parseLessonFromName(pageData.productName);
    }
    if (!lessonInfo) {
      console.log(`${tag} product 정보 없음 → 확인필요`);
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '확인 필요');
      return 'needCheck';
    }
    const { weekly, minutes, hoursNeeded } = lessonInfo;
    let studentSchedule = pageData?.weeklyAvailablePeriods
      ? parseAdminSchedule(pageData.weeklyAvailablePeriods) : {};

    // 시간표 자동 확장 (필요 시)
    if (!validateStudentSchedule(studentSchedule, weekly, hoursNeeded)) {
      const scheduleDays = Object.keys(studentSchedule);
      if (scheduleDays.length >= weekly) {
        const fixResult = await fixAdminSchedule(page, studentSchedule, hoursNeeded, tag);
        if (fixResult.ok) {
          Object.assign(studentSchedule, fixResult.expandedSchedule);
          const detailUrl = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
          await page.goto(detailUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
          await page.waitForTimeout(2000);
        }
      }
    }

    const isExpert = (pageData?.rank || '').toUpperCase() === 'EXPERT';

    // 1차: 변형 필터로 수동매칭
    const filterConfig = hasGyodae
      ? { uncheckUni: '교대', addUni: ['SKY', '서성한', '중경외시'], addMajor: ['교육'] }
      : { uncheckUni: '메디컬', addUni: ['SKY', '서성한', '중경외시'], addMajor: ['의치', '한약수'] };

    const filterResult = await applyUniMajorFilter(page, tag, filterConfig);
    console.log(`${tag} ${uniLabel} 변형 필터: ${filterResult}`);

    const altResult = await tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, 'base');
    if (altResult === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
    if (altResult === 'matched') {
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
      console.log(`${tag} ${uniLabel} 변형 매칭 성공 → "매칭중"`);
      return 'manual';
    }

    // 1�� 실패 → 원래 조건으로 리셋 (페이지 재이동)
    console.log(`${tag} ${uniLabel} 변형 매칭 실패 → 원래 조건으로 리셋`);
    const detailUrl = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
    await page.goto(detailUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);

    // 원래 조건으로 일반 수동매칭 + 임의완화
    for (const rl of ['base', 'relax1', 'relax2', 'relax3']) {
      const r = await tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, rl);
      if (r === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
      if (r === 'matched') {
        await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
        return 'manual';
      }
    }

    // PRO 마지막 시도: 교대/메디컬 해제 + SKY/서성한/중경외시
    if (isExpert) {
      console.log(`${tag} PRO → 교대/메디컬 해제 + SKY/서성한/중경외시 매칭 시도`);
      const detailUrl2 = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
      await page.goto(detailUrl2, { waitUntil: 'domcontentloaded', timeout: 30000 });
      await page.waitForTimeout(2000);
      await applyUniMajorFilter(page, tag, { uncheckUni: '교대', addUni: ['SKY', '서성한', '중경외시'], addMajor: [] });
      await applyUniMajorFilter(page, tag, { uncheckUni: '메디컬', addUni: [], addMajor: [] });
      const proResult = await tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, 'base');
      if (proResult === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
      if (proResult === 'matched') {
        await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
        console.log(`${tag} PRO 매칭 성공 → "매칭중"`);
        return 'manual';
      }
    }

    // 전부 실패
    console.log(`${tag} 모든 매칭 실패 → 조건완화 필요`);
    await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '확인 필요');
    if (memoColLetter) {
      const existingRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}` });
      const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
      if (!existing.includes('조건완화 필요')) {
        const newMemo = existing ? `${existing}\n조건완화 필요` : '조건완화 필요';
        await sheets.spreadsheets.values.update({ spreadsheetId, range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[newMemo]] } });
      }
    }
    return 'noData';
  }

  // === 1단계: AIM ===
  const aimResult = await tryAim(page, tag);
  if (aimResult === 'aimed') {
    await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
    console.log(`${tag} AIM 성공 → "매칭중"`);
    // AIM 성공 후 어드민이 상세 페이지로 리다이렉트할 수 있음 → 네비게이션 완료 대기
    try { await page.waitForLoadState('domcontentloaded', { timeout: 5000 }); } catch {}
    await page.waitForTimeout(1000);
    return 'aimed';
  }

  // === 2단계: 수동매칭 ===
  console.log(`${tag} AIM No Data → 수동매칭 시도`);

  // product 정보: __NEXT_DATA__에서 이미 가져옴 (페이지 이동 불필요)
  let lessonInfo = null;
  if (pageData?.productName) {
    console.log(`${tag} product name: "${pageData.productName}"`);
    lessonInfo = parseLessonFromName(pageData.productName);
  }

  if (!lessonInfo) {
    console.log(`${tag} product 정보 없음 → 확인필요`);
    const checkStatus = '확인 필요';
    await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, checkStatus);
    if (memoColLetter) {
      const existingRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
      });
      const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
      const newMemo = existing ? `${existing}\nproduct 시수 파싱 실패` : 'product 시수 파싱 실패';
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[newMemo]] },
      });
    }
    return 'needCheck';
  }

  const { weekly, minutes, hoursNeeded } = lessonInfo;
  const studentSchedule = pageData?.weeklyAvailablePeriods
    ? parseAdminSchedule(pageData.weeklyAvailablePeriods)
    : {};

  console.log(`${tag} 시수 정보 (어드민): 주${weekly}회 ${minutes}분 (${hoursNeeded}시간 블록 필요)`);
  console.log(`${tag} 학생 시간표 (어드민): ${JSON.stringify(studentSchedule)}`);

  if (!validateStudentSchedule(studentSchedule, weekly, hoursNeeded)) {
    // 요일 수가 주N회와 맞으면 → 시간 자동 확장 후 어드민 수정
    const scheduleDays = Object.keys(studentSchedule);
    if (scheduleDays.length >= weekly) {
      console.log(`${tag} 시간 불일치, 요일 수 ${scheduleDays.length} ≥ 주${weekly}회 → 시간 자동 확장`);
      const fixResult = await fixAdminSchedule(page, studentSchedule, hoursNeeded, tag);
      if (fixResult.ok) {
        // 수정 성공 → 확장된 시간표로 수동매칭 진행
        Object.assign(studentSchedule, fixResult.expandedSchedule);
        console.log(`${tag} 어드민 시간표 수정 완료 → 상세 페이지 재이동`);
        // Submit 후 페이지가 목록으로 돌아갈 수 있으므로 상세 페이지로 다시 이동
        const detailUrl = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
        await page.goto(detailUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
        await page.waitForTimeout(2000);
        // validateStudentSchedule 통과 후 아래 수동매칭으로 fall through
      } else {
        console.log(`${tag} 어드민 시간표 수정 실패: ${fixResult.reason}`);
        // 실패 시 확인필요
        const checkStatus = '확인 필요';
        await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, checkStatus);
        if (memoColLetter) {
          const existingRes = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
          });
          const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
          const newMemo = existing ? `${existing}\n희망시간대와 시수 불일치 (자동수정 실패)` : '희망시간대와 시수 불일치 (자동수정 실패)';
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[newMemo]] },
          });
        }
        return 'needCheck';
      }
    } else {
      console.log(`${tag} 희망시간대와 시수 불일치 (요일 수 ${scheduleDays.length} < 주${weekly}회) → 확인필요`);
      const checkStatus = '확인 필요';
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, checkStatus);
      if (memoColLetter) {
        const existingRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
        });
        const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
        const newMemo = existing ? `${existing}\n희망시간대와 시수 불일치` : '희망시간대와 시수 불일치';
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[newMemo]] },
        });
      }
      return 'needCheck';
    }
  }

  // rank 확인 (EXPERT면 전문 강사 토글)
  const isExpert = (pageData?.rank || '').toUpperCase() === 'EXPERT';

  // 수동 검색 시도 (기본 조건 + 임의완화)
  for (const rl of ['base', 'relax1', 'relax2', 'relax3']) {
    if (rl !== 'base') console.log(`${tag} 수동매칭 실패 → 임의완화 시도 (${rl})`);
    const r = await tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, rl);
    if (r === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
    if (r === 'matched') {
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
      console.log(`${tag} ${rl === 'base' ? '수동매칭' : rl} 성공 → "매칭중"`);
      return 'manual';
    }
  }

  // PRO 마지막 시도: 교대/메디컬 해제 + SKY/서성한/중경외시
  if (isExpert) {
    console.log(`${tag} PRO → 교대/메디컬 해제 + SKY/서성한/중경외시 매칭 시도`);
    const detailUrl = new URL(detailHref, 'https://tutor-admin.qanda.ai').href;
    await page.goto(detailUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);
    await applyUniMajorFilter(page, tag, { uncheckUni: '교대', addUni: ['SKY', '서성한', '중경외시'], addMajor: [] });
    await applyUniMajorFilter(page, tag, { uncheckUni: '메디컬', addUni: [], addMajor: [] });
    const proResult = await tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, 'base');
    if (proResult === 'systemError') return await handleSystemError(sheets, spreadsheetId, monthlyTab, statusColLetter, memoColLetter, target, tag, 'orderStatus 확인 필요');
    if (proResult === 'matched') {
      await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '매칭중');
      console.log(`${tag} PRO 매칭 성공 → "매칭중"`);
      return 'manual';
    }
  }

  // 모든 단계 실패 → 시트에 "확인 필요" + 메모에 "조건완화 필요"
  console.log(`${tag} 모든 매칭 실패 → 조건완화 필요`);
  await updateSheet(sheets, spreadsheetId, monthlyTab, statusColLetter, target.rowNum, '확인 필요');
  if (memoColLetter) {
    const existingRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
    });
    const existing = ((existingRes.data.values || [])[0] || [])[0] || '';
    if (!existing.includes('조건완화 필요')) {
      const newMemo = existing ? `${existing}\n조건완화 필요` : '조건완화 필요';
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${monthlyTab}'!${memoColLetter}${target.rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[newMemo]] },
      });
    }
  }
  return 'noData';
}

// --- 어드민 status 읽기 ---
async function getAdminStatus(page) {
  return page.evaluate(() => {
    const table = document.querySelector('table, .ant-table');
    if (!table) return '';
    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim().toLowerCase());
    const statusIdx = headers.findIndex(h => h === 'status' || h.includes('status'));
    if (statusIdx < 0) return '';
    const rows = [...table.querySelectorAll('tbody tr')].filter(r => {
      const cells = r.querySelectorAll('td');
      return cells.length > 1 && [...cells].some(c => c.textContent.trim());
    });
    if (rows.length === 0) return '';
    const cells = rows[0].querySelectorAll('td');
    return cells.length > statusIdx ? cells[statusIdx].textContent.trim() : '';
  });
}

// --- 시트 상태 업데이트 ---
async function updateSheet(sheets, spreadsheetId, monthlyTab, colLetter, rowNum, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${monthlyTab}'!${colLetter}${rowNum}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
}

// --- 1단계: AIM 시도 ---
async function tryAim(page, tag) {
  try {
    await page.waitForFunction(() => document.body.innerText.includes('AI 매칭 검색'), { timeout: 8000 });
  } catch {
    console.log(`${tag} AI 매칭 검색 섹션 못 찾음`);
    return 'noData';
  }

  // 검색 버튼 클릭
  await page.evaluate(() => {
    const all = [...document.querySelectorAll('*')];
    for (const el of all) {
      if (el.children.length === 0 && el.textContent.trim() === 'AI 매칭 검색') {
        let parent = el.parentElement;
        for (let d = 0; d < 5 && parent; d++) {
          const btn = parent.querySelector('button');
          if (btn) { btn.click(); return; }
          parent = parent.parentElement;
        }
      }
    }
  });
  console.log(`${tag} AI 검색 시작`);
  await page.waitForTimeout(4000);

  // AIM 매칭 시작 클릭
  const aimBtnResult = await page.evaluate(() => {
    const btn = [...document.querySelectorAll('button')].find(b => /aim.*매칭.*시작|매칭.*시작/i.test(b.textContent));
    if (btn) { btn.click(); return `clicked: "${btn.textContent.trim()}"`; }
    const allBtns = [...document.querySelectorAll('button')].map(b => b.textContent.trim()).filter(Boolean);
    return `못 찾음. 버튼 목록: [${allBtns.join(', ')}]`;
  });
  console.log(`${tag} AIM 매칭시작 버튼: ${aimBtnResult}`);
  await page.waitForTimeout(3000);

  // 모달 체크 — 모달이 뜰 때까지 잠시 대기
  let noTarget = false;
  for (let i = 0; i < 5; i++) {
    noTarget = await page.evaluate(() => document.body.innerText.includes('제안서 발송 대상이 없습니다'));
    if (noTarget) break;
    // AIM 진행중 표시가 있으면 아직 처리중
    const aimProgress = await page.evaluate(() => {
      const text = document.body.innerText;
      return text.includes('매칭 중') || text.includes('진행 중') || text.includes('Loading');
    });
    if (!aimProgress) break;
    await page.waitForTimeout(1500);
  }

  if (noTarget) {
    await page.evaluate(() => {
      const btn = [...document.querySelectorAll('.ant-modal button, .ant-btn')].find(b => /ok|확인|닫기/i.test(b.textContent));
      if (btn) btn.click();
    });
    console.log(`${tag} AIM 제안 대상 없음`);
    return 'noData';
  }

  // AIM 결과가 실제로 있는지 확인
  const aimResultCheck = await page.evaluate(() => {
    const modals = document.querySelectorAll('.ant-modal');
    const modalTexts = [...modals].map(m => m.innerText.substring(0, 200));
    const tables = document.querySelectorAll('table');
    const lastTable = tables[tables.length - 1];
    const rowCount = lastTable ? lastTable.querySelectorAll('tbody tr').length : 0;
    return { modalCount: modals.length, modalTexts, tableRowCount: rowCount };
  });
  console.log(`${tag} AIM 결과 확인 — 모달 ${aimResultCheck.modalCount}개, 테이블 행 ${aimResultCheck.tableRowCount}`);

  return 'aimed';
}

// --- 2단계/3단계: 수동매칭 ---
async function tryManualMatch(page, tag, studentSchedule, weekly, hoursNeeded, isExpert, relaxLevel) {
  // 필터 설정
  if (relaxLevel !== 'base') {
    await applyRelaxation(page, relaxLevel, tag);
  }

  // EXPERT면 전문 강사 여부 토글 ON (매 검색 직전 확인)
  if (isExpert) {
    const toggleResult = await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && /전문\s*강사\s*여부/i.test(el.textContent.trim())) {
          let parent = el.parentElement;
          for (let d = 0; d < 5 && parent; d++) {
            const toggle = parent.querySelector('.ant-switch, button[role="switch"], input[type="checkbox"]');
            if (toggle) {
              const isOn = toggle.classList.contains('ant-switch-checked') || toggle.getAttribute('aria-checked') === 'true';
              if (!isOn) { toggle.click(); return 'ON으로 변경'; }
              return '이미 ON';
            }
            parent = parent.parentElement;
          }
        }
      }
      return '토글 못 찾음';
    });
    console.log(`${tag} 전문강사 토글 (${relaxLevel}): ${toggleResult}`);
    await page.waitForTimeout(300);
  }

  // 교대 단독 선택 시 → SKY/서성한/중경외시 추가 + 전공 > 교육 체크
  if (relaxLevel === 'base') {
    const gyodaeResult = await fixGyodaeFilter(page, tag);
    if (gyodaeResult) console.log(`${tag} 교대 필터 보정: ${gyodaeResult}`);
  }

  // "선생님 필터 검색" 섹션의 Search 버튼 클릭
  const searchClicked = await page.evaluate(() => {
    // "Search" 버튼 찾기 (선생님 필터 검색 영역)
    const btns = [...document.querySelectorAll('button')];
    for (const btn of btns) {
      if (/^Search$/i.test(btn.textContent.trim())) {
        btn.click();
        return true;
      }
    }
    return false;
  });
  console.log(`${tag} 수동 검색 시작 (${relaxLevel}) — Search 버튼: ${searchClicked ? 'O' : 'X'}`);
  await page.waitForTimeout(3000);

  // 검색 결과 진단: 몇 건 나왔는지, 상태 분포
  const searchDiag = await page.evaluate(() => {
    const tables = document.querySelectorAll('table');
    const table = tables[tables.length - 1];
    if (!table) return { found: false, msg: 'table 없음' };
    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
    const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
    const idIdx = headers.findIndex(h => /tutor\s*id/i.test(h));
    const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
    const statuses = rows.map(r => {
      const cells = r.querySelectorAll('td');
      const tid = idIdx >= 0 ? cells[idIdx]?.textContent?.trim() : '?';
      const status = talkIdx >= 0 ? cells[talkIdx]?.textContent?.trim() : '?';
      return `${tid}:${status}`;
    });
    return { found: true, total: rows.length, talkIdx, statuses };
  });
  console.log(`${tag} 검색 결과: ${searchDiag.total || 0}건 [${(searchDiag.statuses || []).join(', ')}]`);

  // 검색 결과에서 시간 매칭 시도
  const result = await iterateSearchResults(page, tag, studentSchedule, weekly, hoursNeeded);
  return result;
}

// --- 검색 결과 순회: 전송하기 우선 → sms 전송됨 (최대 2명) ---
const MAX_PROPOSALS_PER_STUDENT = 2;
const MAX_SEND_ATTEMPTS = 10; // 전송 시도 상한 (실패 포함)

async function iterateSearchResults(page, tag, studentSchedule, weekly, hoursNeeded, genderFilter) {
  let sentCount = 0;
  let totalAttempts = 0;
  for (const targetStatus of ['전송하기', 'sms 전송됨']) {
    let pageNum = 1;
    while (true) {
      const result = await tryMatchOnPage(page, tag, studentSchedule, weekly, hoursNeeded, targetStatus, MAX_PROPOSALS_PER_STUDENT - sentCount, genderFilter);
      if (result.systemError) return 'systemError';
      sentCount += result.sent;
      totalAttempts += result.attempts || 0;
      if (sentCount >= MAX_PROPOSALS_PER_STUDENT) {
        console.log(`${tag} 최대 ${MAX_PROPOSALS_PER_STUDENT}명 제안 완료`);
        return 'matched';
      }
      if (totalAttempts >= MAX_SEND_ATTEMPTS) {
        console.log(`${tag} 전송 시도 ${totalAttempts}회 도달 → 중단`);
        return sentCount > 0 ? 'matched' : 'noMatch';
      }

      // 다음 페이지
      const hasNext = await page.evaluate((pn) => {
        const nextPage = document.querySelector(`.ant-pagination-item-${pn + 1}, li[title="${pn + 1}"]`);
        if (nextPage) { nextPage.click(); return true; }
        return false;
      }, pageNum);

      if (!hasNext) break;
      pageNum++;
      await page.waitForTimeout(1500);
    }
  }
  return sentCount > 0 ? 'matched' : 'noMatch';
}

// --- 모달 전부 닫기 ---
async function closeAllModals(page) {
  for (let i = 0; i < 5; i++) {
    const modalCount = await page.evaluate(() => {
      const modals = document.querySelectorAll('.ant-modal-wrap');
      let visible = 0;
      modals.forEach(m => { if (m.style.display !== 'none') visible++; });
      return visible;
    });
    if (modalCount === 0) break;
    await page.keyboard.press('Escape');
    await page.waitForTimeout(600);
  }
  // 모달 closing 애니메이션 대기
  await page.waitForTimeout(300);
}

// --- 검색 결과 테이블 찾기 (timeline 헤더가 있는 테이블) ---
function getSearchTableScript() {
  return `
    (function getSearchTable() {
      const tables = document.querySelectorAll('table');
      for (const t of tables) {
        const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
        if (headers.some(h => /^timeline$/i.test(h))) return t;
      }
      return null;
    })()
  `;
}

// --- 현재 페이지에서 시간 매칭 시도 ---
async function tryMatchOnPage(page, tag, studentSchedule, weekly, hoursNeeded, targetStatus, maxSend, genderFilter) {
  let sent = 0;
  let attempts = 0;
  // 검색 결과 테이블에서 해당 상태의 행 + 튜터ID + 성별 수집
  const candidates = await page.evaluate((status) => {
    const tables = document.querySelectorAll('table');
    let table = null;
    for (const t of tables) {
      const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
      if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
    }
    if (!table) return [];

    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
    const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
    const idIdx = headers.findIndex(h => /tutor\s*id/i.test(h));
    const genderIdx = headers.findIndex(h => /^gender$/i.test(h));

    const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);

    const result = [];
    rows.forEach((row, i) => {
      const cells = row.querySelectorAll('td');
      if (talkIdx >= 0 && cells.length > talkIdx) {
        const cellText = cells[talkIdx].textContent.trim();
        if (cellText.includes(status)) {
          const tutorId = idIdx >= 0 ? (cells[idIdx]?.textContent?.trim() || '?') : '?';
          const gender = genderIdx >= 0 && cells[genderIdx] ? cells[genderIdx].textContent.trim() : '';
          result.push({ rowIdx: i, tutorId, gender });
        }
      }
    });
    return result;
  }, targetStatus);

  if (candidates.length === 0) {
    // 테이블 헤더 디버그
    const tableDebug = await page.evaluate(() => {
      const tables = document.querySelectorAll('table');
      return [...tables].map((t, i) => {
        const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
        return `table[${i}]: [${headers.join(', ')}]`;
      });
    });
    console.log(`${tag} candidates 0건 — ${tableDebug.join(' | ')}`);
  }

  for (const { rowIdx, tutorId, gender } of candidates) {
    // 성별 필터 (tutorPool 모드에서만 적용)
    if (genderFilter && genderFilter.length > 0 && gender) {
      const genderMap = { 'MALE': '남', 'FEMALE': '여', '남': '남', '여': '여', '남자': '남', '여자': '여' };
      const normalizedGender = genderMap[gender.trim()] || gender.trim();
      if (!genderFilter.includes(normalizedGender)) {
        console.log(`${tag} 튜터 ${tutorId} 성별 불일치 (${gender}→${normalizedGender} vs [${genderFilter}]) → 스킵`);
        continue;
      }
    }

    // timeline 클릭 → 시간표 파싱
    const tutorSchedule = await clickTimelineAndGetSchedule(page, rowIdx);
    if (!tutorSchedule) {
      console.log(`${tag} 튜터 ${tutorId} (행 ${rowIdx}) timeline 파싱 실패`);
      await closeAllModals(page);
      continue;
    }

    console.log(`${tag} 튜터 ${tutorId} 시간표: ${JSON.stringify(tutorSchedule)}`);

    // 시간 비교
    if (matchSchedules(studentSchedule, tutorSchedule, weekly, hoursNeeded)) {
      console.log(`${tag} 시간 매칭 성공 — 튜터 ${tutorId} (행 ${rowIdx}, ${targetStatus})`);

      // timeline 모달 닫기 (ESC + 사라짐 대기)
      await closeAllModals(page);

      // 알림톡 전송
      attempts++;
      const didSend = await sendProposal(page, rowIdx, tag, tutorId);
      if (didSend === 'systemError') {
        return { sent, attempts, systemError: true };
      }
      if (didSend === true) {
        sent++;
        if (sent >= maxSend) return { sent, attempts };
        continue;
      }
      console.log(`${tag} 전송 실패 → 다음 선생님`);
      continue;
    } else {
      console.log(`${tag} 튜터 ${tutorId} 시간 불일치`);
      // 모달 닫기 → 다음 선생님
      await closeAllModals(page);
    }
  }

  return { sent, attempts };
}

// --- timeline 클릭 → 선생님 시간표 파싱 ---
async function clickTimelineAndGetSchedule(page, rowIdx) {
  try {
    // 검색 결과 테이블(timeline 헤더)에서 해당 행의 timeline 셀 클릭
    const clickResult = await page.evaluate((idx) => {
      const tables = document.querySelectorAll('table');
      let table = null;
      for (const t of tables) {
        const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
        if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
      }
      if (!table) return 'table 못 찾음';

      const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
      const timelineIdx = headers.findIndex(h => /^timeline$/i.test(h));
      const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
      if (!rows[idx]) return '행 없음';

      const cell = rows[idx].querySelectorAll('td')[timelineIdx];
      if (!cell) return 'timeline 셀 없음';

      const cellText = cell.textContent.trim();
      // 클릭 가능 요소 찾기
      const clickable = cell.querySelector('a, button, span, div');
      if (clickable) {
        clickable.click();
        return 'clicked:' + cellText;
      }
      cell.click();
      return 'cell-clicked:' + cellText;
    }, rowIdx);

    if (clickResult.includes('못') || clickResult.includes('없음')) {
      console.log(`[timeline] 행 ${rowIdx}: ${clickResult}`);
      return null;
    }
    console.log(`[timeline] 행 ${rowIdx}: ${clickResult}`);

    await page.waitForTimeout(1500);

    // 모달 확인 + 시간표 파싱
    const schedule = await page.evaluate(() => {
      const modal = document.querySelector('.ant-modal');
      if (!modal) return { error: '모달 없음' };
      const text = modal.innerText;

      // "HH:00 ~ HH:00" 패턴으로 시간 파싱
      const days = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일'];
      const result = {};
      for (const day of days) {
        // "화요일\n17:00 ~ 18:00\n21:00 ~ 22:00" 같은 패턴
        const dayIdx = text.indexOf(day);
        if (dayIdx < 0) continue;
        // 해당 요일 뒤 ~ 다음 요일 또는 끝까지 텍스트 추출
        const nextDayIdx = days.reduce((min, d) => {
          if (d === day) return min;
          const idx = text.indexOf(d, dayIdx + day.length);
          return idx > 0 && idx < min ? idx : min;
        }, text.length);
        const dayText = text.substring(dayIdx + day.length, nextDayIdx);
        // "HH:00" 패턴에서 시간 추출
        const hours = [...dayText.matchAll(/(\d{1,2}):00/g)].map(m => parseInt(m[1]));
        // 중복 제거 + "~" 앞의 시작 시간만 사용
        const startHours = [];
        const timeSlots = [...dayText.matchAll(/(\d{1,2}):00\s*~\s*(\d{1,2}):00/g)];
        if (timeSlots.length > 0) {
          timeSlots.forEach(m => startHours.push(parseInt(m[1])));
        } else if (hours.length > 0) {
          // fallback: 모든 시간
          hours.forEach(h => { if (!startHours.includes(h)) startHours.push(h); });
        }
        if (startHours.length > 0) result[day] = startHours.sort((a, b) => a - b);
      }
      return Object.keys(result).length > 0 ? { schedule: result } : { error: '시간 파싱 실패', preview: text.substring(0, 300) };
    });

    if (schedule.error) {
      console.log(`[timeline] 행 ${rowIdx} 파싱 에러: ${schedule.error} | ${(schedule.preview || '').substring(0, 100)}`);
      return null;
    }
    return schedule.schedule;
  } catch (e) {
    console.log(`[timeline] 행 ${rowIdx} 예외: ${e.message}`);
    return null;
  }
}

// --- 매칭 제안 전송 ---
async function checkModalOrError(page) {
  return page.evaluate(() => {
    const modals = document.querySelectorAll('.ant-modal');
    for (const m of modals) {
      const text = m.innerText;
      if (/매칭\s*제안.*할까요|제안을\s*할까요/i.test(text)) return 'proposal';
    }
    const errorSelectors = '.ant-notification, .ant-notification-notice, .ant-message, .ant-message-notice, .ant-message-custom-content, .ant-alert, .ant-modal';
    const errorEls = document.querySelectorAll(errorSelectors);
    for (const el of errorEls) {
      const t = el.textContent;
      if (/orderStatus|paid.*아닙|주문.*상태|order_status/i.test(t)) {
        return 'systemError: ' + t.substring(0, 150).replace(/\n/g, ' ');
      }
    }
    return null;
  });
}

async function sendProposal(page, rowIdx, tag, tutorId) {
  // 1) 잔여 모달 확실히 닫기 (타임라인 모달이 전송하기를 가리는 문제 방지)
  await closeAllModals(page);

  // 2) "전송하기" 버튼의 좌표를 구해서 마우스 클릭 (모달 닫힌 후 좌표 계산)
  const sendBtnBox = await page.evaluate((idx) => {
    const tables = document.querySelectorAll('table');
    let table = null;
    for (const t of tables) {
      const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
      if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
    }
    if (!table) return { error: 'search table 못 찾음' };
    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
    const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
    const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
    if (!rows[idx]) return { error: '행 없음' };

    if (talkIdx >= 0) {
      const cell = rows[idx].querySelectorAll('td')[talkIdx];
      if (cell && /전송하기/i.test(cell.textContent)) {
        const clickable = cell.querySelector('button, a, span[style*="cursor"], [role="button"]') || cell;
        clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
        const rect = clickable.getBoundingClientRect();
        return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
      }
    }
    return { error: '전송하기 못 찾음' };
  }, rowIdx);

  if (sendBtnBox.error) {
    console.log(`${tag} 튜터 ${tutorId} 전송하기 실패: ${sendBtnBox.error}`);
    return false;
  }

  // 디버그: 클릭 좌표에 실제로 뭐가 있는지 확인
  const hitInfo = await page.evaluate(({x, y}) => {
    const el = document.elementFromPoint(x, y);
    if (!el) return 'elementFromPoint: null';
    return `elementFromPoint: <${el.tagName.toLowerCase()}> text="${el.textContent.trim().substring(0, 40)}" class="${el.className.toString().substring(0, 60)}"`;
  }, sendBtnBox);
  console.log(`${tag} 튜터 ${tutorId} 클릭 좌표(${Math.round(sendBtnBox.x)},${Math.round(sendBtnBox.y)}) ${hitInfo}`);

  // 방법1: mouse.click
  await page.mouse.click(sendBtnBox.x, sendBtnBox.y);
  console.log(`${tag} 튜터 ${tutorId} 전송하기 mouse.click 완료`);

  // 300ms 대기 후 모달 확인
  await page.waitForTimeout(300);
  let modalType = await checkModalOrError(page);

  // 방법2: mouse.click 실패 시 evaluate click fallback
  if (!modalType) {
    console.log(`${tag} 튜터 ${tutorId} mouse.click으로 모달 안 열림 → evaluate click 시도`);
    await page.evaluate((idx) => {
      const tables = document.querySelectorAll('table');
      let table = null;
      for (const t of tables) {
        const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
        if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
      }
      if (!table) return;
      const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
      const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
      if (talkIdx < 0) return;
      const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
      if (!rows[idx]) return;
      const cell = rows[idx].querySelectorAll('td')[talkIdx];
      if (!cell) return;
      const clickable = cell.querySelector('button, a, span[style*="cursor"], [role="button"]') || cell;
      clickable.click();
    }, rowIdx);
    await page.waitForTimeout(300);
    modalType = await checkModalOrError(page);
  }

  // 방법3: 그래도 안 되면 dispatchEvent로 시도
  if (!modalType) {
    console.log(`${tag} 튜터 ${tutorId} evaluate click도 실패 → dispatchEvent 시도`);
    await page.evaluate((idx) => {
      const tables = document.querySelectorAll('table');
      let table = null;
      for (const t of tables) {
        const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
        if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
      }
      if (!table) return;
      const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
      const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
      if (talkIdx < 0) return;
      const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
      if (!rows[idx]) return;
      const cell = rows[idx].querySelectorAll('td')[talkIdx];
      if (!cell) return;
      const clickable = cell.querySelector('button, a, span[style*="cursor"], [role="button"]') || cell;
      clickable.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
    }, rowIdx);
    await page.waitForTimeout(500);
    modalType = await checkModalOrError(page);
  }

  // 최종: 그래도 실패 시 좀 더 기다려봄
  if (!modalType) {
    for (let i = 0; i < 5; i++) {
      await page.waitForTimeout(400);
      modalType = await checkModalOrError(page);
      if (modalType) break;
    }
  }

  if (!modalType) {
    console.log(`${tag} 튜터 ${tutorId} 모달/알림 안 열림`);
    await closeAllModals(page);
    return false;
  }

  if (modalType.startsWith('systemError')) {
    console.log(`${tag} 튜터 ${tutorId} ${modalType}`);
    await closeAllModals(page);
    return 'systemError';
  }

  // 3) "Level 상관없이 매칭 제안합니다" 체크 — 1번만 클릭
  const isAlreadyChecked = await page.evaluate(() => {
    const modals = document.querySelectorAll('.ant-modal');
    for (const modal of modals) {
      if (!/매칭\s*제안.*할까요/i.test(modal.innerText)) continue;
      const w = modal.querySelector('.ant-checkbox-wrapper');
      return w ? w.classList.contains('ant-checkbox-wrapper-checked') : false;
    }
    return false;
  });

  if (!isAlreadyChecked) {
    const cbBox = await page.evaluate(() => {
      const modals = document.querySelectorAll('.ant-modal');
      for (const modal of modals) {
        if (!/매칭\s*제안.*할까요/i.test(modal.innerText)) continue;
        const cb = modal.querySelector('.ant-checkbox');
        if (cb) {
          cb.scrollIntoView({ behavior: 'instant', block: 'center' });
          const rect = cb.getBoundingClientRect();
          return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
        }
      }
      return null;
    });
    if (cbBox) {
      await page.mouse.click(cbBox.x, cbBox.y);
      console.log(`${tag} 튜터 ${tutorId} Level 체크박스 클릭`);
      await page.waitForTimeout(800);
    }
  }

  // 4) "매칭 제안하기" 버튼 클릭 (마우스 클릭)
  const proposalBtnBox = await page.evaluate(() => {
    const modals = document.querySelectorAll('.ant-modal');
    for (const modal of modals) {
      if (!/매칭\s*제안.*할까요/i.test(modal.innerText)) continue;
      const btns = modal.querySelectorAll('button');
      for (const btn of btns) {
        if (/매칭\s*제안/i.test(btn.textContent.trim()) && !/닫기|취소/i.test(btn.textContent.trim())) {
          const disabled = btn.disabled;
          btn.scrollIntoView({ behavior: 'instant', block: 'center' });
          const rect = btn.getBoundingClientRect();
          return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, text: btn.textContent.trim(), disabled };
        }
      }
    }
    return null;
  });

  if (!proposalBtnBox) {
    console.log(`${tag} 튜터 ${tutorId} "매칭 제안하기" 버튼 못 찾음`);
    await closeAllModals(page);
    return false;
  }

  if (proposalBtnBox.disabled) {
    console.log(`${tag} 튜터 ${tutorId} "${proposalBtnBox.text}" 버튼 disabled → 체크박스 미체크?`);
    await closeAllModals(page);
    return false;
  }

  await page.mouse.click(proposalBtnBox.x, proposalBtnBox.y);
  console.log(`${tag} 튜터 ${tutorId} "${proposalBtnBox.text}" 클릭`);

  // 5) 결과 대기 — 모달 닫힘 + 에러 알림 체크
  await page.waitForTimeout(2000);

  // 시스템 에러 체크 (notification, message, modal 모두)
  const errorCheck = await page.evaluate(() => {
    const text = document.body.innerText;
    // ant-notification
    const notifs = document.querySelectorAll('.ant-notification, .ant-message, .ant-notification-notice');
    for (const n of notifs) {
      const t = n.textContent;
      if (/orderStatus|paid|주문.*상태/i.test(t)) return 'notification: ' + t.substring(0, 100);
    }
    // 새로 뜬 모달
    const modals = document.querySelectorAll('.ant-modal');
    for (const m of modals) {
      const mt = m.innerText;
      if (/orderStatus|paid.*아닙|주문.*상태/i.test(mt)) return 'modal: ' + mt.substring(0, 100);
    }
    // ant-message-notice (상단 메시지 바)
    const msgs = document.querySelectorAll('.ant-message-notice, .ant-message-custom-content');
    for (const m of msgs) {
      const t = m.textContent;
      if (/orderStatus|paid|에러|error|실패/i.test(t)) return 'message: ' + t.substring(0, 100);
    }
    return null;
  });

  if (errorCheck) {
    console.log(`${tag} 튜터 ${tutorId} 시스템 에러 감지: ${errorCheck}`);
    await closeAllModals(page);
    return 'systemError';
  }

  // 성공 판정: 모달 닫힘 OR 테이블에서 해당 튜터 상태가 "전송하기"가 아닌 값으로 변경
  const modalGone = await page.evaluate(() => {
    const modals = document.querySelectorAll('.ant-modal');
    return ![...modals].some(m => /매칭\s*제안.*할까요/i.test(m.innerText));
  });

  if (modalGone) {
    console.log(`${tag} 튜터 ${tutorId} 매칭 제안 완료 (모달 닫힘 확인)`);
    return true;
  }

  // 모달 안 닫혔어도 실제 발송됐을 수 있음 → 테이블 상태 확인
  await closeAllModals(page);
  await page.waitForTimeout(1000);

  const statusAfter = await page.evaluate((idx) => {
    const tables = document.querySelectorAll('table');
    let table = null;
    for (const t of tables) {
      const headers = [...t.querySelectorAll('thead th')].map(h => h.textContent.trim());
      if (headers.some(h => /^timeline$/i.test(h))) { table = t; }
    }
    if (!table) return null;
    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
    const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
    if (talkIdx < 0) return null;
    const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
    if (!rows[idx]) return null;
    return rows[idx].querySelectorAll('td')[talkIdx]?.textContent?.trim() || null;
  }, rowIdx);

  if (statusAfter && !/전송하기/i.test(statusAfter)) {
    console.log(`${tag} 튜터 ${tutorId} 매칭 제안 완료 (테이블 상태: "${statusAfter}")`);
    return true;
  }

  console.log(`${tag} 튜터 ${tutorId} 제안 실패 (테이블 상태: "${statusAfter}")`);
  return false;
}

// --- 특정 튜터풀 매칭 ---
async function tryTutorPoolMatch(page, tag, studentSchedule, weekly, hoursNeeded, tutorIdsStr, studentGenders) {
  // "이름 및 ��이디" 영역의 ID input 찾기 → focus → clipboard paste
  const inputBox = await page.evaluate(() => {
    const all = [...document.querySelectorAll('*')];
    for (const el of all) {
      if (el.children.length === 0 && /이름\s*(및|&)\s*아이디/i.test(el.textContent.trim())) {
        let parent = el.parentElement;
        for (let d = 0; d < 8 && parent; d++) {
          const inputs = parent.querySelectorAll('input[type="text"], input:not([type])');
          if (inputs.length >= 2) {
            const inp = inputs[1]; // ID 필드
            inp.scrollIntoView({ behavior: 'instant', block: 'center' });
            inp.focus();
            inp.click();
            const rect = inp.getBoundingClientRect();
            const placeholders = [...inputs].map(i => i.getAttribute('placeholder') || '?');
            return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, placeholders };
          }
          parent = parent.parentElement;
        }
        return { error: 'inputs not found' };
      }
    }
    return { error: 'label not found' };
  });

  if (inputBox.error) {
    console.log(`${tag} 튜터풀 ID 입력 실패: ${inputBox.error}`);
    return 'noMatch';
  }

  // input 클릭 → 전체 선택 → 클립보드 붙여넣기
  await page.mouse.click(inputBox.x, inputBox.y);
  await page.waitForTimeout(200);
  await page.keyboard.press('Control+A');
  await page.evaluate((text) => { navigator.clipboard.writeText(text); }, tutorIdsStr);
  await page.waitForTimeout(100);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(500);

  // 입력 확인
  const inputValue = await page.evaluate(() => {
    const el = document.activeElement;
    return el ? el.value?.substring(0, 50) + '...' : 'no focus';
  });
  console.log(`${tag} 튜터풀 ID 입���: placeholders=[${inputBox.placeholders}] value="${inputValue}"`);

  if (!inputResult.startsWith('ok')) return 'noMatch';

  // "이름 및 아이디" 영역의 Search 버튼 클릭 (mouse.click)
  const searchBtnBox = await page.evaluate(() => {
    const all = [...document.querySelectorAll('*')];
    for (const el of all) {
      if (el.children.length === 0 && /이름\s*(및|&)\s*아이디/i.test(el.textContent.trim())) {
        let parent = el.parentElement;
        for (let d = 0; d < 8 && parent; d++) {
          const btns = parent.querySelectorAll('button');
          for (const btn of btns) {
            const t = btn.textContent.trim();
            if (/search|검색|조회/i.test(t) || btn.querySelector('.anticon-search') || btn.querySelector('svg')) {
              btn.scrollIntoView({ behavior: 'instant', block: 'center' });
              const rect = btn.getBoundingClientRect();
              return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, text: t || 'icon-btn' };
            }
          }
          parent = parent.parentElement;
        }
        // 디버그
        let p = el.parentElement;
        for (let d = 0; d < 8 && p; d++) {
          const btns = p.querySelectorAll('button');
          if (btns.length > 0) {
            return { error: 'not found — nearby: [' + [...btns].map(b => `"${b.textContent.trim()}"`).join(', ') + ']' };
          }
          p = p.parentElement;
        }
        return { error: 'no buttons nearby' };
      }
    }
    return { error: 'label not found' };
  });

  if (searchBtnBox.error) {
    console.log(`${tag} 튜터풀 Search: ${searchBtnBox.error}`);
    return 'noMatch';
  }

  await page.mouse.click(searchBtnBox.x, searchBtnBox.y);
  console.log(`${tag} 튜터풀 Search 클릭: "${searchBtnBox.text}"`);
  await page.waitForTimeout(5000);

  // 검색 결과 진단
  const searchDiag = await page.evaluate(() => {
    const tables = document.querySelectorAll('table');
    const table = tables[tables.length - 1];
    if (!table) return { total: 0 };
    const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
    const talkIdx = headers.findIndex(h => /알림톡\s*전송/i.test(h));
    const idIdx = headers.findIndex(h => /tutor\s*id/i.test(h));
    const rows = [...table.querySelectorAll('tbody tr')].filter(r => r.querySelectorAll('td').length > 1);
    const statuses = rows.map(r => {
      const cells = r.querySelectorAll('td');
      const tid = idIdx >= 0 ? cells[idIdx]?.textContent?.trim() : '?';
      const status = talkIdx >= 0 ? cells[talkIdx]?.textContent?.trim() : '?';
      return `${tid}:${status}`;
    });
    return { total: rows.length, statuses };
  });
  console.log(`${tag} 튜터풀 검색 결과: ${searchDiag.total}건 [${(searchDiag.statuses || []).join(', ')}]`);

  // 시간 매칭 시도 (성별 필터 적용)
  const result = await iterateSearchResults(page, tag, studentSchedule, weekly, hoursNeeded, studentGenders);
  return result;
}

// --- 대학교/전공 필터 변형 적용 (교대·메디컬 등) ---
async function applyUniMajorFilter(page, tag, config) {
  // config: { uncheckUni: '교대', addUni: ['SKY','서성한','중경외시'], addMajor: ['교육'] }
  const result = await page.evaluate((cfg) => {
    const log = [];
    const allEls = [...document.querySelectorAll('*')];

    // 대학교 그룹 찾기
    let uniGroup = null, majorGroup = null;
    for (const el of allEls) {
      if (el.children.length === 0 && el.tagName === 'H5') {
        if (el.textContent.trim() === '대학교') uniGroup = el;
        if (el.textContent.trim() === '전공') majorGroup = el;
      }
    }

    // 대학교 체크박스 조작
    if (uniGroup) {
      let parent = uniGroup.parentElement;
      for (let d = 0; d < 5 && parent; d++) {
        const cbs = parent.querySelectorAll('.ant-checkbox-wrapper');
        if (cbs.length >= 3) {
          for (const cb of cbs) {
            const text = cb.textContent.trim();
            const isChecked = cb.classList.contains('ant-checkbox-wrapper-checked');
            const input = cb.querySelector('input') || cb;
            // 해제 대상
            if (text === cfg.uncheckUni && isChecked) {
              input.click();
              log.push(`${text} 해제`);
            }
            // 추가 대상
            if (cfg.addUni.includes(text) && !isChecked) {
              input.click();
              log.push(`${text} 체크`);
            }
          }
          break;
        }
        parent = parent.parentElement;
      }
    }

    // 전공 체크박스 조작
    if (majorGroup && cfg.addMajor) {
      let parent = majorGroup.parentElement;
      for (let d = 0; d < 5 && parent; d++) {
        const cbs = parent.querySelectorAll('.ant-checkbox-wrapper');
        if (cbs.length >= 3) {
          for (const cb of cbs) {
            const text = cb.textContent.trim();
            const isChecked = cb.classList.contains('ant-checkbox-wrapper-checked');
            if (cfg.addMajor.includes(text) && !isChecked) {
              const input = cb.querySelector('input') || cb;
              input.click();
              log.push(`전공>${text} 체크`);
            }
          }
          break;
        }
        parent = parent.parentElement;
      }
    }

    return log.length > 0 ? log.join(', ') : '변경 없음';
  }, config);
  return result;
}

// --- 교대 단독 선택 시 대학교 + 전공 필터 보정 ---
async function fixGyodaeFilter(page, tag) {
  const result = await page.evaluate(() => {
    // 대학교 체크박스 그룹 찾기
    const allEls = [...document.querySelectorAll('*')];
    let uniGroup = null;
    let majorGroup = null;
    for (const el of allEls) {
      if (el.children.length === 0 && el.tagName === 'H5') {
        if (el.textContent.trim() === '대학교') uniGroup = el;
        if (el.textContent.trim() === '전공') majorGroup = el;
      }
    }
    if (!uniGroup) return null;

    // 대학교 그룹의 체크박스 상태 확인
    let uniParent = uniGroup.parentElement;
    for (let d = 0; d < 5 && uniParent; d++) {
      const cbs = uniParent.querySelectorAll('.ant-checkbox-wrapper');
      if (cbs.length >= 3) {
        const items = [...cbs].map(cb => ({
          text: cb.textContent.trim(),
          checked: cb.classList.contains('ant-checkbox-wrapper-checked'),
          el: cb,
        }));

        // 교대만 단독 체크인지 확인
        const checkedItems = items.filter(i => i.checked);
        if (checkedItems.length !== 1 || checkedItems[0].text !== '교대') return null;

        // SKY, 서성한, 중경외시 추가 체크
        const toAdd = ['SKY', '서성한', '중경외시'];
        const added = [];
        for (const name of toAdd) {
          const item = items.find(i => i.text === name && !i.checked);
          if (item) {
            const input = item.el.querySelector('input') || item.el;
            input.click();
            added.push(name);
          }
        }

        // 전공 > 교육 체크
        let eduAdded = false;
        if (majorGroup) {
          let majorParent = majorGroup.parentElement;
          for (let md = 0; md < 5 && majorParent; md++) {
            const majorCbs = majorParent.querySelectorAll('.ant-checkbox-wrapper');
            if (majorCbs.length >= 3) {
              for (const mcb of majorCbs) {
                if (mcb.textContent.trim() === '교육' && !mcb.classList.contains('ant-checkbox-wrapper-checked')) {
                  const input = mcb.querySelector('input') || mcb;
                  input.click();
                  eduAdded = true;
                }
              }
              break;
            }
            majorParent = majorParent.parentElement;
          }
        }

        return `대학교 추가: [${added.join(', ')}], 전공 교육: ${eduAdded ? '추가' : '이미 체크 또는 못 찾음'}`;
      }
      uniParent = uniParent.parentElement;
    }
    return null;
  });
  return result;
}

// --- AIM 실패 시 tutorPairingStatus → MATCHING 변경 ---
async function fixPairingStatus(page, tag) {
  try {
    // 현재 URL에서 pairingId 추출 → /admin/tutor-pairing/{pairingId}/update로 직접 이동
    const currentUrl = await page.url();
    const pairingIdMatch = currentUrl.match(/\/tutor-pairing\/(\d+)/);
    if (!pairingIdMatch) {
      console.log(`${tag} URL에서 pairingId 추출 실패: ${currentUrl}`);
      return false;
    }
    const pairingId = pairingIdMatch[1];
    const updateUrl = `https://tutor-admin.qanda.ai/admin/tutor-pairing/${pairingId}/update`;
    console.log(`${tag} Tutor Pairing Update 이동: ${updateUrl}`);

    await page.goto(updateUrl, { waitUntil: 'domcontentloaded', timeout: 15000 });
    await page.waitForFunction(
      () => document.body.innerText.includes('tutorPairingStatus'),
      { timeout: 10000 }
    );
    await page.waitForTimeout(500);

    // 4) 스크롤 + select 클릭
    await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === 'tutorPairingStatus') {
          el.scrollIntoView({ behavior: 'instant', block: 'center' });
          return;
        }
      }
    });
    await page.waitForTimeout(300);

    const box = await page.evaluate(() => {
      const all = [...document.querySelectorAll('label, span, div')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === 'tutorPairingStatus') {
          const formItem = el.closest('.ant-form-item') || el.parentElement?.parentElement;
          if (!formItem) continue;
          const selector = formItem.querySelector('.ant-select-selector');
          if (selector) {
            const rect = selector.getBoundingClientRect();
            return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
          }
        }
      }
      return null;
    });

    if (!box) {
      console.log(`${tag} tutorPairingStatus select 못 찾음`);
      return false;
    }

    await page.mouse.click(box.x, box.y);
    await page.waitForTimeout(500);

    // 5) MATCHING 옵션 선택
    const selected = await page.evaluate(() => {
      const opts = document.querySelectorAll('.ant-select-item-option');
      for (const opt of opts) {
        const title = opt.getAttribute('title') || opt.textContent.trim();
        if (title === 'MATCHING') {
          opt.click();
          return 'selected';
        }
      }
      const allOpts = [...document.querySelectorAll('.ant-select-item-option')].map(o => o.getAttribute('title') || o.textContent.trim());
      return `not found (options: [${allOpts.join(', ')}])`;
    });
    console.log(`${tag} tutorPairingStatus → MATCHING: ${selected}`);

    if (!selected.startsWith('selected')) return false;

    await page.waitForTimeout(300);

    // 6) Submit 클릭
    const submitClicked = await page.evaluate(() => {
      const btns = [...document.querySelectorAll('button')];
      for (const btn of btns) {
        if (/^Submit$/i.test(btn.textContent.trim())) {
          btn.click();
          return true;
        }
      }
      return false;
    });
    if (!submitClicked) {
      console.log(`${tag} Submit 버튼 못 찾음`);
      return false;
    }
    console.log(`${tag} Submit 클릭`);

    await page.waitForTimeout(500);

    // 7) Yes 확인 모달
    const yesClicked = await page.evaluate(() => {
      const btns = document.querySelectorAll('.ant-modal button, .ant-popconfirm button');
      for (const btn of btns) {
        if (/^Yes$/i.test(btn.textContent.trim())) { btn.click(); return true; }
      }
      return false;
    });
    console.log(`${tag} Yes 클릭: ${yesClicked}`);

    await page.waitForTimeout(2000);
    return true;
  } catch (e) {
    console.log(`${tag} fixPairingStatus 에러: ${e.message}`);
    return false;
  }
}

// --- 시간표 자동 확장 + 어드민 수정 ---
async function fixAdminSchedule(page, studentSchedule, hoursNeeded, tag) {
  const dayKeyToFull = { mon: 'monday', tue: 'tuesday', wed: 'wednesday', thu: 'thursday', fri: 'friday', sat: 'saturday', sun: 'sunday' };

  // 각 요일별로 시작 시간부터 hoursNeeded만큼 확장, 추가해야 할 시간 계산
  const toAdd = {}; // { wednesday: [22], friday: [22], ... }
  const expandedSchedule = {}; // 최종 스케줄 (short key)
  for (const [shortDay, hours] of Object.entries(studentSchedule)) {
    const fullDay = dayKeyToFull[shortDay];
    const startHour = Math.min(...hours);
    const neededHours = [];
    for (let h = startHour; h < startHour + hoursNeeded; h++) {
      neededHours.push(h);
    }
    const missingHours = neededHours.filter(h => !hours.includes(h));
    if (missingHours.length > 0) {
      toAdd[fullDay] = missingHours;
    }
    expandedSchedule[shortDay] = neededHours;
  }

  if (Object.keys(toAdd).length === 0) {
    return { ok: true, expandedSchedule };
  }

  console.log(`${tag} 시간 추가 필요: ${JSON.stringify(toAdd)}`);

  try {
    // "신청 내용" Edit 클릭
    const editClicked = await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && /^신청\s*내용$/.test(el.textContent.trim())) {
          let parent = el.parentElement;
          for (let d = 0; d < 5 && parent; d++) {
            const btn = parent.querySelector('button');
            if (btn && /edit/i.test(btn.textContent)) { btn.click(); return true; }
            parent = parent.parentElement;
          }
        }
      }
      return false;
    });
    if (!editClicked) return { ok: false, reason: 'Edit 버튼 못 찾음' };

    // Edit 폼 로드 대기
    await page.waitForFunction(
      () => document.querySelector('label[for="weeklyAvailablePeriods.monday"]'),
      { timeout: 10000 }
    );
    await page.waitForTimeout(500);

    // 각 요일별로 빠진 시간 추가
    for (const [fullDay, hours] of Object.entries(toAdd)) {
      for (const h of hours) {
        const optionText = `${h}:00 ~ ${h + 1}:00`;
        console.log(`${tag} ${fullDay} "${optionText}" 추가 시도`);

        // 1) 스크롤해서 해당 요일 select를 뷰포트에 보이게
        await page.evaluate((day) => {
          const label = document.querySelector(`label[for="weeklyAvailablePeriods.${day}"]`);
          if (label) label.scrollIntoView({ behavior: 'instant', block: 'center' });
        }, fullDay);
        await page.waitForTimeout(300);

        // 2) select-selector의 좌표를 구해서 마우스 클릭 (드롭다운 열기)
        const box = await page.evaluate((day) => {
          const label = document.querySelector(`label[for="weeklyAvailablePeriods.${day}"]`);
          const formItem = label?.closest('.ant-form-item');
          const selector = formItem?.querySelector('.ant-select-selector');
          if (!selector) return null;
          const rect = selector.getBoundingClientRect();
          return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
        }, fullDay);

        if (!box) {
          console.log(`${tag} ${fullDay}: select 못 찾음`);
          continue;
        }

        await page.mouse.click(box.x, box.y);
        await page.waitForTimeout(500);

        // 3) 시간 타이핑으로 검색 필터
        await page.keyboard.type(String(h), { delay: 30 });
        await page.waitForTimeout(800);

        // 4) 드롭다운에서 옵션 클릭
        const selected = await page.evaluate((targetText) => {
          const opts = document.querySelectorAll('.ant-select-item-option');
          for (const opt of opts) {
            const title = opt.getAttribute('title') || opt.textContent.trim();
            if (title === targetText && !opt.classList.contains('ant-select-item-option-selected')) {
              opt.click();
              return 'selected';
            }
            if (title === targetText) return 'already selected';
          }
          return 'not found';
        }, optionText);

        console.log(`${tag} ${fullDay} "${optionText}": ${selected}`);

        // 5) 드롭다운 닫기
        await page.keyboard.press('Escape');
        await page.waitForTimeout(300);
      }
    }

    // Submit 클릭
    const submitClicked = await page.evaluate(() => {
      const btns = [...document.querySelectorAll('button')];
      for (const btn of btns) {
        if (/^Submit$/i.test(btn.textContent.trim())) {
          btn.click();
          return true;
        }
      }
      return false;
    });
    if (!submitClicked) return { ok: false, reason: 'Submit 버튼 못 찾음', expandedSchedule };
    console.log(`${tag} Submit 클릭`);

    await page.waitForTimeout(500);

    // Yes 확인 모달
    const yesClicked = await page.evaluate(() => {
      const btns = document.querySelectorAll('.ant-modal button, .ant-popconfirm button');
      for (const btn of btns) {
        if (/^Yes$/i.test(btn.textContent.trim())) { btn.click(); return true; }
      }
      return false;
    });
    console.log(`${tag} Yes 클릭: ${yesClicked}`);

    await page.waitForTimeout(2000);

    return { ok: true, expandedSchedule };
  } catch (e) {
    console.log(`${tag} fixAdminSchedule 에러: ${e.message}`);
    return { ok: false, reason: e.message };
  }
}

// --- 3단계: 조건 완화 적용 ---
async function applyRelaxation(page, level, tag) {
  if (level === 'relax1') {
    const result = await page.evaluate(() => {
      const log = [];
      // 입시 전형: 정시 체크
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === '입시 전형') {
          let parent = el.parentElement;
          for (let d = 0; d < 8 && parent; d++) {
            const cbs = parent.querySelectorAll('.ant-checkbox-wrapper');
            if (cbs.length > 0) {
              cbs.forEach(cb => {
                const t = cb.textContent.trim();
                const checked = cb.classList.contains('ant-checkbox-wrapper-checked');
                if (!checked) { (cb.querySelector('input') || cb).click(); log.push(t + ' 추가'); }
              });
              break;
            }
            parent = parent.parentElement;
          }
          break;
        }
      }
      // 튜터 스타일 clear — X 버튼으로 개별 항목 제거
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === '튜터 스타일') {
          let parent = el.parentElement;
          for (let d = 0; d < 8 && parent; d++) {
            const select = parent.querySelector('.ant-select');
            if (select) {
              const removes = select.querySelectorAll('.ant-select-selection-item-remove');
              removes.forEach(r => r.click());
              log.push(`튜터스타일: ${removes.length}개 항목 제거`);
              break;
            }
            parent = parent.parentElement;
          }
          break;
        }
      }
      return log;
    });
    console.log(`${tag} 완화1: ${result.join(', ')}`);

  } else if (level === 'relax2') {
    const result = await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === '성적대') {
          let parent = el.parentElement;
          for (let d = 0; d < 8 && parent; d++) {
            const select = parent.querySelector('.ant-select');
            if (select) {
              const removes = select.querySelectorAll('.ant-select-selection-item-remove');
              removes.forEach(r => r.click());
              return `${removes.length}개 항목 제거`;
            }
            parent = parent.parentElement;
          }
        }
      }
      return 'select 못 찾음';
    });
    console.log(`${tag} 완화2 성적대: ${result}`);

  } else if (level === 'relax3') {
    // 1) 고3 가능 과목 값 읽기 + 제거
    const subjectValue = await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === '고3 가능 과목') {
          let parent = el.parentElement;
          for (let d = 0; d < 8 && parent; d++) {
            const select = parent.querySelector('.ant-select');
            if (select) {
              const items = select.querySelectorAll('.ant-select-selection-item');
              const values = [...items].map(i => i.getAttribute('title') || i.textContent.trim());
              const removes = select.querySelectorAll('.ant-select-selection-item-remove');
              removes.forEach(r => r.click());
              return values;
            }
            parent = parent.parentElement;
          }
        }
      }
      return [];
    });
    console.log(`${tag} 완화3 고3과목 제거: [${subjectValue.join(', ')}]`);

    // 2) 과목 드롭다운에 같은 값 추가 — "고3 가능 과목" 바로 위의 "과목" select 찾기
    for (const val of subjectValue) {
      // "고3 가능 과목" 라벨의 위치를 기준으로 바로 위에 있는 "과목" select 찾기
      await page.evaluate(() => {
        const all = [...document.querySelectorAll('*')];
        let go3Y = 99999;
        // 고3 가능 과목 위치
        for (const el of all) {
          if (el.children.length === 0 && el.textContent.trim() === '고3 가능 과목') {
            go3Y = el.getBoundingClientRect().y;
            break;
          }
        }
        // 고3 가능 과목 바로 아래에 있는 "과목" 라벨 중 가장 가까운 것
        let closest = null;
        let closestDist = 99999;
        for (const el of all) {
          if (el.children.length === 0 && /^과목$/.test(el.textContent.trim())) {
            const y = el.getBoundingClientRect().y;
            const dist = y - go3Y;
            if (dist > 0 && dist < closestDist) {
              closestDist = dist;
              closest = el;
            }
          }
        }
        if (closest) closest.scrollIntoView({ behavior: 'instant', block: 'center' });
      });
      await page.waitForTimeout(300);

      // 해당 "과목" 라벨 근처의 select-selector 좌표
      const box = await page.evaluate(() => {
        const all = [...document.querySelectorAll('*')];
        let go3Y = 99999;
        for (const el of all) {
          if (el.children.length === 0 && el.textContent.trim() === '고3 가능 과목') {
            go3Y = el.getBoundingClientRect().y;
            break;
          }
        }
        let closest = null;
        let closestDist = 99999;
        for (const el of all) {
          if (el.children.length === 0 && /^과목$/.test(el.textContent.trim())) {
            const y = el.getBoundingClientRect().y;
            const dist = y - go3Y;
            if (dist > 0 && dist < closestDist) {
              closestDist = dist;
              closest = el;
            }
          }
        }
        if (!closest) return null;
        let parent = closest.parentElement;
        for (let d = 0; d < 8 && parent; d++) {
          const select = parent.querySelector('.ant-select');
          if (select) {
            const selector = select.querySelector('.ant-select-selector');
            if (selector) {
              const rect = selector.getBoundingClientRect();
              return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
            }
          }
          parent = parent.parentElement;
        }
        return null;
      });

      if (!box) {
        console.log(`${tag} 완화3 과목 select 못 찾음`);
        break;
      }

      await page.mouse.click(box.x, box.y);
      await page.waitForTimeout(800);

      // 검색 타이핑으로 필터
      await page.keyboard.type(val.substring(0, 4), { delay: 30 });
      await page.waitForTimeout(800);

      // 드롭다운에서 옵션 선택
      const selected = await page.evaluate((targetVal) => {
        const options = document.querySelectorAll('.ant-select-item-option');
        for (const opt of options) {
          const optText = (opt.getAttribute('title') || opt.textContent).trim();
          if (optText === targetVal && !opt.classList.contains('ant-select-item-option-selected')) {
            opt.click();
            return 'selected';
          }
          if (optText === targetVal) return 'already';
        }
        const allOpts = [...options].map(o => (o.getAttribute('title') || o.textContent).trim());
        return `not found (${allOpts.length}개: [${allOpts.slice(0, 5).join(', ')}])`;
      }, val);
      console.log(`${tag} 완화3 과목 추가 "${val}": ${selected}`);

      await page.keyboard.press('Escape');
      await page.waitForTimeout(300);
    }
  }
  await page.waitForTimeout(300);
}

// --- 외부 API ---
async function runAim(mode) {
  if (mode === '1' || mode === 'new') {
    await runAimForType('new');
  } else if (mode === '2' || mode === 'rematch') {
    await runAimForType('rematch');
  } else {
    await runAimForType('new');
    await runAimForType('rematch');
  }
}

module.exports = { runAim };

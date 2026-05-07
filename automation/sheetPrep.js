const axios = require('axios');
const { newPage, newBackgroundPage } = require('./browser');
const config = require('../config.json');
const { getSheetsApi } = require('../sheets');

// GAS 웹앱으로 Connected Sheets 데이터 소스 새로고침 요청
// 데이터 소스가 없거나 실패해도 예외 던지지 않고 로그만 남김
async function refreshViaGas(spreadsheetId) {
  try {
    console.log('[sheetPrep] 데이터 소스 새로고침 요청...');
    const res = await axios.get(config.gas.url, {
      params: { action: 'refresh', sheetId: spreadsheetId },
      timeout: 120000, // 2분
    });
    const data = res.data || {};
    if (data.ok) {
      console.log(`[sheetPrep] 새로고침 완료 (${data.refreshed}/${data.total})`);
    } else if (data.error === 'no data sources found') {
      console.log('[sheetPrep] 데이터 소스 없음 — 스킵');
    } else {
      console.log(`[sheetPrep] 새로고침 실패: ${data.error || 'unknown'}`);
    }
  } catch (e) {
    console.log('[sheetPrep] 새로고침 요청 에러 — 기존 데이터로 진행:', e.message);
  }
}

const REMATCH_ADMIN_URL = config.adminUrl + '/admin/tutor-pairing?page=1&size=100&filters_rematch=true&filters_tutorPairingStatusList=READY%2CMATCHING%2CAIM_ING%2CUNPAIRED&sort=lastLessonDate%2CASC';

// 현재 월 탭 이름 생성: [신규] 25.04, [재매칭] 25.04
function getMonthlyTabName(type) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  return type === 'new' ? `[신규] ${yy}.${mm}` : `[재매칭] ${yy}.${mm}`;
}

// 오늘 날짜 문자열: "26. 04. 07"
function getTodayString() {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  return `${yy}. ${mm}. ${dd}`;
}

// 월별 탭에서 제외 상태가 아닌 행의 A열에 오늘 날짜 쓰기 (헤더 제외)
// type: 'new' | 'rematch' — 제외 상태와 상태 컬럼이 다름
async function stampDateColumnA(sheets, spreadsheetId, tabName, type) {
  const today = getTodayString();

  // 헤더 읽어서 상태 컬럼 위치 찾기
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabName}'!A1:BZ1`,
  });
  const headerRow = (headerRes.data.values || [])[0] || [];
  console.log(`[sheetPrep] 헤더: ${JSON.stringify(headerRow)}`);

  const statusColName = type === 'new' ? '매칭상태' : 'status';
  // 공백 무시하고 매칭
  const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
  const statusIdx = headerRow.findIndex(h => norm(h) === norm(statusColName));
  console.log(`[sheetPrep] 상태 컬럼 "${statusColName}" 위치: ${statusIdx}`);

  // 공백 정규화 후 비교 (시트 값 변동성 대응)
  const excludeNew = ['매칭완료', '동일튜터유지', '환불'];
  const excludeRematch = ['재매칭 완료', '재매칭완료', '환불/소멸', '환불', '소멸'];
  const excludeList = (type === 'new' ? excludeNew : excludeRematch).map(norm);

  // 데이터 읽기
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabName}'!A2:BZ`,
  });
  const allRows = res.data.values || [];

  // 규칙:
  //  - match_id(B열) 없는 행 → 빈 값 (유령 행 정리)
  //  - 완료 상태(매칭완료/동일튜터유지/환불 등) → 기존 A열 값 그대로 유지
  //  - 미완료 행 → 오늘 날짜로 갱신
  let updatedCount = 0;
  let keptCount = 0;
  let clearedCount = 0;
  const seenStatuses = new Set();
  const dateValues = allRows.map(row => {
    const currentA = row[0] || '';
    const matchId = (row[1] || '').trim(); // B열
    if (!matchId) {
      // 유령 행: A열에만 날짜가 있는 등 → 빈 값으로 정리
      if (currentA) clearedCount++;
      return [''];
    }

    if (statusIdx >= 0) {
      const rawStatus = (row[statusIdx] || '').trim();
      seenStatuses.add(rawStatus);
      if (excludeList.includes(norm(rawStatus))) {
        keptCount++;
        return [currentA]; // 완료 상태 → 기존 값 유지
      }
    }
    // 미완료 → 오늘 날짜로 갱신
    if (currentA !== today) updatedCount++;
    return [today];
  });
  console.log(`[sheetPrep] 시트에서 발견된 상태값: ${JSON.stringify([...seenStatuses])}`);

  if (dateValues.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabName}'!A2:A${allRows.length + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: dateValues },
    });
    console.log(`[sheetPrep] A열: ${updatedCount}행 갱신 → "${today}", ${keptCount}행 유지(완료), ${clearedCount}행 정리(유령)`);
  }
}

// ============================================================
// 신규: [신규]list 쿼리 탭 → [신규] yy.mm 월별 탭
// ============================================================
async function prepNewSheet() {
  console.log('\n[sheetPrep] === 신규 시트 정리 시작 ===');
  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets.new;
  const queryTab = '[신규]list';
  const monthlyTab = getMonthlyTabName('new');

  // 1) GAS 웹앱을 통해 Connected Sheets/BigQuery 새로고침
  await refreshViaGas(spreadsheetId);

  // 2) Sheets API로 쿼리 탭 데이터 읽기 (헤더 제외)
  const queryRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${queryTab}'!A2:BZ`,
  });
  const queryRows = (queryRes.data.values || []).filter(row =>
    row.some(cell => cell && cell.trim() !== '')
  );
  console.log(`[sheetPrep] 쿼리 탭 데이터: ${queryRows.length}행`);

  if (queryRows.length === 0) {
    console.log('[sheetPrep] 쿼리 탭에 데이터 없음, 종료');
    return;
  }

  // 2) 월별 탭 기존 데이터 읽기
  let existingRows = [];
  try {
    const monthlyRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${monthlyTab}'!A2:BZ`,
    });
    existingRows = monthlyRes.data.values || [];
  } catch (e) {
    console.log(`[sheetPrep] 월별 탭 '${monthlyTab}' 없음 또는 비어있음`);
  }

  // 3) 기존 match_id 수집 + B열(match_id) 있는 마지막 행 찾기
  const existingIds = new Set();
  let lastDataRowIdx = -1; // 0-indexed within existingRows
  existingRows.forEach((row, i) => {
    const id = (row[1] || '').trim(); // B열
    if (id) {
      existingIds.add(id);
      lastDataRowIdx = i;
    }
  });
  console.log(`[sheetPrep] 기존 월별 탭: ${existingRows.length}행 (실제 데이터 ${existingIds.size}개, 마지막 데이터 행 ${lastDataRowIdx + 2})`);

  // 4) 중복 제거: 기존에 있는 match_id는 추가하지 않음
  const today = getTodayString();
  const newRows = queryRows
    .filter(row => {
      // list 탭은 A열부터 데이터 시작 → match_id는 row[0]
      const matchId = (row[0] || '').trim();
      return matchId && !existingIds.has(matchId);
    })
    .map(row => [today, ...row]); // A열에 날짜 추가, 원본 데이터는 한 칸 밀어서 B열부터
  console.log(`[sheetPrep] 신규 추가 대상: ${newRows.length}행`);

  if (newRows.length > 0) {
    // 5) 실제 데이터 마지막 행 다음에 update로 직접 쓰기 (유령 행 무시)
    const startRow = lastDataRowIdx + 3; // +1(0→1) +1(헤더) +1(다음 행)
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${monthlyTab}'!A${startRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: newRows },
    });
    console.log(`[sheetPrep] ${newRows.length}행 추가 완료 (행 ${startRow}~${startRow + newRows.length - 1})`);
  } else {
    console.log('[sheetPrep] 추가할 신규 데이터 없음');
  }

  // 6) 전체 비공백 행 A열에 오늘 날짜 쓰기
  await stampDateColumnA(sheets, spreadsheetId, monthlyTab, 'new');
  console.log('[sheetPrep] === 신규 시트 정리 완료 ===');
}

// ============================================================
// 재매칭: 어드민 페이지 copy clipboard → [재매칭] yy.mm 월별 탭
// ============================================================
async function prepRematchSheet() {
  console.log('\n[sheetPrep] === 재매칭 시트 정리 시작 ===');
  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets.rematch;
  const monthlyTab = getMonthlyTabName('rematch');

  // 0) GAS 웹앱으로 데이터 소스 새로고침 (나중에 Connected Sheets 추가될 때 대비)
  await refreshViaGas(spreadsheetId);

  // 1) 어드민에서 데이터 복사 (백그라운드 탭)
  const page = await newBackgroundPage();
  await page.goto(REMATCH_ADMIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(2000);

  // 전체 체크박스 클릭 (헤더의 체크박스)
  const headerCheckbox = page.locator('thead input[type="checkbox"]').first();
  await headerCheckbox.click();
  console.log('[sheetPrep] 전체 체크박스 선택');

  // Copy Clipboard 버튼 클릭
  const copyBtn = page.locator('button', { hasText: 'copy clipboard' })
    .or(page.locator('button', { hasText: 'Copy Clipboard' }))
    .or(page.locator('button', { hasText: 'Copy clipboard' }));
  await copyBtn.click();
  console.log('[sheetPrep] Copy Clipboard 클릭');

  // 클립보드 읽기
  const clipText = await page.evaluate(() => navigator.clipboard.readText());
  await page.close();

  if (!clipText || !clipText.trim()) {
    console.log('[sheetPrep] 클립보드 데이터 없음, 종료');
    return;
  }

  // 2) TSV 파싱 (탭 구분)
  const lines = clipText.trim().split('\n');
  // 첫 줄이 헤더일 수 있으므로 확인
  const dataLines = lines[0].toLowerCase().includes('pairing') ? lines.slice(1) : lines;
  const clipRows = dataLines
    .map(line => line.split('\t'))
    .filter(row => row.some(cell => cell && cell.trim() !== ''));

  console.log(`[sheetPrep] 어드민 복사 데이터: ${clipRows.length}행`);

  if (clipRows.length === 0) {
    console.log('[sheetPrep] 추가할 데이터 없음');
    return;
  }

  // 3) 월별 탭 기존 데이터 + 헤더 읽기
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${monthlyTab}'!A1:BZ1`,
  });
  const monthlyHeader = (headerRes.data.values || [])[0] || [];

  // 클립보드 인덱스 = 월별 인덱스 - 1 (월별 A는 처리일, B부터 클립보드 데이터)
  const findClipIdx = (name) => {
    const idx = monthlyHeader.findIndex(h => (h || '').trim().toLowerCase() === name.toLowerCase());
    return idx > 0 ? idx - 1 : -1;
  };
  const noteIdx = findClipIdx('note');
  const lastLessonIdx = findClipIdx('lastLessonDate');
  console.log(`[sheetPrep] 컬럼 위치 — note(클립): ${noteIdx}, lastLessonDate(클립): ${lastLessonIdx}`);

  // 어제 날짜 (시간 0시로 정규화)
  const yesterday = new Date();
  yesterday.setHours(0, 0, 0, 0);
  yesterday.setDate(yesterday.getDate() - 1);

  const parseDate = (s) => {
    if (!s) return null;
    const m = String(s).trim().match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(String(s).trim());
    return isNaN(d) ? null : d;
  };

  // 4) 클립보드 필터: lastLessonDate <= 어제 OR note == "1"
  const filteredClipRows = clipRows.filter(row => {
    const noteVal = (row[noteIdx] || '').toString().trim();
    const lastDate = parseDate(row[lastLessonIdx]);
    const dateOk = lastDate && lastDate <= yesterday;
    const noteOk = noteVal === '1';
    return dateOk || noteOk;
  });
  console.log(`[sheetPrep] 필터 후 (lastLessonDate≤어제 또는 note=1): ${filteredClipRows.length}/${clipRows.length}행`);

  // 5) 월별 탭 기존 데이터 읽기
  let existingRows = [];
  try {
    const monthlyRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${monthlyTab}'!A2:BZ`,
    });
    existingRows = monthlyRes.data.values || [];
  } catch (e) {
    console.log(`[sheetPrep] 월별 탭 '${monthlyTab}' 없음 또는 비어있음`);
  }

  // 4) 기존 pairing ID 수집 + 마지막 데이터 행 찾기 (B열 = index 1)
  const existingIds = new Set();
  let lastDataRowIdx = -1;
  existingRows.forEach((row, i) => {
    const id = (row[1] || '').trim();
    if (id) {
      existingIds.add(id);
      lastDataRowIdx = i;
    }
  });
  console.log(`[sheetPrep] 기존 월별 탭: ${existingRows.length}행 (실제 데이터 ${existingIds.size}개, 마지막 데이터 행 ${lastDataRowIdx + 2})`);

  // 7) 중복 제거: 기존에 있는 pairing ID는 추가하지 않음
  const today = getTodayString();
  const newRows = filteredClipRows
    .filter(row => {
      const pairingId = (row[0] || '').trim(); // 클립보드 첫 열 = 시트 B열
      return pairingId && !existingIds.has(pairingId);
    })
    .map(row => [today, ...row]); // A열에 오늘 날짜, B열부터 클립보드 데이터
  console.log(`[sheetPrep] 신규 추가 대상: ${newRows.length}행`);

  if (newRows.length > 0) {
    // 6) 실제 데이터 마지막 행 다음에 update로 직접 쓰기 (유령 행 무시)
    const startRow = lastDataRowIdx + 3;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${monthlyTab}'!A${startRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: newRows },
    });
    console.log(`[sheetPrep] ${newRows.length}행 추가 완료 (행 ${startRow}~${startRow + newRows.length - 1})`);
  } else {
    console.log('[sheetPrep] 추가할 신규 데이터 없음');
  }

  // 7) A열 날짜 처리
  await stampDateColumnA(sheets, spreadsheetId, monthlyTab, 'rematch');
  console.log('[sheetPrep] === 재매칭 시트 정리 완료 ===');
}

// ============================================================
// 월별 시트 분리 (월초 1회): 전월 잔여건 → 신규 월별 탭으로 이동
// ============================================================
async function createMonthlyTab(type) {
  const prefix = type === 'new' ? '신규' : '재매칭';
  console.log(`\n[sheetPrep] === ${prefix} 월별 탭 생성 ===`);

  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  const newTab = getMonthlyTabName(type);

  // 전월 탭 이름
  const now = new Date();
  const prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevYY = String(prevMonth.getFullYear()).slice(2);
  const prevMM = String(prevMonth.getMonth() + 1).padStart(2, '0');
  const prevTab = type === 'new'
    ? `[신규] ${prevYY}.${prevMM}`
    : `[재매칭] ${prevYY}.${prevMM}`;

  // 1) 전월 탭에서 잔여건 읽기 (제외 상태가 아닌 것)
  let prevRows = [];
  let headerRow = [];
  try {
    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${prevTab}'!A1:BZ1`,
    });
    headerRow = (headerRes.data.values || [])[0] || [];

    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${prevTab}'!A2:BZ`,
    });
    prevRows = dataRes.data.values || [];
  } catch (e) {
    console.log(`[sheetPrep] 전월 탭 '${prevTab}' 읽기 실패:`, e.message);
    return;
  }

  // 잔여건 필터 (완료/환불 제외)
  const excludeNew = ['매칭완료', '동일튜터유지', '환불'];
  const excludeRematch = ['재매칭 완료', '재매칭완료', '환불/소멸', '환불', '소멸'];
  const excludeList = type === 'new' ? excludeNew : excludeRematch;

  // 상태 컬럼 인덱스 찾기
  const statusColName = type === 'new' ? '매칭상태' : 'status';
  const statusIdx = headerRow.findIndex(h =>
    h.trim().toLowerCase() === statusColName.toLowerCase()
  );

  const remaining = prevRows.filter(row => {
    if (statusIdx < 0) return true;
    const status = (row[statusIdx] || '').trim();
    return !excludeList.includes(status);
  });

  console.log(`[sheetPrep] 전월 잔여건: ${remaining.length}/${prevRows.length}행`);

  // 2) 신규 월별 탭 생성 + 헤더 + 잔여건 쓰기
  try {
    // 탭 추가
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const tabExists = spreadsheet.data.sheets.some(
      s => s.properties.title === newTab
    );

    if (!tabExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: newTab } } }],
        },
      });
      console.log(`[sheetPrep] '${newTab}' 탭 생성`);
    }

    // 헤더 + 잔여건 쓰기
    const allRows = [headerRow, ...remaining];
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${newTab}'!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: allRows },
    });

    console.log(`[sheetPrep] '${newTab}' 탭에 헤더 + ${remaining.length}행 작성 완료`);
  } catch (e) {
    console.error(`[sheetPrep] 탭 생성 에러:`, e.message);
  }

  console.log(`[sheetPrep] === ${prefix} 월별 탭 생성 완료 ===`);
}

module.exports = { prepNewSheet, prepRematchSheet, createMonthlyTab, getMonthlyTabName };

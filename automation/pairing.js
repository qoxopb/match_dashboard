const { newBackgroundPage } = require('./browser');
const { getSheetsApi } = require('../sheets');
const config = require('../config.json');

const ADMIN_URL = (config.adminUrl || 'https://tutor-admin.qanda.ai') + '/admin/tutor-pairing';
const WORKER_COUNT = 5;

function getMonthlyTabName() {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  return `[신규] ${yy}.${mm}`;
}

// 0-indexed → A, B, ..., Z, AA, AB, ...
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

function buildFilterUrl(matchId) {
  return `${ADMIN_URL}?page=1&size=20&filters_matchId=${encodeURIComponent(matchId)}`;
}

async function runPairing(manager) {
  const managerLabel = manager && manager !== '공통' ? ` (담당자: ${manager})` : '';
  console.log(`\n[pairing] === 신규 페어링 생성 시작${managerLabel} ===`);
  const startedAt = Date.now();

  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets.new;
  const monthlyTab = getMonthlyTabName();

  // 1) 헤더 + 데이터 읽기
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${monthlyTab}'!A1:BZ1`,
  });
  const header = (headerRes.data.values || [])[0] || [];
  const matchIdIdx = header.findIndex(h => (h || '').trim().toLowerCase() === 'match_id');
  const statusIdx = header.findIndex(h => (h || '').trim() === '매칭상태');
  console.log(`[pairing] 컬럼 위치 — match_id: ${matchIdIdx}, 매칭상태: ${statusIdx}`);

  if (matchIdIdx < 0 || statusIdx < 0) {
    throw new Error('헤더에서 컬럼을 찾을 수 없습니다 (match_id 또는 매칭상태)');
  }

  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${monthlyTab}'!A2:BZ`,
  });
  const rows = dataRes.data.values || [];

  // 2) 대상 필터: match_id 있고 매칭상태가 빈 값
  const filterByManager = manager && manager !== '공통';
  const targets = [];
  rows.forEach((row, i) => {
    const matchId = (row[matchIdIdx] || '').trim();
    const status = (row[statusIdx] || '').trim();
    if (matchId && status === '') {
      if (filterByManager && (row[41] || '').trim().toLowerCase() !== manager.toLowerCase()) return;
      targets.push({ matchId, rowNum: i + 2 });
    }
  });
  console.log(`[pairing] 대상 ${targets.length}건 / ${WORKER_COUNT} 워커 병렬 처리`);

  if (targets.length === 0) {
    console.log('[pairing] 처리할 대상 없음');
    return;
  }

  // 3) 워커 수 조정 (대상보다 많으면 줄임)
  const workerCount = Math.min(WORKER_COUNT, targets.length);
  const statusColLetter = colNumberToLetter(statusIdx);

  // 4) 페이지 N개 열기
  const pages = [];
  for (let i = 0; i < workerCount; i++) {
    pages.push(await newBackgroundPage());
  }
  console.log(`[pairing] ${workerCount}개 백그라운드 페이지 준비 완료`);

  // 5) 작업 큐 — 각 워커가 atomic하게 다음 작업 가져감
  let nextIdx = 0;
  const counters = { success: 0, exists: 0, updated: 0, error: 0 };

  async function worker(workerId, page) {
    while (true) {
      const myIdx = nextIdx++;
      if (myIdx >= targets.length) break;
      const target = targets[myIdx];
      try {
        const result = await processOne(workerId, page, target, sheets, spreadsheetId, monthlyTab, statusColLetter);
        counters[result] = (counters[result] || 0) + 1;
      } catch (e) {
        console.error(`[pairing][W${workerId}] ${target.matchId} 에러: ${e.message}`);
        counters.error++;
        // 네비게이션 충돌 시 페이지 안정화 대기
        try { await page.waitForLoadState('domcontentloaded', { timeout: 5000 }); } catch {}
        await page.waitForTimeout(1000);
      }
    }
  }

  await Promise.all(pages.map((p, i) => worker(i + 1, p)));

  // 6) 페이지 정리
  await Promise.all(pages.map(p => p.close().catch(() => {})));

  const elapsedMs = Date.now() - startedAt;
  const mins = Math.floor(elapsedMs / 60000);
  const secs = Math.floor((elapsedMs % 60000) / 1000);
  const elapsedStr = mins > 0 ? `${mins}분 ${secs}초` : `${secs}초`;
  console.log(`[pairing] === 완료: 생성 ${counters.success}, 시트갱신 ${counters.updated}, 스킵 ${counters.exists}, 에러 ${counters.error} | 소요시간 ${elapsedStr} ===`);
}

async function processOne(workerId, page, target, sheets, spreadsheetId, monthlyTab, statusColLetter) {
  const tag = `[pairing][W${workerId}] ${target.matchId}`;
  console.log(`${tag} 처리 시작 (행 ${target.rowNum})`);

  // 1) 검색 → 기존 페어링 status 가져오기
  let status = await searchAndGetStatus(page, target.matchId, tag);

  let created = false;
  if (!status) {
    console.log(`${tag} 페어링 없음 → Create 시도`);

    // 2) 리스트 페이지로 이동 후 Create
    await page.goto(ADMIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });

    const createInput = page.locator('input[placeholder="match_id를 입력하세요"]');
    try {
      await createInput.waitFor({ state: 'visible', timeout: 8000 });
    } catch {
      console.log(`${tag} Create input 못찾음`);
      return 'error';
    }
    await createInput.fill(target.matchId);

    const createBtn = page.locator('button', { hasText: 'Create' }).first();
    await createBtn.click();

    // 페이지 이동 또는 짧게 대기
    await Promise.race([
      page.waitForURL(/\/tutor-pairing\/.+/, { timeout: 5000 }).catch(() => null),
      page.waitForTimeout(1500),
    ]);

    // 3) Create 후 다시 검색
    status = await searchAndGetStatus(page, target.matchId, tag);
    created = true;
  }

  if (!status) {
    console.log(`${tag} 검색 결과 없음 → 다음`);
    return 'error';
  }

  console.log(`${tag} status="${status}"${created ? ' (방금 생성)' : ' (기존)'}`);

  const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
  const ns = norm(status);

  if (ns.includes('매칭준비')) {
    return created ? 'success' : 'exists';
  }

  if (ns === '신청서작성') {
    const range = `'${monthlyTab}'!${statusColLetter}${target.rowNum}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['신청서 미작성']] },
    });
    console.log(`${tag} 시트 ${range} = "신청서 미작성"`);
    return 'updated';
  }

  console.log(`${tag} 미처리 status "${status}"`);
  return created ? 'success' : 'exists';
}

// URL 필터로 검색 + 첫 결과의 status 텍스트 반환 (없으면 null)
async function searchAndGetStatus(page, matchId, tag = '') {
  await page.goto(buildFilterUrl(matchId), { waitUntil: 'domcontentloaded', timeout: 30000 });

  // 테이블이 렌더링될 때까지 대기 (스마트 대기)
  try {
    await page.waitForSelector('table tbody, .ant-table-tbody', { timeout: 8000 });
  } catch {
    if (tag) console.log(`${tag} tbody 안 나타남`);
    return null;
  }

  const result = await page.evaluate(() => {
    const table = document.querySelector('table, .ant-table');
    if (!table) return { error: 'table 없음' };

    const headers = Array.from(table.querySelectorAll('thead th')).map(h => (h.textContent || '').trim());
    let statusIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      const t = headers[i].toLowerCase();
      if (t === 'status' || t === '상태' || t.includes('status')) {
        statusIdx = i;
        break;
      }
    }
    if (statusIdx < 0) return { error: 'status 헤더 없음' };

    const tbody = table.querySelector('tbody, .ant-table-tbody');
    if (!tbody) return { error: 'tbody 없음' };

    const dataRows = Array.from(tbody.querySelectorAll('tr')).filter(r => {
      const cells = r.querySelectorAll('td');
      return cells.length > 1 && Array.from(cells).some(c => (c.textContent || '').trim());
    });
    if (dataRows.length === 0) return { empty: true };

    const firstCells = dataRows[0].querySelectorAll('td');
    if (firstCells.length <= statusIdx) return { error: 'cells 부족' };
    return { status: (firstCells[statusIdx].textContent || '').trim() };
  });

  if (result.error) {
    if (tag) console.log(`${tag} status 조회 실패: ${result.error}`);
    return null;
  }
  if (result.empty) return null;
  return result.status;
}

module.exports = { runPairing };

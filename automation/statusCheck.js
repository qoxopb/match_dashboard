const { newBackgroundPage } = require('./browser');
const { getSheetsApi } = require('../sheets');
const config = require('../config.json');

const ADMIN_URL = (config.adminUrl || 'https://tutor-admin.qanda.ai') + '/admin/tutor-pairing';
const WORKER_COUNT = 5;

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

function buildFilterUrl(id, useIdFilter = false) {
  const filterKey = useIdFilter ? 'filters_id' : 'filters_matchId';
  return `${ADMIN_URL}?page=1&size=20&${filterKey}=${encodeURIComponent(id)}`;
}

async function searchAndGetStatus(page, id, tag = '', useIdFilter = false) {
  await page.goto(buildFilterUrl(id, useIdFilter), { waitUntil: 'domcontentloaded', timeout: 30000 });

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

async function runStatusCheckForType(type) {
  const label = type === 'new' ? '신규' : '재매칭';
  console.log(`\n[statusCheck] === ${label} 매칭결과 확인 시작 ===`);
  const startedAt = Date.now();

  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  const monthlyTab = getMonthlyTabName(type);

  // 1) 헤더 읽기
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${monthlyTab}'!A1:BZ1`,
  });
  const header = (headerRes.data.values || [])[0] || [];

  const matchIdColName = type === 'new' ? 'match_id' : 'match ID';
  const statusColName = type === 'new' ? '매칭상태' : 'status';
  const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();

  const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdColName));
  const statusIdx = header.findIndex(h => norm(h) === norm(statusColName));
  const pairingIdIdx = type === 'rematch' ? header.findIndex(h => norm(h) === norm('pairing ID')) : -1;
  console.log(`[statusCheck] 컬럼 위치 — ${matchIdColName}: ${matchIdIdx}, ${statusColName}: ${statusIdx}${pairingIdIdx >= 0 ? ', pairing ID: ' + pairingIdIdx : ''}`);

  if (matchIdIdx < 0 || statusIdx < 0) {
    throw new Error(`헤더에서 컬럼을 찾을 수 없습니다 (${matchIdColName} 또는 ${statusColName})`);
  }

  // 2) 데이터 읽기
  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${monthlyTab}'!A2:BZ`,
  });
  const rows = dataRes.data.values || [];

  // 3) 대상 필터: match_id 있고, 완료 상태 아닌 행
  const excludeNew = ['매칭완료', '동일튜터유지', '환불'];
  const excludeRematch = ['재매칭 완료', '재매칭완료', '환불/소멸', '환불', '소멸'];
  const excludeList = (type === 'new' ? excludeNew : excludeRematch).map(norm);

  const targets = [];
  rows.forEach((row, i) => {
    const matchId = (row[matchIdIdx] || '').trim();
    const status = (row[statusIdx] || '').trim();
    if (matchId && !excludeList.includes(norm(status))) {
      const pairingId = pairingIdIdx >= 0 ? (row[pairingIdIdx] || '').trim() : '';
      targets.push({ matchId, rowNum: i + 2, currentStatus: status, pairingId });
    }
  });
  console.log(`[statusCheck] 대상 ${targets.length}건 / ${WORKER_COUNT} 워커 병렬 처리`);

  if (targets.length === 0) {
    console.log('[statusCheck] 확인할 대상 없음');
    return;
  }

  // 4) 워커 준비
  const workerCount = Math.min(WORKER_COUNT, targets.length);
  const statusColLetter = colNumberToLetter(statusIdx);
  const completedStatus = type === 'new' ? '매칭완료' : '재매칭 완료';

  const pages = [];
  for (let i = 0; i < workerCount; i++) {
    pages.push(await newBackgroundPage());
  }
  console.log(`[statusCheck] ${workerCount}개 백그라운드 페이지 준비 완료`);

  // 5) 작업 큐
  let nextIdx = 0;
  const counters = { updated: 0, skipped: 0, error: 0 };

  async function worker(workerId, page) {
    while (true) {
      const myIdx = nextIdx++;
      if (myIdx >= targets.length) break;
      const target = targets[myIdx];
      const searchId = type === 'rematch' && target.pairingId ? target.pairingId : target.matchId;
      const useIdFilter = type === 'rematch' && !!target.pairingId;
      const tag = `[statusCheck][W${workerId}] ${target.matchId}`;

      try {
        const adminStatus = await searchAndGetStatus(page, searchId, tag, useIdFilter);

        if (!adminStatus) {
          console.log(`${tag} (${searchId}${useIdFilter ? '/id' : '/matchId'}) 검색 결과 없음 → 스킵`);
          counters.skipped++;
          continue;
        }

        if (norm(adminStatus).includes('매칭완료')) {
          const range = `'${monthlyTab}'!${statusColLetter}${target.rowNum}`;
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[completedStatus]] },
          });
          console.log(`${tag} → "${completedStatus}" 업데이트 (행 ${target.rowNum})`);
          counters.updated++;
        } else {
          console.log(`${tag} (${searchId}${useIdFilter ? '/id' : '/matchId'}) status="${adminStatus}" → 스킵`);
          counters.skipped++;
        }
      } catch (e) {
        console.error(`${tag} 에러: ${e.message}`);
        counters.error++;
      }
    }
  }

  await Promise.all(pages.map((p, i) => worker(i + 1, p)));
  await Promise.all(pages.map(p => p.close().catch(() => {})));

  const elapsedMs = Date.now() - startedAt;
  const mins = Math.floor(elapsedMs / 60000);
  const secs = Math.floor((elapsedMs % 60000) / 1000);
  const elapsedStr = mins > 0 ? `${mins}분 ${secs}초` : `${secs}초`;
  console.log(`[statusCheck] === ${label} 완료: 매칭완료 ${counters.updated}, 스킵 ${counters.skipped}, 에러 ${counters.error} | 소요시간 ${elapsedStr} ===`);
}

async function runStatusCheck(mode) {
  if (mode === '1' || mode === 'new') {
    await runStatusCheckForType('new');
  } else if (mode === '2' || mode === 'rematch') {
    await runStatusCheckForType('rematch');
  } else {
    await runStatusCheckForType('new');
    await runStatusCheckForType('rematch');
  }
}

module.exports = { runStatusCheck, runStatusCheckForType };

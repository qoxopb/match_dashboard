const express = require('express');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const notifier = require('node-notifier');
const config = require('./config.json');
const { prepNewSheet, prepRematchSheet, createMonthlyTab } = require('./automation/sheetPrep');
const { runPairing } = require('./automation/pairing');
const { runStatusCheck, runStatusCheckForType } = require('./automation/statusCheck');
const { runAim } = require('./automation/aim');
const { disconnect } = require('./automation/browser');


const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// --- 로그 ---
const { shouldAbort, requestAbort, resetAbort } = require('./abort');
const logs = [];
let running = null;

// 파일 로그 — logs/ 폴더에 날짜별 저장
const logsDir = path.join(__dirname, 'logs');
if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir);

function getLogFileName() {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  return path.join(logsDir, `${yy}${mm}${dd}.log`);
}

function addLog(msg) {
  const time = new Date().toLocaleTimeString('ko-KR');
  logs.push({ time, msg });
  if (logs.length > 2000) logs.shift();
  // 파일에도 기록
  try { fs.appendFileSync(getLogFileName(), `[${time}] ${msg}\n`); } catch {}
}

const origLog = console.log;
const origError = console.error;
console.log = (...args) => { addLog(args.join(' ')); };
console.error = (...args) => { addLog('[ERROR] ' + args.join(' ')); };

// --- 작업 정의 ---
const JOBS = {
  prep:        { name: '시트 정리', fn: async () => { await prepNewSheet(); await prepRematchSheet(); } },
  'prep:new':  { name: '신규 시트 정리', fn: () => prepNewSheet() },
  'prep:rematch': { name: '재매칭 시트 정리', fn: () => prepRematchSheet() },
  pairing:     { name: '페어링 생성', fn: () => runPairing(config.mode) },
  aim:         { name: '매칭 제안 (AIM)', fn: () => runAim('both') },
  'aim:new':   { name: '신규 매칭 제안', fn: () => runAim('new') },
  'aim:rematch': { name: '재매칭 제안', fn: () => runAim('rematch') },
  statusCheck: { name: '매칭결과 확인', fn: () => runStatusCheck('both') },
  'statusCheck:new': { name: '신규 매칭결과 확인', fn: () => runStatusCheck('new') },
  'statusCheck:rematch': { name: '재매칭 결과 확인', fn: () => runStatusCheck('rematch') },
  monthly:     { name: '월별 탭 생성', fn: async () => {
    if (config.mode === '1' || config.mode === 'both') await createMonthlyTab('new');
    if (config.mode === '2' || config.mode === 'both') await createMonthlyTab('rematch');
  }},
};

// 실행 중인 작업 세트 + 대기 큐
const runningJobs = new Set(); // 실행 중인 baseJobId들
const taskQueue = []; // { jobId, resolve }

function getBaseJobId(jobId) {
  // 'wf:xxx' → 워크플로 내부 블록의 실제 jobId 기준이 아닌, 최상위 jobId 그룹
  // 'aim:new', 'aim:rematch', 'aim' → 'aim'
  // 'statusCheck:new' → 'statusCheck'
  // 'wf:123' → 'wf:123' (워크플로 자체는 고유)
  if (jobId.startsWith('wf:')) return 'wf';
  return jobId.split(':')[0];
}

async function runTask(jobId) {
  const job = JOBS[jobId];
  if (!job) return { ok: false, message: `알 수 없는 작업: ${jobId}` };

  const baseId = getBaseJobId(jobId);

  // 같은 종류 실행 중이면 큐에 넣고 대기
  if (runningJobs.has(baseId)) {
    addLog(`[큐] ${job.name} 대기 (${baseId} 실행 중)`);
    await new Promise(resolve => { taskQueue.push({ baseId, resolve }); });
  }

  runningJobs.add(baseId);
  running = [...runningJobs].length > 0 ? [...runningJobs].map(id => { const j = JOBS[id]; return j ? j.name : id; }).join(', ') : null;
  resetAbort();
  addLog(`=== ${job.name} 시작 ===`);
  try {
    await job.fn();
    if (shouldAbort()) {
      addLog(`=== ${job.name} 중단됨 ===`);
      notifier.notify({ title: '매칭 자동화', message: `${job.name} 중단됨` });
      return { ok: true, message: `${job.name} 중단됨` };
    }
    addLog(`=== ${job.name} 완료 ===`);
    notifier.notify({ title: '매칭 자동화', message: `${job.name} 완료` });
    return { ok: true, message: `${job.name} 완료` };
  } catch (err) {
    if (shouldAbort()) {
      addLog(`=== ${job.name} 중단됨 ===`);
      return { ok: true, message: `${job.name} 중단됨` };
    }
    addLog(`=== ${job.name} 에러: ${err.message} ===`);
    notifier.notify({ title: '매칭 자동화', message: `${job.name} 에러: ${err.message}` });
    return { ok: false, message: err.message };
  } finally {
    runningJobs.delete(baseId);
    running = runningJobs.size > 0 ? [...runningJobs].map(id => { const j = JOBS[id]; return j ? j.name : id; }).join(', ') : null;
    // 큐에서 같은 종류 대기 중인 작업 하나 깨우기
    const idx = taskQueue.findIndex(t => t.baseId === baseId);
    if (idx >= 0) taskQueue.splice(idx, 1)[0].resolve();
    resetAbort();
  }
}

// --- 스케줄 관리 (파일 영속화) ---
const schedulesPath = path.join(__dirname, 'schedules.json');
let scheduleIdCounter = 1;
const schedules = {}; // id → { id, jobId, type, label, cronTask, enabled, ... }

function saveSchedulesToFile() {
  const data = Object.values(schedules).map(s => ({
    id: s.id, jobId: s.jobId, type: s.type, label: s.label, enabled: s.enabled, options: s.options,
  }));
  try { fs.writeFileSync(schedulesPath, JSON.stringify(data, null, 2), 'utf-8'); } catch {}
}

function getNextRunTime(type, options) {
  const now = new Date();
  if (type === 'interval') {
    const mins = options.minutes;
    const nextMin = Math.ceil((now.getMinutes() + 1) / mins) * mins;
    const next = new Date(now);
    next.setMinutes(nextMin, 0, 0);
    if (next <= now) next.setMinutes(next.getMinutes() + mins);
    return next.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
  }
  if (type === 'alarm') {
    const hh = String(options.hour).padStart(2, '0');
    const mm = String(options.minute).padStart(2, '0');
    return `${hh}:${mm}`;
  }
  if (type === 'date') {
    const hh = String(options.hour).padStart(2, '0');
    const mm = String(options.minute).padStart(2, '0');
    return `${options.dateStr || ''} ${hh}:${mm}`;
  }
  return '?';
}

function loadSchedulesFromFile() {
  try {
    const data = JSON.parse(fs.readFileSync(schedulesPath, 'utf-8'));
    let maxId = 0;
    for (const s of data) {
      if (s.id > maxId) maxId = s.id;
      const restored = addSchedule(s.jobId, s.type, s.options, s.id, s.enabled, true);
      if (restored) {
        const nextLabel = s.enabled ? getNextRunTime(s.type, s.options) : '비활성';
        console.log(`[스케줄] 복원: ${s.jobId} — ${s.label} (${s.enabled ? 'ON' : 'OFF'}) → 다음 실행: ${nextLabel}`);
      }
    }
    if (maxId >= scheduleIdCounter) scheduleIdCounter = maxId + 1;
  } catch {}
}

function addSchedule(jobId, type, options, forceId, forceEnabled, isRestore) {
  const id = forceId || scheduleIdCounter++;
  if (!forceId && id >= scheduleIdCounter) scheduleIdCounter = id + 1;
  const job = JOBS[jobId];
  if (!job) return null;

  let cronExpr, label;

  if (type === 'interval') {
    const mins = options.minutes;
    cronExpr = `*/${mins} * * * *`;
    label = `${mins}분 간격`;
  } else if (type === 'alarm') {
    const { hour, minute, days } = options;
    const hh = String(hour).padStart(2, '0');
    const mm = String(minute).padStart(2, '0');
    if (days && days.length > 0 && days.length < 7) {
      cronExpr = `${minute} ${hour} * * ${days.join(',')}`;
      const dayNames = ['일','월','화','수','목','금','토'];
      const dayStr = days.map(d => dayNames[d]).join(',');
      label = `${hh}:${mm} (${dayStr})`;
    } else {
      cronExpr = `${minute} ${hour} * * *`;
      label = `${hh}:${mm} (매일)`;
    }
  } else if (type === 'date') {
    const { hour, minute, date, month, repeat, dateStr } = options;
    const hh = String(hour).padStart(2, '0');
    const mm = String(minute).padStart(2, '0');
    if (repeat === 'monthly') {
      cronExpr = `${minute} ${hour} ${date} * *`;
      label = `매월 ${date}일 ${hh}:${mm}`;
    } else if (repeat === 'yearly') {
      cronExpr = `${minute} ${hour} ${date} ${month} *`;
      label = `매년 ${month}/${date} ${hh}:${mm}`;
    } else {
      cronExpr = `${minute} ${hour} ${date} ${month} *`;
      label = `${dateStr} ${hh}:${mm} (1회)`;
    }
  }

  const cronTask = cron.schedule(cronExpr, async () => {
    const sched = schedules[id];
    if (!sched || !sched.enabled) return;
    console.log(`[스케줄] ${job.name} 자동 실행 (${label})`);
    await runTask(jobId);
  }, { scheduled: true });

  const enabled = forceEnabled !== undefined ? forceEnabled : true;
  if (!enabled) cronTask.stop();
  schedules[id] = { id, jobId, jobName: job.name, type, label, cronExpr, enabled, cronTask, options };
  if (!isRestore) {
    console.log(`[스케줄] 등록: ${job.name} — ${label}`);
    // 주기 타입은 등록 즉시 1회 실행
    if (type === 'interval') {
      console.log(`[스케줄] ${job.name} 즉시 실행`);
      runTask(jobId);
    }
  }
  saveSchedulesToFile();
  return id;
}

function removeSchedule(id) {
  const sched = schedules[id];
  if (!sched) return false;
  sched.cronTask.stop();
  delete schedules[id];
  console.log(`[스케줄] 삭제: ${sched.jobName} — ${sched.label}`);
  saveSchedulesToFile();
  return true;
}

function toggleSchedule(id, enabled) {
  const sched = schedules[id];
  if (!sched) return false;
  sched.enabled = enabled;
  if (enabled) sched.cronTask.start(); else sched.cronTask.stop();
  console.log(`[스케줄] ${sched.jobName} — ${sched.label}: ${enabled ? 'ON' : 'OFF'}`);
  saveSchedulesToFile();
  return true;
}

function getScheduleList() {
  return Object.values(schedules).map(s => ({
    id: s.id, jobId: s.jobId, jobName: s.jobName,
    type: s.type, label: s.label, enabled: s.enabled,
  }));
}

// --- API ---

// 즉시 실행
app.post('/api/run/:jobId', async (req, res) => {
  const result = await runTask(req.params.jobId);
  res.json(result);
});

// 스케줄 등록
app.post('/api/schedule', (req, res) => {
  const { jobId, type, options } = req.body;
  const id = addSchedule(jobId, type, options);
  if (id) res.json({ ok: true, id });
  else res.json({ ok: false, message: '등록 실패' });
});

// 스케줄 삭제
app.delete('/api/schedule/:id', (req, res) => {
  const ok = removeSchedule(Number(req.params.id));
  res.json({ ok });
});

// 스케줄 토글
app.patch('/api/schedule/:id', (req, res) => {
  const ok = toggleSchedule(Number(req.params.id), req.body.enabled);
  res.json({ ok });
});

// 스케줄 목록
app.get('/api/schedules', (req, res) => {
  res.json(getScheduleList());
});

// 로그 + 상태
app.get('/api/logs', (req, res) => {
  res.json({ running, logs: logs.map(l => `[${l.time}] ${l.msg}`) });
});

// 파일 로그 조회 (날짜별, ?date=260422 또는 기본 오늘)
app.get('/api/logs/file', (req, res) => {
  try {
    const date = req.query.date || (() => {
      const now = new Date();
      return String(now.getFullYear()).slice(2) + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
    })();
    const filePath = path.join(logsDir, `${date}.log`);
    if (!fs.existsSync(filePath)) return res.json({ date, lines: [] });
    const content = fs.readFileSync(filePath, 'utf-8');
    const lines = content.split('\n').filter(l => l.trim());
    res.json({ date, total: lines.length, lines });
  } catch (e) {
    res.json({ error: e.message });
  }
});

// 로그 파일 목록
app.get('/api/logs/files', (req, res) => {
  try {
    const files = fs.readdirSync(logsDir).filter(f => f.endsWith('.log')).sort().reverse();
    res.json(files);
  } catch (e) { res.json([]); }
});

app.get('/api/status', (req, res) => {
  res.json({ running, runningJobs: [...runningJobs], queue: taskQueue.length, mode: config.mode });
});

// [디버그] 유저메모 시트의 status 분포 + 필터 결과
app.get('/api/userMemo/debug', async (req, res) => {
  try {
    const type = req.query.type || 'new';
    const sheets = await getSheetsApi();
    const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
    const tabName = getUserMemoTabName(type);

    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A1:BZ1`,
    });
    const header = (headerRes.data.values || [])[0] || [];

    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A2:BZ`,
    });
    const rows = dataRes.data.values || [];

    const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
    const matchIdCol = type === 'new' ? 'match_id' : 'match ID';
    const statusCol = type === 'new' ? '매칭상태' : 'status';
    const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdCol));
    const statusIdx = header.findIndex(h => norm(h) === norm(statusCol));

    const statusCounts = {};
    let totalWithMatchId = 0;
    rows.forEach(row => {
      const matchId = matchIdIdx >= 0 ? (row[matchIdIdx] || '').trim() : '';
      if (!matchId) return;
      totalWithMatchId++;
      const status = statusIdx >= 0 ? (row[statusIdx] || '').trim() : '';
      const key = status || '(빈값)';
      statusCounts[key] = (statusCounts[key] || 0) + 1;
    });

    res.json({
      type, tabName,
      headerCols: { matchId: matchIdCol, status: statusCol },
      headerIdx: { matchId: matchIdIdx, status: statusIdx },
      totalRows: rows.length,
      totalWithMatchId,
      statusCounts,
    });
  } catch (e) {
    res.json({ error: e.message });
  }
});

// [워크플로우] 시트 status별 카운트
app.get('/api/wf/count', async (req, res) => {
  try {
    const type = req.query.type || 'new';
    const targetStatus = (req.query.status || '').trim(); // 빈 문자열 = 미처리
    const sheets = await getSheetsApi();
    const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
    const now = new Date();
    const yy = String(now.getFullYear()).slice(2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const tabName = type === 'new' ? `[신규] ${yy}.${mm}` : `[재매칭] ${yy}.${mm}`;

    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A1:BZ1`,
    });
    const header = (headerRes.data.values || [])[0] || [];
    const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
    const matchIdCol = type === 'new' ? 'match_id' : 'match ID';
    const statusCol = type === 'new' ? '매칭상태' : 'status';
    const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdCol));
    const statusIdx = header.findIndex(h => norm(h) === norm(statusCol));

    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A2:BZ`,
    });
    const rows = dataRes.data.values || [];
    const targetNorm = norm(targetStatus);

    let count = 0;
    rows.forEach(row => {
      const matchId = matchIdIdx >= 0 ? (row[matchIdIdx] || '').trim() : '';
      if (!matchId) return;
      const status = statusIdx >= 0 ? (row[statusIdx] || '').trim() : '';
      if (norm(status) === targetNorm) count++;
    });

    res.json({ ok: true, type, status: targetStatus, count });
  } catch (e) {
    res.json({ ok: false, message: e.message, count: 0 });
  }
});

// [디버그] 어드민 상세 페이지의 __NEXT_DATA__ 전체 덤프
app.get('/api/debug/nextdata/:matchId', async (req, res) => {
  const { newBackgroundPage } = require('./automation/browser');
  const page = await newBackgroundPage();
  try {
    const matchId = req.params.matchId;
    await page.goto(`https://tutor-admin.qanda.ai/admin/tutor-pairing?page=1&size=20&filters_matchId=${matchId}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForSelector('table tbody', { timeout: 8000 });
    const detailHref = await page.evaluate(() => {
      const link = document.querySelector('table tbody a[href*="/tutor-pairing/"]');
      return link ? link.getAttribute('href') : null;
    });
    if (!detailHref) return res.json({ error: '상세 링크 못 찾음' });
    await page.goto(new URL(detailHref, 'https://tutor-admin.qanda.ai').href, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForSelector('#__NEXT_DATA__', { state: 'attached', timeout: 5000 });

    const data = await page.evaluate(() => {
      const script = document.getElementById('__NEXT_DATA__');
      if (!script) return null;
      try {
        const json = JSON.parse(script.textContent);
        return json?.props?.pageProps || null;
      } catch (e) { return { error: e.message }; }
    });

    // 추가: "조건완화", "tutor_genders", "tutor_university_level" 키워드를 페이지 innerText에서도 검색
    const innerTextSnippets = await page.evaluate(() => {
      const text = document.body.innerText;
      const targets = ['조건완화', 'tutor_genders', 'tutor_university_level', 'gender', 'university', 'relax'];
      const result = {};
      for (const t of targets) {
        const idx = text.toLowerCase().indexOf(t.toLowerCase());
        if (idx >= 0) {
          result[t] = text.substring(Math.max(0, idx - 30), idx + 200);
        }
      }
      return result;
    });

    res.json({ matchId, pageProps: data, innerTextSnippets });
  } catch (e) {
    res.json({ error: e.message });
  } finally {
    await page.close().catch(() => {});
  }
});

// [임시] 페어링 상세 페이지 구조 확인용
app.get('/api/debug/pairing/:matchId', async (req, res) => {
  const { newBackgroundPage } = require('./automation/browser');
  const page = await newBackgroundPage();
  try {
    const matchId = req.params.matchId;
    await page.goto(`https://tutor-admin.qanda.ai/admin/tutor-pairing?page=1&size=20&filters_matchId=${matchId}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForSelector('table tbody', { timeout: 8000 });
    const detailHref = await page.evaluate(() => {
      const link = document.querySelector('table tbody a[href*="/tutor-pairing/"]');
      return link ? link.getAttribute('href') : null;
    });
    if (!detailHref) return res.json({ error: '상세 링크 못 찾음' });
    await page.goto(new URL(detailHref, 'https://tutor-admin.qanda.ai').href, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(3000);

    // 완화 단계 (쿼리: ?relax=0,1,2,3)
    const relaxLevel = parseInt(req.query.relax || '0');
    const relaxLog = [];

    // 헬퍼 함수 (브라우저 내)
    const helpers = `
      function findSelectNearLabel(p){const all=[...document.querySelectorAll('*')];for(const el of all){if(el.children.length===0&&new RegExp(p,'i').test(el.textContent.trim())){let parent=el.parentElement;for(let d=0;d<8&&parent;d++){const s=parent.querySelector('.ant-select');if(s)return s;parent=parent.parentElement;}}}return null;}
      function clearAntSelect(s){if(!s)return false;s.dispatchEvent(new MouseEvent('mouseenter',{bubbles:true}));const c=s.querySelector('.ant-select-clear');if(c){c.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));c.click();return true;}const r=s.querySelectorAll('.ant-select-selection-item-remove');r.forEach(x=>x.click());return r.length>0;}
    `;

    if (relaxLevel >= 1) {
      // 완화1: 입시 전형 정시+기타 추가, 튜터 스타일 삭제
      const r1 = await page.evaluate((h) => {
        eval(h);
        const log = [];
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
                  if ((t === '정시' || t === '기타') && !checked) {
                    (cb.querySelector('input') || cb).click();
                    log.push(t + ' 추가');
                  }
                });
                break;
              }
              parent = parent.parentElement;
            }
            break;
          }
        }
        const ss = findSelectNearLabel('튜터\\\\s*스타일');
        log.push('튜터스타일: ' + (ss && clearAntSelect(ss) ? 'cleared' : 'fail'));
        return log;
      }, helpers);
      relaxLog.push('relax1: ' + r1.join(', '));
    }
    if (relaxLevel >= 2) {
      await page.waitForTimeout(300);
      const r2 = await page.evaluate((h) => {
        eval(h);
        const s = findSelectNearLabel('^성적대$');
        return s && clearAntSelect(s) ? 'cleared' : 'fail';
      }, helpers);
      relaxLog.push('relax2 성적대: ' + r2);
    }
    if (relaxLevel >= 3) {
      await page.waitForTimeout(300);
      const r3 = await page.evaluate((h) => {
        eval(h);
        const s = findSelectNearLabel('고3\\\\s*가능\\\\s*과목');
        if (!s) return 'select 못 찾음';
        clearAntSelect(s);
        return 'cleared';
      }, helpers);
      relaxLog.push('relax3 고3과목: ' + r3);
    }

    // Search 버튼 클릭
    await page.evaluate(() => {
      const btns = [...document.querySelectorAll('button')];
      for (const btn of btns) {
        if (/^Search$/i.test(btn.textContent.trim())) { btn.click(); return; }
      }
    });
    await page.waitForTimeout(4000);

    // 모든 테이블 구조 덤프
    const tables = await page.evaluate(() => {
      const result = [];
      document.querySelectorAll('table').forEach((table, ti) => {
        const headers = [...table.querySelectorAll('thead th')].map(h => h.textContent.trim());
        const rows = [...table.querySelectorAll('tbody tr')].map(r => {
          const cells = [...r.querySelectorAll('td')].map(c => c.textContent.trim().substring(0, 50));
          return cells;
        });
        result.push({ tableIndex: ti, headers, rowCount: rows.length, rows: rows.slice(0, 5) });
      });
      return result;
    });

    // 필터 영역 DOM 구조 덤프
    const filterDom = await page.evaluate(() => {
      const result = {};
      // 각 필터 항목의 상태 수집
      const labels = ['성별','계열','대학교','입시 전형','성적대','튜터 스타일','고3 가능 과목','과목'];
      for (const label of labels) {
        const all = [...document.querySelectorAll('*')];
        for (const el of all) {
          if (el.children.length === 0 && el.textContent.trim() === label) {
            let parent = el.parentElement;
            for (let d = 0; d < 8 && parent; d++) {
              // 체크박스
              const checks = parent.querySelectorAll('.ant-checkbox-wrapper, .ant-tag-checkable, .ant-tag');
              if (checks.length > 0) {
                result[label] = {
                  type: 'tags/checks',
                  items: [...checks].map(c => ({
                    text: c.textContent.trim().substring(0, 30),
                    selected: c.classList.contains('ant-tag-checkable-checked') || c.classList.contains('ant-checkbox-wrapper-checked') || c.querySelector('input:checked') !== null,
                    classes: [...c.classList].join(' ')
                  }))
                };
                break;
              }
              // 셀렉트
              const selects = parent.querySelectorAll('.ant-select');
              if (selects.length > 0) {
                result[label] = {
                  type: 'select',
                  items: [...selects].map(s => ({
                    value: s.querySelector('.ant-select-selection-item')?.textContent?.trim() || '',
                    classes: [...s.classList].join(' ')
                  }))
                };
                break;
              }
              parent = parent.parentElement;
            }
            break;
          }
        }
      }
      return result;
    });

    res.json({ relaxLog, tables: tables.slice(-2), filterDom });
  } catch (e) {
    res.json({ error: e.message });
  } finally {
    await page.close().catch(() => {});
  }
});

app.post('/api/stop', (req, res) => {
  if (!running) return res.json({ ok: false, message: '실행 중인 작업 없음' });
  requestAbort();
  addLog(`=== ${running} 중단 요청 ===`);
  res.json({ ok: true, message: `${running} 중단 요청됨` });
});

// --- AIM 키워드 설정 API ---
const configPath = path.join(__dirname, 'config.json');

function saveConfig() {
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2), 'utf-8');
}

app.get('/api/aim/keywords', (req, res) => {
  res.json(config.aimKeywords || []);
});

app.post('/api/aim/keywords', (req, res) => {
  const { keyword, action, tutorIds } = req.body;
  if (!keyword || !action) return res.status(400).json({ error: 'keyword, action 필수' });
  if (!config.aimKeywords) config.aimKeywords = [];
  config.aimKeywords.push({ keyword, action, tutorIds: tutorIds || '' });
  saveConfig();
  res.json({ ok: true, keywords: config.aimKeywords });
});

app.delete('/api/aim/keywords/:idx', (req, res) => {
  const idx = parseInt(req.params.idx);
  if (!config.aimKeywords || idx < 0 || idx >= config.aimKeywords.length) {
    return res.status(400).json({ error: '잘못된 인덱스' });
  }
  config.aimKeywords.splice(idx, 1);
  saveConfig();
  res.json({ ok: true, keywords: config.aimKeywords });
});

// --- 유저메모 API ---
const { getSheetsApi } = require('./sheets');
const { fetchDetail: fetchMemoDetail, saveTutorMemo, closePage: closeMemoPage } = require('./automation/userMemo');

function getUserMemoTabName(type) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  return type === 'new' ? `[신규] ${yy}.${mm}` : `[재매칭] ${yy}.${mm}`;
}

// 유저메모 진입 전 상태 동기화 (statusCheck 실행)
app.get('/api/userMemo/sync', async (req, res) => {
  try {
    const type = req.query.type || 'new';
    console.log(`[userMemo] ${type === 'new' ? '신규' : '재매칭'} 상태 동기화 시작`);
    await runStatusCheckForType(type);
    res.json({ ok: true });
  } catch (e) {
    console.error('[userMemo] sync 에러:', e.message);
    res.json({ ok: false, message: e.message });
  }
});

// 목록 조회 (status 공란인 건만)
app.get('/api/userMemo/list', async (req, res) => {
  try {
    const type = req.query.type || 'new';
    const sheets = await getSheetsApi();
    const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
    const tabName = getUserMemoTabName(type);

    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A1:BZ1`,
    });
    const header = (headerRes.data.values || [])[0] || [];

    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${tabName}'!A2:BZ`,
    });
    const rows = dataRes.data.values || [];

    const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
    const matchIdCol = type === 'new' ? 'match_id' : 'match ID';
    const statusCol = type === 'new' ? '매칭상태' : 'status';
    const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdCol));
    const statusIdx = header.findIndex(h => norm(h) === norm(statusCol));

    const items = rows.map((row, i) => {
      const matchId = matchIdIdx >= 0 ? (row[matchIdIdx] || '').trim() : '';
      const status = statusIdx >= 0 ? (row[statusIdx] || '').trim() : '';
      const formData = {};
      header.forEach((h, ci) => { formData[h] = row[ci] || ''; });
      return { matchId, status, rowNum: i + 2, formData };
    }).filter(item => {
      if (!item.matchId) return false;
      const sn = (item.status || '').replace(/\s+/g, '').toLowerCase();
      const filter = req.query.filter || 'all';

      if (filter === 'initial') {
        // 초기편집: status가 공란 또는 READY
        return sn === '' || sn === 'ready';
      }

      // 전체: 기존 필터 + 초기편집 대상 포함
      if (type === 'new') {
        return sn === '' || sn === '매칭중' || sn === '확인필요';
      } else {
        const allowed = ['', 'ready', '첫수업전재매칭', '일반재매칭', '매칭중', '보류', '확인필요'];
        return allowed.includes(sn);
      }
    });

    res.json(items);
  } catch (e) {
    console.error('[userMemo] list 에러:', e.message);
    res.json([]);
  }
});

// 상세 조회 (어드민 스크래핑 + Slack)
app.get('/api/userMemo/detail', async (req, res) => {
  try {
    const matchId = req.query.matchId;
    if (!matchId) return res.json({ ok: false, message: 'matchId 필요' });
    const detail = await fetchMemoDetail(matchId);
    res.json({ ok: true, ...detail });
  } catch (e) {
    console.error('[userMemo] detail 에러:', e.message);
    res.json({ ok: false, message: e.message });
  }
});

// AI 편집 (placeholder)
app.post('/api/userMemo/ai-edit', (req, res) => {
  res.json({ ok: false, message: 'AI 편집 기능 미구현' });
});

// 어드민 저장
app.post('/api/userMemo/save', async (req, res) => {
  try {
    const { matchId, tutorMemo } = req.body;
    if (!matchId || !tutorMemo) return res.json({ ok: false, message: 'matchId와 tutorMemo 필요' });
    const result = await saveTutorMemo(matchId, tutorMemo);
    res.json(result);
  } catch (e) {
    console.error('[userMemo] save 에러:', e.message);
    res.json({ ok: false, message: e.message });
  }
});

// --- 버전 & 배포 ---
const { execSync } = require('child_process');

function getCurrentVersion() {
  try {
    return JSON.parse(fs.readFileSync(path.join(__dirname, 'version.json'), 'utf-8')).version;
  } catch { return '?'; }
}

app.get('/api/version', (req, res) => {
  res.json({ version: getCurrentVersion() });
});

app.get('/api/changelog', (req, res) => {
  try {
    const log = JSON.parse(fs.readFileSync(path.join(__dirname, 'changelog.json'), 'utf-8'));
    res.json(log);
  } catch { res.json([]); }
});

app.get('/api/version/latest', (req, res) => {
  try {
    execSync('git fetch', { cwd: __dirname, encoding: 'utf-8', timeout: 15000 });
    const remote = execSync('git show origin/main:version.json', { cwd: __dirname, encoding: 'utf-8', timeout: 5000 });
    const latest = JSON.parse(remote).version;
    res.json({ version: latest, current: getCurrentVersion(), updateAvailable: latest !== getCurrentVersion() });
  } catch (e) {
    res.json({ version: '?', current: getCurrentVersion(), updateAvailable: false, error: e.message });
  }
});

app.post('/api/deploy', (req, res) => {
  try {
    const output = execSync('git pull', { cwd: __dirname, encoding: 'utf-8', timeout: 30000 });
    addLog(`[deploy] git pull: ${output.trim()}`);
    res.json({ ok: true, message: output.trim() });
    setTimeout(() => process.exit(0), 1000);
  } catch (e) {
    res.json({ ok: false, message: e.message });
  }
});

// --- 워크플로 프리셋 서버 저장 & 실행 ---
const wfPresetsPath = path.join(__dirname, 'wf-presets.json');
let wfPresetsData = {};
try { wfPresetsData = JSON.parse(fs.readFileSync(wfPresetsPath, 'utf-8')); } catch {}

function saveWfPresets() {
  fs.writeFileSync(wfPresetsPath, JSON.stringify(wfPresetsData, null, 2), 'utf-8');
}

async function runWfBlocks(blocks) {
  for (const b of blocks) {
    if (shouldAbort()) return;
    if (b.kind === 'job') {
      const job = JOBS[b.jobId];
      if (job) {
        console.log(`[워크플로] ${b.label || b.jobId} 실행`);
        await job.fn();
      }
    } else if (b.kind === 'loop') {
      for (let r = 0; r < (b.count || 1); r++) {
        if (shouldAbort()) return;
        console.log(`[워크플로] 반복 ${r + 1}/${b.count}`);
        await runWfBlocks(b.children || []);
      }
    } else if (b.kind === 'condLoop') {
      for (let r = 0; r < (b.maxIter || 5); r++) {
        if (shouldAbort()) return;
        const met = await checkWfCondition(b);
        if (met) { console.log(`[워크플로] 조건 만족 — ${r}회차에 종료`); break; }
        console.log(`[워크플로] 조건 반복 ${r + 1}/${b.maxIter}`);
        await runWfBlocks(b.children || []);
      }
    } else if (b.kind === 'userMemo') {
      const memoType = b.memoType || 'all';
      const memoFilter = b.memoFilter || 'initial';
      const typeLabel = { new: '신규', rematch: '재매칭', all: '전체' }[memoType] || memoType;

      // 편집 대상 건수 체크
      const memoCounts = await countMemoTargetsByType(memoType, memoFilter);
      const totalCount = memoCounts.new + memoCounts.rematch;
      if (totalCount === 0) {
        console.log(`[워크플로] 유저메모 편집 대상 0건 → 건너뜀`);
      } else {
        const countDesc = memoType === 'all'
          ? `신규 ${memoCounts.new}건 / 재매칭 ${memoCounts.rematch}건`
          : `${typeLabel} ${totalCount}건`;
        console.log(`[워크플로] 유저메모 편집 대기 (${countDesc})`);
        const dashboardUrl = `http://192.168.0.185:${PORT}/memoLanding.html?memoType=${memoType}&memoFilter=${memoFilter}`;
        await sendSlackNotification(
          `:memo: *유저메모 편집이 필요합니다*\n${countDesc}\n아래 버튼을 눌러 편집을 시작해주세요.`,
          dashboardUrl
        );
        // 대기
        wfPendingTask = { type: memoType, filter: memoFilter, createdAt: Date.now() };
        await new Promise(resolve => { wfPendingTask.resolve = resolve; });
        wfPendingTask = null;
        console.log('[워크플로] 유저메모 편집 완료 → 다음 블록');
      }
    }
  }
}

// --- Slack 알림 ---
const axios = require('axios');

async function sendSlackNotification(text, url) {
  try {
    const token = config.slack && config.slack.botToken;
    const channel = (config.slack && config.slack.notifyChannelId) || (config.slack && config.slack.channelId);
    if (!token || !channel) { console.log('[Slack] 토큰/채널 미설정 → 알림 생략'); return; }
    const blocks = [
      { type: 'section', text: { type: 'mrkdwn', text } },
    ];
    if (url) {
      blocks.push({
        type: 'actions',
        elements: [{
          type: 'button',
          text: { type: 'plain_text', text: '유저메모 편집 열기' },
          url,
          style: 'primary',
        }],
      });
    }
    await axios.post('https://slack.com/api/chat.postMessage', {
      channel, text, blocks,
    }, { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } });
    console.log(`[Slack] 알림 전송: ${text.substring(0, 60)}`);
  } catch (e) {
    console.error(`[Slack] 알림 실패: ${e.message}`);
  }
}

// --- 유저메모 편집 대상 건수 ---
async function countMemoTargetsByType(memoType, memoFilter) {
  try {
    const sheets = await getSheetsApi();
    const counts = { new: 0, rematch: 0 };
    const types = memoType === 'all' ? ['new', 'rematch'] : [memoType];
    for (const type of types) {
      const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
      const tabName = getUserMemoTabName(type);
      const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tabName}'!A1:BZ1` });
      const header = (headerRes.data.values || [])[0] || [];
      const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
      const matchIdCol = type === 'new' ? 'match_id' : 'match ID';
      const statusCol = type === 'new' ? '매칭상태' : 'status';
      const matchIdIdx = header.findIndex(h => norm(h) === norm(matchIdCol));
      const statusIdx = header.findIndex(h => norm(h) === norm(statusCol));
      const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tabName}'!A2:BZ` });
      const rows = dataRes.data.values || [];
      rows.forEach(row => {
        const matchId = matchIdIdx >= 0 ? (row[matchIdIdx] || '').trim() : '';
        if (!matchId) return;
        const sn = statusIdx >= 0 ? (row[statusIdx] || '').replace(/\s+/g, '').toLowerCase() : '';
        if (memoFilter === 'initial') {
          if (sn === '' || sn === 'ready') counts[type]++;
        } else {
          if (type === 'new') {
            if (sn === '' || sn === '매칭중' || sn === '확인필요') counts[type]++;
          } else {
            if (['', 'ready', '첫수업전재매칭', '일반재매칭', '매칭중', '보류', '확인필요'].includes(sn)) counts[type]++;
          }
        }
      });
    }
    return counts;
  } catch (e) {
    console.error(`[countMemoTargets] 에러: ${e.message}`);
    return { new: 0, rematch: 0 };
  }
}

// --- 워크플로 대기 작업 ---
let wfPendingTask = null;

app.get('/api/wf/pending', (req, res) => {
  if (wfPendingTask) {
    res.json({ pending: true, type: wfPendingTask.type, filter: wfPendingTask.filter, editing: wfPendingTask.editing || false });
  } else {
    res.json({ pending: false });
  }
});

app.post('/api/wf/pending/register', (req, res) => {
  if (wfPendingTask) {
    wfPendingTask.registered = (wfPendingTask.registered || 0) + 1;
    console.log(`[워크플로] 유저메모 랜딩 열림 (등록: ${wfPendingTask.registered})`);
  }
  res.json({ ok: true });
});

app.post('/api/wf/pending/unregister', (req, res) => {
  if (wfPendingTask && wfPendingTask.registered > 0) {
    wfPendingTask.registered--;
  }
  res.json({ ok: true });
});

app.post('/api/wf/pending/accept', (req, res) => {
  if (!wfPendingTask) return res.json({ ok: false, message: '대기 중인 작업 없음' });
  if (wfPendingTask.editing) return res.json({ ok: false, message: '다른 기기에서 이미 편집을 시���했습니다.' });
  wfPendingTask.editing = true;
  const editor = req.body && req.body.name || req.headers['x-forwarded-for'] || req.connection.remoteAddress || '알 수 ��음';
  wfPendingTask.editor = editor;
  console.log(`[워크플로] 유저메모 편집 수락 — ${editor}`);
  res.json({ ok: true });
});

app.post('/api/wf/pending/reject', (req, res) => {
  if (wfPendingTask) {
    wfPendingTask.rejected = (wfPendingTask.rejected || 0) + 1;
    if (wfPendingTask.registered > 0) wfPendingTask.registered--;
    console.log(`[워크플로] 유저메모 편집 거절 (거절: ${wfPendingTask.rejected}, 남은 등록: ${wfPendingTask.registered})`);
    // 모두 거절 → 대기 취소
    if (wfPendingTask.registered <= 0 && !wfPendingTask.editing) {
      console.log('[워크플로] 모든 사용자 거절 → 유저메모 편집 건너뜀');
      if (wfPendingTask.resolve) wfPendingTask.resolve();
      wfPendingTask = null;
    }
  }
  res.json({ ok: true });
});

app.post('/api/wf/pending/complete', (req, res) => {
  if (wfPendingTask && wfPendingTask.resolve) {
    console.log('[워크플로] 유저메모 편집 완료 → 워크플로 재개');
    wfPendingTask.resolve();
    wfPendingTask = null;
    res.json({ ok: true });
  } else {
    res.json({ ok: false, message: '대기 중인 작업 없음' });
  }
});

async function countSheetStatus(type, condStatus) {
  const sheets = await getSheetsApi();
  const spreadsheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const tabName = type === 'new' ? `[신규] ${yy}.${mm}` : `[재매칭] ${yy}.${mm}`;
  const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tabName}'!A1:BZ1` });
  const header = (headerRes.data.values || [])[0] || [];
  const norm = s => (s || '').replace(/\s+/g, '').toLowerCase();
  const matchIdIdx = header.findIndex(h => norm(h) === norm(type === 'new' ? 'match_id' : 'match ID'));
  const statusIdx = header.findIndex(h => norm(h) === norm(type === 'new' ? '매칭상태' : 'status'));
  const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tabName}'!A2:BZ` });
  const rows = dataRes.data.values || [];
  const targetNorm = norm(condStatus || '');
  let count = 0;
  rows.forEach(row => {
    const matchId = matchIdIdx >= 0 ? (row[matchIdIdx] || '').trim() : '';
    if (!matchId) return;
    const status = statusIdx >= 0 ? (row[statusIdx] || '').trim() : '';
    if (norm(status) === targetNorm) count++;
  });
  return count;
}

async function checkWfCondition(cond) {
  try {
    const type = cond.condType || 'new';
    let count;
    if (type === 'both') {
      const [n, r] = await Promise.all([countSheetStatus('new', cond.condStatus), countSheetStatus('rematch', cond.condStatus)]);
      count = n + r;
    } else {
      count = await countSheetStatus(type, cond.condStatus);
    }
    const op = cond.condOp || '<=';
    const val = parseInt(cond.condValue) || 0;
    if (op === '<=') return count <= val;
    if (op === '>=') return count >= val;
    return count === val;
  } catch (e) { console.error('[워크플로] 조건 체크 에러:', e.message); }
  return false;
}

function registerWfJob(id, name, workspace) {
  JOBS['wf:' + id] = {
    name: `워크플로: ${name}`,
    fn: () => runWfBlocks(workspace),
  };
}

// 저장된 프리셋을 시작 시 등록
for (const [id, preset] of Object.entries(wfPresetsData)) {
  registerWfJob(id, preset.name, preset.workspace);
}

// 저장된 스케줄 복원
loadSchedulesFromFile();

app.post('/api/wf/presets/:id', (req, res) => {
  const { id } = req.params;
  const { name, workspace } = req.body;
  wfPresetsData[id] = { name, workspace };
  saveWfPresets();
  registerWfJob(id, name, workspace);
  res.json({ ok: true });
});

app.delete('/api/wf/presets/:id', (req, res) => {
  const { id } = req.params;
  delete wfPresetsData[id];
  saveWfPresets();
  delete JOBS['wf:' + id];
  Object.values(schedules).forEach(s => {
    if (s.jobId === 'wf:' + id) removeSchedule(s.id);
  });
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`서버 시작: http://localhost:${PORT}`);
});

process.on('SIGINT', async () => {
  await disconnect();
  process.exit(0);
});

const cron = require('node-cron');
const config = require('./config.json');
const { runPairing } = require('./automation/pairing');
const { runStatusCheck } = require('./automation/statusCheck');
const { prepNewSheet, prepRematchSheet, createMonthlyTab } = require('./automation/sheetPrep');
const { disconnect } = require('./automation/browser');

// --- 수동 실행 ---
// node index.js --run prep          시트 정리 (신규+재매칭)
// node index.js --run prep:new      신규 시트 정리만
// node index.js --run prep:rematch  재매칭 시트 정리만
// node index.js --run monthly       월별 탭 생성 (전월 잔여건 이월)
// node index.js --run pairing       페어링 생성
// node index.js --run statusCheck   상태 확인
// node index.js --run all           시트 정리 → 페어링 → 상태 확인
const args = process.argv.slice(2);
const runIdx = args.indexOf('--run');

if (runIdx !== -1) {
  const task = args[runIdx + 1] || 'all';
  (async () => {
    try {
      const mode = config.mode;

      // 시트 정리
      if (task === 'prep' || task === 'prep:new' || task === 'all') {
        await prepNewSheet();
      }
      if (task === 'prep' || task === 'prep:rematch' || task === 'all') {
        await prepRematchSheet();
      }

      // 월별 탭 생성 (월초 1회)
      if (task === 'monthly') {
        if (mode === '1' || mode === 'both') await createMonthlyTab('new');
        if (mode === '2' || mode === 'both') await createMonthlyTab('rematch');
      }

      // 페어링 생성
      if (task === 'pairing' || task === 'all') {
        await runPairing(mode);
      }

      // 상태 확인
      if (task === 'statusCheck' || task === 'all') {
        await runStatusCheck(mode);
      }
    } catch (err) {
      console.error('[index] 실행 에러:', err.message);
    } finally {
      await disconnect();
      process.exit(0);
    }
  })();
} else {
  // --- 스케줄 모드 ---
  console.log('[index] 스케줄 모드 시작');
  console.log(`[index] 모드: ${config.mode}`);

  const { pairing, statusCheck } = config.jobs;

  if (pairing.enabled && pairing.schedule) {
    cron.schedule(pairing.schedule, async () => {
      console.log(`\n[cron] pairing 실행 (${new Date().toLocaleString('ko-KR')})`);
      try {
        await runPairing(config.mode);
      } catch (err) {
        console.error('[cron] pairing 에러:', err.message);
      }
    });
    console.log(`[index] pairing 스케줄 등록: ${pairing.schedule}`);
  }

  if (statusCheck.enabled && statusCheck.schedule) {
    cron.schedule(statusCheck.schedule, async () => {
      console.log(`\n[cron] statusCheck 실행 (${new Date().toLocaleString('ko-KR')})`);
      try {
        await runStatusCheck(config.mode);
      } catch (err) {
        console.error('[cron] statusCheck 에러:', err.message);
      }
    });
    console.log(`[index] statusCheck 스케줄 등록: ${statusCheck.schedule}`);
  }

  console.log('[index] 대기 중... (Ctrl+C로 종료)');

  // 깔끔한 종료
  process.on('SIGINT', async () => {
    console.log('\n[index] 종료 중...');
    await disconnect();
    process.exit(0);
  });
}

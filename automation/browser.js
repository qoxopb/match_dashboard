const { chromium } = require('playwright');

let browser = null;
let context = null;

async function connect() {
  if (browser && browser.isConnected()) return context;

  browser = await chromium.connectOverCDP('http://127.0.0.1:9222');
  context = browser.contexts()[0];

  if (!context) {
    throw new Error('Chrome에 열린 컨텍스트가 없습니다. 브라우저에서 아무 탭이나 열어주세요.');
  }

  console.log('[browser] CDP 연결 성공');
  return context;
}

async function newPage() {
  const ctx = await connect();
  return ctx.newPage();
}

// 백그라운드 탭(포커스 안 뺏음) 생성 — CDP Target.createTarget의 background:true 사용
async function newBackgroundPage() {
  const ctx = await connect();
  const existingPages = ctx.pages();

  // 기존 페이지가 있어야 CDP session을 만들 수 있음
  let session;
  if (existingPages.length > 0) {
    session = await ctx.newCDPSession(existingPages[0]);
  } else {
    // 페이지가 하나도 없으면 일반 newPage로 fallback
    return ctx.newPage();
  }

  await session.send('Target.createTarget', {
    url: 'about:blank',
    background: true,
  });

  // 새 페이지가 컨텍스트에 등록될 때까지 대기
  for (let i = 0; i < 30; i++) {
    await new Promise(r => setTimeout(r, 100));
    const pages = ctx.pages();
    if (pages.length > existingPages.length) {
      // 가장 최근 추가된 페이지 반환
      const newOnes = pages.filter(p => !existingPages.includes(p));
      if (newOnes.length > 0) return newOnes[0];
    }
  }
  throw new Error('백그라운드 페이지 생성 실패');
}

async function disconnect() {
  if (browser) {
    browser.close();
    browser = null;
    context = null;
  }
}

module.exports = { connect, newPage, newBackgroundPage, disconnect };

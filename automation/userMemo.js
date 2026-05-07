const { newBackgroundPage } = require('./browser');
const axios = require('axios');
const config = require('../config.json');

const ADMIN_URL = config.adminUrl + '/admin/tutor-pairing';

let slackChannelName = null;

// 서버 시작 시 채널명 미리 캐시
(async () => {
  try {
    const res = await axios.get('https://slack.com/api/conversations.info', {
      headers: { Authorization: `Bearer ${config.slack.botToken}` },
      params: { channel: config.slack.channelId },
    });
    if (res.data.ok) slackChannelName = res.data.channel.name;
  } catch {}
})();

async function closePage() {
  // no-op — 매번 새 페이지 생성/닫기 방식
}

// --- 어드민 스크래핑 (pairing.js와 동일 방식) ---
async function scrapeAdmin(matchId) {
  const empty = { counselMemo: '', userMemo: '', createdAt: '' };
  const page = await newBackgroundPage();

  try {
    // 1) 검색 페이지
    await page.goto(
      `${ADMIN_URL}?page=1&size=20&filters_matchId=${encodeURIComponent(matchId)}`,
      { waitUntil: 'domcontentloaded', timeout: 30000 }
    );
    await page.waitForSelector('table tbody, .ant-table-tbody', { timeout: 8000 });

    // 2) 상세 페이지 href 추출 → 이동
    const detailHref = await page.evaluate(() => {
      const link = document.querySelector('table tbody a[href*="/tutor-pairing/"]');
      return link ? link.getAttribute('href') : null;
    });
    if (!detailHref) {
      return empty;
    }

    await page.goto(
      new URL(detailHref, config.adminUrl).href,
      { waitUntil: 'domcontentloaded', timeout: 30000 }
    );

    // 3) __NEXT_DATA__ 먼저 읽기 (페이지 로드 직후 존재, 빠름)
    try { await page.waitForSelector('#__NEXT_DATA__', { state: 'attached', timeout: 5000 }); } catch {}
    const nextData = await page.evaluate(() => {
      const script = document.getElementById('__NEXT_DATA__');
      if (!script) return null;
      try {
        const json = JSON.parse(script.textContent);
        const pp = json?.props?.pageProps;
        if (!pp) return null;
        return {
          pageProps: pp,
          applicant: pp.applicant || null,
          product: pp.product || null,
        };
      } catch { return null; }
    });
    // __NEXT_DATA__에서 대부분의 데이터 추출 (빠름)
    const app = (nextData && nextData.applicant) || {};
    const prod = (nextData && nextData.product) || {};
    const pp = (nextData && nextData.pageProps) || {};


    const userId = String(app.userId || '');
    const userName = app.userName || '';
    const createdAt = app.createdAt || '';
    const kakaoTalkId = app.kakaoTalkId || '';
    const isRematch = !!app.isRematch;
    const tutorMemo = app.tutorMemo || '';
    // userMemos = 학생이 작성한 메모 배열, 각 항목은 { tutorMemo, createdAt }
    // 키 이름이 tutorMemo지만 실제로는 학생 작성 본문임
    let userMemo = '';
    if (Array.isArray(app.userMemos) && app.userMemos.length > 0) {
      userMemo = app.userMemos
        .filter(m => m && (m.tutorMemo || '').trim())
        .map(m => `${m.createdAt || ''}\n${m.tutorMemo}`)
        .join('\n─────────────────────\n');
    }
    const lessonInfo = prod.name
      ? (prod.name.replace(/\s*수업\s*/, '').replace(/,\s*/g, ' ') + (prod.totalMonths ? ` ${prod.totalMonths}개월` : ''))
      : '';

    // 4) 상담메모만 innerText에서 스크래핑 (Show 버튼 클릭 필요)
    try {
      await page.waitForFunction(
        () => document.body.innerText.includes('상담 메모'),
        { timeout: 5000 }
      );
    } catch {}

    const showBtns = page.locator('button:has-text("Show")');
    const showCount = await showBtns.count();
    for (let i = 0; i < showCount; i++) {
      try { await showBtns.nth(i).click({ force: true, timeout: 1000 }); } catch {}
    }

    const counselMemo = await page.evaluate(() => {
      const text = document.body.innerText;
      let memo = '';
      const allCounselMatches = [...text.matchAll(/상담\s*메모/g)];
      const ceIdx = text.indexOf('매칭 제안 전송 리스트');
      if (allCounselMatches.length > 0) {
        const lastMatch = allCounselMatches[allCounselMatches.length - 1];
        const csIdx = lastMatch.index;
        const headerEnd = text.indexOf('\n', csIdx);
        if (headerEnd > 0) {
          const raw = text.substring(headerEnd, ceIdx > csIdx ? ceIdx : headerEnd + 5000).trim();
          const lines = raw.split('\n');
          const entries = [];
          let i = 0;
          while (i < lines.length) {
            const l = lines[i].trim();
            const tsM = l.match(/^(?:Updated|Created)\s*At\s*:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})/i);
            if (tsM) {
              let author = '';
              for (let j = i - 1; j >= 0; j--) {
                const prev = lines[j].trim();
                if (/^(Edit|Submit|Show|Hide|상담\s*메모)$/i.test(prev)) continue;
                if (/^\d{1,2}$/.test(prev)) continue;
                if (/^(Updated|Created)\s*At/i.test(prev)) break;
                author = prev;
                break;
              }
              const bodyLines = [];
              for (let j = i + 1; j < lines.length; j++) {
                const next = lines[j].trim();
                if (/^(Updated|Created)\s*At\s*:/i.test(next)) break;
                if (/^(Edit|Submit|Show|Hide)$/i.test(next)) continue;
                if (/^\d{1,2}$/.test(next)) continue;
                if (j + 1 < lines.length && /^(Updated|Created)\s*At\s*:/i.test(lines[j + 1].trim())) break;
                bodyLines.push(next);
              }
              entries.push({ author, ts: tsM[1], body: bodyLines.join('\n').trim() });
            }
            i++;
          }
          memo = entries.map(e =>
            `${e.author}${e.author ? ' | ' : ''}${e.ts}\n${e.body}`
          ).join('\n─────────────────────\n');
        }
      }
      return memo;
    });

    return {
      userId, userName, createdAt, kakaoTalkId, isRematch,
      tutorMemo, userMemo, counselMemo, lessonInfo,
      applicant: app, product: prod,
      detailUrl: page.url(),
    };
  } finally {
    await page.close().catch(() => {});
  }
}

// --- 튜터메모 어드민 저장 ---
async function saveTutorMemo(matchId, tutorMemo) {
  const page = await newBackgroundPage();

  try {
    // 1) 검색 → 상세 페이지 href 추출
    await page.goto(
      `${ADMIN_URL}?page=1&size=20&filters_matchId=${encodeURIComponent(matchId)}`,
      { waitUntil: 'domcontentloaded', timeout: 30000 }
    );
    await page.waitForSelector('table tbody, .ant-table-tbody', { timeout: 8000 });

    const detailHref = await page.evaluate(() => {
      const link = document.querySelector('table tbody a[href*="/tutor-pairing/"]');
      return link ? link.getAttribute('href') : null;
    });
    if (!detailHref) throw new Error('상세 링크 못 찾음');

    // 2) 상세 페이지 이동
    await page.goto(
      new URL(detailHref, config.adminUrl).href,
      { waitUntil: 'domcontentloaded', timeout: 30000 }
    );

    // 3) "신청 내용" 근처 Edit 버튼 클릭 (JS로 직접)
    await page.evaluate(() => {
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && /^신청\s*내용$/.test(el.textContent.trim())) {
          let parent = el.parentElement;
          for (let d = 0; d < 5 && parent; d++) {
            const btn = parent.querySelector('button');
            if (btn) { btn.click(); return; }
            parent = parent.parentElement;
          }
        }
      }
    });

    // 페이지 변경 대기
    await page.waitForFunction(
      () => document.body.innerText.includes('tutorMemo'),
      { timeout: 10000 }
    );
    await page.waitForTimeout(500);

    // 4) tutorMemo 라벨 옆 textarea — JS로 값 덮어쓰기
    const fillResult = await page.evaluate((memo) => {
      // tutorMemo 라벨 찾기
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        if (el.children.length === 0 && el.textContent.trim() === 'tutorMemo') {
          // DOM 순서로 이후 textarea 찾기
          const allAfter = [];
          let found = false;
          document.querySelectorAll('*').forEach(n => {
            if (n === el) found = true;
            if (found && n.tagName === 'TEXTAREA') allAfter.push(n);
          });
          if (allAfter.length > 0) {
            const ta = allAfter[0];
            const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
            setter.call(ta, memo);
            ta.dispatchEvent(new Event('input', { bubbles: true }));
            ta.dispatchEvent(new Event('change', { bubbles: true }));
            return 'ok';
          }
        }
      }
      // 디버그 정보
      const tas = document.querySelectorAll('textarea');
      return `not_found|textareas:${tas.length}|url:${location.href}`;
    }, tutorMemo);

    if (!fillResult.startsWith('ok')) throw new Error(`tutorMemo 필드 못 찾음 (${fillResult})`);

    // 5) Submit 클릭
    const submitBtn = page.locator('button:has-text("Submit"), button[type="submit"]').first();
    await submitBtn.click({ force: true });

    // 6) 확인 모달 Yes 클릭
    await page.waitForTimeout(500);
    const yesBtn = page.locator('.ant-modal button:has-text("Yes"), .ant-popconfirm button:has-text("Yes")').first();
    try {
      await yesBtn.waitFor({ timeout: 3000 });
      await yesBtn.click({ force: true });
    } catch {}

    await page.waitForTimeout(1000);
    return { ok: true };
  } finally {
    await page.close().catch(() => {});
  }
}

// --- Slack 검색 (conversations.history 기반) ---
async function fetchSlackRequests(matchId, createdAt) {
  const token = config.slack.botToken;
  const channelId = config.slack.channelId;
  const searchStr = String(matchId);

  // createdAt 이후 타임스탬프 계산
  let oldest = undefined;
  if (createdAt) {
    try {
      const parsed = createdAt.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
      if (parsed) {
        oldest = String(new Date(+parsed[1], +parsed[2] - 1, +parsed[3]).getTime() / 1000);
      }
    } catch {}
  }

  // 최대 3페이지 (200*3=600 메시지)까지 조회
  const matches = [];
  let cursor = undefined;
  for (let page = 0; page < 3; page++) {
    const params = { channel: channelId, limit: 200 };
    if (oldest) params.oldest = oldest;
    if (cursor) params.cursor = cursor;

    const res = await axios.get('https://slack.com/api/conversations.history', {
      headers: { Authorization: `Bearer ${token}` },
      params,
    });

    if (!res.data.ok) {
      console.log(`[userMemo] Slack API 에러: ${res.data.error}`);
      return [];
    }

    const msgs = res.data.messages || [];
    for (const m of msgs) {
      if (m.text && m.text.includes(searchStr)) {
        matches.push(m);
      }
    }

    // 다음 페이지
    if (res.data.has_more && res.data.response_metadata?.next_cursor) {
      cursor = res.data.response_metadata.next_cursor;
    } else {
      break;
    }
  }

  return matches.map(m => ({
    text: m.text,
    permalink: `${config.slack.workspaceUrl}/archives/${channelId}/p${m.ts.replace('.', '')}`,
    date: new Date(parseFloat(m.ts) * 1000).toLocaleString('ko-KR'),
  }));
}

// --- 통합: 어드민 + Slack 병렬 실행 ---
async function fetchDetail(matchId) {
  // 어드민 스크래핑과 Slack 검색을 병렬로 시작
  // 단, Slack은 createdAt 필터가 필요하므로 어드민 없이도 먼저 검색 (날짜 필터 없이)
  // 어드민 결과 나오면 날짜로 재필터

  const [adminResult, slackRaw] = await Promise.all([
    scrapeAdmin(matchId),
    fetchSlackRequests(matchId, null).catch(() => []),
  ]);

  // createdAt으로 Slack 결과 필터
  let slackFiltered = slackRaw;
  if (adminResult.createdAt && slackRaw.length > 0) {
    const parsed = adminResult.createdAt.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (parsed) {
      const afterDate = new Date(+parsed[1], +parsed[2] - 1, +parsed[3]);
      slackFiltered = slackRaw.filter(m => {
        const msgDate = new Date(m.date);
        return !isNaN(msgDate) && msgDate >= afterDate;
      });
    }
  }

  return {
    userId: adminResult.userId || '',
    userName: adminResult.userName || '',
    lessonInfo: adminResult.lessonInfo || '',
    counselMemo: adminResult.counselMemo || '',
    userMemo: adminResult.userMemo || '',
    tutorMemo: adminResult.tutorMemo || '',
    createdAt: adminResult.createdAt || '',
    isRematch: adminResult.isRematch || false,
    kakaoTalkId: adminResult.kakaoTalkId || '',
    detailUrl: adminResult.detailUrl || '',
    applicant: adminResult.applicant || null,
    product: adminResult.product || null,
    request: slackFiltered,
    error: adminResult.error,
  };
}

module.exports = { fetchDetail, saveTutorMemo, closePage };

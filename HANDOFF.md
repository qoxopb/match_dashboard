# 콴다과외 매칭 자동화 — 핸드오프

> Claude Code에서 작업한 내역. claude.ai에 붙여넣어서 이어 작업하세요.

## 개요

로컬 Node.js 대시보드. 콴다과외 매칭 업무 자동화 (시트 정리, 페어링 생성, 매칭 제안, 매칭결과 확인, 유저메모 편집). 브라우저로 `http://localhost:3000` 접속해서 사용.

**폴더 구조:**
```
matching-dashboard/
├── win/     ← Windows 버전 (start.vbs로 실행)
├── mac/     ← Mac 버전 (start.sh로 실행)
└── HANDOFF.md
```

각 폴더 안에 동일한 코드. OS별 시작 스크립트만 다름.
- **Windows**: `start.vbs` 더블클릭 또는 `start.ps1` 실행
- **Mac**: `chmod +x start.sh && ./start.sh`
- 시작 스크립트가 node_modules 없으면 자동으로 `npm install` 실행 (Node.js만 설치되어 있으면 됨)

**스택**
- Express 5
- Playwright (CDP로 기존 크롬에 attach — `--remote-debugging-port=9222`)
- googleapis (Google Sheets)
- node-cron (스케줄러)
- node-notifier (알림)
- axios (Slack API)
- Vanilla HTML/JS 대시보드

## 파일 구조

```
matching-dashboard/
├── server.js              # Express 서버, 작업/스케줄 관리, REST API
├── config.json            # 모드, 시트 ID, GAS URL, Slack 토큰
├── sheets.js              # Google Sheets 헬퍼
├── abort.js               # 중단 요청 플래그 + activePage 관리
├── credentials.json       # Google API credentials (gitignore)
├── oauth-token.json       # OAuth 토큰 (gitignore)
├── automation/
│   ├── browser.js         # CDP 크롬 attach, newBackgroundPage (포커스 안 뺏는 백그라운드 탭)
│   ├── sheetPrep.js       # prepNewSheet, prepRematchSheet, createMonthlyTab
│   ├── pairing.js         # runPairing(mode) — 5워커 병렬
│   ├── statusCheck.js     # runStatusCheck(mode) — 5워커 병렬
│   ├── aim.js             # runAim(mode) — 매칭 제안 (AIM + 수동 + 임의완화 3단계)
│   └── userMemo.js        # 유저메모 어드민 스크래핑 + 저장 + Slack 검색
├── public/
│   ├── index.html         # 메인 대시보드 (사이드바 + 기능 카드)
│   └── userMemo.html      # 유저메모 편집 전용 페이지
├── start.vbs / start.ps1 / start.bat  # Windows 자동 실행 스크립트
└── install-startup.bat / install-startup.ps1
```

## 주요 작업 (JOBS)

`server.js`의 `JOBS` 객체에 정의됨:

| jobId | 이름 | 설명 |
|---|---|---|
| `prep` / `prep:new` / `prep:rematch` | 시트 정리 | sheetPrep |
| `pairing` | 페어링 생성 | runPairing |
| `aim` / `aim:new` / `aim:rematch` | 매칭 제안 (AIM) | runAim — 자동 AI 매칭 + 수동매칭 + 임의완화 |
| `statusCheck` / `statusCheck:new` / `statusCheck:rematch` | 매칭결과 확인 | runStatusCheck |
| `monthly` | 월별 탭 생성 | createMonthlyTab |

`config.mode` 가 `"1"` / `"2"` / `"both"` — 현재 `"both"`.

## 매칭 제안 (AIM) 핵심 로직 — `automation/aim.js`

매치아이디 1건당 처리 흐름:

1. **어드민 status 확인**
   - `신청서 미작성` → 스킵
   - `매칭 완료` → 시트에 반영 후 다음
   - `AIM 진행중` → 매칭 중단 후 재시도

2. **1단계 AIM** — `tryAim()`
   - "AI 매칭 검색" 버튼 → "AIM 매칭 시작" 버튼 클릭
   - "제안서 발송 대상이 없습니다" 모달 체크
   - 성공 시 시트 status → `매칭중`

3. **2단계 수동매칭** — `tryManualMatch(relax='base')`
   - product name에서 시수 파싱 (주 N회, NN분)
   - 학생 시간표 (어드민 `weeklyAvailablePeriods`) vs 선생님 timeline 비교
   - 시간 매칭 성공한 선생님에게 "전송하기" 클릭 → 모달의 "매칭 제안하기" 클릭
   - **EXPERT면 "전문 강사 여부" 토글 ON** — 매 검색 직전 확인 (이미 ON이면 건드리지 않음)
   - 학생당 최대 2명까지 제안

4. **3단계 임의완화** (수동매칭 실패 시)
   - **relax1**: 입시 전형 정시+기타 추가, 튜터 스타일 삭제
   - **relax2**: 성적대 해제
   - **relax3**: 고3 가능 과목 → 과목으로 이동

5. **시수 파싱 실패 / 희망시간 불일치** → status `확인 필요` + 메모 컬럼에 사유 기록

## 유저메모 편집 — `public/userMemo.html` + `automation/userMemo.js`

`?type=new` 또는 `?type=rematch` 쿼리로 신규/재매칭 분기.

**목록 필터링** (`/api/userMemo/list`):
- 신규: `매칭중`, `확인 필요` 만
- 재매칭: `첫수업전재매칭`, `일반 재매칭`, `매칭중`, `보류`, `확인 필요` 만

**상세 데이터 (어드민 `__NEXT_DATA__` 기반)**:
- `pageProps.applicant` — 학생 정보, 시간표, tutorStyles 등
- `pageProps.product` — 수업명, 시수, 기간
- `pageProps.user`, `pageProps.match`, `pageProps.pairing` — 추가 정보
- `applicant.userMemos[]` — 학생 작성 메모 (각 항목 `{tutorMemo, createdAt}` ← 키 이름 헷갈리지만 학생 작성)
- `applicant.tutorMemo` — 운영팀이 어드민에 저장한 튜터메모 본문

**수업신청서 표시 필드** (어드민 `applicant.tutorStyles` 기반):
- `created_at_kst` ← `applicant.createdAt`
- `tutor_university_level` ← `tutorStyles.tutorUniversities` (HIGHEST→SKY, HIGHER→서성한 등)
- `tutor_genders` ← `tutorStyles.tutorGenders` (MALE→남자)
- `tutor_pass_types` ← `tutorStyles.tutorPassTypes` (EARLY_DECISION→수시)
- `tutor_tracks` ← `tutorStyles.tutorTracks` (LIBERAL_ARTS→문과)
- `liked_tutor_styles` ← `tutorStyles.likedTutorStyles` (WARM→따뜻한 선생님)
- `living_abroad` ← `applicant.livingAbroad` (해외거주 시 강조 + 카카오톡 ID 표시)
- 가능 시간 ← `applicant.weeklyAvailablePeriods` (요일별 시간 배열)

**상단 뱃지**:
- 타입 뱃지: `PRO 신규/재매칭` (rank=EXPERT) 또는 `일반 신규/재매칭`
- status 뱃지: 시트 status 그대로 표시 (매칭중=주황, 매칭완료=초록, 확인필요/보류=빨강, 미처리=파랑)

**참고 자료 3열**:
1. 상담메모 (innerText 스크래핑, Show 버튼 클릭 필요)
2. 튜터메모(어드민 현재값) + 유저메모 (위/아래로)
3. Request (Slack 메시지)

**저장**: `saveTutorMemo(matchId, tutorMemo)` — 어드민 신청 내용 Edit → tutorMemo textarea 덮어쓰기 → Submit → Yes

## 재매칭 status 형식 (변경됨)

기존 `[3]재매칭완료`, `[5]환불/소멸` 등 번호 접두사 형식에서 → 번호 없는 일반 텍스트로 변경됨:
- `첫수업전재매칭`
- `일반 재매칭`
- `매칭중`
- `재매칭 완료`
- `보류`
- `환불/소멸`
- `확인 필요`

영향 받은 파일 (전부 수정 완료):
- `aim.js`: `excludeRematch`, `completeStatus`, `checkStatus`
- `statusCheck.js`: `excludeRematch`, `completedStatus`
- `sheetPrep.js`: `excludeRematch` (두 군데)
- `userMemo.html`: status 뱃지 색상 매핑

## 스케줄러

`server.js`에 in-memory 스케줄 관리. 두 가지 모드:

- **interval**: `*/N * * * *` 형태. 5/10/20/30/40/50/60/90/120분 chip
- **alarm**: 특정 시각 + 요일. `MM HH * * days` 형태

**서버 재시작 시 사라짐** (영속화 안 되어 있음).

## REST API

- `POST /api/run/:jobId` — 즉시 실행
- `POST /api/stop` — 현재 작업 중단 요청
- `POST /api/schedule` `{jobId, type, options}` — 등록
- `DELETE /api/schedule/:id` — 삭제
- `PATCH /api/schedule/:id` `{enabled}` — ON/OFF
- `GET /api/schedules` — 목록
- `GET /api/logs` — `{running, logs[]}` (최대 500줄)
- `GET /api/status` — `{running, mode}`
- `GET /api/userMemo/list?type=new|rematch` — 처리 대상 목록
- `GET /api/userMemo/detail?matchId=...` — 어드민 + Slack 병렬 조회
- `POST /api/userMemo/save` — 어드민에 튜터메모 저장
- `GET /api/debug/nextdata/:matchId` — 어드민 상세 페이지의 `__NEXT_DATA__` 전체 덤프 (디버깅용)
- `GET /api/debug/pairing/:matchId?relax=0~3` — 페어링 페이지 필터 디버깅

## 대시보드 UI (`public/index.html`)

사이드바 + 메인 컨텐츠 SPA. 메뉴: **기능** / **당일 현황** / **대시보드**

**기능 페이지** (현재 유일하게 컨텐츠 있음):
- 카드: 시트 정리 / 페어링 생성 / 유저메모 편집하기 (새 창) / 매칭 제안하기 (AIM) / 매칭결과 확인 / 월별 탭 관리
- 각 카드: 즉시 실행 버튼 (신규/재매칭/전체) + 스케줄 설정 (접기/펼치기)
- 실행 로그 박스 (3초마다 폴링)

**당일 현황 / 대시보드**: 빈 페이지

## Windows 자동 실행

- `start.vbs` → `start.ps1` 을 hidden PowerShell로 실행
- `start.ps1`:
  1. 9222 안 열려있으면 크롬 `--remote-debugging-port=9222 --user-data-dir=C:\chrome-debug --start-minimized` 실행
  2. 3000 안 열려있으면 `node server.js` hidden 프로세스
  3. `http://localhost:3000` 브라우저로 열기
  4. 토스트 알림 (UTF-8 BOM 필수 — cp949 디코딩 방지)

## 주요 이슈/결정 사항

### 1. CDP attach 방식
사용자 로그인 세션 유지를 위해 이미 떠있는 크롬에 attach. 토스트에 "최소화된 크롬 브라우저를 닫지 마세요" 안내.

### 2. `newBackgroundPage()` — 포커스 안 뺏기
CDP `Target.createTarget` 의 `background:true` 옵션 사용. 새 탭이 활성화되지 않아서 사용자 작업 방해 안 함.

### 3. `__NEXT_DATA__` selector
스크립트 태그라 hidden 상태. `waitForSelector('#__NEXT_DATA__', { state: 'attached' })` 필수 (`visible` 쓰면 timeout).

### 4. 로그가 터미널에 안 나옴
`server.js` 에서 `console.log` 를 override해서 메모리에만 저장. 대시보드 UI에서만 확인 가능.

### 5. EXPERT 토글
수동매칭에서 `applicant.rank === 'EXPERT'` 면 매 단계(base/relax1/2/3) 검색 직전 "전문 강사 여부" 토글 ON. 이미 ON이면 그대로 유지 (반복 클릭으로 OFF 되지 않게).

### 6. 시트 ID
`config.json`:
- `new`: `1EM5v9fzdkm_YMV07hXcUELAZWQsMmt8OP5QimtIXjvs`
- `rematch`: `1rMEnyGVtYUnBXbY1enPLIQYhUDewAqKziwTLuW-PebQ`

### 7. 시수는 어드민 product name에서 파싱
시트가 아니라 어드민 `product.name` ("주 2회, 60분 수업") 에서 정규식으로 파싱. 시트 의존 X.

## 다음 할 일

1. **AIM 모든 단계 실패 시 조건완화 발송** — `aim.js:339` TODO
2. **당일 현황 페이지** — 오늘 실행 횟수, 처리 건수, 결과 요약
3. **대시보드 페이지** — 그래프/통계
4. (선택) 스케줄 영속화
5. (선택) AI 편집 기능 (`/api/userMemo/ai-edit` placeholder)

---

**가장 최근 변경 (2026-04-21)**:
- **Win/Mac 이중 폴더 구조**: matching-dashboard/win + mac 으로 분리, 동일 코드 + OS별 시작 스크립트
- **AIM 키워드 설정**: 대시보드에서 키워드 등록 → exclude(제외) 또는 tutorPool(특정튜터풀) 액션. exclude가 항상 우선.
- **특정튜터풀 매칭**: 키워드 매칭 시 "이름 및 아이디" ID 필드에 튜터ID 붙여넣기(clipboard paste) → Search → 시간+성별 매칭
- **교대/메디컬 포함 시**: AIM 건너뛰기 → 변형 필터(교대/메디컬 해제, SKY/서성한/중경외시 추가) → 실패 시 원래 조건 복귀
- **PRO 마지막 시도**: 모든 매칭 실패 후 PRO면 교대/메디컬 해제 + SKY/서성한/중경외시 추가 매칭 한번 더
- **전문강사 토글(isExpert)**: `applicant.tutorRank` (어드민 __NEXT_DATA__) 기반으로 변경 (시트 의존 제거)
- **조건완화 필요 시 시트 업데이트**: status "확인 필요" + 메모 "조건완화 필요"
- **sendProposal 재작성**: mouse.click 기반, 체크박스 1회만 클릭, orderStatus 에러 감지 (notification/message/modal 전체 스캔)
- **재매칭 pairingId 검색**: aim.js + statusCheck.js 모두 재매칭은 filters_id 사용
- **유저메모 편집**: 전체/초기편집 토글 (기본: 초기편집), 전체 버튼 시 신규→재매칭 자동 이동, 시트메모 표시
- **Slack API**: search.messages → conversations.history(botToken) 로 변경 (search:read scope 불필요)
- **relax3 과목 추가**: "고3 가능 과목" 바로 아래의 "과목" select를 정확히 찾도록 수정

**이전 변경 (2026-04-17)**:
- **시간표 자동 확장**: 시수-시간대 불일치 시 요일 ��가 주N회와 맞으면 각 요일의 시작 시간부터 필요한 타임블록만큼 확장 → 어드민 weeklyAvailablePeriods 자동 수정 (`fixAdminSchedule`)
  - 예: 주3회 90분, 수/금/일 21시만 선택 → 수/금/일 21~22시로 확장
  - ant-select 드롭다운 조작: scrollIntoView → mouse.click → keyboard.type → 옵션 클릭
  - 요일 수 < 주N회이면 기존대로 "확인 필요"
- **AIM 실패 자동 복구**: 어드민 상태가 "AIM 실패"면 `/admin/tutor-pairing/{pairingId}/update`로 이동 → tutorPairingStatus를 MATCHING으로 변경 → 상세 페이지 복귀 → 매칭 재시도 (`fixPairingStatus`)
- **교대 단독 선택 시 필터 보정**: 대학교 체크박스에 교대만 단독 체크된 경우 → SKY/서성한/중경외시 추가 + 전공 > 교육 체크 후 검색 (`fixGyodaeFilter`, base 단계에서 실행)
- **수동매칭 timeline 테이블 선택 버그 수정**: 페이지에 timeline 헤더 테이블이 여러 개 있을 때 첫 번째가 아닌 **마지막** 테이블 사용 (table[9]는 다른 섹션, table[10]이 실제 검색 결과)
- **수동매칭 디버그 로그 강화**: 튜터별 timeline 파싱 결과·시간표·매칭 성공/실패 로그 추가

**이전 변경 (2026-04-16)**:
- 매칭 제안 (AIM) 기능 안정화: AIM 판정 타이밍 개선, 수동매칭 모달의 "매칭 제안하기" 버튼 매칭 패턴 확장, 모든 모달 디버그 로그 추가
- EXPERT 전문강사 토글 매 단계 검증 (이미 ON이면 유지)
- 재매칭 status 형식 변경 반영 (`[3]재매칭완료` → `재매칭 완료` 등) — aim.js / statusCheck.js / sheetPrep.js 전체 수정
- 유저메모 편집: 신규/재매칭 모두 지원, 수업신청서·가능시간·상단정보 전부 어드민 `__NEXT_DATA__` 기반으로 전환 (시트 의존 제거 → 재매칭 시트에 컬럼 없어도 작동)
- 유저메모 목록 필터: 신규=매칭중/확인필요, 재매칭=첫수업전재매칭/일반재매칭/매칭중/보류/확인필요
- 튜터메모(어드민 현재값) 별도 항목으로 표시 (유저메모 위)
- status 뱃지 추가 (PRO/신규 뱃지 옆)
- `__NEXT_DATA__` selector를 `state: 'attached'` 로 수정 (script 태그 hidden 이슈)
- `/api/debug/nextdata/:matchId` 디버그 엔드포인트 추가

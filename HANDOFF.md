# 콴다과외 매칭 자동화 — 핸드오프

> Claude Code에서 작업한 내역. claude.ai에 붙여넣어서 이어 작업하세요.

## 개요

로컬 Node.js 대시보드. 콴다과외 매칭 업무 자동화 (시트 정리, 페어링 생성, 매칭 제안, 매칭결과 확인, 유저메모 편집). 공용PC(Windows)에서 서버 실행, 같은 네트워크 기기에서 `http://192.168.0.185:3000` 접속.

**배포 구조:**
- 개발: `iris.kwon` PC에서 코드 수정 → git push (`C:\Users\iris.kwon\match_dashboard`)
- 공용PC 원격 업데이트: `Invoke-RestMethod -Method Post "http://192.168.0.185:3000/api/deploy"` — 공용PC 직접 안 건드려도 됨
- 서버: VBS + BAT 루프로 hidden background 실행
- Chrome CDP: `--remote-debugging-port=9333` (별도 프로필)
- **레포 public** (`github.com/qoxopb/match_dashboard`) — config/credentials는 gitignore라 시크릿 없음
- **로그 원격 조회**: `Invoke-RestMethod "http://192.168.0.185:3000/api/logs/file?date=YYMMDD"`

**스택**
- Express 5
- Playwright (CDP로 기존 크롬에 attach — port 9333)
- googleapis (Google Sheets)
- node-cron (스케줄러)
- node-notifier (알림)
- axios (Slack API)
- Vanilla HTML/JS 대시보드

## 파일 구조

```
match_dashboard/
├── server.js              # Express 서버, 작업/스케줄/큐 관리, REST API
├── config.json            # 모드, 시트 ID, Slack 토큰 등 (gitignore)
├── config.example.json    # config 템플릿
├── sheets.js              # Google Sheets 헬퍼
├── abort.js               # 중단 요청 플래그 + activePage 관리
├── credentials.json       # Google API credentials (gitignore)
├── oauth-token.json       # OAuth 토큰 (gitignore)
├── version.json           # 현재 버전
├── changelog.json         # 버전별 변경 이력
├── wf-presets.json        # 워크플로 프리셋 데이터 (gitignore, 서버 저장)
├── schedules.json         # 스케줄 영속화 (gitignore, 서버 재시작 시 복원)
├── automation/
│   ├── browser.js         # CDP 크롬 attach, newBackgroundPage
│   ├── sheetPrep.js       # prepNewSheet, prepRematchSheet, createMonthlyTab
│   ├── pairing.js         # runPairing(mode) — 5워커 병렬
│   ├── statusCheck.js     # runStatusCheck(mode), runStatusCheckForType — 5워커 병렬
│   ├── aim.js             # runAim(mode) — 매칭 제안 (AIM + 수동 + 임의완화 3단계)
│   └── userMemo.js        # 유저메모 어드민 스크래핑 + 저장 + Slack 검색
├── public/
│   ├── index.html         # 메인 대시보드 (기능/워크플로/당일현황/대시보드/설정)
│   ├── userMemo.html      # 유저메모 편집 전용 페이지
│   └── memoGate.html      # 유저메모 편집 게이트 (Slack 링크 → 편집자 확인)
└── .gitignore
```

## 주요 작업 (JOBS)

`server.js`의 `JOBS` 객체에 정의됨 + 워크플로 프리셋이 동적 등록됨:

| jobId | 이름 | 설명 |
|---|---|---|
| `prep` / `prep:new` / `prep:rematch` | 시트 정리 | sheetPrep |
| `pairing` | 페어링 생성 | runPairing |
| `send` / `send:new` / `send:rematch` | 매칭 제안 | runAim — AIM + 수동매칭 + 임의완화 |
| `statusCheck` / `statusCheck:new` / `statusCheck:rematch` | 매칭결과 확인 | runStatusCheck |
| `monthly` | 월별 탭 생성 | createMonthlyTab |
| `wf:{presetId}` | 워크플로 프리셋 | 동적 등록, runWfBlocks로 실행 |

> `aim/aim:new/aim:rematch`는 구 jobId — `normalizeJobId()`로 `send`로 매핑됨 (하위호환)

## 작업 큐 시스템

`server.js`의 `runTask`:
- **다른 jobId** → 동시 실행 허용 (예: 시트 정리 + 매칭결과 확인)
- **같은 jobId** → 큐에 넣고 순차 대기
- **워크플로 프리셋**은 모두 `wf` 그룹으로 순차 실행
- `runningJobs` (Set, lockId 기준) + `taskQueue` 배열로 관리
- UI: 실행 중인 버튼만 개별 비활성화 (다른 버튼은 활성 유지)
- 상태바: 동시 실행 중인 작업명 콤마로 나열

## 워크플로 시스템

### 블록 종류
- `job` — 단일 작업 (시트 정리, 페어링, 매칭 제안 등)
- `loop` — N회 반복, 자식 블록 포함
- `condLoop` — 조건 반복 (시트 status 건수 기준), 신규/재매칭/전체
- `userMemo` — 유저메모 편집 요청 (Slack 알림 → 대기 → 완료 시 재개)

### 블록 중첩
반복/조건 블록 안에 또 반복/조건 블록 가능 (재귀 구조)

### 프리셋
- localStorage에 UI 상태, `wf-presets.json`에 서버 데이터 저장
- 프리셋 저장 시 서버에도 동기화 (스케줄 실행용)
- 프리셋 목록 드래그 드롭 순서 변경 가능

### 스케줄
- **주기**: N분 간격 (등록 즉시 1회 실행)
- **시간**: 특정 시각 + 요일
- **일자**: 특정 날짜 (1회/매월/매년)
- `schedules.json`에 영속화 → 서버 재시작 시 자동 복원
- 복원 시 다음 실행 예정 시간 로그 표시
- 프리셋 목록에 재생/일시정지 버튼 (스케줄 일괄 ON/OFF)

### 유저메모 편집 블록 (human-in-the-loop)
1. 편집 대상 건수 체크 → 0건이면 자동 건너뜀
2. Slack 채널에 알림 발송 (건수 + 태그 대상 멘션 + 편집 링크)
3. 서버 pending 상태로 대기
4. Slack 링크 클릭 → `memoGate.html` → 편집자 선택 → accept → 유저메모 편집 페이지
5. 다른 사람이 이미 accept → "이미 편집 중" 안내 → 5초 후 닫기
6. 모두 거절 → pending 해제, 워크플로 건너뜀
7. 편집 완료 (모든 항목 처리 + 탭 닫힘) → 서버에 complete → 워크플로 재개
8. Slack 메시지 업데이트: 수락 시 "OO님이 편집 시작" / 완료 시 "편집 완료" (원본 메시지 수정, 별도 메시지 안 보냄)

## 매칭 제안 (AIM) 핵심 로직 — `automation/aim.js`

매치아이디 1건당 처리 흐름 (`processOne`):

1. **어드민 status 확인**
   - `신청서 미작성` → 스킵
   - `매칭 완료` → 시트에 반영 후 다음
   - `AIM 진행중` → 매칭 중단 후 재시도
   - `AIM 실패` → tutorPairingStatus를 MATCHING으로 변경 후 재시도

2. **특정 튜터풀 모드** — 키워드 매칭 시 해당 튜터 ID 직접 검색

3. **교대/메디컬 포함 시** — AIM 건너뛰고 변형 수동매칭 우선 (변형 필터 → 실패 시 원래 조건 + 임의완화 → PRO면 추가 시도)

4. **1단계 AIM** — AI 매칭 시작 버튼 클릭
   - AIM 성공 후 `waitForLoadState`로 리다이렉트 완료 대기 (네비게이션 충돌 방지)

5. **2단계 수동매칭** — 시간표 비교 후 전송하기
   - EXPERT면 "전문 강사 여부" 토글 ON
   - 학생당 최대 2명까지 제안 (`MAX_PROPOSALS_PER_STUDENT = 2`)
   - 전송 시도 상한 10회 (`MAX_SEND_ATTEMPTS = 10`)

6. **3단계 임의완화** (relax1~3)

7. **sendProposal 성공 감지**: 모달 닫힘 OR 테이블 상태 변경 ("전송하기" → "응답대기" 등)

8. **navigation interrupted 에러 시 1회 자동 재시도**

9. **시트 제외 status**: 신규 `['매칭완료','동일튜터유지','동일 튜터 유지','환불','보류']`, 재매칭 `['재매칭 완료','재매칭완료','환불/소멸','환불','소멸','보류']`

10. **재매칭 검색**: `pairingId`가 있으면 `filters_id`로, 없으면 `filters_matchId`로 검색

## 유저메모 편집 — `public/userMemo.html`

- 진입 시 statusCheck 자동 실행 (어드민 상태 동기화 후 목록 표시)
- 모든 항목 처리 완료 → 5초 카운트다운 → 탭 자동 닫기
- autoNext: 신규 완료 → 재매칭 자동 이동
- 워크플로에서 온 경우 (`fromWorkflow=true`) 완료 시 서버에 complete 신호

## 설정 탭

- **로그인 권한**: Google 계정 인증 (추후 구현)
- **Slack 태그 대상**: User ID + 별칭 관리. 유저메모 알림 시 @멘션

## 대시보드 UI (`public/index.html`)

사이드바 메뉴: **기능** / **워크플로** / **당일 현황** / **대시보드** / **설정**

### 기능 탭
- 카드: 시트 정리 / 페어링 생성 / 유저메모 편집 / 매칭 제안 / 매칭결과 확인 / 월별 탭 관리
- AIM 키워드 설정 모달 (exclude/tutorPool)
- 실행 로그 (3초 폴링)
- 상태바 (실행 중 작업명 표시 / 대기 중 + 중단 버튼)
- 항목별 스케줄 설정 토글 없음 (워크플로 탭에서만 스케줄 관리)

### 워크플로 탭
- 3열 레이아웃: 편집 영역(설정바 + 팔레트 + 워크스페이스) | 프리셋 목록 | 스케줄 설정
- 블록 드래그 드롭 (팔레트 → 워크스페이스, 워크스페이스 내 순서 변경, 컨테이너 간 이동)
- 드롭 위치 인디케이터 (파란 선)
- 실행 로그 (클라이언트 + 서버 로그 폴링)
- 상태바 (기능 탭과 동기화)

### 당일 현황 / 대시보드
빈 페이지 (추후 구현)

## REST API

### 작업 실행
- `POST /api/run/:jobId` — 즉시 실행
- `POST /api/stop` — 중단 요청 (pending 대기도 즉시 해제)

### 스케줄
- `POST /api/schedule` `{jobId, type, options}` — 등록
- `DELETE /api/schedule/:id` — 삭제
- `PATCH /api/schedule/:id` `{enabled}` — ON/OFF
- `GET /api/schedules` — 목록

### 워크플로 프리셋
- `POST /api/wf/presets/:id` `{name, workspace}` — 저장
- `DELETE /api/wf/presets/:id` — 삭제

### 워크플로 대기
- `GET /api/wf/pending` — pending 상태 조회
- `GET /api/wf/count?type=new|rematch&status=` — 시트 status별 건수 (빈값=미처리)
- `POST /api/wf/pending/register` — 랜딩 페이지 등록
- `POST /api/wf/pending/unregister` — 등록 해제
- `POST /api/wf/pending/accept` `{name}` — 편집 수락
- `POST /api/wf/pending/reject` — 편집 거절
- `POST /api/wf/pending/complete` — 편집 완료

### 유저메모
- `GET /api/userMemo/sync?type=new|rematch` — statusCheck 실행 (목록 조회 전)
- `GET /api/userMemo/list?type=...&filter=initial|all` — 처리 대상 목록
- `GET /api/userMemo/detail?matchId=...` — 어드민 + Slack 조회
- `POST /api/userMemo/save` — 어드민 저장

### 설정
- `POST /api/config/set` `{path, value}` — config 값 설정
- `GET /api/config/slack-tags` — Slack 태그 대상 조회
- `POST /api/config/slack-tags` `{tagTargets}` — Slack 태그 대상 저장

### 로그
- `GET /api/logs` — 메모리 로그 (running 상태 + runningJobs Set + queue 수 포함)
- `GET /api/logs/file?date=YYMMDD` — 날짜별 파일 로그
- `GET /api/logs/files` — 로그 파일 목록

### 버전/배포
- `GET /api/version` — 현재 버전
- `GET /api/version/latest` — git fetch 후 원격 버전 비교
- `GET /api/changelog` — 변경 이력
- `POST /api/deploy` — git pull + 서버 재시작

### 디버그
- `GET /api/debug/nextdata/:matchId` — 어드민 상세 페이지 `__NEXT_DATA__` 전체 덤프
- `GET /api/debug/pairing/:matchId?relax=0~3` — 페어링 상세 페이지 + 완화 단계별 DOM 구조
- `GET /api/userMemo/debug?type=new|rematch` — 시트 status 분포 + 필터 결과

## config.json 구조

```json
{
  "mode": "both",
  "sheets": {
    "new": "시트ID",
    "rematch": "시트ID"
  },
  "gas": { "url": "GAS URL" },
  "slack": {
    "botToken": "xoxb-...",
    "channelId": "기존 채널",
    "notifyChannelId": "C05KGLU5NG7",
    "tagTargets": [{"id": "U...", "name": "별칭"}]
  },
  "aimKeywords": [{"keyword": "...", "action": "exclude|tutorPool", "tutorIds": "..."}]
}
```

## 주요 이슈/결정 사항

### CDP 포트 9333
공용PC에서 `chrome.exe --remote-debugging-port=9333 --user-data-dir=C:\chrome-cdp-profile`. PC 재시작 시 수동으로 다시 실행 필요 (자동 시작 미설정).

### sendProposal 성공 감지
모달 닫힘만으로 판정하면 false negative 발생 (실제 발송됐는데 실패 판정 → 2명 제한 안 걸림). 테이블 상태 변경("전송하기" → 다른 값)도 확인.

### 타임라인 모달 가림 문제
sendProposal 전 closeAllModals 실행 + 좌표 계산을 모달 닫은 후 수행. closeAllModals는 `.ant-modal-wrap` visible 체크, ESC 5회, 애니메이션 대기.

### 시수는 어드민 product name에서 파싱
시트가 아니라 어드민 `product.name` ("주 2회, 60분 수업")에서 정규식으로 파싱.

### 브라우저 팝업 차단
HTTP 환경에서 사용자 클릭 없이 window.open 불가. Slack 알림 + memoGate 링크로 해결.

### 공용PC 배포
레포 public (`github.com/qoxopb/match_dashboard`). 코드 push 후 `Invoke-RestMethod -Method Post "http://192.168.0.185:3000/api/deploy"` 로 원격 업데이트 가능. 서버 다운 시엔 API 호출 불가 — 공용PC에서 직접 `git pull` 필요.

### adminUrl / slack.workspaceUrl
코드에 하드코딩된 URL을 `config.json`으로 분리. `config.json`은 gitignore라 공용PC에 필드가 없으면 fallback(`https://tutor-admin.qanda.ai`)으로 동작. 추가하려면 공용PC `config.json`에 `"adminUrl"`, `"slack.workspaceUrl"` 직접 추가.

### ADMIN_BASE 치환 주의
`replace_all`로 `config.adminUrl` → `ADMIN_BASE` 치환 시 정의 줄(`const ADMIN_BASE = config.adminUrl || ...`)까지 바뀌어 자기참조(`ADMIN_BASE = ADMIN_BASE`) ReferenceError 발생 → 서버 시작 크래시. 향후 상수 치환 시 정의 줄 제외하고 치환할 것.

---

**현재 버전: v1.4.3 (2026-05-08, 정상 작동 확인)**

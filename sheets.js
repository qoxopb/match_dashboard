const axios = require('axios');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const config = require('./config.json');

// --- GAS 웹앱으로 시트 읽기 ---
async function readSheet(type) {
  // type: 'new' | 'rematch'
  const sheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  const url = `${config.gas.url}?sheetId=${sheetId}`;

  const res = await axios.get(url);
  return res.data; // GAS가 반환하는 행 배열
}

// --- Google Sheets API로 시트 쓰기 ---
function getAuth() {
  const credPath = path.join(__dirname, 'credentials.json');
  if (!fs.existsSync(credPath)) {
    throw new Error('credentials.json이 없습니다. Google Sheets API 키를 프로젝트 루트에 넣어주세요.');
  }

  const creds = JSON.parse(fs.readFileSync(credPath, 'utf8'));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/bigquery.readonly',
    ],
  });
  return auth;
}

async function getSheetsApi() {
  const auth = getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function updateCell(sheetId, range, value) {
  const sheets = await getSheetsApi();

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });

  console.log(`[sheets] ${sheetId} ${range} → "${value}" 업데이트 완료`);
}

// 특정 match_id의 상태 컬럼 업데이트
async function updateStatus(type, matchId, rowIndex, status) {
  const sheetId = config.sheets[type === 'new' ? 'new' : 'rematch'];
  // 신규: '매칭상태' 컬럼, 재매칭: 'status' 컬럼
  // rowIndex는 시트 기준 행 번호 (1-based, 헤더 포함)
  // 컬럼 위치는 실제 시트에 맞게 조정 필요
  const col = type === 'new' ? 'B' : 'B'; // TODO: 실제 컬럼 확인 후 수정
  const tabName = type === 'new' ? '[신규] 25.04' : '[재매칭] 25.04';
  const range = `'${tabName}'!${col}${rowIndex}`;

  await updateCell(sheetId, range, status);
}

module.exports = { readSheet, getSheetsApi, updateCell, updateStatus };

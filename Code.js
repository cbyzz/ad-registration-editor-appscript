// 테스트
// const ADMIN_EMAIL = 'choi.byoungyoul@nbt.com';
// const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_TEST_WEBHOOK_URL');;



//. 실제 라이브
const ADMIN_EMAIL = 'choi.byoungyoul@nbt.com,operation@nbt.com,sales@nbt.com,adison.cs@nbt.com';
const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
const SYSTEM_URL = PropertiesService.getScriptProperties().getProperty('SYSTEM_URL');

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// 변경 시트 
const EXTERNAL_DATA_SHEET_ID = PropertiesService.getScriptProperties().getProperty('EXTERNAL_DATA_SHEET_ID');
const externalSs = SpreadsheetApp.openById(EXTERNAL_DATA_SHEET_ID);


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetByGid(spreadsheet, gid) {
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == gid) {
      return sheets[i];
    }
  }
  return null;
}

// 광고주 목록(List)과 카테고리 매핑 정보(Map)를 모두 반환
function getExternalAdvertisersData() {
  const sheet = getSheetByGid(externalSs, 1564465491); 
  if (!sheet || sheet.getLastRow() < 3) return { list: [], map: {} };
  
  // C열(광고주)과 D열(카테고리) 데이터를 함께 가져옴
  const range = sheet.getRange('C3:D' + sheet.getLastRow());
  const values = range.getValues();
  
  const list = [];
  const map = {};

  values.forEach(row => {
    const advertiser = row[0]; // C열
    const category = row[1];   // D열
    
    if (advertiser) {
      list.push(advertiser);
      map[advertiser] = category || ''; // 카테고리가 비어있을 경우 대비
    }
  });

  return {
    list: list.sort(),
    map: map
  };
}

// '거래처' 목록을 외부 시트에서 가져오는 함수
function getExternalClients() {
  const sheet = getSheetByGid(externalSs, 1564465491);
  if (!sheet || sheet.getLastRow() < 3) return [];

  const lastRow = sheet.getLastRow();
  const ranges = ['C3:C' + lastRow, 'G3:G' + lastRow, 'H3:H' + lastRow, 'I3:I' + lastRow];
  let allClients = [];

  ranges.forEach(rangeString => {
    const values = sheet.getRange(rangeString).getValues().flat().filter(String);
    allClients = allClients.concat(values);
  });
  
  // 중복을 제거하고 정렬하여 반환
  return [...new Set(allClients)].sort();
}

function doGet(e) {
  Logger.log("[doGet] 웹앱 요청 시작");
  Logger.log("[doGet] 파라미터: " + JSON.stringify(e?.parameter || {}));
  Logger.log("VS Code에서 수정해서 보낸 메시지입니다! 🚀");
  // '이 광고 담당하기' 처리 로직 (기존과 동일)
  if (e && e.parameter && e.parameter.action === 'confirm' && e.parameter.id) {
    const adId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordConfirmation(adId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete' && e.parameter.id && e.parameter.adId) {
    const registrationId = e.parameter.id;
    const adId = e.parameter.adId;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCompletion(registrationId, adId, completerEmail);
    // processCompletion 함수는 객체를 반환하므로, 메시지만 추출하여 사용합니다.
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_mod' && e.parameter.id) {
    const modId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordModificationConfirmation(modId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_dsp' && e.parameter.id) {
    const dspId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordDspConfirmation(dspId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_dsp' && e.parameter.id && e.parameter.adId) {
    const registrationId = e.parameter.id;
    const adId = e.parameter.adId;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processDspCompletion(registrationId, adId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  // '수정 완료' 처리 로직
  if (e && e.parameter && e.parameter.action === 'complete_mod' && e.parameter.id) {
    const modId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processModificationCompletion(modId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_dsp_mod' && e.parameter.id) {
    const dspModId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordDspModificationConfirmation(dspModId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_dsp_mod' && e.parameter.id) {
    const dspModId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processDspModificationCompletion(dspModId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_cs' && e.parameter.id) {
    const csId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordCashslideConfirmation(csId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_cs' && e.parameter.id && e.parameter.adId) {
    const registrationId = e.parameter.id;
    const adId = e.parameter.adId;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCashslideCompletion(registrationId, adId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_cs_mod' && e.parameter.id) {
    const csModId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordCashslideModificationConfirmation(csModId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_cs_mod' && e.parameter.id) {
    const csModId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCashslideModificationCompletion(csModId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }


  if (e && e.parameter && e.parameter.action === 'confirm_cx' && e.parameter.id) {
    const cxId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordCxConfirmation(cxId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  // ▼▼▼ [추가] CX팀 완료 처리 ▼▼▼
  if (e && e.parameter && e.parameter.action === 'complete_cx' && e.parameter.id) {
    const cxId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCxCompletion(cxId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_bd' && e.parameter.id) {
    const bdId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordBdConfirmation(bdId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  // ▼▼▼ [추가] 오퍼월사업팀 완료 처리 ▼▼▼
  if (e && e.parameter && e.parameter.action === 'complete_bd' && e.parameter.id) {
    const bdId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processBdCompletion(bdId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_other' && e.parameter.id) {
    const otherId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordOtherConfirmation(otherId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_other' && e.parameter.id) {
    const otherId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processOtherCompletion(otherId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_other' && e.parameter.id) {
    const otherId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processOtherCompletion(otherId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  // ▼▼▼ [추가] 쿠폰 발급 요청 담당하기 및 완료 처리 ▼▼▼
  if (e && e.parameter && e.parameter.action === 'confirm_coupon' && e.parameter.id) {
    const couponId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordCouponConfirmation(couponId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_coupon' && e.parameter.id) {
    const couponId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCouponCompletion(couponId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'confirm_copy' && e.parameter.id) {
    const copyId = e.parameter.id;
    const approverEmail = Session.getActiveUser().getEmail();
    const resultMessage = recordCopyCreationConfirmation(copyId, approverEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage}</h1>`);
  }

  if (e && e.parameter && e.parameter.action === 'complete_copy' && e.parameter.id) {
    const copyId = e.parameter.id;
    const completerEmail = Session.getActiveUser().getEmail();
    const resultMessage = processCopyCreationCompletion(copyId, completerEmail);
    return HtmlService.createHtmlOutput(`<h1>${resultMessage.message}</h1>`);
  }

  // 기본 웹앱 로드 로직 (기존과 동일)
  const userEmail = Session.getActiveUser().getEmail();
  if (!isAuthorizedUser(userEmail)) {
    return HtmlService.createHtmlOutput(`<h1>접근 권한이 없습니다.</h1><p>관리자에게 문의하세요. (${userEmail})</p>`);
  }
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('광고 등록 요청 시스템');
  return html;
}

function isAuthorizedUser(email) {
  if (!email.endsWith('@nbt.com')) {
    return false;
  }
  const userSheet = ss.getSheetByName('사용자');
  if (!userSheet) { 
    return false;
  }
  const userList = userSheet.getRange('A2:A' + userSheet.getLastRow()).getValues().flat();
  return userList.includes(email);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}


/**
 * ID를 기반으로 해당 행의 정보를 효율적으로 찾습니다. (요청 타입에 따라 시트 분기)
 * @param {string} id - 찾을 고유 ID.
 * @param {string} type - '광고' 또는 '수정'. 기본값은 '광고'.
 * @returns {object|null} - 찾은 경우 {sheet, rowIndex, headers, rowData}, 못 찾은 경우 null.
 */
function findRowById(id, type = '광고') {
  if (!id || typeof id !== 'string' || !id.includes('-')) {
    return null;
  }
  const cleanId = id.trim();
  const lastHyphenIndex = cleanId.lastIndexOf('-');
  const userNameWithPrefix = cleanId.substring(0, lastHyphenIndex);
  // 'choi.byoungyoul-mod' 같은 경우를 대비해 사용자 이름만 추출
  const userName = userNameWithPrefix.replace('-mod', '');

  const sheetName = (type === '수정') ? `${userName} - 수정` : `${userName} - 광고`;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return null;
  }

  const idColumn = sheet.getRange('A:A');
  const textFinder = idColumn.createTextFinder(cleanId).matchEntireCell(true);
  const foundCell = textFinder.findNext();

  if (foundCell) {
    const rowIndex = foundCell.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return {
      sheet: sheet,
      rowIndex: rowIndex,
      headers: headers,
      rowData: rowData
    };
  }

  return null;
}

function getAdDataById(adId) {
  const found = findRowById(adId);
  if (found) {
    const adData = {};
    found.headers.forEach((header, index) => {
      let value = found.rowData[index];
      if (value instanceof Date) {
        try {
          // ▼▼▼▼▼ [수정] 시간 필드 포맷팅 로직 추가 ▼▼▼▼▼
          if (header.endsWith('라이브 시작 시간') || header.endsWith('라이브 종료 시간')) {
            value = Utilities.formatDate(value, "Asia/Seoul", "HH:mm");
          } else {
            value = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm");
          }
          // ▲▲▲▲▲ [수정] ▲▲▲▲▲
        } catch(e) {
          value = '날짜 형식 오류';
        }
      }
      adData[header] = value;
    });
    return adData;
  }
  return null;
}



function logUserAction(userEmail, action, details) {
  try {
    const logSheetName = '활동 로그';
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName, 0);
      const headers = ['시간', '사용자', '작업', '대상 ID', '상세 내용'];
      logSheet.appendRow(headers);
      logSheet.getRange('1:1').setFontWeight('bold').setBackground('#f3f3f3');
      logSheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    const logData = [
      timestamp,
      userEmail,
      action,
      details.targetId || '',
      details.message || ''
    ];
    
    logSheet.appendRow(logData);

  } catch (e) {
    console.error(`로깅 실패: ${e.toString()}`);
  }
}

function submitModificationShare(formData) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    
    const sheetName = '수정 공유 로그';
    let sheet = ss.getSheetByName(sheetName);
    const headers = ['등록일시', '등록자', '대상 광고 ID', '대상 광고명', '주요 공유사항', '총물량', '일물량', 'ON/OFF', '노출 대상', '이미지 소재'];
   

    if (!sheet) {
      sheet = ss.insertSheet(sheetName, 0);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    
    const newRowData = [
      formattedTimestamp, userEmail, formData['대상 광고 ID'] || '', formData['대상 광고명'] || '',
      formData['주요 공유사항'] || '', formData['총물량'] || '', formData['일물량'] || '',
      formData['ON / OFF'] || '', formData['노출 대상'] || '', formData['이미지 소재'] || ''
    ];

    sheet.appendRow(newRowData);

    const targetAdName = formData['대상 광고명'].split('\n')[0];
    const subject = `[광고 수정 공유]${targetAdName ? ' ' + targetAdName.split('\n')[0] : ''}`;
    const ccEmails = formData.ccRecipients || '';

    // ▼▼▼ [수정] 메일 본문에 시스템 링크를 추가합니다. ▼▼▼
    let body = `<p>안녕하세요, 운영팀.</p>
                <p><b>${userEmail}</b>님께서 광고 수정 사항을 공유했습니다.</p>
                <p>해당 내용은 '수정 공유 로그' 시트에 저장되었습니다.</p>
                <p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p>
                <hr>
                <h3>공유 내용</h3>
                <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;
    // ▲▲▲ [수정] ▲▲▲
    
    headers.slice(2).forEach(field => {
      let clientKey = field;
      if (field === 'ON/OFF') {
        clientKey = 'ON / OFF'; // 시트 헤더 'ON/OFF'일 때, 클라이언트 키 'ON / OFF' 사용
      }
      if (formData[clientKey]) { 
         const value = formData[clientKey].replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${field}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${value}</td></tr>`;
      }
    });
    body += `</table>`;

    GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails }); // cc 옵션 추가
    
    try {
      const slackMessage = { 'text': `${subject}` };
      const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    } catch (e) {
      console.error(`수정 공유 슬랙 발송 실패 (광고: ${targetAdName}): ${e.toString()}`);
    }

    logUserAction(userEmail, '수정 공유', {
      targetId: formData['대상 광고 ID'],
      message: `광고 '${targetAdName}' 수정 공유 및 시트 저장`
    });

    return { success: true, message: '수정 내용이 성공적으로 공유 및 저장되었습니다.' };
  } catch (e) {
    console.error(`submitModificationShare Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  }
}

function submitCouponRequest(formData) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    const sheetName = '쿠폰 발급 요청';
    let sheet = ss.getSheetByName(sheetName);
    
    // 관리 및 데이터 컬럼 정의
    const headers = [
      'id', 'timestamp', 'registrant', 'status', 'manager', 'manager_timestamp', 'completion_timestamp', // 시스템 관리용
      'subject', // 메일 제목 (검색용)
      'target_ad_id', 'target_ad_name', 'amount', 'coupon_name', 'quantity', 'expiry_date', 'additional_request'
    ];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName, 0);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // 1. ID 생성
    const idPrefix = `coupon-${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;

    // 2. 제목 생성
    const targetAdName = formData['대상 광고명'].split('\n')[0];
    const subject = `[쿠폰 발급 요청] ${targetAdName} (${uniqueId})`;

    // 3. 알림 발송
    sendCouponNotification(userEmail, uniqueId, subject, formData);

    // 4. 시트 저장
    const newRow = [
      uniqueId, formattedTimestamp, userEmail, '등록 요청 완료', '', '', '',
      subject,
      formData['대상 광고 ID'], formData['대상 광고명'],
      formData['쿠폰 금액'], formData['쿠폰 명'], formData['쿠폰 발급 수량'],
      formData['쿠폰 만료 일자'], formData['추가 요청 사항']
    ];

    sheet.appendRow(newRow);

    logUserAction(userEmail, '쿠폰 발급 요청', {
      targetId: uniqueId,
      message: subject
    });

    return { success: true, message: `쿠폰 발급 요청이 완료되었습니다. (ID: ${uniqueId})` };
  } catch (e) {
    console.error(`submitCouponRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function sendCouponNotification(senderEmail, id, subject, formData) {
  const ccEmails = formData.ccRecipients || '';
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_coupon&id=${id}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_coupon&id=${id}`;

  let body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 쿠폰 발급을 요청했습니다.</p>
              <p><b>ID: ${id}</b></p>
              <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
                <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 요청 담당하기 ]</a>
                <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 처리 완료 ]</a>
                <br><br>
                <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
                <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">시스템 바로가기</a>
              </div>
              <hr>
              <h3>요청 내용</h3>
              <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;

  const fieldMap = {
    '대상 광고 ID': '대상 광고 ID',
    '대상 광고명': '대상 광고명',
    '쿠폰 금액': '쿠폰 금액',
    '쿠폰 명': '쿠폰 명',
    '쿠폰 발급 수량': '쿠폰 발급 수량',
    '쿠폰 만료 일자': '쿠폰 만료 일자',
    '추가 요청 사항': '추가 요청 사항'
  };

  for (const [key, label] of Object.entries(fieldMap)) {
    if (formData[key]) {
      const value = String(formData[key]).replace(/\n/g, '<br>');
      body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${label}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${value}</td></tr>`;
    }
  }
  body += `</table>`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails });
  
  try {
    const slackMessage = { 'text': `${subject}` };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) });
  } catch (e) {
    console.error(`쿠폰 요청 슬랙 발송 실패: ${e.toString()}`);
  }
}

function findCouponRowById(id) {
  const sheet = ss.getSheetByName("쿠폰 발급 요청");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true);
  const foundCell = textFinder.findNext();
  if (foundCell) {
    const rowIndex = foundCell.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { sheet, rowIndex, headers, rowData };
  }
  return null;
}

function recordCouponConfirmation(id, approverEmail) {
  const found = findCouponRowById(id);
  if (!found) return `쿠폰 요청 ID: ${id} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const timestampColIndex = headers.indexOf('manager_timestamp');

  if (managerColIndex === -1) return '필수 컬럼(manager)이 없습니다.';

  const currentManager = rowData[managerColIndex];
  if (currentManager && currentManager !== '') {
    return `처리 실패: 이 건(ID: ${id})은 이미 ${currentManager} 님이 담당하고 있습니다.`;
  }

  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  if (timestampColIndex > -1) {
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  }

  try {
    const subject = rowData[headers.indexOf('subject')];
    if (subject) {
        const threads = GmailApp.search(`subject:"${subject}"`, 0, 1);
        if (threads && threads.length > 0) {
            threads[0].replyAll("", {
                htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>쿠폰 요청 ID: ${id}</b> 건의 담당자로 지정되어 처리를 진행합니다.</p><p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>`
            });
        }
    }
  } catch (e) {
    console.error(`쿠폰 담당자 알림 발송 오류: ${e.toString()}`);
  }

  return `쿠폰 요청 ID: ${id} 건의 담당자로 ${approverEmail}님이 지정되었습니다.`;
}

function processCouponCompletion(id, completerEmail) {
  const found = findCouponRowById(id);
  if (!found) return { success: false, message: `쿠폰 요청 ID(${id})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  const completionDateColIndex = headers.indexOf('completion_timestamp');

  if (statusColIndex === -1) return { success: false, message: 'status 컬럼을 찾을 수 없습니다.' };

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '완료') {
    return { success: false, message: `이미 완료 처리된 건입니다. (ID: ${id})` };
  }

  sheet.getRange(rowIndex, statusColIndex + 1).setValue('완료');
  if (completionDateColIndex > -1) {
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  }

  logUserAction(completerEmail, '쿠폰 요청 완료', { targetId: id });
  return { success: true, message: `쿠폰 요청 건(ID: ${id})이 성공적으로 완료 처리되었습니다.` };
}






function submitModificationRequest(formData) {
  // ▼▼▼▼▼ [추가] ID 중복 생성을 막기 위해 LockService를 추가합니다. ▼▼▼▼▼
  const lock = LockService.getUserLock();
 lock.waitLock(30000);
  // ▲▲▲▲▲ [추가] ▲▲▲▲▲

try {
 const userEmail = Session.getActiveUser().getEmail();
 const userName = userEmail.split('@')[0];
 const sheetName = `${userName} - 수정`;
 let sheet = ss.getSheetByName(sheetName);

 const masterHeaderOrder = [
 '등록ID', '등록일시', '등록자', '상태', '담당자', '담당자 확인 일시', '메일 스레드 ID', '수정 완료 일시', '반려 일시', '반려 사유',
 '주요 요청사항', '대상 캠페인 ID', '대상 광고 ID', '대상 광고명', '예약 반영 시점',
 '광고주 연동 토큰 값', '매체', '단가', '총물량', '리워드', '일물량',
 '광고 집행 시작 일시', '광고 집행 종료 일시', '광고 노출 중단 시작일시', '광고 노출 중단 종료일시', '광고 참여 시작 후 완료 인정 유효기간 (일단위)',
 '트래커', '완료 이벤트 이름', '트래커 추가 정보 입력',
 'URL - 기본', 'URL - AOS', 'URL - IOS', 'URL - PC',
    '기본 URL',
    '상세전용랜딩 URL',
 '소재 경로', '적용 필요 항목',
 '라이브 시작 시간', '라이브 종료 시간', 'adid 타겟팅 모수파일', '데모타겟1', '데모타겟2',
 '2차 액션 팝업 사용', '2차 액션 팝업 이미지 링크', '2차 액션 팝업 타이틀', '2차 액션 팝업 액션 버튼명', '2차 액션 팝업 랜딩 URL',
 '문구 - 타이틀', '문구 - 서브', '문구 - 상세화면 상단 타이틀', '문구 - 서브1 상단', '문구 - 서브1 하단',
 '액션 버튼', '문구 - 서브2', '노출 대상', '기타', '광고 타입별 추가',
   '쿠키오븐 CPS_최소 결제 금액', '쿠키오븐 CPS_파트너 광고주 타입', '쿠키오븐 CPS_파트너 광고주 ID', '쿠키오븐 CPS_참여 경로 유형(app/web)',
   '네이버페이 알림받기_(메타) NF 광고주 연동 타입', '네이버페이 알림받기_(메타) NF 광고주 연동 ID', '네이버페이 알림받기_URL',
   '네이버페이 CPS_본광고_URL', '네이버페이 CPS_본광고_최소 결제 금액', '네이버페이 CPS_본광고_(목록) 리워드 조건 설명', '네이버페이 CPS_본광고_(목록) 리워드 텍스트', '네이버페이 CPS_본광고_(메타) NF 광고주 연동 ID', '네이버페이 CPS_본광고_(메타) 클릭 리워드 지급 금액',
   '네이버페이 CPS_부스팅_복사 필요한 광고 ID', '네이버페이 CPS_부스팅_URL & 상세 전용 랜딩 URL', '네이버페이 CPS_부스팅_문구 - 서브1 하단', '네이버페이 CPS_부스팅_최소 결제 금액', '네이버페이 CPS_부스팅_(목록) 리워드 조건 설명', '네이버페이 CPS_부스팅_부스팅 옵션', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_추천 세팅 여부', '네이버페이 CPS_부스팅_placement 세팅 정보 기본', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_카테고리',
   'CPQ_CPQ 뷰', 'CPQ_랜딩 형태', 'CPQ_임배디드 연결 형태', 'CPQ_유튜브 ID / 네이버 TV CODE', 'CPQ_이미지', 'CPQ_이미지 연결 링크', 'CPQ_퀴즈', 'CPQ_정답', 'CPQ_정답 placeholder 텍스트', 'CPQ_오답 alert 메시지', 'CPQ_사전 랜딩(딥링크) 사용', 'CPQ_사전 랜딩 실행 필수', 'CPQ_사전 랜딩 URL', 'CPQ_사전 랜딩 버튼 텍스트', 'CPQ_사전 랜딩 미실행 alert 메시지',
   'CPA SUBSCRIBE_구독 대상 이름', 'CPA SUBSCRIBE_이미지 인식에 사용할 식별자', 'CPA SUBSCRIBE_광고주 계정 식별자1', 'CPA SUBSCRIBE_광고주 계정 식별자2', 'CPA SUBSCRIBE_광고주 계정 식별자3', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL AOS', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL IOS'
 ];

if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(masterHeaderOrder);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
      Logger.log(`Sheet "${sheetName}" created with headers.`);
    } else {
      // 시트가 이미 있는 경우, 누락된 컬럼 확인 및 추가
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('Current Headers: ' + JSON.stringify(currentHeaders)); // 현재 헤더 로그
      const missingHeaders = masterHeaderOrder.filter(h => !currentHeaders.includes(h));
      Logger.log('Missing Headers: ' + JSON.stringify(missingHeaders)); // 누락된 헤더 로그

      if (missingHeaders.length > 0) {
        try {
          // 누락된 헤더를 시트의 마지막 열 다음에 추가
          sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
          Logger.log(`Successfully added missing headers: ${missingHeaders.join(', ')}`); // 성공 로그
          // 변경사항이 시트에 즉시 반영되도록 강제
          SpreadsheetApp.flush();
        } catch (e) {
          Logger.log(`Error adding missing headers: ${e.toString()}`); // 에러 로그
        }
      } else {
        Logger.log('No missing headers found.'); // 누락 헤더 없음 로그
      }
    }

    // --- ▼▼▼ [수정] finalHeaders 정의 위치 변경 ▼▼▼ ---
    // 누락된 컬럼이 추가된 *후에* 최종 헤더 목록을 다시 가져옴
    const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Final Headers for data mapping: ' + JSON.stringify(finalHeaders)); // 최종 헤더 로그

  if (formData['광고 타입별 추가'] === 'CPA SUBSCRIBE') {
   const subscriptionTarget = formData['CPA SUBSCRIBE_구독 대상 이름'];
   const autoGeneratedTargets = ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우', '좋아요', '채널추가', '유튜브_좋아요', '언론사 구독', '틱톡', 'X(트위터)'];
  
   if (autoGeneratedTargets.includes(subscriptionTarget)) {
    const id1 = formData['CPA SUBSCRIBE_광고주 계정 식별자1'];
    const id2 = formData['CPA SUBSCRIBE_광고주 계정 식별자2'];
    const id3 = formData['CPA SUBSCRIBE_광고주 계정 식별자3'];
    const identifiers = [id1, id2, id3].filter(id => id && id.trim() !== '');

    if (identifiers.length > 0) {
     const identifierPart = `(${identifiers.map(id => `{${id}:text}`).join(' || ')})`;
     let conditionPart = '';

     switch (subscriptionTarget) {
      case '유튜브 구독(채널메인)': conditionPart = "({구독중:text} || {구독 중:text} || {구독충:text} || {subscribed:text}) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text})"; break;
      case '유튜브 구독(특정영상)': conditionPart = "(({구독중:text} || {구독 중:text} || {구독충:text} || {subscribed:text}) || ({youtube_subscribe_alarm_all:customml} || {youtube_subscribe_alarm:customml} || {youtube_subscribe_no_alarm:customml})) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text})"; break;
      case '팔로우': conditionPart = "({follow_white:customml} || {follow_black:customml} || {팔로잉 ~:text} || {팔로잉 아:text} || {팔로잉 v:text} || {팔로잉~:text} || {팔... ~:text} || {팔로... ~:text} || {팔...~:text} || {팔로...~:text} || {팔... :text} || {팔로... :text} || {팔...:text} || {팔로...:text} || {following v:text} || {following:text} || {팔로잉.*팔로잉:regex} || {following.*following:regex}) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text})"; break;
      case '유튜브_좋아요': conditionPart = "({like:customml}) && (({구독중:text} || {구독 중:text} || {구독충:text} || {subscribed:text}) || ({youtube_subscribe_alarm_all:customml} || {youtube_subscribe_alarm:customml} || {youtube_subscribe_no_alarm:customml})) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text})"; break;
      case '좋아요': conditionPart = "({기본:text} || {즐겨찾기:text} || {좋아요:text} || {liked:text} || {팔로우:text}) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text} && !{취소:text})"; break;
      case '채널추가': conditionPart = "({kakao_channel:customml} || {kakao_channel_dark:customml} || {추가한 채널:text} || {추가한채널:text} || {추가완료:text} || {추가 완료:text} || {채널을 추가해:text} || {추가해 주셔서:text}) && (!{캡쳐하기:text} && !{적립받기:text}) && (!{예시:text} && !{ch +:text} && !{ch+:text} && !{취소:text})"; break;
      case '언론사 구독': conditionPart = "(!{뉴스판:text} && !{네이버 메인:text} && !{네이버메인:text} && !{스크린샷:text} && !{구독 이벤트:text} && !{구독이벤트:text})"; break;
      case '틱톡': conditionPart = "({tiktok_subscribe_humanicon1:customml} || {tiktok_subscribe_humanicon2:customml} || {tiktok_subscribe_sendicon:customml}) && (!{캡쳐하기:text} && !{적립받기:text} && !{예시:text})"; break;
      case 'X(트위터)': conditionPart = "({following.*following:regex} || {팔로잉.*팔로잉:regex}) || {twitter_X_subscribe_alarm:customml} && (!{캡쳐하기:text} && !{적립받기:text} && !{팔로우하기:text} && !{예시:text} && !{test:text} && !{가입하기:text})"; break;
     }
     if (conditionPart) {
      formData['CPA SUBSCRIBE_이미지 인식에 사용할 식별자'] = `${identifierPart} && ${conditionPart}`;
     }
    }
   }
  }
  
    const idPrefix = `${userName}-mod-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;
 const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

 const subject = `[광고 수정 요청] ${formData['대상 광고명'].split('\n')[0]}`;
 const uniqueSubject = `${subject} (ID: ${uniqueId})`;
  const messageId = sendModificationRequestNotification(userEmail, uniqueId, uniqueSubject, formData);

formData['등록ID'] = uniqueId;
    formData['등록일시'] = formattedTimestamp;
    formData['등록자'] = userEmail;
    formData['상태'] = '수정 요청 완료';
    formData['메일 스레드 ID'] = messageId;

    // --- newRow 생성 (이제 finalHeaders가 최신 상태이므로 수정 없음) ---
    const newRow = finalHeaders.map(header => {
      switch(header) {
        case '라이브 시작 시간':
        case '라이브 종료 시간':
          const timeValue = formData[header];
          return timeValue ? `'${timeValue}` : ''; // 텍스트로 저장
        // ▼▼▼ [추가] 새로 추가된 컬럼 값 처리 ▼▼▼
        case '광고 노출 중단 시작일시':
        case '광고 노출 중단 종료일시':
          const dateTimeValue = formData[header];
          return dateTimeValue || ''; // formData에 값이 있으면 사용, 없으면 빈 문자열
        // ▲▲▲ [추가] ▲▲▲
        default:
          if (Array.isArray(formData[header])) {
            return formData[header].join(', ');
          }
          return formData[header] || '';
      }
    });
    // --- newRow 생성 끝 ---

    try {
        sheet.appendRow(newRow);
        Logger.log('Successfully appended new row data.'); // 행 추가 성공 로그
    } catch (e) {
        Logger.log(`Error appending row: ${e.toString()}`); // 행 추가 에러 로그
    }

 logUserAction(userEmail, '수정 요청', {
 targetId: uniqueId,
 message: `광고 수정 '${subject}' 요청`
 });

 return { success: true, message: `광고 수정 요청이 완료되었습니다. (ID: ${uniqueId})` };
} catch (e) {
 console.error(`submitModificationRequest Error: ${e.toString()}`);
 return { success: false, message: `수정 요청 처리 중 오류가 발생했습니다: ${e.message}` };
} finally {
    // ▼▼▼▼▼ [추가] try...catch 작업이 끝나면 반드시 잠금을 해제합니다. ▼▼▼▼▼
    lock.releaseLock();
    // ▲▲▲▲▲ [추가] ▲▲▲▲▲
  }
}

// Code.gs 파일에서 이 함수 전체를 교체해주세요.

function sendModificationRequestNotification(senderEmail, modId, subject, data) {
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_mod&id=${modId}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_mod&id=${modId}`;
  const ccEmails = data.ccRecipients || '';

  let body = `<p>안녕하세요, 운영팀.</p>
    <p><b>${senderEmail}</b>님께서 광고 수정을 요청했습니다.</p>
    <p><b>수정 ID: ${modId}</b></p>
    <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
      <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 수정 담당하기 ]</a>
      <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 수정 완료 ]</a>
      <br><br>
      <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
      <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">광고 등록 시스템 바로가기</a>
    </div>
    <hr><h3>요청 내용</h3>
    <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;

  const fieldOrder = [
  '주요 요청사항', '대상 캠페인 ID', '대상 광고 ID', '대상 광고명', '예약 반영 시점',
  '광고주 연동 토큰 값', '매체', '단가', '총물량', '리워드', '일물량',
  '광고 집행 시작 일시', '광고 집행 종료 일시', '광고 노출 중단 시작일시', '광고 노출 중단 종료일시', '광고 참여 시작 후 완료 인정 유효기간 (일단위)',
  '트래커', '완료 이벤트 이름', '트래커 추가 정보 입력', 'URL - 기본', 'URL - AOS', 'URL - IOS', 'URL - PC',
    '기본 URL',
    '상세전용랜딩 URL',
    '소재 경로', '적용 필요 항목',
  '라이브 시작 시간', '라이브 종료 시간', 'adid 타겟팅 모수파일', '데모타겟1', '데모타겟2',
  '2차 액션 팝업 사용', '2차 액션 팝업 이미지 링크', '2차 액션 팝업 타이틀', '2차 액션 팝업 액션 버튼명', '2차 액션 팝업 랜딩 URL',
  '문구 - 타이틀', '문구 - 서브', '문구 - 상세화면 상단 타이틀', '문구 - 서브1 상단', '문구 - 서브1 하단',
  '액션 버튼', '문구 - 서브2', '노출 대상', '기타', '광고 타입별 추가',
    // 광고 타입별 필드 순서 정의
    '쿠키오븐 CPS_최소 결제 금액', '쿠키오븐 CPS_파트너 광고주 타입', '쿠키오븐 CPS_파트너 광고주 ID', '쿠키오븐 CPS_참여 경로 유형(app/web)',
    '네이버페이 알림받기_(메타) NF 광고주 연동 타입', '네이버페이 알림받기_(메타) NF 광고주 연동 ID', '네이버페이 알림받기_URL',
    '네이버페이 CPS_본광고_URL', '네이버페이 CPS_본광고_최소 결제 금액', '네이버페이 CPS_본광고_(목록) 리워드 조건 설명', '네이버페이 CPS_본광고_(목록) 리워드 텍스트', '네이버페이 CPS_본광고_(메타) NF 광고주 연동 ID', '네이버페이 CPS_본광고_(메타) 클릭 리워드 지급 금액',
    '네이버페이 CPS_부스팅_복사 필요한 광고 ID', '네이버페이 CPS_부스팅_URL & 상세 전용 랜딩 URL', '네이버페이 CPS_부스팅_문구 - 서브1 하단', '네이버페이 CPS_부스팅_최소 결제 금액', '네이버페이 CPS_부스팅_(목록) 리워드 조건 설명', '네이버페이 CPS_부스팅_부스팅 옵션', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_추천 세팅 여부', '네이버페이 CPS_부스팅_placement 세팅 정보 기본', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_카테고리',
    'CPQ_CPQ 뷰', 'CPQ_랜딩 형태', 'CPQ_임배디드 연결 형태', 'CPQ_유튜브 ID / 네이버 TV CODE', 'CPQ_이미지', 'CPQ_이미지 연결 링크', 'CPQ_퀴즈', 'CPQ_정답', 'CPQ_정답 placeholder 텍스트', 'CPQ_오답 alert 메시지', 'CPQ_사전 랜딩(딥링크) 사용', 'CPQ_사전 랜딩 실행 필수', 'CPQ_사전 랜딩 URL', 'CPQ_사전 랜딩 버튼 텍스트', 'CPQ_사전 랜딩 미실행 alert 메시지',
    'CPA SUBSCRIBE_구독 대상 이름', 'CPA SUBSCRIBE_이미지 인식에 사용할 식별자', 'CPA SUBSCRIBE_광고주 계정 식별자1', 'CPA SUBSCRIBE_광고주 계정 식별자2', 'CPA SUBSCRIBE_광고주 계정 식별자3', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL AOS', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL IOS'
  ];

    const fieldsToEscape = ['주요 요청사항', '문구 - 타이틀', '문구 - 서브', '문구 - 상세화면 상단 타이틀', '문구 - 서브1 상단', '문구 - 서브1 하단', '문구 - 서브2'];
    fieldsToEscape.push('CPA SUBSCRIBE_가이드 메세지');
    fieldsToEscape.push('CPA SUBSCRIBE 후지급_가이드 메세지');

  fieldOrder.forEach(key => {
    if (data[key]) {
      let value = data[key];

      if (key === '네이버페이 알림받기_URL' && value && !value.includes('click_key')) {
        value = `${value}?click_key={click_key}&ad_start_date={ad_start_at}&campaign_id={campaign_id}`;
      }
    
      if (data['광고 타입별 추가'] === '네이버페이 스마트스토어 CPS') {
        const boostingOption = data['네이버페이 CPS_부스팅_부스팅 옵션'];
        const priorityMap = { '부스팅_A': 1, '부스팅_B': 5, '부스팅_C': 15 };
        const priority = priorityMap[boostingOption] || 0;

        if (key === '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_추천 세팅 여부' && value === '세팅 O') {
        value = `네이버마케팅_추천(nvmarketing_best) : 우선순위 ${priority}`;
        }
        if (key === '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_카테고리') {
        const categoryMap = { '건강': 'nvmarketing_health', '식품': 'nvmarketing_food', '생활': 'nvmarketing_living', '뷰티': 'nvmarketing_beauty', '기타': 'nvmarketing_etc' };
        if (categoryMap[value]) {
          value = `네이버마케팅_${value}(${categoryMap[value]}) : 우선순위 ${priority}`;
        }
        }
        if (key === '네이버페이 CPS_부스팅_placement 세팅 정보 기본') {
        const basePlacementOptions = {
          '네이버쇼핑(nvshopping)': priority,
          '네이버마케팅(nvmarketing)': priority,
          '네이버마케팅_네앱(nvmarketing_nvapp)': priority,
          '쇼핑주문배송 구매 확정 띠배너(nvshopping_order_card)': 0,
          '쇼핑주문배송 하단 추천 영역(nvshopping_order_bottom)': 0,
          '(신)결제홈 결제내역 카드(historycard)': 0
        };
        const selectedOptions = Array.isArray(value) ? value : String(value).split(',').map(s => s.trim());
        value = selectedOptions.map(opt => `${opt} : 우선순위 ${basePlacementOptions[opt]}`).join('\n');
        }
      }

      const isEscapeTarget = fieldsToEscape.includes(key);
      const displayValue = isEscapeTarget ? String(value).replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>') : String(value).replace(/\n/g, '<br>');
     

            if (fieldsToEscape.includes(key)) {
        value = String(value).replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }
      body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${key.replace(/_/g, ' ')}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${displayValue}</td></tr>`;
    }
  });

  body += `</table>`;

  try {
    GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails }); // cc 옵션 추가
  } catch (e) {
    console.error(`수정 요청 이메일 발송 실패 (ID: ${modId}): ${e.toString()}`);
  }

  try {
    const slackMessage = { 'text': `${subject}` };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  } catch (e) {
    console.error(`수정 요청 슬랙 발송 실패 (ID: ${modId}): ${e.toString()}`);
    // 슬랙 발송이 실패해도 전체 프로세스가 중단되지 않도록 여기서 오류를 잡아줍니다.
  }

  // ▼▼▼ [핵심 수정] 스레드 ID가 아닌, 스레드에 포함된 첫 번째 메시지의 ID를 저장합니다. ▼▼▼
  Utilities.sleep(2000); // Gmail 검색이 안정적으로 되도록 대기 시간을 2초로 늘립니다.
  const threads = GmailApp.search(`subject:"${subject}" in:sent`, 0, 1);
  if (threads && threads.length > 0) {
    const messages = threads[0].getMessages();
    if (messages && messages.length > 0) {
      return messages[0].getId(); // 첫 번째 메시지의 ID를 반환
    }
  }
  return null; // 실패 시 null 반환
}

// Code.gs 파일의 recordModificationConfirmation 함수 전체를 이 코드로 교체하세요.

/**
 * 수정 요청 건에 대한 담당자를 지정하고, 원본 요청 스레드에 답장합니다.
 * @param {string} modId - 수정 요청 ID.
 * @param {string} approverEmail - 담당자 이메일.
 * @returns {string} 결과 메시지.
 */
function recordModificationConfirmation(modId, approverEmail) {
  const found = findRowById(modId, '수정');
  if (!found) return `수정 ID: ${modId} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const approverColIndex = headers.indexOf('담당자');
  const statusColIndex = headers.indexOf('상태');

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '스킵처리') {
    return `처리 실패: 이 수정 건(ID: ${modId})은 이미 스킵 처리되어 담당자로 지정할 수 없습니다.`;
  }
  
  const currentApprover = rowData[approverColIndex];
  if (currentApprover && currentApprover !== '') {
    return `처리 실패: 이 수정 건(ID: ${modId})은 이미 ${currentApprover} 님이 담당하고 있습니다.`;
  }
  
  sheet.getRange(rowIndex, approverColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  
  const timestampColIndex = headers.indexOf('담당자 확인 일시');
  if (timestampColIndex > -1) {
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(formattedTimestamp);
  }

  try {
    const searchQuery = `"수정 ID: ${modId}"`;
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads && threads.length > 0) {
      threads[0].replyAll("", { 
        htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>수정 ID: ${modId}</b> 건의 담당자로 지정되어 수정을 진행합니다.</p><p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p><p>감사합니다.</p>`
      });
    } else {
        console.error(`담당자 지정 알림 실패: 수정 ID ${modId}에 대한 메일 스레드를 찾을 수 없습니다.`);
    }
  } catch (e) {
    console.error(`수정 담당자 지정 메일 발송 중 오류 발생: ${e.toString()}`);
  }
  // ▲▲▲ [핵심 수정] ▲▲▲
  
  return `ID: ${modId} 수정 건의 담당자로 ${approverEmail}님이 지정되었습니다. 이 창은 닫아도 됩니다.`;
}

function processModificationCompletion(modId, completerEmail) {
  const found = findRowById(modId, '수정');
  if (!found) return { success: false, message: `수정 ID(${modId})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('상태');
  
  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '수정 완료') {
    return { success: false, message: `이미 수정 완료 처리된 건입니다. (ID: ${modId})` };
  }

  // --- ▼▼▼ [수정] 담당자 자동 지정 로직 추가 ▼▼▼ ---
  const managerColIndex = headers.indexOf('담당자');
  const confirmDateColIndex = headers.indexOf('담당자 확인 일시');
  const now = new Date();
  
  const currentManager = (managerColIndex > -1) ? rowData[managerColIndex] : '';
  
  // '담당자' 필드가 비어있는 경우 (수정 담당하기를 건너뛴 경우) 자동 지정
  if (managerColIndex > -1 && currentManager === '') { 
    
    // 1. 담당자 지정 (수정 완료를 누른 사용자)
    sheet.getRange(rowIndex, managerColIndex + 1).setValue(completerEmail);
    
    // 2. 담당자 확인 일시 지정 (수정 완료 시간 - 5분)
    if (confirmDateColIndex > -1) {
      const confirmedTime = new Date(now.getTime() - (5 * 60 * 1000));
      const confirmedTimestamp = Utilities.formatDate(confirmedTime, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(rowIndex, confirmDateColIndex + 1).setValue(confirmedTimestamp);
    }
  }
  // --- ▲▲▲ [수정] 담당자 자동 지정 로직 추가 ▲▲▲ ---
  
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('수정 완료');
  
  const completionDateColIndex = headers.indexOf('수정 완료 일시');
  if (completionDateColIndex > -1) {
    // 수정 완료 일시는 현재 시간으로 기록
    const timestamp = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(timestamp);
  }

  logUserAction(completerEmail, '수정 완료 처리', {
    targetId: modId,
    message: `수정 ID '${modId}' 완료 처리`
  });

  return { success: true, message: `수정 건(ID: ${modId})이 성공적으로 완료 처리되었습니다. 이 창은 닫아도 됩니다.` };
}


function getModificationDataById(modId) {
  const found = findRowById(modId, '수정');
  if (found) {
    const modData = {};
    found.headers.forEach((header, index) => {
      let value = found.rowData[index];
      if (value instanceof Date) {
        try {
          // ▼▼▼▼▼ [수정] 필드 이름에 따라 날짜 형식을 다르게 적용합니다. ▼▼▼▼▼
          if (header.endsWith('라이브 시작 시간') || header.endsWith('라이브 종료 시간')) {
            value = Utilities.formatDate(value, "Asia/Seoul", "HH:mm");
          } else if (header.endsWith('일자')) { // '일자'로 끝나는 필드
            value = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
          } else { // '일시'로 끝나는 필드 (시간 포함)
            value = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm");
          }
          // ▲▲▲▲▲ [수정] ▲▲▲▲▲
        } catch(e) {
          value = '날짜 형식 오류';
        }
      }
      modData[header] = value;
    });
    return modData;
  }
  return null;
}



function processModificationSkip(modId) {
  const found = findRowById(modId, '수정'); // '수정' 타입으로 검색
  if (found) {
    const skipperEmail = Session.getActiveUser().getEmail();
    const statusColIndex = found.headers.indexOf('상태');
    found.sheet.getRange(found.rowIndex, statusColIndex + 1).setValue('스킵처리');

    const threadIdColIndex = found.headers.indexOf('메일 스레드 ID');
    const threadId = (threadIdColIndex > -1) ? found.rowData[threadIdColIndex] : null;
    if (threadId) {
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.replyAll("", {
            htmlBody: `<p>안녕하세요,</p><p>요청하신 <b>수정 ID: ${modId}</b> 건이 <b>스킵 처리</b>되었음을 알려드립니다.</p><p>감사합니다.</p><p>- 처리자: ${skipperEmail}</p>`,
          });
        }
      } catch (e) {
        console.error(`수정 스킵 알림 메일 발송 실패(ID: ${modId}): ${e.toString()}`);
      }
    }

    const adName = found.rowData[found.headers.indexOf('대상 광고명')] || modId;
    const subject = String(adName).split('\n')[0];

    const slackMessage = { 'text': `[수정 스킵 처리] - ${subject} (ID: ${modId})` };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
    try {
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    } catch(e) {
      console.error(`수정 스킵 알림 슬랙 발송 실패 (ID: ${modId}): ${e.toString()}`);
    }

    logUserAction(skipperEmail, '수정 스킵 처리', {
      targetId: modId,
      message: `수정 ID '${modId}' 스킵 처리`
    });

    return { success: true, message: `수정 ID(${modId})가 성공적으로 스킵 처리되었습니다.` };
  }
  return { success: false, message: `수정 ID(${modId})를 찾을 수 없습니다.` };
}

function processModificationRejection(modId, reason) {
  try {
    const rejectorEmail = Session.getActiveUser().getEmail(); // 현재 사용자 (운영팀)
    const found = findRowById(modId, '수정');
    if (!found) {
      return { success: false, message: `수정 ID(${modId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    // 시트 상태 업데이트 (기존과 동일)
    const statusColIndex = headers.indexOf('상태');
    const rejectionDateColIndex = headers.indexOf('반려 일시');
    const rejectionReasonColIndex = headers.indexOf('반려 사유');
    const registrantColIndex = headers.indexOf('등록자');

    if ([statusColIndex, rejectionDateColIndex, rejectionReasonColIndex, registrantColIndex].includes(-1)) {
      return { success: false, message: '시트에서 필수 컬럼(상태, 반려 일시, 반려 사유, 등록자)을 찾을 수 없습니다.' };
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('반려');
    sheet.getRange(rowIndex, rejectionDateColIndex + 1).setValue(timestamp);
    sheet.getRange(rowIndex, rejectionReasonColIndex + 1).setValue(reason);

    const registrantEmail = rowData[registrantColIndex]; // 원본 요청자 (영업팀)
    
    if (registrantEmail) {
      // ▼▼▼ [수정] 메일 본문에 시스템 링크를 추가하고 검색 방식으로 변경합니다. ▼▼▼
      const subject = `[광고 등록 시스템] 요청하신 수정(ID: ${modId})이 반려되었습니다.`;
      let emailBody = `<p>안녕하세요, ${registrantEmail.split('@')[0]}님.</p>
                       <p>요청하신 수정(ID: <b>${modId}</b>)이 아래와 같은 사유로 반려되었습니다.</p>`;
      if (reason) {
        emailBody += `<p style="margin-top:20px;"><b>반려 사유:</b></p>
                      <div style="padding: 12px; border: 1px solid #ddd; background-color: #f9f9f9; border-radius: 5px; margin-top: 5px;">
                        ${reason.replace(/\n/g, '<br>')}
                      </div>`;
      }
      emailBody += `<p style="margin-top:20px;">수정 후 재요청하시거나 담당자(${rejectorEmail})에게 문의해주세요.</p>
                    <p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p>
                    <p>감사합니다.</p>`;

      const mailOptions = { 
        htmlBody: emailBody,
        cc: registrantEmail
      };
      
      const searchQuery = `"수정 ID: ${modId}"`;
      const threads = GmailApp.search(searchQuery, 0, 1);

      if (threads && threads.length > 0) {
        threads[0].replyAll('', mailOptions);
      } else {
        console.error(`Could not find thread for modId: ${modId}. Sending a new email as a fallback.`);
        GmailApp.sendEmail(registrantEmail, subject, '', mailOptions);
      }
      // ▲▲▲ [수정] ▲▲▲
    }

    logUserAction(rejectorEmail, '수정 반려 처리', {
      targetId: modId,
      message: `수정 ID '${modId}' 반려 처리. 사유: ${reason}`
    });

    return { success: true, message: `수정 ID(${modId})가 성공적으로 반려 처리 및 메일 발송되었습니다.` };
  } catch (e) {
    console.error(`Error in processModificationRejection: ${e.toString()}`);
    return { success: false, message: '수정 반려 처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}

function submitOtherRequest(formData) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const ccEmails = formData.ccRecipients || '';
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    const sheetName = '기타 요청';
    let sheet = ss.getSheetByName(sheetName);
    
    // 관리 및 데이터 컬럼 정의
    const headers = [
      'id', 'timestamp', 'registrant', 'status', 'manager', 'manager_timestamp', 'completion_timestamp', // 시스템 관리용
      'request_type', 'advertiser', 'subject', 'content', // 주요 정보
      'campaign_name', 'campaign_id', 'priority', 'image_path', // 세부 정보
      'popup_start', 'popup_end', 'popup_type', 'popup_group',
      'banner_start', 'banner_end', 'banner_new_end', 'banner_text', 'banner_type', 'banner_bg_color', 'banner_group'
    ];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName, 0);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      // 기존 시트가 있다면 헤더 확인 (필요 시 마이그레이션 로직 추가 가능, 여기선 생략)
    }
    
    // 1. ID 생성
    const idPrefix = `other-${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;

    // 2. 제목 생성
    const requestType = formData['선택항목'];
    const campaignName = formData['캠페인명'] || '';
    const advertiser = formData['광고주명'] || '';
    
    let subject;
    if (requestType === '팝업요청') {
      subject = `[팝업 등록 요청] ${campaignName}`;
    } else if (requestType === '배너요청') {
      subject = `[배너 등록 요청] ${campaignName}`;
    } else if (requestType === '채널링요청') {
      subject = `[채널링 등록 요청] ${campaignName}`;
    } else {
      const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyMMdd");
      subject = `[기타 등록 요청] ${advertiser}_${today}`;
    }

    subject = `${subject} (${uniqueId})`;

    // 3. 알림 발송 (HTML 생성 및 메일/슬랙 전송)
    sendOtherRequestNotification(userEmail, uniqueId, subject, formData);

    // 4. 시트 저장
    const newRow = [
      uniqueId, formattedTimestamp, userEmail, '등록 요청 완료', '', '', '', // 시스템 컬럼 초기값
      requestType, advertiser, subject, formData['요청사항'],
      campaignName, formData['캠페인 ID'], formData['우선순위'], formData['이미지 경로'],
      formData['팝업 노출 시작 일시'], formData['팝업 노출 종료 일시'], formData['팝업 표시 타입'], formData['팝업 그룹'],
      formData['배너 노출 시작 일시'], formData['배너 노출 종료 일시'], formData['배너 NEW 표시 종료일시'], formData['배너 텍스트'], formData['배너 표시 타입'], formData['배너 배경색상'], formData['배너 그룹']
    ];

    sheet.appendRow(newRow);

    logUserAction(userEmail, '기타 요청', {
      targetId: uniqueId,
      message: `${subject}`
    });

    return { success: true, message: `기타 요청이 완료되었습니다. (ID: ${uniqueId})` };
  } catch (e) {
    console.error(`submitOtherRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function sendOtherRequestNotification(senderEmail, id, subject, formData) {
  const ccEmails = formData.ccRecipients || '';
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_other&id=${id}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_other&id=${id}`;

  let body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 기타 요청을 등록했습니다.</p>
              <p><b>ID: ${id}</b></p>
              <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
                <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 요청 담당하기 ]</a>
                <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 처리 완료 ]</a>
                <br><br>
                <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
                <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">시스템 바로가기</a>
              </div>
              <hr>
              <h3>요청 내용</h3>
              <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;

  const emailFieldOrder = [
      '광고주명', '요청사항', '선택항목', 
      '캠페인명', '캠페인 ID',
      '팝업 노출 시작 일시', '팝업 노출 종료 일시', '팝업 표시 타입', 
      '배너 노출 시작 일시', '배너 노출 종료 일시', '배너 NEW 표시 종료일시',
      '배너 텍스트', '배너 표시 타입', '배너 배경색상', 
      '우선순위', '이미지 경로',
      '팝업 그룹', '배너 그룹'
  ];

  emailFieldOrder.forEach(field => {
    if (formData[field]) {
      let value = String(formData[field]);
      if (field === '요청사항') {
        value = value.replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }
      body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${field}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${value.replace(/\n/g, '<br>')}</td></tr>`;
    }
  });
  body += `</table>`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails });
  
  try {
    const slackMessage = { 'text': `${subject}` };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) });
  } catch (e) {
    console.error(`기타 요청 슬랙 발송 실패: ${e.toString()}`);
  }
}

function findOtherRowById(id) {
  const sheet = ss.getSheetByName("기타 요청");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true);
  const foundCell = textFinder.findNext();
  if (foundCell) {
    const rowIndex = foundCell.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { sheet, rowIndex, headers, rowData };
  }
  return null;
}

function recordOtherConfirmation(id, approverEmail) {
  const found = findOtherRowById(id);
  if (!found) return `기타 요청 ID: ${id} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const timestampColIndex = headers.indexOf('manager_timestamp');

  if (managerColIndex === -1) return '필수 컬럼(manager)이 없습니다.';

  const currentManager = rowData[managerColIndex];
  if (currentManager && currentManager !== '') {
    return `처리 실패: 이 건(ID: ${id})은 이미 ${currentManager} 님이 담당하고 있습니다.`;
  }

  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  if (timestampColIndex > -1) {
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  }

  // 메일 답장 (제목으로 검색)
  try {
    const subject = rowData[headers.indexOf('subject')]; // 저장된 제목 사용
    if (subject) {
        const threads = GmailApp.search(`subject:"${subject}"`, 0, 1);
        if (threads && threads.length > 0) {
            threads[0].replyAll("", {
                htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>기타 요청 ID: ${id}</b> 건의 담당자로 지정되어 처리를 진행합니다.</p><p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>`
            });
        }
    }
  } catch (e) {
    console.error(`기타 요청 담당자 알림 발송 오류: ${e.toString()}`);
  }

  return `기타 요청 ID: ${id} 건의 담당자로 ${approverEmail}님이 지정되었습니다.`;
}

function processOtherCompletion(id, completerEmail) {
  const found = findOtherRowById(id);
  if (!found) return { success: false, message: `기타 요청 ID(${id})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  const completionDateColIndex = headers.indexOf('completion_timestamp');

  if (statusColIndex === -1) return { success: false, message: 'status 컬럼을 찾을 수 없습니다.' };

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '완료') {
    return { success: false, message: `이미 완료 처리된 건입니다. (ID: ${id})` };
  }

  sheet.getRange(rowIndex, statusColIndex + 1).setValue('완료');
  if (completionDateColIndex > -1) {
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  }

  logUserAction(completerEmail, '기타 요청 완료', { targetId: id });
  return { success: true, message: `기타 요청 건(ID: ${id})이 성공적으로 완료 처리되었습니다.` };
}


function getNextSequentialId(sheet, prefix) {
  if (sheet.getLastRow() < 2) {
    return 1; // 헤더만 있는 경우 1번부터 시작
  }
  
  const ids = sheet.getRange("A2:A" + sheet.getLastRow()).getValues()
                   .flat()
                   .filter(id => id && id.startsWith(prefix));

  if (ids.length === 0) {
    return 1; // 해당 접두사를 가진 ID가 하나도 없는 경우 1번부터 시작
  }

  const numbers = ids.map(id => {
    const numberPart = id.substring(prefix.length);
    return parseInt(numberPart, 10) || 0;
  });

  const maxNumber = Math.max(...numbers);
  return maxNumber + 1;
}

function submitCxRequest(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheetName = "CX팀";
    let sheet = ss.getSheetByName(sheetName);

    // ▼▼▼ [수정] 담당자 지정/완료 처리를 위한 필수 컬럼 추가 ▼▼▼
    const headers = [
      'id', 
      'timestamp', 
      'registrant', 
      'status',              // 상태 (대기/처리중/완료)
      'manager',             // 담당자
      'manager_timestamp',   // 담당자 지정 일시
      'completion_timestamp',// 완료 일시
      'auto_generated_title', 
      'request_content'
    ];
    // ▲▲▲ [수정] ▲▲▲

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      // 기존 시트가 있다면 헤더 확인 후 누락된 컬럼 추가 로직이 필요할 수 있음
      // (현재는 새 컬럼이 뒤에 붙는 구조가 아니라 중간에 삽입되므로, 
      //  테스트 중인 'CX팀' 시트를 삭제하고 다시 생성하는 것을 권장합니다.)
    }

    // 1. 제목 자동 생성 로직
    const today = new Date();
    const yymmdd = Utilities.formatDate(today, "Asia/Seoul", "yyMMdd");
    const baseTitle = `[애디슨오퍼월_광고생성요청] CS지급용 광고 생성요청_${yymmdd}`;
    
    let count = 0;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dateValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
      const todayStr = Utilities.formatDate(today, "Asia/Seoul", "yyyy-MM-dd");
      
      count = dateValues.filter(date => {
        if (date instanceof Date) return Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd") === todayStr;
        if (typeof date === 'string' && date.length >= 10) return date.substring(0, 10) === todayStr;
        return false;
      }).length;
    }

    const finalTitle = count === 0 ? baseTitle : `${baseTitle}_(${count + 1})`;

    // 2. ID 생성
    const idPrefix = `cx-${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;
    const subjectWithId = `${finalTitle} (${uniqueId})`;
    const formattedTimestamp = Utilities.formatDate(today, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    // 3. 알림 발송
    sendCxNotification(userEmail, uniqueId, subjectWithId, formData);

    // 4. 시트 저장 (새로운 헤더 순서에 맞춤)
    const newRow = [
      uniqueId,           // id
      formattedTimestamp, // timestamp
      userEmail,          // registrant
      '등록 요청 완료',     // status (초기 상태)
      '',                 // manager (초기 공란)
      '',                 // manager_timestamp (초기 공란)
      '',                 // completion_timestamp (초기 공란)
      subjectWithId,         // auto_generated_title
      formData['요청 내용'] // request_content
    ];
    sheet.appendRow(newRow);

    logUserAction(userEmail, 'CX팀 요청', {
      targetId: uniqueId,
      message: subjectWithId
    });

    return { success: true, message: `CX팀 요청이 완료되었습니다. (ID: ${uniqueId})` };

  } catch (e) {
    console.error(`submitCxRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function sendCxNotification(senderEmail, id, subject, formData) {
  const ccEmails = formData.ccRecipients || '';
  const requestContent = formData['요청 내용']
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>')
    // URL 자동 링크 변환
    .replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank">$1</a>');


  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_cx&id=${id}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_cx&id=${id}`;
  
let body = `<p>안녕하세요, 운영팀.</p>
  <p><b>${senderEmail}</b>님께서 CX팀 요청을 등록했습니다.</p>
  <p><b>ID: ${id}</b></p>
  
  <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
    <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 요청 담당하기 ]</a>
    <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 처리 완료 ]</a>
    <br><br>
    <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
    <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">시스템 바로가기</a>
  </div>
  <hr>
  <h3>${subject}</h3>
  <div style="padding: 15px; border: 1px solid #e0e0e0; background-color: #f9f9f9; border-radius: 5px;">
    ${requestContent}
  </div>`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails });

  try {
    const slackMessage = { 'text': subject };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(slackMessage) });
  } catch (e) {
    console.error(`CX 슬랙 발송 실패: ${e.toString()}`);
  }
}



function submitBdRequest(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheetName = "오퍼월사업팀";
    let sheet = ss.getSheetByName(sheetName);

    // 영문 헤더 (status, mail_thread_id 제외, 소재 제외)
    const headers = [
      'id', 
      'timestamp', 
      'registrant', 
      'status',              // 상태
      'manager',             // 담당자
      'manager_timestamp',   // 담당자 지정 일시
      'completion_timestamp',// 완료 일시
      'auto_generated_title', 
      'request_title', 
      'request_content'
    ];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // 1. 제목 생성
    const today = new Date();
    const yymmdd = Utilities.formatDate(today, "Asia/Seoul", "yyMMdd");
    const requestTitle = formData['요청제목'];
    const finalTitle = `[오퍼월사업팀_요청] ${requestTitle}_${yymmdd}`;

    // 2. ID 생성
    const idPrefix = `bd-${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;
    const subjectWithId = `${finalTitle} (${uniqueId})`;
    const formattedTimestamp = Utilities.formatDate(today, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    // 3. 첨부파일 처리 (Blob 변환)
    let blobs = [];
    if (formData['attachments_json']) {
      const filesData = JSON.parse(formData['attachments_json']);
      blobs = filesData.map(file => {
        const decoded = Utilities.base64Decode(file.data);
        return Utilities.newBlob(decoded, file.type, file.name);
      });
    }

    // 4. 알림 발송
    sendBdNotification(userEmail, uniqueId, subjectWithId, formData, blobs);

    // 5. 시트 저장 (소재 제외)
    const newRow = [
      uniqueId,           // id
      formattedTimestamp, // timestamp
      userEmail,          // registrant
      '등록 요청 완료',     // status (초기 상태)
      '',                 // manager (초기 공란)
      '',                 // manager_timestamp (초기 공란)
      '',                 // completion_timestamp (초기 공란)
      subjectWithId,         // auto_generated_title
      requestTitle,       // request_title
      formData['요청내용'] // request_content
    ];
    sheet.appendRow(newRow);

    logUserAction(userEmail, '오퍼월사업팀 요청', {
      targetId: uniqueId,
      message: subjectWithId
    });

    return { success: true, message: `오퍼월사업팀 요청이 완료되었습니다. (ID: ${uniqueId})` };

  } catch (e) {
    console.error(`submitBdRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function sendBdNotification(senderEmail, id, subject, formData, blobs) {
  const ccEmails = formData.ccRecipients || '';
  const requestContent = formData['요청내용']
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>')
    .replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank">$1</a>');

    const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_bd&id=${id}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_bd&id=${id}`;

let body = `<p>안녕하세요, 운영팀.</p>
  <p><b>${senderEmail}</b>님께서 오퍼월사업팀 요청을 등록했습니다.</p>
  <p><b>ID: ${id}</b></p>
  
  <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
    <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 요청 담당하기 ]</a>
    <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 처리 완료 ]</a>
    <br><br>
    <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
    <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">시스템 바로가기</a>
  </div>
  <hr>
  <h3>${subject}</h3>
  <p><b>요청 제목:</b> ${formData['요청제목']}</p>
  <div style="padding: 15px; border: 1px solid #e0e0e0; background-color: #f9f9f9; border-radius: 5px;">
    ${requestContent}
  </div>
  <br>
  <p>※ 첨부파일은 이 메일에 포함되어 있습니다.</p>`;
  const mailOptions = { 
    htmlBody: body, 
    cc: ccEmails,
    attachments: blobs // 첨부파일 추가
  };

  const bdRecipients = 'choi.byoungyoul@nbt.com,operation@nbt.com,sales@nbt.com,biz.dev@nbt.com,cx@nbt.com';
  // const bdRecipients = 'choi.byoungyoul@nbt.com';
  GmailApp.sendEmail(bdRecipients, subject, '', mailOptions);

  try {
    const slackMessage = { 'text': subject };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(slackMessage) });
  } catch (e) {
    console.error(`BD 슬랙 발송 실패: ${e.toString()}`);
  }
}


function findCxRowById(cxId) {
  const sheet = ss.getSheetByName("CX팀");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(cxId).matchEntireCell(true);
  const foundCell = textFinder.findNext();
  if (foundCell) {
    const rowIndex = foundCell.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { sheet, rowIndex, headers, rowData };
  }
  return null;
}

/**
 * CX팀 담당자 지정
 */
function recordCxConfirmation(cxId, approverEmail) {
  const found = findCxRowById(cxId);
  if (!found) return `CX 요청 ID: ${cxId} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const timestampColIndex = headers.indexOf('manager_timestamp');

  if (managerColIndex === -1 || statusColIndex === -1) return '필수 컬럼이 시트에 없습니다.';

  const currentManager = rowData[managerColIndex];
  if (currentManager && currentManager !== '') {
    return `처리 실패: 이 건(ID: ${cxId})은 이미 ${currentManager} 님이 담당하고 있습니다.`;
  }

  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  
  if (timestampColIndex > -1) {
    const now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(now);
  }
try {
    // 본문에 ID가 포함되어 있으므로 ID로 스레드를 검색합니다.
    const searchQuery = `"${cxId}"`; 
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads && threads.length > 0) {
      threads[0].replyAll("", {
        htmlBody: `<p>안녕하세요,</p>
        <p><b>${approverEmail}</b> 님이 <b>CX 요청 ID: ${cxId}</b> 건의 담당자로 지정되어 처리를 진행합니다.</p>
        <p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>
        <p>감사합니다.</p>`
      });
    } else {
      console.log(`CX 담당자 지정 알림 실패: ${cxId} 관련 메일을 찾을 수 없습니다.`);
    }
  } catch (e) {
    console.error(`CX 담당자 지정 알림 발송 중 오류: ${e.toString()}`);
  }
  // ▲▲▲ [추가] ▲▲▲

  return `CX 요청 ID: ${cxId} 건의 담당자로 ${approverEmail}님이 지정되었습니다.`;
}

/**
 * CX팀 완료 처리
 */
function processCxCompletion(cxId, completerEmail) {
  const found = findCxRowById(cxId);
  if (!found) return { success: false, message: `CX 요청 ID(${cxId})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  const completionDateColIndex = headers.indexOf('completion_timestamp');

  if (statusColIndex === -1) return { success: false, message: 'status 컬럼을 찾을 수 없습니다.' };

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '완료') {
    return { success: false, message: `이미 완료 처리된 건입니다. (ID: ${cxId})` };
  }

  const now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('완료');
  
  if (completionDateColIndex > -1) {
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(now);
  }

  logUserAction(completerEmail, 'CX 요청 완료', { targetId: cxId });
  return { success: true, message: `CX 요청 건(ID: ${cxId})이 성공적으로 완료 처리되었습니다.` };
}


function findBdRowById(bdId) {
  const sheet = ss.getSheetByName("오퍼월사업팀");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(bdId).matchEntireCell(true);
  const foundCell = textFinder.findNext();
  if (foundCell) {
    const rowIndex = foundCell.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { sheet, rowIndex, headers, rowData };
  }
  return null;
}

function recordBdConfirmation(bdId, approverEmail) {
  const found = findBdRowById(bdId);
  if (!found) return `오퍼월사업팀 요청 ID: ${bdId} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const timestampColIndex = headers.indexOf('manager_timestamp');

  if (managerColIndex === -1 || statusColIndex === -1) return '필수 컬럼이 시트에 없습니다.';

  const currentManager = rowData[managerColIndex];
  if (currentManager && currentManager !== '') {
    return `처리 실패: 이 건(ID: ${bdId})은 이미 ${currentManager} 님이 담당하고 있습니다.`;
  }

  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  
  if (timestampColIndex > -1) {
    const now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(now);
  }

try {
    const searchQuery = `"${bdId}"`;
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads && threads.length > 0) {
      threads[0].replyAll("", {
        htmlBody: `<p>안녕하세요,</p>
        <p><b>${approverEmail}</b> 님이 <b>오퍼월사업팀 요청 ID: ${bdId}</b> 건의 담당자로 지정되어 처리를 진행합니다.</p>
        <p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>
        <p>감사합니다.</p>`
      });
    } else {
      console.log(`BD 담당자 지정 알림 실패: ${bdId} 관련 메일을 찾을 수 없습니다.`);
    }
  } catch (e) {
    console.error(`BD 담당자 지정 알림 발송 중 오류: ${e.toString()}`);
  }
  // ▲▲▲ [추가] ▲▲▲

  return `오퍼월사업팀 요청 ID: ${bdId} 건의 담당자로 ${approverEmail}님이 지정되었습니다.`;
}

/**
 * 오퍼월사업팀 완료 처리
 */
function processBdCompletion(bdId, completerEmail) {
  const found = findBdRowById(bdId);
  if (!found) return { success: false, message: `오퍼월사업팀 요청 ID(${bdId})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  const completionDateColIndex = headers.indexOf('completion_timestamp');

  if (statusColIndex === -1) return { success: false, message: 'status 컬럼을 찾을 수 없습니다.' };

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '완료') {
    return { success: false, message: `이미 완료 처리된 건입니다. (ID: ${bdId})` };
  }

  const now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('완료');
  
  if (completionDateColIndex > -1) {
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(now);
  }

  logUserAction(completerEmail, '오퍼월사업팀 요청 완료', { targetId: bdId });
  return { success: true, message: `오퍼월사업팀 요청 건(ID: ${bdId})이 성공적으로 완료 처리되었습니다.` };
}


function submitCopyCreationRequest(formData) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheetName = '복사 생성 요청';
    let sheet = ss.getSheetByName(sheetName);

    // 영문 헤더 정의
    const headers = [
      'id', 'timestamp', 'registrant', 'status', 'manager', 'manager_timestamp', 'completion_timestamp',
      'mail_thread_id',
      'request_details', // 주요 요청사항
      'campaign_id',     // 캠페인 ID
      'target_ad_id_to_modify',   // 수정 필요 광고 ID
      'target_ad_name_to_modify', // 수정 필요 광고명
      'source_ad_id',    // 복사 대상 광고 ID
      'modification_options_json' // 공통 항목 (선택) - JSON으로 저장
    ];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName, 0);
      sheet.appendRow(headers);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const idPrefix = `copy-${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    // 공통 항목(선택값) 추출 및 JSON 변환
    const modificationOptions = {};
    for (const key in formData) {
      if (!['주요 요청사항', '캠페인 ID', '수정 필요 광고 ID', '수정 필요 광고명', '복사 대상 광고 ID', 'ccRecipients'].includes(key)) {
        if (formData[key]) modificationOptions[key] = formData[key];
      }
    }

    // 이메일 제목 생성
    const targetAdName = formData['수정 필요 광고명'] || '(광고명 미입력)';
    const yymmdd = Utilities.formatDate(new Date(), "Asia/Seoul", "yyMMdd");
    const subject = `[광고 수정,생성_요청] ${targetAdName}_${yymmdd} (ID: ${uniqueId})`;

    // 알림 발송
    const messageId = sendCopyCreationNotification(userEmail, uniqueId, subject, formData, modificationOptions);

    const newRow = [
      uniqueId, formattedTimestamp, userEmail, '등록 요청 완료', '', '', '',
      messageId,
      formData['주요 요청사항'],
      formData['캠페인 ID'],
      formData['수정 필요 광고 ID'],
      formData['수정 필요 광고명'],
      formData['복사 대상 광고 ID'],
      JSON.stringify(modificationOptions, null, 2)
    ];

    sheet.appendRow(newRow);

    logUserAction(userEmail, '복사 생성 요청', { targetId: uniqueId, message: subject });

    return { success: true, message: `복사 생성 요청이 완료되었습니다. (ID: ${uniqueId})` };
  } catch (e) {
    console.error(`submitCopyCreationRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function sendCopyCreationNotification(senderEmail, id, subject, formData, modificationOptions) {
  const ccEmails = formData.ccRecipients || '';
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_copy&id=${id}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_copy&id=${id}`;

  let body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 복사 생성을 요청했습니다.</p>
              <p><b>ID: ${id}</b></p>
              <div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
                <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 요청 담당하기 ]</a>
                <a href="${completionUrl}" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">[ 처리 완료 ]</a>
                <br><br>
                <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
                <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">시스템 바로가기</a>
              </div>
              <hr>
              <h3>요청 내용</h3>
              <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;

  const mainFields = [
    '주요 요청사항', '캠페인 ID', '복사 대상 광고 ID', '수정 필요 광고 ID', '수정 필요 광고명'
  ];

  mainFields.forEach(key => {
    if (formData[key]) {
      let value = String(formData[key]);
      if (key === '주요 요청사항') {
        value = value.replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }
      body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${key}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${value.replace(/\n/g, '<br>')}</td></tr>`;
    }
  });

  if (Object.keys(modificationOptions).length > 0) {
     body += `<tr><td colspan="2" style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f0f0f0; font-weight: bold; text-align: center;">공통 항목 (선택)</td></tr>`;
     
     // 이메일에 표시할 순서대로 필드명을 정의합니다.
     const orderedKeys = [
       '광고주 연동 토큰 값', '매체', '단가', '총물량', '리워드', '일물량',
       '광고 집행 시작 일시', '광고 집행 종료 일시', '광고 노출 중단 시작일시', '광고 노출 중단 종료일시',
       '광고 참여 시작 후 완료 인정 유효기간 (일단위)', '트래커', '완료 이벤트 이름', '트래커 추가 정보 입력',
       'URL - 기본', 'URL - AOS', 'URL - IOS', 'URL - PC',
       '기본 URL', '상세전용랜딩 URL', // 네이버페이 CPC 전용 하위 필드
       '소재 경로', '적용 필요 항목', '라이브 시작 시간', '라이브 종료 시간',
       'adid 타겟팅 모수파일', '데모타겟1', '데모타겟2',
       '2차 액션 팝업 사용', '2차 액션 팝업 이미지 링크', '2차 액션 팝업 타이틀', '2차 액션 팝업 액션 버튼명', '2차 액션 팝업 랜딩 URL',
       '문구 - 타이틀', '문구 - 서브', '문구 - 상세화면 상단 타이틀', '문구 - 서브1 상단', '문구 - 서브1 하단',
       '액션 버튼', '문구 - 서브2', '노출 대상', '기타', '광고 타입별 추가',
       // 광고 타입별 추가 필드들
       '쿠키오븐 CPS_최소 결제 금액', '쿠키오븐 CPS_파트너 광고주 타입', '쿠키오븐 CPS_파트너 광고주 ID', '쿠키오븐 CPS_참여 경로 유형(app/web)',
       '네이버페이 알림받기_(메타) NF 광고주 연동 타입', '네이버페이 알림받기_(메타) NF 광고주 연동 ID', '네이버페이 알림받기_URL',
       '네이버페이 CPS_본광고_URL', '네이버페이 CPS_본광고_최소 결제 금액', '네이버페이 CPS_본광고_(목록) 리워드 조건 설명', '네이버페이 CPS_본광고_(목록) 리워드 텍스트', '네이버페이 CPS_본광고_(메타) NF 광고주 연동 ID', '네이버페이 CPS_본광고_(메타) 클릭 리워드 지급 금액',
       '네이버페이 CPS_부스팅_복사 필요한 광고 ID', '네이버페이 CPS_부스팅_URL & 상세 전용 랜딩 URL', '네이버페이 CPS_부스팅_문구 - 서브1 하단', '네이버페이 CPS_부스팅_최소 결제 금액', '네이버페이 CPS_부스팅_(목록) 리워드 조건 설명', '네이버페이 CPS_부스팅_부스팅 옵션', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_추천 세팅 여부', '네이버페이 CPS_부스팅_placement 세팅 정보 기본', '네이버페이 CPS_부스팅_placement 세팅 정보 옵션_카테고리',
       'CPQ_CPQ 뷰', 'CPQ_랜딩 형태', 'CPQ_임배디드 연결 형태', 'CPQ_유튜브 ID / 네이버 TV CODE', 'CPQ_이미지', 'CPQ_이미지 연결 링크', 'CPQ_퀴즈', 'CPQ_정답', 'CPQ_정답 placeholder 텍스트', 'CPQ_오답 alert 메시지', 'CPQ_사전 랜딩(딥링크) 사용', 'CPQ_사전 랜딩 실행 필수', 'CPQ_사전 랜딩 URL', 'CPQ_사전 랜딩 버튼 텍스트', 'CPQ_사전 랜딩 미실행 alert 메시지',
       'CPA SUBSCRIBE_구독 대상 이름', 'CPA SUBSCRIBE_이미지 인식에 사용할 식별자', 'CPA SUBSCRIBE_광고주 계정 식별자1', 'CPA SUBSCRIBE_광고주 계정 식별자2', 'CPA SUBSCRIBE_광고주 계정 식별자3', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL AOS', 'CPA SUBSCRIBE_구독 페이지 랜딩 URL IOS'
     ];

     // 1. 정해진 순서대로 출력 (데이터가 있는 경우에만)
     orderedKeys.forEach(key => {
       if (modificationOptions.hasOwnProperty(key)) {
         let val = modificationOptions[key];
         if (Array.isArray(val)) val = val.join(', '); // 배열인 경우 문자열 변환
         let displayVal = String(val).replace(/</g, '&lt;').replace(/>/g, '&gt;');
         body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${key}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${displayVal.replace(/\n/g, '<br>')}</td></tr>`;
         
         delete modificationOptions[key]; // 출력한 키는 삭제하여 중복 방지
       }
     });

     // 2. 순서 목록에 없지만 데이터에 남아있는 항목들 출력 (예외 처리)
     for (const [key, val] of Object.entries(modificationOptions)) {
       let displayVal = String(val).replace(/</g, '&lt;').replace(/>/g, '&gt;');
       body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${key}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${displayVal.replace(/\n/g, '<br>')}</td></tr>`;
     }
  }

  body += `</table>`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails });

  try {
    const slackMessage = { 'text': `${subject}` };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) });
  } catch (e) {
    console.error(`복사 생성 슬랙 발송 실패: ${e.toString()}`);
  }
  
  Utilities.sleep(2000);
  const threads = GmailApp.search(`subject:"${subject}" in:sent`, 0, 1);
  if (threads && threads.length > 0) return threads[0].getId();
  return null;
}

function findCopyCreationRowById(id) {
  const sheet = ss.getSheetByName("복사 생성 요청");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true);
  const foundCell = textFinder.findNext();
  if (foundCell) {
    return { sheet, rowIndex: foundCell.getRow(), rowData: sheet.getRange(foundCell.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0], headers: sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] };
  }
  return null;
}

function recordCopyCreationConfirmation(id, approverEmail) {
  const found = findCopyCreationRowById(id);
  if (!found) return `요청 ID: ${id} 건을 찾을 수 없습니다.`;
  const { sheet, rowIndex, headers, rowData } = found;
  const managerIndex = headers.indexOf('manager');
  const statusIndex = headers.indexOf('status');
  const timeIndex = headers.indexOf('manager_timestamp');

  if (rowData[managerIndex]) return `이미 ${rowData[managerIndex]}님이 담당 중입니다.`;

  sheet.getRange(rowIndex, managerIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusIndex + 1).setValue('처리중');
  sheet.getRange(rowIndex, timeIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));

  // 메일 답장
  const threadId = rowData[headers.indexOf('mail_thread_id')];
  if (threadId) {
    try {
      GmailApp.getThreadById(threadId).replyAll("", { htmlBody: `<p><b>${approverEmail}</b> 님이 <b>ID: ${id}</b> 건의 담당자로 지정되었습니다.</p><p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>` });
    } catch(e) { console.error(e); }
  }
  return `ID: ${id} 담당자로 지정되었습니다.`;
}

function processCopyCreationCompletion(id, completerEmail) {
  const found = findCopyCreationRowById(id);
  if (!found) return { success: false, message: `ID(${id})를 찾을 수 없습니다.` };
  const { sheet, rowIndex, headers } = found;
  const statusIndex = headers.indexOf('status');
  const timeIndex = headers.indexOf('completion_timestamp');

  sheet.getRange(rowIndex, statusIndex + 1).setValue('완료');
  sheet.getRange(rowIndex, timeIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  logUserAction(completerEmail, '복사 생성 완료', { targetId: id });
  return { success: true, message: `ID: ${id} 완료 처리되었습니다.` };
}

function getCopyCreationDataById(id) {
  const found = findCopyCreationRowById(id);
  if (found) {
    const data = {};
    found.headers.forEach((header, index) => {
      let value = found.rowData[index];
      if (value instanceof Date) {
        try {
          value = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm");
        } catch(e) {
          value = '날짜 형식 오류';
        }
      }
      data[header] = value;
    });
    return data;
  }
  return null;
}

/**
 * 복사 생성 요청을 스킵 처리합니다.
 */
function processCopyCreationSkip(id) {
  try {
    const skipperEmail = Session.getActiveUser().getEmail();
    const found = findCopyCreationRowById(id);
    if (!found) return { success: false, message: `ID(${id})를 찾을 수 없습니다.` };

    const { sheet, rowIndex, headers, rowData } = found;
    const statusIndex = headers.indexOf('status');
    const threadId = rowData[headers.indexOf('mail_thread_id')];
    const targetAdName = rowData[headers.indexOf('target_ad_name_to_modify')] || id;

    sheet.getRange(rowIndex, statusIndex + 1).setValue('스킵처리');

    if (threadId) {
      try {
        GmailApp.getThreadById(threadId).replyAll("", {
          htmlBody: `<p>안녕하세요,</p><p>요청하신 <b>복사 생성 ID: ${id}</b> 건이 <b>스킵 처리</b>되었음을 알려드립니다.</p><p>감사합니다.</p><p>- 처리자: ${skipperEmail}</p>`
        });
      } catch (e) { console.error(e); }
    }

    try {
      const slackMessage = { 'text': `[복사 생성 스킵] ${targetAdName} (ID: ${id})` };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) });
    } catch (e) { console.error(e); }

    logUserAction(skipperEmail, '복사 생성 스킵', { targetId: id });
    return { success: true, message: `ID(${id})가 성공적으로 스킵 처리되었습니다.` };
  } catch (e) {
    return { success: false, message: `오류 발생: ${e.toString()}` };
  }
}

/**
 * 복사 생성 요청을 반려 처리합니다.
 */
function processCopyCreationRejection(id, reason) {
  try {
    const rejectorEmail = Session.getActiveUser().getEmail();
    const found = findCopyCreationRowById(id);
    if (!found) return { success: false, message: `ID(${id})를 찾을 수 없습니다.` };

    const { sheet, rowIndex, headers, rowData } = found;
    const statusIndex = headers.indexOf('status');
    const registrantEmail = rowData[headers.indexOf('registrant')];
    const threadId = rowData[headers.indexOf('mail_thread_id')];
    const targetAdName = rowData[headers.indexOf('target_ad_name_to_modify')] || id;

    sheet.getRange(rowIndex, statusIndex + 1).setValue('반려');
    
    // 반려 알림 메일 (ID 입력 없는 단순 완료 처리와 대칭되는 개념)
    if (registrantEmail) {
      const subject = `[광고 등록 시스템] 요청하신 복사 생성(ID: ${id})이 반려되었습니다.`;
      let body = `<p>안녕하세요, ${registrantEmail.split('@')[0]}님.</p><p>요청하신 <b>ID: ${id}</b> 건이 반려되었습니다.</p>`;
      if (reason) body += `<p><b>반려 사유:</b> ${reason.replace(/\n/g, '<br>')}</p>`;
      body += `<p>수정 후 재요청하시거나 담당자(${rejectorEmail})에게 문의해주세요.</p><p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>`;

      if (threadId) {
        try {
          GmailApp.getThreadById(threadId).replyAll("", { htmlBody: body, cc: registrantEmail });
        } catch (e) {
          GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: body });
        }
      } else {
        GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: body });
      }
    }

    try {
      const slackMessage = { 'text': `[복사 생성 반려] ${targetAdName} (ID: ${id})` };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) });
    } catch (e) { console.error(e); }

    logUserAction(rejectorEmail, '복사 생성 반려', { targetId: id, message: `사유: ${reason}` });
    return { success: true, message: `ID(${id})가 성공적으로 반려 처리되었습니다.` };
  } catch (e) {
    return { success: false, message: `오류 발생: ${e.toString()}` };
  }
}

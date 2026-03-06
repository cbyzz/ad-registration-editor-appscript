function checkFolderAndProcessPdf() {
  // 모니터링할 구글 드라이브 폴더 ID (URL의 folders/ 뒤에 있는 값)
  const POLICY_FOLDER_ID = '1jVy591ILRYoZqS2uXW74a_t_RF227-Dz';
  const sheetName = '제한업종 DB';
  
  try {
    const folder = DriveApp.getFolderById(POLICY_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.PDF); // PDF 파일만 가져오기
    
    let latestFile = null;
    let maxUpdatedTime = 0;

    // 폴더 내 파일 중 가장 최근에 업데이트된 파일 찾기
    while (files.hasNext()) {
      const file = files.next();
      const updatedTime = file.getLastUpdated().getTime();
      
      if (updatedTime > maxUpdatedTime) {
        maxUpdatedTime = updatedTime;
        latestFile = file;
      }
    }

    if (!latestFile) {
      Logger.log("폴더에 PDF 파일이 없습니다.");
      return;
    }

    // 이전에 마지막으로 처리했던 시간 가져오기
    const props = PropertiesService.getScriptProperties();
    const lastProcessedTime = parseInt(props.getProperty('LAST_PROCESSED_TIME') || '0', 10);

    // 저장된 시간보다 더 최근에 수정된 파일이 있을 경우에만 실행
    if (maxUpdatedTime > lastProcessedTime) {
      Logger.log(`새로운/수정된 파일 감지됨: ${latestFile.getName()}`);
      
      const base64Data = Utilities.base64Encode(latestFile.getBlob().getBytes());
      const mimeType = latestFile.getMimeType();
      
      // Gemini API 호출하여 JSON 데이터 추출
      const jsonData = extractKeywordsFromPdfBase64(base64Data, mimeType);
      
      if (jsonData && jsonData.length > 0) {
        saveKeywordsToSheet(sheetName, jsonData);
        
        // 성공적으로 시트에 저장 완료 후, 마지막 처리 시간 업데이트
        props.setProperty('LAST_PROCESSED_TIME', maxUpdatedTime.toString());
        Logger.log(`DB 업데이트 완료. 총 ${jsonData.length}개 키워드 추출됨.`);
      } else {
        Logger.log("파싱된 데이터가 없어 업데이트를 건너뜁니다.");
      }
    } else {
      Logger.log("새로 업데이트된 내용이 없습니다. (Gemini를 호출하지 않음)");
    }
  } catch (e) {
    Logger.log("폴더 확인 및 PDF 처리 중 에러 발생: " + e.toString());
  }
}

// 3. Gemini API 호출 함수
function extractKeywordsFromPdfBase64(base64Data, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`;
  
  const prompt = `
다음 첨부된 PDF는 광고 심사 정책에 관한 문서입니다.
이 문서에서 '불가 업종' 또는 제한되는 업종의 키워드들을 추출해주세요.

[조건]
1. 각 플랫폼/서비스 명(예: 번개장터, 알바몬, 직방 등)을 개별 키워드로 분리하세요.
2. 대분류 카테고리(예: 중고거래, 구인구직 플랫폼 등)를 category 필드에 넣으세요.
3. 예외 허용 조건(예: '단, 일반 사업주가 구인 목적으로...')이 있다면 해당 키워드의 exception 필드에 적어주세요. 명시된 예외가 없으면 빈 문자열("")로 두세요.
4. [중요] 추출한 키워드가 붙여 쓴 명사라면, 띄어쓰기가 포함된 변형 키워드도 반드시 별도의 항목으로 추가하세요. (예: '번개장터' 추출 시 '번개 장터'도 별도로 추가, '중고거래' 추출 시 '중고 거래'도 추가)
5. [중요] 문서 내에 따옴표(' ' 또는 " ")로 둘러싸인 항목(예: '지금여기', '우리동네')이 있다면, 따옴표 안의 내용을 분리하지 말고 하나의 단일 키워드로 추출하세요. (추출 시 겉의 따옴표 기호는 제거하세요)
6. 반드시 아래와 같은 JSON 배열 형식으로만 응답하세요. 마크다운(\`\`\`json 등)은 절대 포함하지 마세요.

[
  {"category": "중고거래", "keyword": "번개장터", "exception": ""},
  {"category": "중고거래", "keyword": "번개 장터", "exception": ""},
  {"category": "커뮤니티 & 모임 서비스", "keyword": "지금여기", "exception": ""},
  {"category": "커뮤니티 & 모임 서비스", "keyword": "지금 여기", "exception": ""},
  {"category": "구인구직 플랫폼", "keyword": "알바몬", "exception": "단, 일반 사업주가 구인 목적으로 위 서비스 랜딩페이지를 광고로 활용 가능"}
]
  `;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64Data } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json" // 응답을 JSON 형식으로 강제
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode === 200) {
    const json = JSON.parse(response.getContentText());
    const resultText = json.candidates[0].content.parts[0].text;
    return JSON.parse(resultText);
  } else {
    Logger.log("API 오류: " + response.getContentText());
    return null;
  }
}

// 4. 시트에 데이터 저장 함수
function saveKeywordsToSheet(sheetName, data) {
  let sheet = ss.getSheetByName(sheetName);
  
  // 시트가 없으면 새로 생성
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['대분류', '제한 키워드', '예외 조건', '업데이트 일시']);
    sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  // 기존 데이터 초기화 (1행 헤더 제외)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  }
  
  const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  // 배열 데이터를 2차원 배열로 변환
  const rows = data.map(item => [
    item.category, 
    item.keyword, 
    item.exception || '', 
    timestamp
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}
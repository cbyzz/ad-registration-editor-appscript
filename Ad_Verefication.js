function addToVerificationQueue(type, headers, rowData, adId) {
  try {
    const QUEUE_SHEET_NAME = '검증 대기열';
    let queueSheet = ss.getSheetByName(QUEUE_SHEET_NAME);
    
    const queueHeaders = [
      '검증상태', '검증유형', '대기열등록일시', 
      '등록ID', '광고ID(어드민)', '대상광고ID', '대상광고명', '등록자',
      '원본데이터(JSON)'
    ];
    
    if (!queueSheet) {
      queueSheet = ss.insertSheet(QUEUE_SHEET_NAME);
      queueSheet.appendRow(queueHeaders);
      queueSheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      queueSheet.setFrozenRows(1);
    }
    
    // 원본 데이터를 JSON으로 변환 (빈 값 제외)
    const originalData = {};
    headers.forEach((header, index) => {
      const value = rowData[index];
      if (value !== '' && value !== null && value !== undefined) {
        // Date 객체 처리
        if (value instanceof Date) {
          originalData[header] = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
        } else {
          originalData[header] = value;
        }
      }
    });
    
    // 메타 필드 추출
    const registrationId = originalData['등록ID'] || '';
    const targetAdId = originalData['대상 광고 ID'] || '';
    const targetAdName = originalData['대상 광고명'] || '';
    const registrant = originalData['등록자'] || '';
    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    
    const newRow = [
      '대기', type, timestamp,
      registrationId, adId || '', targetAdId, targetAdName, registrant,
      JSON.stringify(originalData)
    ];
    
    queueSheet.appendRow(newRow);
    
  } catch (e) {
    console.error(`검증 대기열 등록 실패: ${e.toString()}`);
  }
}
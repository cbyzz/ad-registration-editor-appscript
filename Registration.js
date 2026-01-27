function submitData(formData, subType) {
  Logger.log("[submitData] 광고 제출 시작 - 타입: " + subType);
  Logger.log("[submitData] 사용자 이메일: " + Session.getActiveUser().getEmail());
  Logger.log("[submitData] 폼 데이터 키: " + Object.keys(formData).join(", "));
  // 동시 요청으로 인한 ID 중복 생성을 막기 위해 LockService 추가
  const lock = LockService.getUserLock();
  // 다른 요청이 lock을 해제할 때까지 최대 30초 대기
  lock.waitLock(30000); 
  
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheetName = `${userName} - 광고`;
    let sheet = ss.getSheetByName(sheetName);

    formData['광고 타입'] = subType;

    if (subType === 'CPA TRACKER') {
      if (formData['완료 이벤트 조건(JSON)']) {
        formData['완료 이벤트 조건(event_parameters JSON 파라미터)'] = formData['완료 이벤트 조건(JSON)'];
      }
      if (formData['완료 이벤트 조건(개별)']) {
        formData['완료 이벤트 조건 (개별 파라미터)'] = formData['완료 이벤트 조건(개별)'];
      }
    }
    
    // --- 원본 값 임시 저장을 위한 변수 선언 ---
    let tempOriginalActionAmount = null;
    let tempOriginalCategory = null;
    let tempOriginalRecSetting = null;

    if (subType === '테스트 광고' && (formData['선택항목'] === '서버 연동' || formData['선택항목'] === 'CPA TRACKER')) {
      formData['재참여 타입'] = '무한참여';
      formData['OS'] = 'OS 전체';
      formData['총물량'] = '5건';
      formData['일물량'] = '5건';
      formData['단가'] = '100원';

      const startDate = new Date();
      const endDate = new Date();
      endDate.setDate(startDate.getDate() + 6);

      formData['광고 집행 시작'] = Utilities.formatDate(startDate, "Asia/Seoul", "yyyy-MM-dd 00:00");
      formData['광고 집행 종료'] = Utilities.formatDate(endDate, "Asia/Seoul", "yyyy-MM-dd 23:59");
      
      const medium = formData['집행매체1'];
      if (medium === '네이버페이') formData['리워드'] = '60원';
      else if (medium === '애디슨 네트워크') formData['리워드'] = '30원';
      else if (medium === '쿠키오븐') formData['리워드'] = '1개';

      if (formData['선택항목'] === '서버 연동') {
      const adNetwork = formData['광고 네트워크 연동 매체'];
      if (adNetwork) {
        // const adNetworkIdMap = {
        //   'TNK': 'wtwkm5nx2MPrqy6ckJcYXsYe', 'AD POPCORN': '3aprX9MkiN5GDPgZfNGioro2',
        //   'PINCRUX': 'AyqRBwPiak2FvkFwQwVE9HmN', 'SUCOMM': 'EETKG8XuqPqyUPcu2RwxDoRx',
        //   'OHC': 'd8UZjbYeYp19NW4tgzV5XPGk', 'IVE': 'nVJtAyjTguM2wm8R1WBnm8Ss',
        //   'BUZZVIL': 'ja2WCwj4gbR3xukwWJPtqvzq', 'ADISON DSP': 'TB47DMobi5VMNP5Uke15Qd8t',
        //   '나스미디어NAP': 'JkgjfkoSmcM2nJiVTEhaUHU8', 'CAULY': 'AUSgJwbr28843L2zfpbPk23E'
        // };
        // formData['광고주 연동 토큰 값'] = adNetworkIdMap[adNetwork] || '';
        }
      } else if (formData['선택항목'] === 'CPA TRACKER') {
        formData['테스트 광고 타입'] = 'CPA TRACKER'; // 숨겨진 기본값 추가
        const jsonSubFields = [
            '완료 이벤트 조건(JSON)-이벤트 타입', '완료 이벤트 조건(JSON)-이벤트 이름',
            '완료 이벤트 조건(JSON)-value', '완료 이벤트 조건(JSON)-from', '완료 이벤트 조건(JSON)-to'
        ];
        const individualSubFields = [
            '완료 이벤트 조건(개별)-파라미터 타입', '완료 이벤트 조건(개별)-파라미터 이름',
            '완료 이벤트 조건(개별)-value', '완료 이벤트 조건(개별)-from', '완료 이벤트 조건(개별)-to'
        ];

        const hasJsonSubFieldData = jsonSubFields.some(field => formData[field] && String(formData[field]).trim() !== '');
        if (hasJsonSubFieldData) {
            formData['완료 이벤트 조건(event_parameters JSON 파라미터)'] = 'TRUE';
        }

        const hasIndividualSubFieldData = individualSubFields.some(field => formData[field] && String(formData[field]).trim() !== '');
        if (hasIndividualSubFieldData) {
            formData['완료 이벤트 조건 (개별 파라미터)'] = 'TRUE';
        }
      }
    } 
    else if (subType === '네이버페이 알림받기') {
      formData['재참여 타입'] = '단일 참여';
      formData['리포트 타입'] = '일반';
      formData['광고 정산 타입'] = '기타';

      if (formData['태깅'] === '스토어알림') {
        formData['광고주 연동 토큰 값'] = 'EYccpaU5WYY3v66zMnWHqAYu';
        formData['(메타) NF 광고주 연동 타입'] = '스토어 알림받기 / 게임 스토어 알림 (store_alarm)';
      } else if (formData['태깅'] === '라방알림') {
        formData['광고주 연동 토큰 값'] = 'Amk947c373B4phwBTg4MSYbS';
        formData['(메타) NF 광고주 연동 타입'] = '쇼핑라이브 알림받기 (live_alarm)';
      }
      
      const originalUrl = formData['URL'];
      if (originalUrl && !originalUrl.includes('click_key')) {
          formData['URL'] = `${originalUrl}?click_key={click_key}&ad_start_date={ad_start_at}&campaign_id={campaign_id}`;
      }

      formData['이벤트 상세안내 이미지 경로'] = 'G:\\공유 드라이브\\000. 소재저장소\\002. 애디슨오퍼월\\000. 디폴트_상세설명\\006. 네이버페이_스토어알림받기_상세설명이미지_디폴트\\스토어알림받기&쇼핑라이브알림받기(2개 상품 공용)';

    } else if (subType === '쿠키오븐 스마트스토어 CPS') {
      tempOriginalActionAmount = formData['액션명'];
      formData['액션명'] = `${tempOriginalActionAmount}만원 이상 구매`;
      formData['태깅'] = '구매형';
      formData['재참여 타입'] = '일일한번참여';
      formData['리포트 타입'] = 'CPS-RewardFail'; 
      formData['정산 타입 Block 처리 설정(부정 결제 취소 유저 대상)'] = '실패 내역 검색 기간 : 30일\nBlock 처리 기준 : 2회\nBlock 처리 대상 정산타입 : CPS\nBlock 기간 : 30일';
      formData['광고 정산 타입'] = 'CPS';
      formData['전체 목록 노출 여부'] = '미노출';
      const adDetailType = formData['광고 상세 타입'];
      switch (adDetailType) {
        case '도착보장 상품 구매': formData['광고주 연동 토큰 값'] = 'oNXSPE24dLSZqM13uZYXLnHg'; formData['파트너 광고주 타입'] = '도착보장 상품 구매 (ARRIVAL_GUARANTEE)'; break;
        case '특정 판매자 상품 구매': formData['광고주 연동 토큰 값'] = 'jLP7VcGm9AdXZPS3NhcT8J63'; formData['파트너 광고주 타입'] = '스스 / 브스 특정 판매자 상품 구매 (NORMAL_SELLER)'; break;
        case '도착보장 내 특정 판매자 상품 구매': formData['광고주 연동 토큰 값'] = 'tNwtuQn1wuiMMBYZYyqrRpQc'; formData['파트너 광고주 타입'] = '도착보장 내 특정 판매자 상품 구매 (ARRIVAL_GUARANTEE_SELLER)'; break;
        case '특정 판매자 특정 상품 구매': formData['광고주 연동 토큰 값'] = 'xvsSa1XQeLA34bojmhAAEYU2'; formData['파트너 광고주 타입'] = '스스 / 브스 특정 판매자 특정 상품 구매 (NORMAL_SELLER_PRODUCT)'; break;
      }

    } else if (subType === 'CPA SUBSCRIBE') {
      const subscriptionTarget = formData['구독 대상 이름'];
      let defaultGuideMessage = '';
      let defaultButtonMessage = '';
      switch (subscriptionTarget) {
        case '팔로우': defaultGuideMessage = '인스타그램 계정을 <팔로우> 해주세요'; break;
        case '좋아요': defaultGuideMessage = '페이스북 계정을 <좋아요> 해주세요'; break;
        case '채널추가': defaultGuideMessage = '카카오톡 채널 계정을 추가해주세요'; break;
        case '유튜브 구독(채널메인)': case '유튜브 구독(특정영상)': defaultGuideMessage = '유튜브 계정을 <구독> 해주세요'; break;
        case '언론사 구독': defaultGuideMessage = '네이버뉴스를 <구독> 해주세요'; break;
        case '라이브방송 참여하기': defaultGuideMessage = '쇼핑라이브 방송을 시청하고 인증 해주세요'; break;
        case '쇼핑라이브 알림받기': defaultGuideMessage = '쇼핑라이브 <알림받기>를 해주세요'; break;
        case '유튜브_좋아요': defaultGuideMessage = '유튜브 영상을 보고 <좋아요> 및 <구독> 해주세요'; break;
        case '네이버_플레이스_저장': defaultGuideMessage = '네이버 플레이스를 <저장> 해주세요. 만약 저장 버튼이 보이지 않으시면 검색 결과에 나온 플레이스 명을 클릭해주세요.'; break;
        case '틱톡': defaultGuideMessage = '틱톡 계정을 <팔로우> 해주세요'; break;
        case 'X(트위터)': defaultGuideMessage = 'X(트위터) 계정을 <팔로우> 해주세요'; break;
      }
      switch (subscriptionTarget) {
        case '팔로우': defaultButtonMessage = '팔로우 하러가기'; break;
        case '좋아요': defaultButtonMessage = '좋아요 하러가기'; break;
        case '채널추가': defaultButtonMessage = '추가 하러가기'; break;
        case '유튜브 구독(채널메인)': case '유튜브 구독(특정영상)': defaultButtonMessage = '구독 하러가기'; break;
        case '언론사 구독': defaultButtonMessage = '구독 하러가기'; break;
        case '라이브방송 참여하기': defaultButtonMessage = '라이브 방송 보러가기'; break;
        case '쇼핑라이브 알림받기': defaultButtonMessage = '알림 받으러가기'; break;
        case '유튜브_좋아요': defaultButtonMessage = '컨텐츠 좋아요 + 구독 하러가기'; break;
        case '네이버_플레이스_저장': defaultButtonMessage = '저장하러 가기'; break;
        case '틱톡': defaultButtonMessage = '팔로우 하러가기'; break;
        case 'X(트위터)': defaultButtonMessage = '팔로우 하러가기'; break;
      }
      if (defaultGuideMessage) { formData['가이드 메세지'] = defaultGuideMessage; }
      if (defaultButtonMessage) { formData['버튼 메세지'] = defaultButtonMessage; }

      const autoGeneratedTargets = ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우', '좋아요', '채널추가', '유튜브_좋아요', '언론사 구독', '틱톡', 'X(트위터)'];
      if (autoGeneratedTargets.includes(subscriptionTarget)) {
        const id1 = formData['광고주 계정 식별자1'];
        const id2 = formData['광고주 계정 식별자2'];
        const id3 = formData['광고주 계정 식별자3'];
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
          if (conditionPart) { formData['이미지 인식에 사용할 식별자'] = `${identifierPart} && ${conditionPart}`; }
        }
      }

      if (formData['구독 페이지 랜딩 URL']) {
        const instagramId = formData['구독 페이지 랜딩 URL'];
        formData['AOS 랜딩 URL'] = `intent://instagram.com/_u/${instagramId}#Intent;package=com.instagram.android;scheme=https;end`;
        formData['IOS 랜딩 URL'] = `instagram://user?username=${instagramId}`;
      }

    } else if (subType === 'CPA SUBSCRIBE 후지급') {
      const subscriptionTarget = formData['구독 대상 이름'];
      let defaultButtonMessage = '';

      switch (subscriptionTarget) {
        case '팔로우': defaultButtonMessage = '팔로우 하고 24시간 뒤 적립 받기'; break;
        case '좋아요': defaultButtonMessage = '좋아요 하고 24시간 뒤 적립 받기'; break;
        case '채널추가': defaultButtonMessage = '추가 하고 24시간 뒤 적립 받기'; break;
        case '유튜브 구독(채널메인)': case '유튜브 구독(특정영상)': defaultButtonMessage = '구독 하고 24시간 뒤 적립 받기'; break;
        case '언론사 구독': defaultButtonMessage = '구독 하고 24시간 뒤 적립 받기'; break;
        case '라이브방송 참여하기': defaultButtonMessage = '라이브 방송 참여하고 24시간 뒤 적립 받기'; break;
        case '쇼핑라이브 알림받기': defaultButtonMessage = '알림 받기 하고 24시간 뒤 적립 받기'; break;
        case '유튜브_좋아요': defaultButtonMessage = '컨텐츠 좋아요 + 구독 하고 24시간 뒤 적립 받기'; break;
        case '네이버_플레이스_저장': defaultButtonMessage = '저장하고 24시간 뒤 적립 받기'; break;
        case '틱톡': defaultButtonMessage = '팔로우 하고 24시간 뒤 적립 받기'; break;
        case 'X(트위터)': defaultButtonMessage = '팔로우 하고 24시간 뒤 적립 받기'; break;
      }
      
      const guideMessageMap = {
        '팔로우': '인스타그램 계정을 <팔로우> 해주세요', '좋아요': '페이스북 계정을 <좋아요> 해주세요',
        '채널추가': '카카오톡 채널 계정을 추가해주세요', '유튜브 구독(채널메인)': '유튜브 계정을 <구독> 해주세요',
        '유튜브 구독(특정영상)': '유튜브 계정을 <구독> 해주세요', '언론사 구독': '네이버뉴스를 <구독> 해주세요',
        '라이브방송 참여하기': '쇼핑라이브 방송을 시청하고 인증 해주세요', '쇼핑라이브 알림받기': '쇼핑라이브 <알림받기>를 해주세요',
        '유튜브_좋아요': '유튜브 영상을 보고 <좋아요> 및 <구독> 해주세요', '네이버_플레이스_저장': '네이버 플레이스를 <저장> 해주세요. 만약 저장 버튼이 보이지 않으시면 검색 결과에 나온 플레이스 명을 클릭해주세요.',
        '틱톡': '틱톡 계정을 <팔로우> 해주세요', 'X(트위터)': 'X(트위터) 계정을 <팔로우> 해주세요'
      };
      
      if (guideMessageMap[subscriptionTarget]) { formData['가이드 메세지'] = guideMessageMap[subscriptionTarget]; }
      if (defaultButtonMessage) { formData['버튼 메세지'] = defaultButtonMessage; }


      const autoGeneratedTargets_post = ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우', '좋아요', '채널추가', '유튜브_좋아요', '언론사 구독', '틱톡', 'X(트위터)'];
      if (autoGeneratedTargets_post.includes(subscriptionTarget)) {
        const id1 = formData['광고주 계정 식별자1'];
        const id2 = formData['광고주 계정 식별자2'];
        const id3 = formData['광고주 계정 식별자3'];
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
          if (conditionPart) { formData['이미지 인식에 사용할 식별자'] = `${identifierPart} && ${conditionPart}`; }
        }
      }

      if (formData['구독 페이지 랜딩 URL']) {
        const instagramId = formData['구독 페이지 랜딩 URL'];
        formData['AOS 랜딩 URL'] = `intent://instagram.com/_u/${instagramId}#Intent;package=com.instagram.android;scheme=https;end`;
        formData['IOS 랜딩 URL'] = `instagram://user?username=${instagramId}`;
      }

    } else if (subType === '네이버페이 스마트스토어 CPS') {
        formData['재참여 타입'] = '일일한번참여';
        formData['리포트 타입'] = 'CPS';
        formData['정산 타입 Block 처리 설정(부정 결제 취소 유저 대상)'] = '실패 내역 검색 기간 : 30일\nBlock 처리 기준 : 2회\nBlock 처리 대상 정산타입 : CPS\nBlock 기간 : 30일';
        formData['집행매체1'] = '네이버페이';
        formData['광고 정산 타입'] = 'CPS';
        formData['전체 목록 노출 여부'] = '미노출';
        formData['탭'] = '선택안함';
        formData['문구 - 서브2'] = '* 본 이벤트는 선착순으로 조기 종료될 수 있어요\n* 이벤트 참여하기 클릭 후 24시간 내 "처음 구매금액 조건을 충족하는 결제"와 매칭돼요\n* 참여하기 클릭 후 24시간 이내에도 선착순 조기 종료 시에는 적립이 안돼요\n* 네이버 쇼핑 이벤트 참여 시 로그인한 네이버 ID로 구매 시에만 포인트가 지급돼요\n* 아래 주의사항을 꼭 읽어보세요!';
        formData['(목록) 리워드 조건 설명'] = '선착순';
        formData['placement 세팅 정보'] = '결제내역 CPS 띠배너(paymenthistory_card) : 우선순위 0';
        const cpsTagging = formData['태깅'];
        if (cpsTagging === '일반_브랜드펀딩') {
            formData['광고주 연동 토큰 값'] = 'QQSiBPT5B78MVt8z8SxgEqNw';
            formData['(메타) NF 광고주 연동 타입'] = '일반 브랜드펀딩(store_BRCPS)';
        } else if (cpsTagging === '일반_플랫폼펀딩') {
            formData['광고주 연동 토큰 값'] = 'i4A3Rmvxr5tmnx8rFb72Rq8R';
            formData['(메타) NF 광고주 연동 타입'] = '일반 플랫폼펀딩(store_NVCPS)';
        } else if (cpsTagging === '쇼핑앱_일반_플랫폼펀딩') {
            formData['광고주 연동 토큰 값'] = 'ryp6iE4vwTJBaSpgQPV2EAEz';
            formData['(메타) NF 광고주 연동 타입'] = '쇼핑앱_일반_플랫폼펀딩(store_shoppingapp_NVCPS)';
        }
        const boostingTag = formData['부스팅 CPC 광고 태깅'];
        formData['부스팅 CPC 광고 재참여 타입'] = '단일참여';
        formData['부스팅 CPC 광고 리포트 타입'] = '부스팅 CPC : CPC';
        formData['부스팅 CPC 광고 광고주 연동 토큰 값'] = '부스팅 CPC : 적용 X';
        formData['부스팅 CPC 광고 집행매체1'] = '네이버페이';
        formData['부스팅 CPC 광고 광고 정산 타입'] = '부스팅 CPC : CPC';
        formData['부스팅 CPC 광고 전체 목록 노출 여부'] = '노출';
        formData['부스팅 CPC 광고 단가'] = '부스팅 CPC : 0원';
        const rewardMap = { '부스팅_A': '1원', '부스팅_B': '5원', '부스팅_C': '15원' };
        formData['부스팅 CPC 광고 리워드'] = rewardMap[boostingTag] || '';
        formData['부스팅 CPC 광고 URL'] = 'https://ofw.adison.co/u/naverpay/ads/상위광고번호';
        formData['부스팅 CPC 광고 상세 랜딩 전용 URL'] = 'https://ofw.adison.co/u/naverpay/ads/상위광고번호';
        formData['부스팅 CPC 광고 탭'] = '쇼핑';
        formData['부스팅 CPC 광고 문구 - 타이틀'] = formData['문구 - 타이틀'];
        formData['부스팅 CPC 광고 문구 - 서브'] = formData['문구 - 서브'];
        formData['부스팅 CPC 광고 문구 - 서브1 상단'] = formData['문구 - 서브1 상단'];
        formData['부스팅 CPC 광고 문구 - 서브1 하단'] = formData['문구 - 서브1 하단'];
        formData['부스팅 CPC 광고 문구 - 서브2'] = formData['문구 - 서브2'];
        formData['부스팅 CPC 광고 액션 버튼'] = '참여하고 {REWARD_STR} 적립';
        formData['부스팅 CPC 광고 (목록) 리워드 조건 설명'] = formData['(목록) 리워드 조건 설명'];
        formData['부스팅 CPC 광고 (목록) 리워드 텍스트'] = formData['(목록) 리워드 텍스트'];
        formData['부스팅 CPC 광고 (메타) NF 광고주 연동 ID'] = formData['(메타) NF 광고주 연동 ID'];
        const priorityMap = { '부스팅_A': 1, '부스팅_B': 5, '부스팅_C': 15 };
        const priority = priorityMap[boostingTag] || 0;
        let placementInfo = [];
        placementInfo.push(`네이버쇼핑(nvshopping) : 우선순위 ${priority}`);
        placementInfo.push(`네이버마케팅(nvmarketing) : 우선순위 ${priority}`);
        placementInfo.push(`네이버마케팅_네앱(nvmarketing_nvapp) : 우선순위 ${priority}`);
        placementInfo.push(`쇼핑주문배송 구매 확정 띠배너(nvshopping_order_card) : 우선순위 0`);
        placementInfo.push(`쇼핑주문배송 하단 추천 영역(nvshopping_order_bottom) : 우선순위 0`);
        placementInfo.push(`(신)결제홈 결제내역 카드(historycard) : 우선순위 0`);
        formData['부스팅 CPC 광고 placement 세팅 정보 기본'] = placementInfo.join('\n');
        
        tempOriginalRecSetting = formData['부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부'];
        if (tempOriginalRecSetting === '세팅 O') {
          formData['부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부'] = `네이버마케팅_추천(nvmarketing_best) : 우선순위 ${priority}`;
        }
        
        tempOriginalCategory = formData['부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리'];
        const categoryMap = { '건강': 'nvmarketing_health', '식품': 'nvmarketing_food', '생활': 'nvmarketing_living', '뷰티': 'nvmarketing_beauty', '기타': 'nvmarketing_etc' };
        if (categoryMap[tempOriginalCategory]) {
          formData['부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리'] = `네이버마케팅_${tempOriginalCategory}(${categoryMap[tempOriginalCategory]}) : 우선순위 ${priority}`;
        }
    } else if (subType === '완전 정률 - 쿠키오븐 스마트스토어 CPS') {
      formData['액션명'] = '구매형'; // 액션명 기본값 설정
      formData['태깅'] = '구매형';
      formData['재참여 타입'] = '일일한번참여';
      formData['리포트 타입'] = 'CPS-RewardFail';
      formData['정산 타입 Block 처리 설정(부정 결제 취소 유저 대상)'] = '실패 내역 검색 기간 : 30일\nBlock 처리 기준 : 2회\nBlock 처리 대상 정산타입 : CPS\nBlock 기간 : 30일';
      formData['광고 정산 타입'] = 'CPS';
      formData['전체 목록 노출 여부'] = '미노출';

      // 광고 상세 타입에 따른 광고주 연동 토큰 값 및 파트너 타입 설정 (기존 CPS와 동일)
      const adDetailType = formData['광고 상세 타입'];
      switch (adDetailType) {
        case '도착보장 상품 구매': formData['광고주 연동 토큰 값'] = 'oNXSPE24dLSZqM13uZYXLnHg'; formData['파트너 광고주 타입'] = '도착보장 상품 구매 (ARRIVAL_GUARANTEE)'; break;
        case '특정 판매자 상품 구매': formData['광고주 연동 토큰 값'] = 'jLP7VcGm9AdXZPS3NhcT8J63'; formData['파트너 광고주 타입'] = '스스 / 브스 특정 판매자 상품 구매 (NORMAL_SELLER)'; break;
        case '도착보장 내 특정 판매자 상품 구매': formData['광고주 연동 토큰 값'] = 'tNwtuQn1wuiMMBYZYyqrRpQc'; formData['파트너 광고주 타입'] = '도착보장 내 특정 판매자 상품 구매 (ARRIVAL_GUARANTEE_SELLER)'; break;
        case '특정 판매자 특정 상품 구매': formData['광고주 연동 토큰 값'] = 'xvsSa1XQeLA34bojmhAAEYU2'; formData['파트너 광고주 타입'] = '스스 / 브스 특정 판매자 특정 상품 구매 (NORMAL_SELLER_PRODUCT)'; break;
      }

      // 문구 - 서브2 자동 계산 및 formData에 다시 저장 (클라이언트 계산 값을 서버에서도 확인/저장)
      const rate = parseFloat(formData['광고비 R/S율']) || 0;
      const minAmount = parseFloat(formData['최소 결제 금액']) || 0;
      const maxAmount = parseFloat(formData['최대 인정 결정 금액']) || 0;

      const A = (rate / 100) * 0.7 * 100;
      const B = minAmount / 10000;
      const C = maxAmount / 10000;
      const formattedA = A.toFixed(1);

      formData['문구 - 서브2'] = `[결제 금액 및 지급 쿠키 관련 안내 사항]
*결제 금액의 ${formattedA}%에 해당하는 금액이 쿠키로 지급됩니다.
*1만 원 이상 주문 건에 대해서만 쿠키가 지급됩니다.
*주문 금액 기준으로 ${C}만 원까지만 쿠키가 지급됩니다. (${C}만 원을 초과하여 주문해도 지급되는 쿠키는 ${C}만 원에 해당하는 금액 기준으로 적용됩니다.)
*단, 쿠키는 1개당 100원으로 계산되며, 10원 단위는 버림 처리 후 쿠키가 지급됩니다.

[참여 관련 안내 사항]
*배송료를 제외한 실 결제금액 기준으로 지급 쿠키가 계산됩니다.
*예약구매, 무통장입금, 후불결제 시에는 쿠키 지급이 되지 않습니다.
*네이버웹툰/시리즈에 로그인한 네이버ID로 구매시에만 쿠키가 지급 됩니다.
*참여하기 버튼 클릭 후 24시간 이내 구매가 완료 되어야 하며, 가장 마지막으로 클릭한 이벤트 기준으로 쿠키가 지급 됩니다.
*결제 취소 시(부분취소 및 판매자 이슈로 인한 취소 포함) 지급된 쿠키는 전량 회수됩니다.
*쿠키 사용 후 구매 취소 시 이벤트 참여가 제한될 수 있으며, 사용하신 쿠키 금액이 청구될 수 있습니다.
*지급된 쿠키는 유효기간이 존재합니다. 하단 ‘상세 안내’를 꼭 참고 부탁드립니다.`;

    }

    const mediumConsolidationTypes = ['CPA S2S', 'CPI', 'CPA TRACKER', 'CPA SUBSCRIBE', 'CPA SUBSCRIBE 후지급', 'CPC', '애드네트워크 연동형', 'CPQ'];
    if (mediumConsolidationTypes.includes(subType)) {
      if (formData['집행매체1'] === '쿠키오븐') {
        formData['집행매체2'] = formData['집행매체2_쿠키오븐'];
      } else if (formData['집행매체1'] === '애디슨 네트워크') {
        if (formData['집행매체2_애디슨'] === '특정 매체만 진행') {
          formData['집행매체2'] = formData['집행매체2_애디슨_특정'];
        } else {
          formData['집행매체2'] = formData['집행매체2_애디슨'];
        }
      }
      delete formData['집행매체2_쿠키오븐'];
      delete formData['집행매체2_애디슨'];
      delete formData['집행매체2_애디슨_특정'];
    }
    
    const osConsolidationTypes = ['CPA S2S', 'CPA TRACKER', 'CPA SUBSCRIBE', 'CPC', '애드네트워크 연동형'];
    if (osConsolidationTypes.includes(subType)) {
      if (formData['OS_DROPDOWN']) {
        formData['OS'] = formData['OS_DROPDOWN'];
      } else if (formData['OS_CHECKBOX']) {
        formData['OS'] = formData['OS_CHECKBOX'];
      }
      delete formData['OS_DROPDOWN'];
      delete formData['OS_CHECKBOX'];
    }
    
    // 시트 생성 및 헤더 관리 로직
    const allFieldsData = getAllFormFields();
    // FormFields.gs에 정의된 모든 광고 타입의 headerOrder를 가져와 통합하고 중복을 제거합니다.
    const allFormHeaders = Object.values(allFieldsData['광고']).reduce((acc, currentType) => {
      if (currentType.headerOrder) {
        acc.push(...currentType.headerOrder);
      }
      return acc;
    }, []);

    // 고유한 헤더 목록을 만듭니다. (순서 유지를 위해 Set 사용 후 다시 배열로 변환)
    const uniqueFormHeaders = [...new Set(allFormHeaders)];
    
    // 기본적으로 필요한 시스템 컬럼들을 정의합니다.
    const baseHeaders = ['등록ID', '등록일시', '등록자', '상태', '담당자', '담당자 확인 일시', '메일 스레드 ID', '광고 타입', '반려 일시', '반려 사유', '광고 ID', '완료 일시', 'AOS 랜딩 URL', 'IOS 랜딩 URL', '물량 OS 선택'];
    
    // 기본 컬럼과 모든 폼 컬럼을 합쳐 최종 마스터 헤더 목록을 생성합니다.
    const masterHeaders = baseHeaders.concat(uniqueFormHeaders.filter(h => !baseHeaders.includes(h)));

    if (!sheet) {
      // 시트가 없으면 새로 생성하고 마스터 헤더를 첫 행에 추가합니다.
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(masterHeaders);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      // 시트가 이미 있는 경우, 누락된 컬럼이 있는지 확인하고 추가합니다.
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const missingHeaders = masterHeaders.filter(h => !currentHeaders.includes(h));

      if (missingHeaders.length > 0) {
        // 누락된 헤더가 있으면 시트의 마지막 열 다음에 추가합니다.
        sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      }
    }
    
    const idPrefix = `${userName}-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = `${idPrefix}${nextId}`;

    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let subject;
    if (subType === '테스트 광고') {
        subject = `[테스트 광고 등록 요청] ${formData['캠페인명']}_${formData['집행매체1']}`;
    } else if (subType === '네이버페이 CPS구매확정형_금액X') {
        // 1. 본광고 고정값 및 히든 필드 주입
        formData['태깅'] = '[일반_브랜드펀딩_구매확정형]';
        formData['리포트 타입'] = 'CPS본광고 : CPS-NpayDecided';
        formData['광고주 CS 담당자 이름'] = '알림받기와 동일하게 NBT -> 네이버페이로 전달 및 처리';
        formData['정산 타입 Block 처리 설정(부정 결제 취소 유저 대상)'] = '실패 내역 검색 기간 : 30일\nBlock 처리 기준 : 2회\nBlock 처리 대상 정산타입 : 네이버페이 CPS 구매확정(네이버페이 only) CPS\nBlock 기간 : 30일';
        formData['적립 허용 광고주 연동 선택'] = '5ZUcab3wRh5moR9vnaaLaL9Q';
        formData['집행매체1'] = '네이버페이';
        formData['광고 정산 타입'] = 'CPS본광고 : 네이버페이 CPS 구매확정(네이버페이 only)';
        formData['물량 (총/일)'] = '무제한';
        formData['최소 결제 금액'] = '1';
        
        formData['탭'] = '선택안함';
        formData['(목록) 리워드 조건 설명'] = '선착순';
        formData['(목록) 리워드 텍스트'] = '10%';
        formData['(메타) NF 광고주 연동 타입'] = 'store_BRCPS_decided';
        
        // 메타 ID 프리픽스 처리 (사용자 입력값 앞에 붙임)
        // const userMetaId = formData['(메타) NF 광고주 연동 ID'];
        // const fullMetaId = 'store_BRCPS_decided' + userMetaId;
        // formData['(메타) NF 광고주 연동 ID'] = fullMetaId; // 본광고에 저장

        formData['placement 세팅 정보'] = '결제내역 CPS 띠배너(paymenthistory_card) : 우선순위 0\n네이버페이 큐레이션 페이지 - 기본 (nf_curation_default_list) : 우선순위 0';

        // 2. 부스팅 광고 필드 자동 채우기 및 고정값 설정
        formData['부스팅 CPC 광고 태깅'] = '[CPC_일반_브랜드펀딩_구매확정형_부스팅]';
        formData['부스팅 CPC 광고 재참여 타입'] = '단일참여';
        formData['부스팅 CPC 광고 리포트 타입'] = '부스팅 CPC : CPC';
        // ▼▼▼ [수정] 요청하신 기본값 적용 ▼▼▼
        formData['부스팅 CPC 광고 적립 허용 광고주 연동 선택'] = '부스팅 CPC : 적용 X'; // 기존 값에서 변경
        formData['부스팅 CPC 광고 집행매체1'] = '네이버페이';
        formData['부스팅 CPC 광고 광고 정산 타입'] = '부스팅 CPC : CPC'; // 항목 추가 및 기본값 설정
        
        formData['부스팅 CPC 광고 전체 목록 노출 여부'] = '노출';
        formData['부스팅 CPC 광고 단가'] = '부스팅 CPC : 0원';
        formData['부스팅 CPC 광고 리워드'] = '1원';
        
        formData['부스팅 CPC 광고 URL'] = 'https://ofw.adison.co/u/naverpay/ads/상위광고번호';
        formData['부스팅 CPC 광고 상세 랜딩 전용 URL'] = 'https://ofw.adison.co/u/naverpay/ads/상위광고번호';
        formData['부스팅 CPC 광고 탭'] = '쇼핑'; // 기본값 "쇼핑"으로 변경
        formData['부스팅 CPC 광고 (목록) 리워드 조건 설명'] = '선착순';
        
        // 본광고 데이터 복사
        formData['부스팅 CPC 광고 문구 - 타이틀'] = formData['문구 - 타이틀'];
        formData['부스팅 CPC 광고 문구 - 서브'] = formData['문구 - 서브'];
        formData['부스팅 CPC 광고 문구 - 서브1 상단'] = formData['문구 - 서브1 상단'];
        formData['부스팅 CPC 광고 문구 - 서브1 하단'] = formData['문구 - 서브1 하단'];
        formData['부스팅 CPC 광고 문구 - 서브2'] = formData['문구 - 서브2'];
        
        formData['부스팅 CPC 광고 (목록) 리워드 텍스트'] = '10%';
        formData['부스팅 CPC 광고 (메타) NF 광고주 연동 ID'] = formData['(메타) NF 광고주 연동 ID'];

        // 3. placement 세팅 정보 구성
        let placementInfo = [];
        placementInfo.push(`네이버쇼핑(nvshopping) : 우선순위 1`);
        placementInfo.push(`네이버마케팅(nvmarketing) : 우선순위 1`);
        placementInfo.push(`네이버마케팅_네앱(nvmarketing_nvapp) : 우선순위 1`);
        placementInfo.push(`쇼핑주문배송 구매 확정 띠배너(nvshopping_order_card) : 우선순위 0`);
        placementInfo.push(`쇼핑주문배송 하단 추천 영역(nvshopping_order_bottom) : 우선순위 0`);
        placementInfo.push(`(신)결제홈 결제내역 카드(historycard) : 우선순위 0`);
        
        formData['부스팅 CPC 광고 placement 세팅 정보 기본'] = placementInfo.join('\n');

        if (formData['부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부'] === '세팅 O') {
            formData['부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부'] = `네이버마케팅_추천(nvmarketing_best) : 우선순위 1`;
        }
        
        // (2) 카테고리 설정
        const categoryOption = formData['부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리'];
        const categoryMap = {
            '건강': '네이버마케팅_건강(nvmarketing_health)',
            '식품': '네이버마케팅_식품(nvmarketing_food)',
            '생활': '네이버마케팅_생활(nvmarketing_living)',
            '뷰티': '네이버마케팅_뷰티(nvmarketing_beauty)',
            '기타': '네이버마케팅_기타(nvmarketing_etc)'
        };
        
        if (categoryMap[categoryOption]) {
            formData['부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리'] = `${categoryMap[categoryOption]} : 우선순위 1`;
        }
        // ▲▲▲ [수정] ▲▲▲

        formData['부스팅 CPC 광고 placement 세팅 정보 기본'] = placementInfo.join('\n');

        // 4. 이메일/슬랙 제목 생성
        const brand = formData['브랜드'];
        const cpsSubject = `[일반_브랜드펀딩_구매확정형] ${brand}_구매확정형_네이버페이`;
        const boostingSubject = `[CPC_일반_브랜드펀딩_구매확정형_부스팅] ${brand}_구매확정형_네이버페이`;
        
        subject = { cps: cpsSubject, boosting: boostingSubject };

    } else if (subType === '네이버페이 스마트스토어 CPS') {
        tempOriginalActionAmount = formData['액션명'];
        const actionNameDisplay = `${tempOriginalActionAmount}만원 이상 구매`;

        const cpsTagging = formData['태깅'];
        const boostingTagging = formData['부스팅 CPC 광고 태깅'];
        const brand = formData['브랜드'];
        
        const cpsSubject = `[${cpsTagging}] ${brand}_${actionNameDisplay}_네이버페이`;
        const boostingSubject = `[CPC_${cpsTagging}_${boostingTagging}] ${brand}_${actionNameDisplay}_네이버페이`;
        
        subject = { cps: cpsSubject, boosting: boostingSubject };

        formData['태깅'] = cpsTagging;
        formData['액션명'] = actionNameDisplay;
    } else if (subType === '네이버페이 입점비') {
    // 숨겨진 필드의 기본값을 formData에 주입합니다.
    formData['CPC 입점비 집행매체1'] = '네이버페이';
    formData['CPC 입점비 단가'] = '0원';
    formData['CPC 입점비 총물량'] = '무제한';
    formData['CPC 입점비 리워드'] = '1원';
    formData['CPC 입점비 일물량'] = '무제한';
    formData['CPC 입점비 상세 자동참여 여부'] = '체크';
    formData['CPC 입점비 노출 유지 여부'] = '체크';
    formData['CPC 입점비 (메타) 클릭 리워드 지급 금액'] = '1원';
    formData['CPC 입점비 (메타) 클릭 리워드 종료 일자'] = '2100-12-31 23:59';
    formData['CPC 입점비 placement'] = '혜택_그룹_신규입점 (benefit_group_new)';

    // CPC 입점비 광고 제목 생성
    const cpcTitle = `[CPC_입점비] ${formData['CPC 입점비 문구 - 타이틀']}_네이버페이`;

    if (formData['네이버페이 입점비 본 광고 신규 등록 여부'] === '등록 O') {
      // (기존 본 광고 제목 생성 로직은 그대로 유지)
      const mainAdType = formData['본 광고 타입'];
      const mainAdTagging = formData['태깅'];
      let mainAdParts = [ formData['브랜드'], formData['액션명'], formData['집행매체1'] ];
      if (formData['집행매체2']) mainAdParts.push(formData['집행매체2']);
      if (formData['OS']) {
        const os = formData['OS'];
        if (!((formData['집행매체1'] === '애디슨 네트워크' || formData['집행매체1'] === '쿠키오븐') && (os === '전체' || os === '모두'))){
           mainAdParts.push(os);
        }
      }
      const mainAdSubject = `[${mainAdTagging}] ` + mainAdParts.filter(Boolean).join('_');
      subject = { mainAd: mainAdSubject, feeAd: cpcTitle };
    } else {
      subject = cpcTitle;
    }

  }  else if (subType === '멀티미션') {
    // 숨겨진 필드의 기본값을 formData에 주입
    formData['재참여 타입'] = '단일참여';
    formData['영업 담당자'] = '엔비티 영업실';
    formData['리포트 타입'] = '일반';
    formData['타임존'] = 'Asia / Seoul';
    formData['캠페인 정산 타입'] = 'CPE';
    formData['우선순위'] = '90(수동) -> 국내 어드민 적용 필요';
    formData['화폐 단위'] = 'KRW';
    formData['노출 지역/국가'] = 'South Korea';
    formData['Target PubApp (Is Targeted)'] = 'TRUE';
    formData['캠페인 오너/매니저'] = 'NBT';
    formData['로컬'] = 'KOREAN';
    formData['광고 포맷'] = 'FEED';

    // 캠페인명(제목) 생성
    const category = formData['카테고리'] || '';
    const client = formData['거래처'] || '';
    const title = formData['타이틀 문구(목록)'] || '';
    const os = formData['OS'] === '전체' ? '' : formData['OS'];
    
    const subjectParts = [category, '[멀티미션]', client, title, os].filter(Boolean);
    subject = subjectParts.join('_');

  } else if (subType === '완전 정률 - 쿠키오븐 스마트스토어 CPS') {
      const advertiser = formData['브랜드'] || '브랜드없음';
      const os = formData['OS'] === '전체' ? '' : formData['OS'];
      const subjectParts = ['[구매형]', advertiser, '정률', '쿠키오븐', os].filter(Boolean); // '정률' 추가
      subject = subjectParts.join('_');

  } else if (subType === '테스트 광고') {
        subject = `[테스트 광고 등록 요청]`; // 간단한 제목으로 설정
  } else {
      const medium1 = formData['집행매체1'];
      const medium2 = formData['집행매체2'];
      let os = formData['OS'] || '';

      if (subType === '애드네트워크 연동형') {
          const adNetworkType = formData['애드네트워크 연동형 - 광고 타입'];
          const tagging = formData['태깅'] ? `[${formData['태깅']}]` : '';
          let osPart = '';
          if (adNetworkType === 'CPI') {
              osPart = 'AOS';
          } else {
              if ((medium1 === '애디슨 네트워크' || medium1 === '쿠키오븐') && (os === '모두' || os === '전체')) {
                  osPart = '';
              } else if (medium1 === '네이버페이') {
                  if (os === '모두' || os === '전체') {
                      osPart = '';
                  } else {
                      osPart = os.split(',').map(item => item.trim() === 'AOS+IOS' ? '모바일' : item.trim()).join(',');
                  }
              } else {
                  osPart = os;
              }
          }
          let medium2Part = '';
          if ((medium1 === '쿠키오븐' || medium1 === '애디슨 네트워크') && medium2 === '전체') {
              medium2Part = '';
          } else {
              medium2Part = medium2;
          }
          const underscoreParts = [ formData['브랜드'], formData['액션명'], medium1, medium2Part, osPart ].filter(Boolean);
          subject = tagging + ' ' + underscoreParts.join('_');
      } else if (subType === '쿠키오븐 스마트스토어 CPS') {
          let medium2Part = formData['집행매체2'] === '전체' ? '' : formData['집행매체2'];
          let osPart = formData['OS'] === '전체' ? '' : formData['OS'];
          const underscoreParts = [ formData['브랜드'], formData['액션명'], formData['집행매체1'], medium2Part, osPart ].filter(Boolean);
          subject = `[${formData['태깅']}] ` + underscoreParts.join('_');
      } else if (subType === '네이버페이 알림받기') {
        const underscoreParts = [ formData['브랜드'], formData['액션명'], '네이버페이' ].filter(Boolean);
        subject = `[${formData['태깅']}] ` + underscoreParts.join('_');
      } else if (subType === 'CPI') {
          let medium2Part = '';
          if (medium2 === '전체') {
            medium2Part = '';
          } else {
            medium2Part = medium2;
          }
          const underscoreParts = [ formData['브랜드'], medium1, medium2Part, 'AOS' ].filter(Boolean);
          subject = '[설치형] ' + underscoreParts.join('_');
      } else if (['CPA S2S', 'CPA TRACKER', 'CPA SUBSCRIBE', 'CPA SUBSCRIBE 후지급', 'CPC', 'CPQ'].includes(subType)) {
          let medium2Part = '';
          if (medium1 === '네이버페이' || ((medium1 === '쿠키오븐' || medium1 === '애디슨 네트워크') && medium2 === '전체')) {
              medium2Part = '';
          } else {
              medium2Part = medium2;
          }
          let osPart = '';
          if ((medium1 === '애디슨 네트워크' || medium1 === '쿠키오븐') && (os === '전체' || os === '모두')) {
              osPart = '';
          } else if (medium1 === '네이버페이') {
              if (os === '전체' || os === '모두') {
                  osPart = '';
              } else {
                  osPart = os.split(',').map(item => item.trim() === 'AOS+IOS' ? '모바일' : item.trim()).join(',');
              }
          } else {
              osPart = os;
          }
          const underscoreParts = [ formData['브랜드'], formData['액션명'], medium1, medium2Part, osPart ].filter(Boolean);
          subject = `[${formData['태깅']}] ` + underscoreParts.join('_');
      } else {
        const tagging = formData['태깅'] ? `[${formData['태깅']}]` : '';
        const underscoreParts = [ formData['브랜드'], formData['액션명'], formData['집행매체1'], formData['집행매체2'], formData['OS'] ].filter(Boolean);
        subject = tagging;
        if (underscoreParts.length > 0) {
          subject += (subject ? ' ' : '') + underscoreParts.join('_');
        }
      }
    }

    if (typeof subject === 'string') {
      subject = `${subject} (ID: ${uniqueId})`;
    }
    
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    const messageId = sendNotification(userEmail, '광고', uniqueId, formData, subType, subject);
    formData['메일 스레드 ID'] = messageId;
    
    if (tempOriginalActionAmount !== null) {
        if (subType === '네이버페이 스마트스토어 CPS' || subType === '쿠키오븐 스마트스토어 CPS') {
            formData['액션명'] = tempOriginalActionAmount;
        }
    }
    if (tempOriginalRecSetting !== null) {
        formData['부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부'] = tempOriginalRecSetting;
    }
    if (tempOriginalCategory !== null) {
      if (subType === '네이버페이 스마트스토어 CPS') {
        formData['부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리'] = tempOriginalCategory;
      }
    }

    const newRow = sheetHeaders.map(header => {
      if (header.endsWith('라이브 시작 시간') || header.endsWith('라이브 종료 시간')) {
      const timeValue = formData[header];
      // 값이 있을 경우, 앞에 ' 를 붙여서 일반 텍스트로 저장되도록 합니다.
      return timeValue ? `'${timeValue}` : '';
    }
      switch(header) {
        case '등록ID': return uniqueId;
        case '등록일시': return formattedTimestamp;
        case '등록자': return userEmail;
        case '상태': return '등록 요청 완료';
        case '라이브 시작 시간':
        case '라이브 종료 시간':
          const timeValue = formData[header];
          return timeValue ? `'${timeValue}` : '';
        default: return formData[header] || '';
      }
    });
    
    sheet.appendRow(newRow);
    
    const logSubject = (typeof subject === 'object') ? `[CPS] ${subject.cps} / [부스팅] ${subject.boosting}` : subject;
    logUserAction(userEmail, '광고 등록', {
      targetId: uniqueId,
      message: `광고 '${logSubject}' 등록 요청`
    });
    
    return { success: true, message: `광고 등록 요청이 완료되었습니다. (ID: ${uniqueId})` };

  } catch (e) {
    console.error(`submitData Error: ${e.toString()}`);
    return { success: false, message: `광고 등록 처리 중 오류가 발생했습니다: ${e.message}` };
  } finally { 
    // 다른 요청이 작업을 계속할 수 있도록 lock을 반드시 해제
    lock.releaseLock();
  }
}

function sendNotification(senderEmail, type, id, data, subType, subject) {
  let mainSubject, body;
  const ccEmails = data.ccRecipients || '';

  // 네이버페이 스마트스토어 CPS 또는 네이버페이 입점비와 같이 2개의 광고가 동시에 생성될 경우
  if (typeof subject === 'object' && (subject.cps || subject.mainAd)) {
      const mainAdTitle = subject.cps || subject.mainAd;
      const subAdTitle = subject.boosting || subject.feeAd;
      
      mainSubject = `[복합 요청] ${mainAdTitle} / ${subAdTitle} (ID: ${id})`;
      body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 새로운 ${type} 등록을 요청했습니다.</p>
              <p><b>ID: ${id}</b></p>
              <hr>
              <h3>[본 광고] ${mainAdTitle}</h3>
              <h3>[입점비 광고] ${subAdTitle}</h3>
              <hr>`;
  } else {
      mainSubject = subject; // 기존: 단일 광고 요청
  }

  mainSubject = `[광고 등록 요청] ${mainSubject}`;

  // body가 위에서 정의되지 않은 경우(단일 요청) 초기화
  if (!body) {
    body = `<p>안녕하세요, 운영팀.</p>
            <p><b>${senderEmail}</b>님께서 새로운 ${type} 등록을 요청했습니다.</p>
            <p><b>ID: ${id}</b></p>`;
  }

  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm&id=${id}`;
  body += `<div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
             <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 광고 담당하기 ]</a>
             <br><br>
             <div style="display: inline-block; vertical-align: middle;">
               <form action="${ScriptApp.getService().getUrl()}" method="get" target="_blank" style="margin:0; padding:0;">
                 <input type="hidden" name="action" value="complete">
                 <input type="hidden" name="id" value="${id}">
                 <input type="text" name="adId" placeholder="광고 ID 입력" required style="padding: 8px; border: 1px solid #ccc; border-radius: 4px; margin-right: 5px;">
                 <button type="submit" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; border: none; cursor: pointer;">[ 광고 등록 완료 처리 ]</button>
               </form>
             </div>
             <br><br>
             <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
             <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">광고 등록 시스템 바로가기</a>
           </div>`;


  body += `<hr><h3>등록 내용</h3>`;
  body += `<table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;
  

   const fieldsToEscape = ['특이사항 메모장', '문구 - 타이틀', '문구 - 서브', '문구 - 상세화면 상단 타이틀', '문구 - 서브1 상단', '문구 - 서브1 하단', '문구 - 서브2', '요청사항', '가이드 메세지', '버튼 메세지'];

  const addRow = (header, value) => {
    if (value === undefined || value === null || value === '') return '';
    
    let displayValue = String(value);

    if (subType === '네이버페이 CPS구매확정형_금액X') {
        if (header === '(메타) NF 광고주 연동 ID' || header === '부스팅 CPC 광고 (메타) NF 광고주 연동 ID') {
            displayValue = 'store_BRCPS_decided_' + displayValue;
        }
    }

    if (subType === '완전 정률 - 쿠키오븐 스마트스토어 CPS' && header === '광고비 R/S율') {
      const rate = parseFloat(value) || 0;
      // 미리보기 형식으로 displayValue 재정의
      displayValue = `${rate} % (어드민 : ${rate / 100})`;
      // 중요: 원본 value 변수는 변경하지 않음 (addRow 함수 내부 스코프)
    }

    if (header === '부스팅 CPC 광고 placement 세팅 정보 옵션_추천 세팅 여부') {
        if (value === '세팅 O') {
            displayValue = '네이버마케팅_추천(nvmarketing_best) : 우선순위 1';
        }
        // '세팅 X'인 경우 displayValue는 그대로 '세팅 X'가 됩니다.
    }

    if (header === '부스팅 CPC 광고 placement 세팅 정보 옵션_카테고리') {
        const categoryMap = {
            '건강': '네이버마케팅_건강(nvmarketing_health) : 우선순위 1',
            '식품': '네이버마케팅_식품(nvmarketing_food) : 우선순위 1',
            '생활': '네이버마케팅_생활(nvmarketing_living) : 우선순위 1',
            '뷰티': '네이버마케팅_뷰티(nvmarketing_beauty) : 우선순위 1',
            '기타': '네이버마케팅_기타(nvmarketing_etc) : 우선순위 1'
        };
        
        if (categoryMap[value]) {
            displayValue = categoryMap[value]; // 선택값(예: 건강)을 상세 문구로 덮어씌움
        }
    }

    if (header === '가이드 메세지') {
      displayValue = displayValue.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // 지정된 필드에 대해서만 HTML 태그를 안전한 텍스트로 변환
    if (fieldsToEscape.includes(header)) {
      displayValue = displayValue.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }
    
    if (header.includes('설정') || header.includes('메모장') || header.includes('서브2') || header.includes('placement') || header.includes('요청사항') || header.includes('instruction 문구')) {
      displayValue = displayValue.replace(/\n/g, '<br>');
    }
    
    return `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${header}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${displayValue}</td></tr>`;
  };




  // '테스트 광고' 타입일 때만 특별한 순서로 이메일 본문 생성
  if (subType === '테스트 광고') {
 let emailFieldOrder = [];
    const selection = data['선택항목'];

    // ▼▼▼ [수정] '선택항목'에 따라 이메일 본문 순서를 다르게 지정 ▼▼▼
    if (selection === '서버 연동') {
      emailFieldOrder = [
        '특이사항 메모장', '선택항목', '캠페인명', '광고주', '재참여 타입', '테스트 광고 타입', 
        '광고 네트워크 연동 매체', '광고주 연동토큰', '집행매체1', 'OS', '총물량', '일물량',
        '광고 집행 시작', '광고 집행 종료', '단가', '리워드',
        'URL_기본', 'URL_AOS', 'URL_IOS', 'URL_PC'
      ];
    } else if (selection === 'CPA TRACKER') {
   emailFieldOrder = [
        '특이사항 메모장', '선택항목', '캠페인명', '광고주', '재참여 타입', '테스트 광고 타입',
        '집행매체1', 'OS', '총물량', '일물량', '광고 집행 시작', '광고 집행 종료', '단가', '리워드',
        '트래커', '완료 이벤트 이름', 
        '완료 이벤트 조건(JSON)', '완료 이벤트 조건(JSON)-이벤트 타입', '완료 이벤트 조건(JSON)-이벤트 이름',
        '완료 이벤트 조건(JSON)-value', '완료 이벤트 조건(JSON)-from', '완료 이벤트 조건(JSON)-to',
        '완료 이벤트 조건(개별)', '완료 이벤트 조건(개별)-파라미터 타입', '완료 이벤트 조건(개별)-파라미터 이름',
        '완료 이벤트 조건(개별)-value', '완료 이벤트 조건(개별)-from', '완료 이벤트 조건(개별)-to',
        '최소 결제 금액', 'URL_기본', 'URL_AOS', 'URL_IOS', 'URL_PC'
      ];
    }
    // ▲▲▲ [수정] ▲▲▲

emailFieldOrder.forEach(header => {
      let value;
      if (header === '광고주 연동토큰') {
        // 'GreenP'가 아닐 때만 광고주 연동토큰 행을 메일에 추가합니다.
        if (data['광고 네트워크 연동 매체'] !== 'GreenP') {
          value = data['광고주 연동 토큰 값'];
          body += addRow(header, value);
        }
      } else {
        value = data[header];
        body += addRow(header, value);
      }
    });

  } else {
    // 그 외 모든 광고 타입은 기존 방식(FormFields.gs의 headerOrder)을 따름
    const allFieldsData = getAllFormFields();
    let headersToShow = [];

    if (subType === '네이버페이 입점비' && data['네이버페이 입점비 본 광고 신규 등록 여부'] === '등록 O') {
        const mainAdType = data['본 광고 타입'];
        const mainAdFields = allFieldsData[type]?.[mainAdType]?.headerOrder || [];
        const feeAdFields = allFieldsData[type]?.[subType]?.headerOrder || [];
        headersToShow = ['광고 타입', ...mainAdFields, ...feeAdFields];
    } else {
        const fieldInfo = allFieldsData[type]?.[subType];
        headersToShow = ['광고 타입', ...(fieldInfo ? fieldInfo.headerOrder : Object.keys(data))];
    }
    
    [...new Set(headersToShow)].forEach(header => {
      if (header === '미션 정보 붙여넣기' && data[header]) {
        const missionText = data[header];
        const missions = missionText.trim().split('\n').map(line => line.split('\t'));
        let totalCost = 0;
        let missionTableHtml = `<table style="width:100%; border-collapse: collapse; font-size: 11px;"><thead><tr style="background-color:#f0f0f0;"><th style="border: 1px solid #ccc; padding: 5px;">미션</th><th style="border: 1px solid #ccc; padding: 5px;">이벤트코드</th><th style="border: 1px solid #ccc; padding: 5px;">안내문구</th><th style="border: 1px solid #ccc; padding: 5px;">단가</th><th style="border: 1px solid #ccc; padding: 5px;">소요시간</th></tr></thead><tbody>`;
        missions.forEach(mission => {
          missionTableHtml += `<tr><td style="border: 1px solid #ccc; padding: 5px;">${mission[0] || ''}</td><td style="border: 1px solid #ccc; padding: 5px;">${mission[1] || ''}</td><td style="border: 1px solid #ccc; padding: 5px;">${mission[2] || ''}</td><td style="border: 1px solid #ccc; padding: 5px;">${mission[3] || ''}</td><td style="border: 1px solid #ccc; padding: 5px;">${mission[4] || ''}</td></tr>`;
          totalCost += parseFloat(mission[3]) || 0;
        });
        missionTableHtml += `</tbody><tfoot><tr style="font-weight: bold;"><td colspan="3" style="border: 1px solid #ccc; padding: 5px; text-align: right;">단가 총합:</td><td colspan="2" style="border: 1px solid #ccc; padding: 5px;">${totalCost.toLocaleString()}</td></tr></tfoot></table>`;
        body += addRow(header, missionTableHtml);
      } else {
        let currentHeader = header;
        if(header === 'OS_DROPDOWN' || header === 'OS_CHECKBOX') currentHeader = 'OS';
        if(header.startsWith('집행매체2_')) currentHeader = '집행매체2';
        
        const value = (header === '광고 타입') ? subType : data[header];
        
        if (!header.startsWith('집행매체2_') || (header.startsWith('집행매체2_') && data[header])) {
           body += addRow(currentHeader, value);
        }

        if (header === '구독 페이지 랜딩 URL') {
          if (data['AOS 랜딩 URL']) { body += addRow('AOS 랜딩 URL', data['AOS 랜딩 URL']); }
          if (data['IOS 랜딩 URL']) { body += addRow('IOS 랜딩 URL', data['IOS 랜딩 URL']); }
        }
      }
    });
  }

  body += `</table>`;

  GmailApp.sendEmail(ADMIN_EMAIL, mainSubject, '', {
    htmlBody: body,
    cc: ccEmails
  });

  try {
    const slackMessage = { 'text': `${mainSubject}` };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  } catch (e) {
    console.error(`신규 등록 슬랙 발송 실패 (ID: ${id}): ${e.toString()}`);
  }
  
  Utilities.sleep(1000);
  const searchQuery = `subject:"${mainSubject}" in:sent`;
  const threads = GmailApp.search(searchQuery, 0, 1);
  
  if (threads && threads.length > 0) {
    return threads[0].getId();
  }
  
  return null;
}

// Registration.gs 파일에서 기존 recordConfirmation과 sendAssignmentConfirmationEmail 함수를
// 모두 삭제하고 아래 코드로 교체하세요.

function recordConfirmation(adId, approverEmail) {
  try {
    const found = findRowById(adId);
    if (!found) {
      return `ID: ${adId} 광고를 찾을 수 없습니다.`;
    }

    const { sheet, rowIndex, headers, rowData } = found;
    const approverColIndex = headers.indexOf('담당자');
    const statusColIndex = headers.indexOf('상태');

    const currentStatus = rowData[statusColIndex];
    if (currentStatus === '스킵처리') {
      return `처리 실패: 이 광고 요청 건(ID: ${adId})은 이미 스킵 처리되어 담당자로 지정할 수 없습니다.`;
    }

    if (approverColIndex === -1 || statusColIndex === -1) {
      return `오류: 시트에 '담당자' 또는 '상태' 컬럼이 없습니다.`;
    }

    const currentApprover = rowData[approverColIndex];
    if (currentApprover && currentApprover !== '') {
      return `처리 실패: 이 광고(ID: ${adId})는 이미 ${currentApprover} 님이 담당하고 있습니다.`;
    }

    const timestampColIndex = headers.indexOf('담당자 확인 일시');
    sheet.getRange(rowIndex, approverColIndex + 1).setValue(approverEmail);
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
    if (timestampColIndex > -1) {
      const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(rowIndex, timestampColIndex + 1).setValue(formattedTimestamp);
    }

    const searchQuery = `"ID: ${adId}"`;
    const threads = GmailApp.search(searchQuery, 0, 1);
    
    if (threads && threads.length > 0) {
      threads[0].replyAll("", {
        htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>ID: ${adId}</b> 광고의 담당자로 지정되어 등록을 진행합니다.</p><p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p><p>감사합니다.</p>`
      });
    } else {
      console.error(`담당자 지정 알림 실패: 광고 ID ${adId}에 대한 메일 스레드를 찾을 수 없습니다.`);
    }

    return `ID: ${adId} 광고의 담당자로 ${approverEmail}님이 지정되었습니다. 이 창은 닫아도 됩니다.`;

  } catch (e) {
    return `담당자 지정 처리 중 오류가 발생했습니다: ${e.toString()}`;
  }
}


function sendAssignmentConfirmationEmail(threadId, adId, approverEmail) {
 if (!threadId) return;
 try {
  const thread = GmailApp.getThreadById(threadId);
  if (thread) {
   thread.replyAll("", { // .reply()를 .replyAll()로 변경
    htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>ID: ${adId}</b> 광고의 담당자로 지정되어 등록을 진행합니다.</p><p>감사합니다.</p>`
        // replyTo 옵션 제거
   });
  }
 } catch (e) {
  console.error(`담당자 지정 알림 메일 발송 실패 (Thread ID: ${threadId}): ${e.toString()}`);
 }
}


function processCompletion(registrationId, adId, completerEmail) {
  try {
    if (!registrationId || !adId) {
      return { success: false, message: '등록 ID와 광고 ID를 모두 입력해야 합니다.' };
    }

    const found = findRowById(registrationId);
    if (!found) {
      return { success: false, message: `등록 ID(${registrationId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    const statusColIndex = headers.indexOf('상태');
    const adIdColIndex = headers.indexOf('광고 ID');
    const completionDateColIndex = headers.indexOf('완료 일시');
    const registrantColIndex = headers.indexOf('등록자'); // threadIdColIndex는 더 이상 필요 없음

    if (statusColIndex === -1 || adIdColIndex === -1 || completionDateColIndex === -1 || registrantColIndex === -1) {
      return { success: false, message: '시트에서 필수 컬럼(상태, 광고 ID, 완료 일시, 등록자)을 찾을 수 없습니다.' };
    }
    
    const currentStatus = rowData[statusColIndex];
    if (currentStatus === '등록 완료') {
        return { success: false, message: `이미 등록 완료 처리된 광고입니다. (ID: ${registrationId})` };
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('등록 완료');
    sheet.getRange(rowIndex, adIdColIndex + 1).setValue("'" + adId);
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(timestamp);


    logUserAction(completerEmail, '등록 완료 처리', {
      targetId: registrationId,
      message: `광고 ID '${registrationId}'를 광고 ID '${adId}'로 등록 완료 처리`
    });

    return { success: true, message: `광고(등록ID: ${registrationId})가 성공적으로 등록 완료 처리되었습니다. 이 창은 닫아도 됩니다.` };
  } catch (e) {
    console.error(`Completion Error: ${e.toString()}`);
    return { success: false, message: '완료 처리 중 오류가 발생했습니다: ' + e.message };
  }
}



function processSkip(adId) {
  const found = findRowById(adId);
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
          thread.replyAll("", { // .reply() -> .replyAll()로 변경
            htmlBody: `<p>안녕하세요,</p><p>요청하신 <b>ID: ${adId}</b> 건이 <b>스킵 처리</b>되었음을 알려드립니다.</p><p>감사합니다.</p><p>- 처리자: ${skipperEmail}</p>`,
          });
        }
      } catch (e) {
        console.error(`스킵 알림 메일 발송 실패(ID: ${adId}): ${e.toString()}`);
      }
    }

    const brand = found.rowData[found.headers.indexOf('브랜드')] || '';
    const actionName = found.rowData[found.headers.indexOf('액션명')] || '';
    const subject = [brand, actionName].filter(Boolean).join('_') || adId;

    const slackMessage = { 'text': `[스킵 처리] ${subject} (ID: ${adId})` };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
    try {
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    } catch(e) {
      console.error(`스킵 알림 슬랙 발송 실패 (ID: ${adId}): ${e.toString()}`);
    }

    logUserAction(skipperEmail, '스킵 처리', {
      targetId: adId,
      message: `광고 ID '${adId}' 스킵 처리`
    });

    return { success: true, message: `ID(${adId})가 성공적으로 스킵 처리되었습니다.` };
  }
  return { success: false, message: `ID(${adId})를 찾을 수 없습니다.` };
}

// Registration.gs 파일의 processRejection 함수 전체를 이 코드로 교체하세요.

function processRejection(adId, reason) {
  try {
    const rejectorEmail = Session.getActiveUser().getEmail(); // 현재 사용자 (운영팀)

    const found = findRowById(adId);
    if (!found) {
      return { success: false, message: `ID(${adId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    // 시트 상태 업데이트 (기존과 동일)
    const statusColIndex = headers.indexOf('상태');
    const rejectionDateColIndex = headers.indexOf('반려 일시');
    const rejectionReasonColIndex = headers.indexOf('반려 사유');
    const registrantColIndex = headers.indexOf('등록자');

    if ([statusColIndex, rejectionDateColIndex, rejectionReasonColIndex, registrantColIndex].includes(-1)) {
      return { success: false, message: '시트에서 필수 컬럼(상태, 반려 일시, 반려 사유, 등록자)을 찾을 수 없습니다. 관리자에게 문의하세요.' };
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('반려');
    sheet.getRange(rowIndex, rejectionDateColIndex + 1).setValue(timestamp);
    sheet.getRange(rowIndex, rejectionReasonColIndex + 1).setValue(reason);

    const registrantEmail = rowData[registrantColIndex]; // 원본 요청자 (영업팀)
    
    if (registrantEmail) {
      const subject = `[광고 등록 시스템] 요청하신 광고(ID: ${adId})가 반려되었습니다.`;
      let emailBody = `<p>안녕하세요, ${registrantEmail.split('@')[0]}님.</p>
                       <p>요청하신 광고(ID: <b>${adId}</b>)가 아래와 같은 사유로 반려되었습니다.</p>`;
      
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
      
      const searchQuery = `"ID: ${adId}"`;
      const threads = GmailApp.search(searchQuery, 0, 1);

      if (threads && threads.length > 0) {
        threads[0].replyAll('', mailOptions);
      } else {
        console.error(`Could not find thread for adId: ${adId}. Sending a new email as a fallback.`);
        GmailApp.sendEmail(registrantEmail, subject, '', mailOptions);
      }
      // ▲▲▲ [수정] ▲▲▲
    }

    logUserAction(rejectorEmail, '반려 처리', {
      targetId: adId,
      message: `광고 ID '${adId}' 반려 처리. 사유: ${reason}`
    });

    return { success: true, message: `ID(${adId})가 성공적으로 반려 처리되었습니다.` };
  } catch (e) {
    console.error(`Error in processRejection: ${e.toString()}`);
    return { success: false, message: '반려 처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}
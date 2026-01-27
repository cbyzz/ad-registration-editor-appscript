function getCashslideAllFormFields() {
  // 등록 폼 필드 가져오기
  const registrationFields = getCashslideFields(); 
  
  // 수정 폼 필드 가져오기
  const modificationFields = getCashslideModificationFields();

  // 두 데이터를 하나의 객체로 합쳐서 반환
  const allFields = {
    registration: registrationFields,
    modification: modificationFields
  };
  
  // 광고주 목록은 공통으로 사용하므로 최상위에 추가
  allFields.registration.dropdowns.advertisers = getExternalAdvertisersData().list;
  
  return allFields;
}

function submitCashslideRegistration(formData) {
    const lock = LockService.getUserLock();
    lock.waitLock(30000);

    try {
        const userEmail = Session.getActiveUser().getEmail();
        const userName = userEmail.split('@')[0];
        const sheetName = "캐시슬라이드 광고";
        let sheet = ss.getSheetByName(sheetName);

        // ▼▼▼ [수정] 시트에 정의할 전체 헤더 목록 (순서는 초기 생성 시에만 중요) ▼▼▼
        const intendedHeaders = [
            'id', 'timestamp', 'registrant', 'status', 'manager', 'manager_timestamp', 'mail_thread_id',
            'rejection_timestamp', 'rejection_reason', 'ad_id', 'completion_timestamp',
            'ad_type', 'request_details', 'advertiser_admin_id', 'advertiser_name',
            'campaign_name', 'ad_type_option', 'webview_enabled', 'webview_template_type', 'webview_template_list', 'webview_url',
            'webview_top_overlay_height', 'webview_top_overlay_color', 'webview_bottom_overlay_height', 'webview_bottom_overlay_color',
            'creative_path', 'creative_path_image', 'creative_path_video',
            'autoplay_on_first_slide', 'dynamic_creative_loop',
            'slot', // 맥스뷰 슬랏
            'cover_template_setting', 'aspect_ratio', 'video_streaming', 'allow_margin', 'mid_roll_reward', // 맥스뷰 Teaser/Native 공통
            'cover_title_1', 'cover_title_2', 'cover_advertiser_title', // 맥스뷰 Native 전용
            'demo_target_1', 'demo_target_2', 'retarget_cluster', 'app_package_retargeting', 'app_package_detargeting',
            'lbs_latitude', 'lbs_longitude', 'lbs_radius', 'tag',
            'frequency_criteria', 'frequency_type', 'frequency_daily_limit', 'frequency_max_limit', // 'frequency_type' 추가
            'tracker', 'priority', 'seg_filter', 'exposure_settings_json',
            'live_start_date', 'live_end_date', 'landing_url', 'frequency_timeboard', 'frequency_cpmc',
            'daily_volume_timeboard', 'daily_volume_cpmc', 'priority_timeboard', 'priority_cpmc',
            'ad_start_date', 'ad_end_date', 'title', 'product_info', 'detail_page_url',
            'hns_ad_type', 'slot_priority', 'frequency', 'targeting', 'app_targeting', 'cluster', // <-- cluster 포함
             'daily_volume', 'deeplink_info_json'
        ];
        // ▲▲▲ [수정] ▲▲▲

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.appendRow(intendedHeaders); // 시트 생성 시에는 의도한 순서대로 헤더 추가
            sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold").setFrozenRows(1);
        } else {
            const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            const missingHeaders = intendedHeaders.filter(h => !currentHeaders.includes(h));
            if (missingHeaders.length > 0) {
                // 누락된 헤더만 시트의 맨 뒤에 추가 (기존 열 순서는 유지)
                sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
            }
        }

        // --- 데이터 준비 로직 (이전과 동일) ---
        const idPrefix = `cs-${userName}-`;
        const nextId = getNextSequentialId(sheet, idPrefix);
        const uniqueId = idPrefix + nextId;
        const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
        const subType = formData.cashslideSubType;
        const advertiser = formData['광고주'] || '';
        const campaign = formData['캠페인'] || '';
        let subject = '';
        let dataToSave = {
            id: uniqueId, timestamp: formattedTimestamp, registrant: userEmail, status: '등록 요청 완료',
            ad_type: subType, ccRecipients: formData.ccRecipients
        };
        let exposureSettings = [];
        let fileAttachment = null;
        let deeplinkInfo = [];

    if (subType === '캐슬_노출형(홈앤쇼핑)') {
        const hnsSubType = formData['홈앤쇼핑 광고 타입'];
        let campaignName = formData['캠페인'];
        if (hnsSubType === '기획전') {
          campaignName = '캐시슬라이드_CPS_기획전';
        } else if (hnsSubType === '프로모션') {
          campaignName = '캐시슬라이드_CPS_프로모션';
        } else {
          campaignName = "캐시슬라이드_CPS_" + hnsSubType;
        }

        let startDate = '';

        const pastedData = (formData['세부 정보 (복사 붙여넣기)'] || '').trim().split('\n');
        const headers = ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'];
        
        if (pastedData.length > 0 && pastedData[0] !== "") {
            pastedData.forEach(line => {
                const values = line.split('\t');
                let rowObj = {};
                headers.forEach((header, i) => rowObj[header] = values[i] || '');
                deeplinkInfo.push(rowObj);
            });
            if (deeplinkInfo.length > 0) {
              startDate = formatDate_yyMMdd(deeplinkInfo[0]['시작일자']);
            }
        }
        
        const titleCore = `[잠금화면] ${campaignName}_${startDate}`;
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

Object.assign(dataToSave, {
                request_details: formData['요청사항'],
                advertiser_name: '홈앤쇼핑', // 고정값
                campaign_name: campaignName,
                tracker: 'Singular', // 고정값
                hns_ad_type: hnsSubType,
                slot_priority: formData['슬랏 / 우선순위'],
                frequency: formData['프리퀀시'],
                targeting: formData['타겟팅'],
                app_targeting: formData['앱타겟팅'],
                cluster: formData['클러스터'], // <-- 클러스터 데이터 추가
                daily_volume: formData['일물량'],
                creative_path: formData['소재경로'],
                deeplink_info_json: JSON.stringify(deeplinkInfo, null, 2)
            });

    } else if (subType === '캐슬_노출형(라방패키지)') {
        const liveStartDateStr = formData['라이브 시작일시'];
        const liveStartDate = new Date(liveStartDateStr);
        const datePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "yy/MM/dd");
        const timePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "HH'시'");
        
        const titleCoreParts = [ `[라방패키지] ${advertiser}`, campaign, `${datePart}_${timePart}_라방패키지` ];
        const titleCore = titleCoreParts.filter(Boolean).join('_');
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

        Object.assign(dataToSave, {
            request_details: formData['요청사항'], advertiser_admin_id: formData['광고주 어드민 ID (신규 or 기존 ID)'], advertiser_name: advertiser,
            campaign_name: campaign, live_start_date: liveStartDateStr, live_end_date: formData['라이브 종료일시'],
            tag: Array.isArray(formData['태그']) ? formData['태그'].join(', ') : formData['태그'],
            landing_url: formData['랜딩 URL'], webview_url: formData['웹뷰 URL'],
            webview_template_type: '4. 웹뷰광고', webview_template_list: '40. 웹뷰 풀사이즈 (상하단 조절 가능)',
            webview_top_overlay_height: '100dp', webview_top_overlay_color: '#000000',
            webview_bottom_overlay_height: '100dp', webview_bottom_overlay_color: '#000000',
            frequency_timeboard: '노출 1회', frequency_cpmc: 'X',
            daily_volume_timeboard: '무제한', daily_volume_cpmc: '노출 80,000 / 클릭 12,000',
            priority_timeboard: '30', priority_cpmc: '12'
        });
    } else if (subType === '캐슬_노출형(라방패키지)_유튜브') {
        const liveStartDateStr = formData['라이브 시작일시'];
        const liveStartDate = new Date(liveStartDateStr);
        const datePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "yy/MM/dd");
        const timePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "HH'시'");
        
        // 제목 생성 방식 변경 (사용자 요청 포맷 적용)
        const titleCoreParts = [ 
            `[라방패키지] ${advertiser}`, // "[라방패키지] 광고주"
            campaign,                     // "캠페인"
            '유튜브',                     // "유튜브"
            `${datePart}~${timePart}`,    // "날짜~시간"
            '라방패키지'                  // "라방패키지"
        ]; 
        const titleCore = titleCoreParts.filter(Boolean).join('_');
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

        Object.assign(dataToSave, {
            request_details: formData['요청사항'], advertiser_admin_id: formData['광고주 어드민 ID (신규 or 기존 ID)'], advertiser_name: advertiser,
            campaign_name: campaign, live_start_date: liveStartDateStr, live_end_date: formData['라이브 종료일시'],
            tag: Array.isArray(formData['태그']) ? formData['태그'].join(', ') : formData['태그'],
            landing_url: formData['랜딩 URL'], webview_url: formData['웹뷰 URL'],
            
            // --- 변경된 기본값 ---
            webview_template_type: '4. 웹뷰광고', 
            webview_template_list: '44. 웹뷰 풀사이즈 (상단 여백 / 웹 비율 조절 가능)',
            webview_top_overlay_height: '100dp',
            aspect_ratio: '85%', 
            
            // --- 기존과 동일한 값 ---
            frequency_timeboard: '노출 1회', frequency_cpmc: 'X',
            daily_volume_timeboard: '무제한', daily_volume_cpmc: '노출 80,000 / 클릭 12,000',
            priority_timeboard: '30', priority_cpmc: '12'
        });

    } else if (subType === '캐슬_라이브적립') {
        const liveStartDateStr = formData['라이브 시작일시'];
        const liveStartDate = new Date(liveStartDateStr);
        const datePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "yy/MM/dd");
        const timePart = Utilities.formatDate(liveStartDate, "Asia/Seoul", "HH'시'");

        const titleCore = `[라이브적립] ${advertiser}_${datePart}_${timePart}`;
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

        const adStartDate = new Date(liveStartDate.getTime() - (2 * 24 * 60 * 60 * 1000));
        const adEndDate = new Date(liveStartDate.getTime() + (60 * 60 * 1000));

        Object.assign(dataToSave, {
          advertiser_name: advertiser,
          live_start_date: liveStartDateStr,
          title: formData['타이틀'],
          product_info: formData['상품정보'],
          landing_url: formData['랜딩 URL'],
          detail_page_url: formData['상세 페이지 URL'],
          ad_start_date: Utilities.formatDate(adStartDate, "Asia/Seoul", "yyyy/MM/dd HH:mm"),
          ad_end_date: Utilities.formatDate(adEndDate, "Asia/Seoul", "yyyy/MM/dd HH:mm")
        });
        
        if (formData.creative_file_data) {
          fileAttachment = Utilities.newBlob(
            Utilities.base64Decode(formData.creative_file_data),
            formData.creative_file_type,
            formData.creative_file_name
          );
          dataToSave.creative_path = formData.creative_file_name;
        }

    } else if (subType === '캐슬_노출형(오토뷰)') {
        exposureSettings = formData.dynamicTableData || [];
        
        let firstSlot = '';
        if (exposureSettings.length > 0) {
            firstSlot = exposureSettings[0]['슬랏'].replace('CPV-', '');
        }
        let startDate = formatDate_yyMMdd(exposureSettings.length > 0 ? exposureSettings[0]['라이브 시작 일시'] : '');
        let endDate = formatDate_yyMMdd(exposureSettings.length > 0 ? exposureSettings[0]['라이브 종료 일시'] : '');
        
        const titleCore = `[CPV-${firstSlot}] ${advertiser}_오토뷰_${startDate}~${endDate}`;
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

        Object.assign(dataToSave, {
            request_details: formData['요청사항'],
            advertiser_admin_id: formData['광고주 어드민 ID (신규 or 기존 ID)'],
            advertiser_name: advertiser,
            campaign_name: campaign,
            creative_path_image: formData['이미지 소재경로'],
            creative_path_video: formData['영상 소재경로'],
            autoplay_on_first_slide: formData['첫 슬라이드시 자동 재생 여부'],
            dynamic_creative_loop: formData['동적소재 반복재생 여부'],
            demo_target_1: formData['데모타겟1'], demo_target_2: formData['데모타겟2'],
            retarget_cluster: formData['리타겟 클러스터'],
            app_package_retargeting: formData['앱 패키지명 - 리타겟팅'],
            app_package_detargeting: formData['앱 패키지명 - 디타겟팅'],
            lbs_latitude: formData['LBS 타겟팅 - 위도'], lbs_longitude: formData['LBS 타겟팅 - 경도'], lbs_radius: formData['LBS 타겟팅 - 범위'],
            tag: Array.isArray(formData['태그']) ? formData['태그'].join(', ') : formData['태그'],
            frequency_criteria: formData['프리퀀시 기준'],
            frequency_type: formData['프리퀀시 타입'],
            frequency_daily_limit: formData['일일 프리퀀시 제한 횟수'],
            frequency_max_limit: formData['최대 프리퀀시 제한 횟수'],
            tracker: formData['트래커'],
            priority: formData['노출 우선순위'],
            seg_filter: formData['Seg Filter'],
            exposure_settings_json: JSON.stringify(exposureSettings, null, 2)
        });
    } else if (subType === '캐슬_노출형(맥스뷰)') {
        exposureSettings = formData.dynamicTableData || [];
        
        let firstSlot = '';
        if (exposureSettings.length > 0) {
            firstSlot = exposureSettings[0]['슬랏'] || ''; // 예: CPV-01
        }
        
        let startDate = formatDate_yyMMdd(exposureSettings.length > 0 ? exposureSettings[0]['라이브 시작 일시'] : '');
        let endDate = formatDate_yyMMdd(exposureSettings.length > 0 ? exposureSettings[0]['라이브 종료 일시'] : '');
        
        const titleCore = `[${firstSlot || 'CPV-??'}] ${advertiser}_맥스뷰_${startDate}~${endDate}`;
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;


        Object.assign(dataToSave, {
            request_details: formData['요청사항'],
            advertiser_admin_id: formData['광고주 어드민 ID (신규 or 기존 ID)'],
            // slot: formData['슬랏'],
            advertiser_name: advertiser,
            campaign_name: campaign,
            creative_path: formData['소재경로'],
            cover_template_setting: formData['커버 템플릿 설정'],
            cover_title_1: formData['커버) 광고 타이틀 1번째 줄'],
            cover_title_2: formData['커버) 광고 타이틀 2번째 줄'],
            cover_advertiser_title: formData['커버) 광고주 타이틀'],
            aspect_ratio: formData['선택형 화면비율'],
            dynamic_creative_loop: formData['동적소재 반복재생 여부'],
            autoplay_on_first_slide: formData['첫 슬라이드시 자동 재생 여부'],
            video_streaming: formData['동영상 스트리밍 여부'],
            allow_margin: formData['여백허용'],
            mid_roll_reward: formData['동영상 재생중간 적립 설정'],
            demo_target_1: formData['데모타겟1'], demo_target_2: formData['데모타겟2'],
            retarget_cluster: formData['리타겟 클러스터'],
            app_package_retargeting: formData['앱 패키지명 - 리타겟팅'],
            app_package_detargeting: formData['앱 패키지명 - 디타겟팅'],
            lbs_latitude: formData['LBS 타겟팅 - 위도'], lbs_longitude: formData['LBS 타겟팅 - 경도'], lbs_radius: formData['LBS 타겟팅 - 범위'],
            tag: Array.isArray(formData['태그']) ? formData['태그'].join(', ') : formData['태그'],
            frequency_criteria: formData['프리퀀시 기준'],
            frequency_type: formData['프리퀀시 타입'],
            frequency_daily_limit: formData['일일 프리퀀시 제한 횟수'],
            frequency_max_limit: formData['최대 프리퀀시 제한 횟수'],
            tracker: formData['트래커'],
            priority: formData['노출 우선순위'],
            seg_filter: formData['Seg Filter'],
            exposure_settings_json: JSON.stringify(exposureSettings, null, 2)
        });
    } else { // '캐슬_노출형(기본)', '캐슬_노출형(맥스뷰)'
        exposureSettings = formData.dynamicTableData || [];
        
        let startDateStr = exposureSettings.length > 0 ? exposureSettings[0]['라이브 시작 일시'] : '';
        let endDateStr = exposureSettings.length > 0 ? exposureSettings[0]['라이브 종료 일시'] : '무제한';

        const startDate = formatDate_yyMMdd(startDateStr);
        const endDate = formatDate_yyMMdd(endDateStr);

        const webviewPart = formData['웹뷰 진행여부'] === '웹뷰 진행' ? '_웹뷰' : '';

        const titleCoreParts = [ `[잠금화면] ${advertiser}`, campaign, `${startDate}~${endDate}${webviewPart}` ];
        const titleCore = titleCoreParts.filter(Boolean).join('_');
        subject = `[캐시슬라이드 등록 요청] ${titleCore} (ID: ${uniqueId})`;

        Object.assign(dataToSave, {
            request_details: formData['요청사항'], advertiser_admin_id: formData['광고주 어드민 ID (신규 or 기존 ID)'],
            advertiser_name: advertiser, campaign_name: campaign, ad_type_option: formData['광고 타입'],
            webview_enabled: formData['웹뷰 진행여부'], webview_url: formData['웹뷰 URL'],
            webview_top_overlay_height: formData['상단 오버레이 영역 높이'], webview_top_overlay_color: formData['상단 오버레이 색'],
            webview_bottom_overlay_height: formData['하단 오버레이 영역 높이'], webview_bottom_overlay_color: formData['하단 오버레이 색'],
            creative_path: formData['소재경로'], demo_target_1: formData['데모타겟1'], demo_target_2: formData['데모타겟2'],
            retarget_cluster: formData['리타겟 클러스터'], app_package_retargeting: formData['앱 패키지명 - 리타겟팅'], app_package_detargeting: formData['앱 패키지명 - 디타겟팅'],
            lbs_latitude: formData['LBS 타겟팅 - 위도'], lbs_longitude: formData['LBS 타겟팅 - 경도'], lbs_radius: formData['LBS 타겟팅 - 범위'],
            tag: Array.isArray(formData['태그']) ? formData['태그'].join(', ') : formData['태그'],
            frequency_criteria: formData['프리퀀시 기준'], frequency_type: formData['프리퀀시 타입'], frequency_daily_limit: formData['일일 프리퀀시 제한 횟수'], frequency_max_limit: formData['최대 프리퀀시 제한 횟수'],
            tracker: formData['트래커'], priority: formData['노출 우선순위'], seg_filter: formData['Seg Filter'],
            exposure_settings_json: JSON.stringify(exposureSettings, null, 2)
        });
        
        if (dataToSave.webview_enabled === '웹뷰 진행') {
            dataToSave.webview_template_type = '4. 웹뷰광고';
            dataToSave.webview_template_list = '40. 웹뷰 풀사이즈 (상하단 조절 가능)';
        }
    }

    const messageId = sendCashslideNotification(userEmail, uniqueId, subject, dataToSave, exposureSettings, fileAttachment);
    dataToSave.mail_thread_id = messageId;
const currentSheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        const newRow = currentSheetHeaders.map(header => dataToSave[header] || ''); // 해당 헤더 이름으로 dataToSave에서 값을 찾고, 없으면 빈 문자열('') 사용
    sheet.appendRow(newRow);

    logUserAction(userEmail, '캐시슬라이드 등록 요청', { targetId: uniqueId, message: subject });
    return { success: true, message: `캐시슬라이드 광고 등록 요청이 완료되었습니다. (ID: ${uniqueId})` };

  } catch (e) {
    console.error(`submitCashslideRegistration Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


function sendCashslideNotification(senderEmail, uniqueId, subject, data, exposureSettings, fileAttachment) {
  let body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 캐시슬라이드 광고 등록을 요청했습니다.</p>
              <p><b>등록 ID: ${uniqueId}</b></p>`;
  
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_cs&id=${uniqueId}`;
  body += `<div style="margin-top: 15px; margin-bottom: 15px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
             <a href="${confirmationUrl}" style="background-color: #007bff; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; margin-right: 10px;">[ 이 광고 담당하기 ]</a>
             <br><br>
             <div style="display: inline-block; vertical-align: middle;">
               <form action="${ScriptApp.getService().getUrl()}" method="get" target="_blank" style="margin:0; padding:0;">
                 <input type="hidden" name="action" value="complete_cs">
                 <input type="hidden" name="id" value="${uniqueId}">
                 <input type="text" name="adId" placeholder="광고 ID 입력" required style="padding: 8px; border: 1px solid #ccc; border-radius: 4px; margin-right: 5px;">
                 <button type="submit" style="background-color: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; border: none; cursor: pointer;">[ 광고 등록 완료 처리 ]</button>
               </form>
             </div>
             <br><br>
             <a href="${ss.getUrl()}" style="color: #0056b3; text-decoration: none; margin-right: 15px;">스프레드시트 바로가기</a>
             <a href="${SYSTEM_URL}" style="color: #0056b3; text-decoration: none;">광고 등록 시스템 바로가기</a>
           </div>`;

  body += `<hr><h3>요청 내용</h3>
           <table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">`;

  const headerMap = {
    // 공통 및 기본
    ad_type: '광고 타입', request_details: '요청사항', advertiser_admin_id: '광고주 어드민 ID', advertiser_name: '광고주', campaign_name: '캠페인', ad_type_option: '광고 타입(옵션)',
    // 웹뷰
    webview_enabled: '웹뷰 진행여부', webview_template_type: '템플릿 타입', webview_template_list: '템플릿 목록', webview_url: '웹뷰 URL',
    webview_top_overlay_height: '상단 오버레이 영역 높이', webview_top_overlay_color: '상단 오버레이 색', webview_bottom_overlay_height: '하단 오버레이 영역 높이', webview_bottom_overlay_color: '하단 오버레이 색',
    // 타겟팅
    demo_target_1: '데모타겟1', demo_target_2: '데모타겟2', retarget_cluster: '리타겟 클러스터',
    app_package_retargeting: '앱 패키지명 - 리타겟팅', app_package_detargeting: '앱 패키지명 - 디타겟팅', lbs_latitude: 'LBS 타겟팅 - 위도',
    lbs_longitude: 'LBS 타겟팅 - 경도', lbs_radius: 'LBS 타겟팅 - 범위', tag: '태그',
    // 프리퀀시 및 기타 설정
    frequency_criteria: '프리퀀시 기준', frequency_type: '프리퀀시 타입',
    frequency_daily_limit: '일일 프리퀀시 제한 횟수', frequency_max_limit: '최대 프리퀀시 제한 횟수',
    tracker: '트래커', priority: '노출 우선순위', seg_filter: 'Seg Filter', creative_path: '소재경로',
    // 라방패키지 & 라이브적립
    live_start_date: '라이브 시작일시', live_end_date: '라이브 종료일시', ad_start_date: '광고 시작 일자 (자동계산)', ad_end_date: '광고 종료 일자 (자동계산)',
    title: '타이틀', product_info: '상품정보', landing_url: '랜딩 URL', detail_page_url: '상세 페이지 URL',
    frequency_timeboard: '타임보드 라인', frequency_cpmc: '프리퀀시 (CPMC 3번 라인)',
    daily_volume_timeboard: '일물량 (타임보드 라인)', daily_volume_cpmc: '일물량 (CPMC 3번 라인)',
    priority_timeboard: '노출 우선순위 (타임보드 라인)', priority_cpmc: '노출 우선순위 (CPMC 3번 라인)',
    // 홈앤쇼핑
    hns_ad_type: '홈앤쇼핑 광고 타입', slot_priority: '슬랏 / 우선순위', frequency: '프리퀀시', targeting: '타겟팅', app_targeting: '앱타겟팅', 'cluster': '클러스터', daily_volume: '일물량',
    // 오토뷰
    creative_path_image: '이미지 소재경로', creative_path_video: '영상 소재경로',
    autoplay_on_first_slide: '첫 슬라이드시 자동 재생 여부', dynamic_creative_loop: '동적소재 반복재생 여부',
    // 맥스뷰
    cover_template_setting: '커버 템플릿 설정',
    aspect_ratio: '선택형 화면비율',
    video_streaming: '동영상 스트리밍 여부',
    allow_margin: '여백허용',
    mid_roll_reward: '동영상 재생중간 적립 설정',
    cover_title_1: '커버) 광고 타이틀 1번째 줄',
    cover_title_2: '커버) 광고 타이틀 2번째 줄',
    aspect_ratio: '웹뷰 가로세로 비율',
    cover_advertiser_title: '커버) 광고주 타이틀'
  };

  if (data.ad_type === '캐슬_노출형(홈앤쇼핑)') {
    const hnsFieldOrder = [
      'hns_ad_type', 'request_details', 'advertiser_name', 'campaign_name', 'tracker',
      'slot_priority', 'frequency', 'targeting', 'app_targeting', 'cluster', 'daily_volume', 'creative_path'
    ];
    hnsFieldOrder.forEach(key => {
      if (data[key]) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${headerMap[key]}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${String(data[key]).replace(/\n/g, '<br>')}</td></tr>`;
      }
    });
  } else if (data.ad_type === '캐슬_노출형(오토뷰)') {
    const autoviewFieldOrder = [
      'request_details', 'advertiser_admin_id', 'advertiser_name', 'campaign_name',
      'creative_path_image', 'creative_path_video', 'autoplay_on_first_slide', 'dynamic_creative_loop',
      'demo_target_1', 'demo_target_2', 'retarget_cluster', 'app_package_retargeting', 'app_package_detargeting',
      'lbs_latitude', 'lbs_longitude', 'lbs_radius', 'tag', 'frequency_criteria', 'frequency_type',
      'frequency_daily_limit', 'frequency_max_limit', 'tracker', 'priority', 'seg_filter'
    ];
    autoviewFieldOrder.forEach(key => {
      if (data[key]) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${headerMap[key]}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${String(data[key]).replace(/\n/g, '<br>')}</td></tr>`;
      }
    });
    body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">상품</td><td style="padding: 8px; border: 1px solid #e0e0e0;">오토뷰</td></tr>`;
  } else if (data.ad_type === '캐슬_노출형(맥스뷰)') {
    const maxviewFieldOrder = [
      'request_details', 'advertiser_admin_id', 'product', 'advertiser_name', 'campaign_name',
      'creative_path', 'cover_template_setting',
      'cover_title_1', 'cover_title_2', 'cover_advertiser_title',
      'aspect_ratio', 'dynamic_creative_loop', 'autoplay_on_first_slide', 'video_streaming', 'allow_margin', 'mid_roll_reward',
      'demo_target_1', 'demo_target_2', 'retarget_cluster', 'app_package_retargeting', 'app_package_detargeting',
      'lbs_latitude', 'lbs_longitude', 'lbs_radius', 'tag', 'frequency_criteria', 'frequency_type',
      'frequency_daily_limit', 'frequency_max_limit', 'tracker', 'priority', 'seg_filter'
    ];
    maxviewFieldOrder.forEach(key => {
        if (key === 'product') {
            body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">상품</td><td style="padding: 8px; border: 1px solid #e0e0e0;">맥스뷰</td></tr>`;
        } else if (data[key]) {
            body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${headerMap[key]}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${String(data[key]).replace(/\n/g, '<br>')}</td></tr>`;
        }
    });
  } else if (data.ad_type === '캐슬_노출형(라방패키지)' || data.ad_type === '캐슬_노출형(라방패키지)_유튜브') {
    
    // Helper to add a row if data[key] exists
    const addRow = (key) => {
      if (data[key] && headerMap[key]) {
        let value = String(data[key]).replace(/\n/g, '<br>');
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${headerMap[key]}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${value}</td></tr>`;
      }
    };

    // 1. 요청하신 순서대로 기본 필드 키 정의
    const fieldKeyOrder = [
      'request_details', 'advertiser_admin_id', 'advertiser_name', 'campaign_name',
      'live_start_date', 'live_end_date', 'tag', 'landing_url', 'webview_url'
    ];
    
    // 2. 순서대로 기본 필드 행 추가
    fieldKeyOrder.forEach(key => addRow(key));

    // 3. 웹뷰 템플릿 (복합 항목)
    let webviewHtml = '';
    
    // --- [핵심 수정] 템플릿 타입 항목 추가 ---
    if (data.webview_template_type) {
        webviewHtml += `• ${headerMap['webview_template_type']}: ${data.webview_template_type}<br>`;
    }
    // --- [핵심 수정] ---

    if (data.ad_type === '캐슬_노출형(라방패키지)_유튜브') {
        // 유튜브 버전
        if (data.webview_template_list) webviewHtml += `• ${headerMap['webview_template_list']}: ${data.webview_template_list}<br>`;
        if (data.webview_top_overlay_height) webviewHtml += `• ${headerMap['webview_top_overlay_height']}: ${data.webview_top_overlay_height}<br>`;
        if (data.aspect_ratio) webviewHtml += `• ${headerMap['aspect_ratio']}: ${data.aspect_ratio}`;
    } else {
        // 오리지널 라방패키지
        if (data.webview_template_list) webviewHtml += `• ${headerMap['webview_template_list']}: ${data.webview_template_list}<br>`;
        if (data.webview_top_overlay_height) webviewHtml += `• ${headerMap['webview_top_overlay_height']}: ${data.webview_top_overlay_height}<br>`;
        if (data.webview_top_overlay_color) webviewHtml += `• ${headerMap['webview_top_overlay_color']}: ${data.webview_top_overlay_color}<br>`;
        if (data.webview_bottom_overlay_height) webviewHtml += `• ${headerMap['webview_bottom_overlay_height']}: ${data.webview_bottom_overlay_height}<br>`;
        if (data.webview_bottom_overlay_color) webviewHtml += `• ${headerMap['webview_bottom_overlay_color']}: ${data.webview_bottom_overlay_color}`;
    }
    if (webviewHtml) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">웹뷰 템플릿</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${webviewHtml.replace(/<br>$/, '')}</td></tr>`;
    }

    // 4. 프리퀀시 (복합 항목)
    let freqHtml = '';
    if (data.frequency_timeboard) freqHtml += `• ${headerMap['frequency_timeboard']}: ${data.frequency_timeboard}<br>`;
    if (data.frequency_cpmc) freqHtml += `• ${headerMap['frequency_cpmc']}: ${data.frequency_cpmc}`;
    if (freqHtml) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">프리퀀시</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${freqHtml.replace(/<br>$/, '')}</td></tr>`;
    }
    
    // 5. 일물량 (복합 항목)
    let dailyVolHtml = '';
    if (data.daily_volume_timeboard) dailyVolHtml += `• ${headerMap['daily_volume_timeboard']}: ${data.daily_volume_timeboard}<br>`;
    if (data.daily_volume_cpmc) dailyVolHtml += `• ${headerMap['daily_volume_cpmc']}: ${data.daily_volume_cpmc}`;
    if (dailyVolHtml) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">일물량</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${dailyVolHtml.replace(/<br>$/, '')}</td></tr>`;
    }

    // 6. 노출 우선순위 (복합 항목)
    let priorityHtml = '';
    if (data.priority_timeboard) priorityHtml += `• ${headerMap['priority_timeboard']}: ${data.priority_timeboard}<br>`;
    if (data.priority_cpmc) priorityHtml += `• ${headerMap['priority_cpmc']}: ${data.priority_cpmc}`;
    if (priorityHtml) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">노출 우선순위</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${priorityHtml.replace(/<br>$/, '')}</td></tr>`;
    }

  } else {
    for (const key in headerMap) {
      if (data[key]) {
        body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${headerMap[key]}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${String(data[key]).replace(/\n/g, '<br>')}</td></tr>`;
      }
    }
  }
  
  body += `</table><br style="clear:both;">`;

  if (data.ad_type === '캐슬_노출형(홈앤쇼핑)' && data.deeplink_info_json) {
    try {
      const deeplinkData = JSON.parse(data.deeplink_info_json);
      if (deeplinkData.length > 0) {
        body += `<hr><h3>세부 정보 (복사 붙여넣기)</h3>`;
        const headers = ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'];
        body += `<div style="overflow-x:auto;"><table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;"><thead><tr style="background-color:#f3f3f3;">`;
        headers.forEach(h => body += `<th style="padding: 8px; border: 1px solid #e0e0e0; white-space: nowrap;">${h}</th>`);
        body += `</tr></thead><tbody>`;
        deeplinkData.forEach(row => {
          body += `<tr>`;
          headers.forEach(h => body += `<td style="padding: 8px; border: 1px solid #e0e0e0; white-space: nowrap;">${row[h] || ''}</td>`);
          body += `</tr>`;
        });
        body += `</tbody></table></div><br style="clear:both;">`;
      }
    } catch (e) {
      console.error('deeplink_info_json 파싱 실패: ' + e.toString());
      body += `<hr><h3>세부 정보 (복사 붙여넣기)</h3><p>데이터를 테이블로 변환하는데 실패했습니다.</p>`;
    }
  }

  const exposureTypes = ['캐슬_노출형(기본)', '캐슬_노출형(오토뷰)', '캐슬_노출형(맥스뷰)'];
  
  if (exposureTypes.includes(data.ad_type)) {
    body += `<hr><h3>기본 노출형 설정</h3>`;

    if (exposureSettings && exposureSettings.length > 0) {
      body += `<table align="left" cellpadding="8" style="border-collapse: collapse; border: 1px solid #e0e0e0; font-size: 12px; font-family: sans-serif;">
                 <thead><tr style="background-color:#f3f3f3;">
                   <th style="padding: 8px; border: 1px solid #e0e0e0;">슬랏</th><th style="padding: 8px; border: 1px solid #e0e0e0;">캠페인명</th>
                   <th style="padding: 8px; border: 1px solid #e0e0e0;">광고주 어드민 ID</th><th style="padding: 8px; border: 1px solid #e0e0e0;">라이브 시작/종료</th>
                   <th style="padding: 8px; border: 1px solid #e0e0e0;">데일리 라이브 시간</th><th style="padding: 8px; border: 1px solid #e0e0e0;">일물량</th>
                   <th style="padding: 8px; border: 1px solid #e0e0e0;">URL</th>
                 </tr></thead><tbody>`;
      exposureSettings.forEach(s => {
        body += `<tr>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['슬랏']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['캠페인명']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['광고주 어드민 ID']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['라이브 시작 일시']} ~ ${s['라이브 종료 일시']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['데일리 라이브 시작 시간']} ~ ${s['데일리 라이브 종료 시간']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['일물량']}</td>
                   <td style="padding: 8px; border: 1px solid #e0e0e0;">${s['URL']}</td>
                 </tr>`;
      });
      body += `</tbody></table>`;
    } else {
      body += `<p>입력된 항목이 없습니다.</p>`;
    }
  }

  const mailOptions = {
    htmlBody: body,
    cc: data.ccRecipients
  };
  if (fileAttachment) {
    mailOptions.attachments = [fileAttachment];
  }

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', mailOptions);
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ 'text': subject }) });
  } catch (e) {
    console.error(`캐시슬라이드 등록 슬랙 발송 실패 (ID: ${uniqueId}): ${e.toString()}`);
  }

  Utilities.sleep(2000);
  const threads = GmailApp.search(`subject:"${subject}" in:sent`, 0, 1);
  return threads.length > 0 ? threads[0].getMessages()[0].getId() : null;
}



/**
 * 캐시슬라이드 광고 ID를 기반으로 해당 행의 정보를 찾습니다.
 * @param {string} adId - 찾을 캐시슬라이드 고유 ID.
 * @returns {object|null} - 찾은 경우 {sheet, rowIndex, headers, rowData}, 못 찾은 경우 null.
 */
function findCashslideRowById(adId) {
  const sheetName = "캐시슬라이드 광고";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  const idColumn = sheet.getRange('A:A');
  const textFinder = idColumn.createTextFinder(adId).matchEntireCell(true);
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
 * 캐시슬라이드 광고 ID를 기반으로 해당 행의 데이터를 객체 형태로 가져옵니다.
 * @param {string} adId - 조회할 캐시슬라이드 광고 ID.
 * @returns {object|null} - 찾은 데이터 객체 또는 null.
 */
function getCashslideAdDataById(adId) {
  const found = findCashslideRowById(adId);
  if (found) {
    const adData = {};
    found.headers.forEach((header, index) => {
      let value = found.rowData[index];
      if (value instanceof Date) {
        try {
          value = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm");
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


/**
 * 캐시슬라이드 광고 등록 요청을 스킵 처리합니다.
 * @param {string} adId - 스킵할 캐시슬라이드 요청 ID.
 * @returns {object} 처리 결과 객체.
 */
function processCashslideSkip(adId) {
  try {
    const skipperEmail = Session.getActiveUser().getEmail();
    const found = findCashslideRowById(adId);
    if (!found) {
      return { success: false, message: `캐시슬라이드 ID(${adId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    const statusColIndex = headers.indexOf('status');
    const threadIdColIndex = headers.indexOf('mail_thread_id');
    const campaignNameIndex = headers.indexOf('campaign_name');

    // 1. 시트 상태를 '스킵처리'로 업데이트
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('스킵처리');

    const threadId = rowData[threadIdColIndex];
    const campaignName = rowData[campaignNameIndex] || adId;

    // 2. 원본 요청 메일 스레드에 스킵 처리되었음을 회신
    if (threadId) {
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.replyAll("", {
            htmlBody: `<p>안녕하세요,</p><p>요청하신 <b>캐시슬라이드 ID: ${adId}</b> 건이 <b>스킵 처리</b>되었음을 알려드립니다.</p><p>감사합니다.</p><p>- 처리자: ${skipperEmail}</p>`,
          });
        }
      } catch (e) {
        console.error(`캐시슬라이드 스킵 알림 메일 발송 실패(ID: ${adId}): ${e.toString()}`);
      }
    }

    // 3. 슬랙으로 알림 발송
    try {
      const slackMessage = { 'text': `[캐시슬라이드 스킵 처리] ${campaignName} (ID: ${adId})` };
      const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    } catch (e) {
      console.error(`캐시슬라이드 스킵 알림 슬랙 발송 실패 (ID: ${adId}): ${e.toString()}`);
    }

    // 4. 활동 로그 기록
    logUserAction(skipperEmail, '캐시슬라이드 스킵 처리', {
      targetId: adId,
      message: `캐시슬라이드 ID '${adId}' 스킵 처리`
    });

    return { success: true, message: `캐시슬라이드 ID(${adId})가 성공적으로 스킵 처리되었습니다.` };
  } catch (e) {
    console.error(`Error in processCashslideSkip: ${e.toString()}`);
    return { success: false, message: '캐시슬라이드 스킵 처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}


/**
 * 캐시슬라이드 광고 등록 요청을 반려 처리합니다.
 * @param {string} adId - 반려할 캐시슬라이드 요청 ID.
 * @param {string} reason - 반려 사유.
 * @returns {object} 처리 결과 객체.
 */
function processCashslideRejection(adId, reason) {
  try {
    const rejectorEmail = Session.getActiveUser().getEmail();
    const found = findCashslideRowById(adId);
    if (!found) {
      return { success: false, message: `캐시슬라이드 ID(${adId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    const statusColIndex = headers.indexOf('status');
    const rejectionDateColIndex = headers.indexOf('rejection_timestamp');
    const rejectionReasonColIndex = headers.indexOf('rejection_reason');
    const registrantColIndex = headers.indexOf('registrant');
    const threadIdColIndex = headers.indexOf('mail_thread_id');
    const campaignNameIndex = headers.indexOf('campaign_name');

    // 1. 시트 상태 업데이트
    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('반려');
    if (rejectionDateColIndex > -1) sheet.getRange(rowIndex, rejectionDateColIndex + 1).setValue(timestamp);
    if (rejectionReasonColIndex > -1) sheet.getRange(rowIndex, rejectionReasonColIndex + 1).setValue(reason);

    const registrantEmail = rowData[registrantColIndex];
    const threadId = rowData[threadIdColIndex];
    const campaignName = rowData[campaignNameIndex] || adId;

    // 2. 메일 알림 발송 (전체 회신)
    if (registrantEmail && threadId) {
      const subject = `[광고 등록 시스템] 요청하신 캐시슬라이드 광고(ID: ${adId})가 반려되었습니다.`;
      let emailBody = `<p>안녕하세요, ${registrantEmail.split('@')[0]}님.</p>
                       <p>요청하신 캐시슬라이드 광고(ID: <b>${adId}</b>)가 아래와 같은 사유로 반려되었습니다.</p>`;
      if (reason) {
        emailBody += `<p style="margin-top:20px;"><b>반려 사유:</b></p>
                      <div style="padding: 12px; border: 1px solid #ddd; background-color: #f9f9f9; border-radius: 5px; margin-top: 5px;">
                        ${reason.replace(/\n/g, '<br>')}
                      </div>`;
      }
      emailBody += `<p style="margin-top:20px;">수정 후 재요청하시거나 담당자(${rejectorEmail})에게 문의해주세요.</p>
                    <p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p>
                    <p>감사합니다.</p>`;
    try {
        const searchQuery = `"${adId}"`;
        const threads = GmailApp.search(searchQuery, 0, 1);

        if (threads && threads.length > 0) {
          threads[0].replyAll('', { htmlBody: emailBody, cc: registrantEmail });
        } else {
          console.warn(`캐시슬라이드 반려: 스레드 찾기 실패(${adId}). 새 메일 발송.`);
          GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
        }
      } catch (e) {
         console.error(`캐시슬라이드 반려 메일 발송 실패 (ID: ${adId}): ${e.toString()}`);
         // 최후의 수단으로 새 메일 발송 시도
         GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
      }
      // ▲▲▲ [수정] ▲▲▲
    }

    // 3. 슬랙 알림 발송
    try {
      const slackMessage = { 'text': `[캐시슬라이드 등록 반려] - ${campaignName} (ID: ${adId})` };
      const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(slackMessage) };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    } catch (e) {
      console.error(`캐시슬라이드 반려 슬랙 발송 실패 (ID: ${adId}): ${e.toString()}`);
    }

    // 4. 활동 로그 기록
    logUserAction(rejectorEmail, '캐시슬라이드 반려 처리', {
      targetId: adId,
      message: `캐시슬라이드 ID '${adId}' 반려 처리. 사유: ${reason}`
    });

    return { success: true, message: `캐시슬라이드 ID(${adId})가 성공적으로 반려 처리되었습니다.` };
  } catch (e) {
    console.error(`Error in processCashslideRejection: ${e.toString()}`);
    return { success: false, message: '캐시슬라이드 반려 처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}


/**
 * 날짜 문자열을 yy/MM/dd 형식으로 변환합니다.
 * @param {string} dateStr - 변환할 날짜 문자열.
 * @returns {string} 변환된 날짜 문자열 또는 '무제한'.
 */
function formatDate_yyMMdd(dateStr) {
  if (!dateStr || String(dateStr).toLowerCase() === '무제한') {
    return '무제한';
  }
  try {
    let normalizedDateStr = String(dateStr).split(' ')[0]; // 시간 정보 제거
    normalizedDateStr = normalizedDateStr.replace(/\//g, '-'); // '/'를 '-'로 변경
    
    const parts = normalizedDateStr.split('-');
    if (parts.length === 3) {
      if (parts[0].length === 2) {
        parts[0] = '20' + parts[0]; // yy-mm-dd 형식을 20yy-mm-dd로 변경
      }
      normalizedDateStr = parts.join('-');
    }

    const date = new Date(normalizedDateStr);
    if (isNaN(date.getTime())) {
        return '';
    }

    return Utilities.formatDate(date, "Asia/Seoul", "yy/MM/dd");
  } catch (e) {
    return '';
  }
}


/**
 * 캐시슬라이드 광고 요청 건에 대한 담당자를 지정하고 알림을 보냅니다.
 * @param {string} csId - 캐시슬라이드 요청 ID.
 * @param {string} approverEmail - 담당자 이메일.
 * @returns {string} 결과 메시지.
 */
function recordCashslideConfirmation(csId, approverEmail) {
  const found = findCashslideRowById(csId);
  if (!found) return `캐시슬라이드 ID: ${csId} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const threadIdColIndex = headers.indexOf('mail_thread_id');

  const currentStatus = rowData[statusColIndex];
  if (currentStatus === '스킵처리') {
    return `처리 실패: 이 CS 요청 건(ID: ${csId})은 이미 스킵 처리되어 담당자로 지정할 수 없습니다.`;
  }

  const currentManager = rowData[managerColIndex];
  if (currentManager && currentManager !== '') {
    return `처리 실패: 이 캐시슬라이드 건(ID: ${csId})은 이미 ${currentManager} 님이 담당하고 있습니다.`;
  }
  
  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  
  const timestampColIndex = headers.indexOf('manager_timestamp');
  if (timestampColIndex > -1) {
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(formattedTimestamp);
  }

try {
    // 제목이나 본문에 해당 ID가 포함된 보낸 메일함 검색
    const searchQuery = `"${csId}"`; 
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads && threads.length > 0) {
      threads[0].replyAll("", { 
        htmlBody: `<p>안녕하세요,</p><p><b>${approverEmail}</b> 님이 <b>캐시슬라이드 ID: ${csId}</b> 건의 담당자로 지정되어 등록을 진행합니다.</p><p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p><p>감사합니다.</p>`
      });
    } else {
      console.log(`캐시슬라이드 담당자 지정 알림 실패: ${csId} 관련 메일을 찾을 수 없습니다.`);
    }
  } catch (e) {
    console.error(`캐시슬라이드 담당자 지정 알림 발송 중 오류: ${e.toString()}`);
  }
  // ▲▲▲ [수정] ▲▲▲
  
  return `캐시슬라이드 ID: ${csId} 건의 담당자로 ${approverEmail}님이 지정되었습니다. 이 창은 닫아도 됩니다.`;
}

/**
 * 캐시슬라이드 광고 등록 완료 처리를 수행합니다.
 * @param {string} registrationId - 캐시슬라이드 요청 ID.
 * @param {string} adId - 등록된 실제 광고 ID.
 * @param {string} completerEmail - 완료 처리자 이메일.
 * @returns {object} 결과 객체 {success, message}.
 */
function processCashslideCompletion(registrationId, adId, completerEmail) {
  const found = findCashslideRowById(registrationId);
  if (!found) return { success: false, message: `캐시슬라이드 ID(${registrationId})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  
  if (rowData[statusColIndex] === '등록 완료') {
    return { success: false, message: `이미 등록 완료 처리된 건입니다. (ID: ${registrationId})` };
  }
  
  const adIdColIndex = headers.indexOf('ad_id');
  const completionDateColIndex = headers.indexOf('completion_timestamp');
  const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  sheet.getRange(rowIndex, statusColIndex + 1).setValue('등록 완료');
  sheet.getRange(rowIndex, adIdColIndex + 1).setValue("'" + adId);
  if (completionDateColIndex > -1) {
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(timestamp);
  }

  logUserAction(completerEmail, '캐시슬라이드 등록 완료 처리', {
    targetId: registrationId,
    message: `캐시슬라이드 ID '${registrationId}'를 광고 ID '${adId}'로 등록 완료 처리`
  });

  return { success: true, message: `캐시슬라이드 건(ID: ${registrationId})이 성공적으로 완료 처리되었습니다. 이 창은 닫아도 됩니다.` };
}



/**
 * 캐시슬라이드 수정 요청 데이터를 시트에 저장하고 알림을 보냅니다.
 */
function submitCashslideModificationRequest(formData) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheetName = "캐시슬라이드 수정";
    let sheet = ss.getSheetByName(sheetName);

    const masterHeaders = [
      'id', 'timestamp', 'registrant', 'status', 'manager', 'manager_timestamp', 'mail_thread_id', 'rejection_timestamp', 'rejection_reason', 'completion_timestamp',
      'ad_id', 'ad_name', 'request_details', 'advertiser_account_id', 'live_start_date', 'live_end_date', 'daily_live_start_time', 'daily_live_end_time',
      'creative_path', 'daily_volume', 'landing_url', 'demo_target_1', 'demo_target_2', 'retarget_cluster',
      'app_package_retargeting', 'app_package_detargeting', 'lbs_latitude', 'lbs_longitude', 'lbs_radius',
      'frequency_criteria', 'frequency_type', 'daily_frequency_limit', 'max_frequency_limit', 'tracker', 'priority', 'seg_filter'
    ];
    
    // 헤더와 폼 데이터 키 매핑
    const headerMap = {
      'id': null, 'timestamp': null, 'registrant': null, 'status': null, 'manager': null, 'manager_timestamp': null, 'mail_thread_id': null, 'rejection_timestamp': null, 'rejection_reason': null, 'completion_timestamp': null,
      'ad_id': '광고ID', 'ad_name': '광고명', 'request_details': '요청사항', 'advertiser_account_id': '광고주 계정 ID', 
      'live_start_date': '라이브 시작일시', 'live_end_date': '라이브 종료일시', 'daily_live_start_time': '데일리 라이브 시작 시간', 'daily_live_end_time': '데일리 라이브 종료 시간',
      'creative_path': '소재경로', 'daily_volume': '일물량', 'landing_url': '랜딩 URL', 'demo_target_1': '데모타겟1', 'demo_target_2': '데모타겟2', 
      'retarget_cluster': '리타겟 클러스터', 'app_package_retargeting': '앱 패키지명 - 리타겟팅', 'app_package_detargeting': '앱 패키지명 - 디타겟팅', 
      'lbs_latitude': 'LBS 타겟팅 - 위도', 'lbs_longitude': 'LBS 타겟팅 - 경도', 'lbs_radius': 'LBS 타겟팅 - 범위',
      'frequency_criteria': '프리퀀시 기준', 'frequency_type': '프리퀀시 타입', 'daily_frequency_limit': '일일 프리퀀시 제한 횟수', 
      'max_frequency_limit': '최대 프리퀀시 제한 횟수', 'tracker': '트래커', 'priority': '노출 우선순위', 'seg_filter': 'Seg Filter'
    };

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(masterHeaders);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold").setFrozenRows(1);
    }

    const idPrefix = `cs-${userName}-mod-`;
    const nextId = getNextSequentialId(sheet, idPrefix);
    const uniqueId = idPrefix + nextId;
    
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    const adName = formData['광고명'] ? formData['광고명'].split('\n')[0] : '이름 없음';
    const subject = `[캐시슬라이드 수정 요청] [캐슬_잠금화면] ${adName} 수정 (ID: ${uniqueId})`;
    
    const messageId = sendCashslideModificationNotification(userEmail, uniqueId, subject, formData);
    
    const dataToSave = {};
    for (const header of masterHeaders) {
      const formKey = headerMap[header];
      if (formKey && formData[formKey]) {
        dataToSave[header] = formData[formKey];
      }
    }
    
    dataToSave.id = uniqueId;
    dataToSave.timestamp = formattedTimestamp;
    dataToSave.registrant = userEmail;
    dataToSave.status = '수정 요청 완료';
    dataToSave.mail_thread_id = messageId;
    
    const newRow = masterHeaders.map(header => dataToSave[header] || '');
    sheet.appendRow(newRow);

    logUserAction(userEmail, '캐시슬라이드 수정 요청', { targetId: uniqueId, message: `광고 '${adName}' 수정 요청` });
    
    return { success: true, message: `캐시슬라이드 수정 요청이 완료되었습니다. (ID: ${uniqueId})` };

  } catch (e) {
    console.error(`submitCashslideModificationRequest Error: ${e.toString()}`);
    return { success: false, message: `처리 중 오류가 발생했습니다: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 캐시슬라이드 수정 요청 알림(이메일, 슬랙)을 보냅니다.
 */
function sendCashslideModificationNotification(senderEmail, modId, subject, data) {
  const confirmationUrl = `${ScriptApp.getService().getUrl()}?action=confirm_cs_mod&id=${modId}`;
  const completionUrl = `${ScriptApp.getService().getUrl()}?action=complete_cs_mod&id=${modId}`;
  const ccEmails = data.ccRecipients || '';

  let body = `<p>안녕하세요, 운영팀.</p>
              <p><b>${senderEmail}</b>님께서 캐시슬라이드 광고 수정을 요청했습니다.</p>
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
    '광고ID', '광고명', '요청사항', '광고주 계정 ID', '라이브 시작일시', '라이브 종료일시', '데일리 라이브 시작 시간', '데일리 라이브 종료 시간',
    '소재경로', '일물량', '랜딩 URL', '데모타겟1', '데모타겟2', '리타겟 클러스터',
    '앱 패키지명 - 리타겟팅', '앱 패키지명 - 디타겟팅', 'LBS 타겟팅 - 위도', 'LBS 타겟팅 - 경도', 'LBS 타겟팅 - 범위',
    '프리퀀시 기준', '프리퀀시 타입', '일일 프리퀀시 제한 횟수', '최대 프리퀀시 제한 횟수', '트래커', '노출 우선순위', 'Seg Filter'
  ];

  fieldOrder.forEach(key => {
    if (data[key]) {
      const displayValue = String(data[key]).replace(/\n/g, '<br>');
      body += `<tr><td style="padding: 8px; border: 1px solid #e0e0e0; background-color: #f9f9f9; font-weight: bold; white-space: nowrap;">${key}</td><td style="padding: 8px; border: 1px solid #e0e0e0;">${displayValue}</td></tr>`;
    }
  });
  body += `</table>`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', { htmlBody: body, cc: ccEmails });
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ 'text': subject }) });
  } catch (e) {
    console.error(`캐시슬라이드 수정 요청 슬랙 발송 실패 (ID: ${modId}): ${e.toString()}`);
  }

  Utilities.sleep(2000);
  const threads = GmailApp.search(`subject:"${subject}" in:sent`, 0, 1);
  if (threads.length > 0) {
    return threads[0].getMessages()[0].getId();
  }
  return null;
}

/**
 * 캐시슬라이드 수정 ID로 행 정보를 찾습니다.
 */
function findCashslideModRowById(modId) {
  const sheet = ss.getSheetByName("캐시슬라이드 수정");
  if (!sheet) return null;
  const textFinder = sheet.getRange('A:A').createTextFinder(modId).matchEntireCell(true);
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
 * 캐시슬라이드 수정 담당자를 지정합니다.
 */
function recordCashslideModificationConfirmation(modId, approverEmail) {
  const found = findCashslideModRowById(modId);
  if (!found) return `캐시슬라이드 수정 ID: ${modId} 건을 찾을 수 없습니다.`;
  
  const { sheet, rowIndex, headers, rowData } = found;
  const managerColIndex = headers.indexOf('manager');
  const statusColIndex = headers.indexOf('status');
  const threadIdColIndex = headers.indexOf('mail_thread_id');

  if (rowData[managerColIndex]) {
    return `처리 실패: 이 건(ID: ${modId})은 이미 ${rowData[managerColIndex]} 님이 담당하고 있습니다.`;
  }
  
  sheet.getRange(rowIndex, managerColIndex + 1).setValue(approverEmail);
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('처리중');
  
  const timestampColIndex = headers.indexOf('manager_timestamp');
  if (timestampColIndex > -1) {
    // ▼▼▼ [수정] 날짜 형식을 지정하여 저장합니다. ▼▼▼
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, timestampColIndex + 1).setValue(formattedTimestamp);
    // ▲▲▲ [수정] ▲▲▲
  }

  const threadId = rowData[threadIdColIndex];
  if (threadId) {
    try {
      GmailApp.getThreadById(threadId).replyAll("", { 
        htmlBody: `<p><b>${approverEmail}</b> 님이 <b>캐시슬라이드 수정 ID: ${modId}</b> 건의 담당자로 지정되었습니다.</p><p><a href="${SYSTEM_URL}">시스템 바로가기</a></p>`
      });
    } catch (e) {
      console.error(`캐시슬라이드 수정 담당자 지정 알림 실패: ${e.toString()}`);
    }
  }
  
  return `ID: ${modId} 건의 담당자로 ${approverEmail}님이 지정되었습니다.`;
}

/**
 * 캐시슬라이드 수정 완료 처리를 합니다.
 */
function processCashslideModificationCompletion(modId, completerEmail) {
  const found = findCashslideModRowById(modId);
  if (!found) return { success: false, message: `캐시슬라이드 수정 ID(${modId})를 찾을 수 없습니다.` };

  const { sheet, rowIndex, headers, rowData } = found;
  const statusColIndex = headers.indexOf('status');
  
  if (rowData[statusColIndex] === '수정 완료') {
    return { success: false, message: `이미 수정 완료 처리된 건입니다. (ID: ${modId})` };
  }
  
  const completionDateColIndex = headers.indexOf('completion_timestamp');
  sheet.getRange(rowIndex, statusColIndex + 1).setValue('수정 완료');
  if (completionDateColIndex > -1) {
    // ▼▼▼ [수정] 날짜 형식을 지정하여 저장합니다. ▼▼▼
    const formattedTimestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, completionDateColIndex + 1).setValue(formattedTimestamp);
    // ▲▲▲ [수정] ▲▲▲
  }

  logUserAction(completerEmail, '캐시슬라이드 수정 완료', { targetId: modId });
  return { success: true, message: `캐시슬라이드 수정 건(ID: ${modId})이 성공적으로 완료 처리되었습니다.` };
}

/**
 * 캐시슬라이드 수정 ID를 기반으로 해당 행의 데이터를 객체 형태로 가져옵니다.
 * @param {string} modId - 조회할 캐시슬라이드 수정 요청 ID.
 * @returns {object|null} - 찾은 데이터 객체 또는 null.
 */
function getCashslideModificationDataById(modId) {
  const found = findCashslideModRowById(modId);
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
 * 캐시슬라이드 수정 요청을 스킵 처리합니다.
 * @param {string} modId - 스킵할 캐시슬라이드 수정 요청 ID.
 * @returns {object} 처리 결과 객체.
 */
function processCashslideModificationSkip(modId) {
  try {
    const skipperEmail = Session.getActiveUser().getEmail();
    const found = findCashslideModRowById(modId);
    if (!found) {
      return { success: false, message: `캐시슬라이드 수정 ID(${modId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    const statusColIndex = headers.indexOf('status');
    const threadIdColIndex = headers.indexOf('mail_thread_id');
    const adNameIndex = headers.indexOf('ad_name');
    // ▼▼▼ [추가] 등록자 이메일 컬럼 인덱스 찾기 ▼▼▼
    const registrantColIndex = headers.indexOf('registrant'); 

    // 1. 시트 상태 업데이트
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('스킵처리');

    const threadId = rowData[threadIdColIndex];
    const adName = String(rowData[adNameIndex] || modId).split('\n')[0];
    const registrantEmail = rowData[registrantColIndex]; // 등록자 이메일 가져오기

    // 2. 메일 알림 발송 (검색 로직 포함하여 강화)
    if (registrantEmail) {
      const subject = `[광고 등록 시스템] 요청하신 캐시슬라이드 수정(ID: ${modId})이 스킵 처리되었습니다.`;
      const emailBody = `<p>안녕하세요,</p>
                         <p>요청하신 <b>캐시슬라이드 수정 ID: ${modId}</b> 건이 <b>스킵 처리</b>되었음을 알려드립니다.</p>
                         <p>감사합니다.</p>
                         <p>- 처리자: ${skipperEmail}</p>`;
      
      let thread = null;

      // (1) 저장된 ID로 스레드 조회 시도
      if (threadId) {
        try {
          thread = GmailApp.getThreadById(threadId);
        } catch (e) {
          console.warn(`저장된 ID(${threadId}) 조회 실패. 검색으로 전환합니다.`);
        }
      }

      // (2) 실패 시 제목(ID)으로 검색 시도
      if (!thread) {
        try {
          const threads = GmailApp.search(`"${modId}"`, 0, 1);
          if (threads.length > 0) {
            thread = threads[0];
          }
        } catch (e) {
          console.warn(`제목 검색 실패: ${e.toString()}`);
        }
      }

      // (3) 답장 또는 새 메일 발송
      if (thread) {
        try {
          thread.replyAll("", { htmlBody: emailBody });
        } catch (e) {
          console.error(`답장 실패, 새 메일 발송: ${e.toString()}`);
          GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
        }
      } else {
        console.log('스레드 찾기 실패, 새 메일 발송');
        GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
      }
    }

    // 3. 슬랙 알림 발송
    try {
      const slackMessage = { 'text': `[캐시슬라이드 수정 스킵] ${adName} (ID: ${modId})` };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(slackMessage) });
    } catch (e) {
      console.error(`캐시슬라이드 수정 스킵 알림 슬랙 발송 실패 (ID: ${modId}): ${e.toString()}`);
    }

    logUserAction(skipperEmail, '캐시슬라이드 수정 스킵', { targetId: modId });
    return { success: true, message: `캐시슬라이드 수정 ID(${modId})가 성공적으로 스킵 처리되었습니다.` };
  } catch (e) {
    console.error(`Error in processCashslideModificationSkip: ${e.toString()}`);
    return { success: false, message: '처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}

/**
 * 캐시슬라이드 수정 요청을 반려 처리합니다.
 * @param {string} modId - 반려할 캐시슬라이드 수정 요청 ID.
 * @param {string} reason - 반려 사유.
 * @returns {object} 처리 결과 객체.
 */
function processCashslideModificationRejection(modId, reason) {
  try {
    const rejectorEmail = Session.getActiveUser().getEmail();
    const found = findCashslideModRowById(modId);
    if (!found) {
      return { success: false, message: `캐시슬라이드 수정 ID(${modId})를 찾을 수 없습니다.` };
    }

    const { sheet, rowIndex, headers, rowData } = found;

    const statusColIndex = headers.indexOf('status');
    const rejectionDateColIndex = headers.indexOf('rejection_timestamp');
    const rejectionReasonColIndex = headers.indexOf('rejection_reason');
    const registrantColIndex = headers.indexOf('registrant');
    const threadIdColIndex = headers.indexOf('mail_thread_id');
    const adNameIndex = headers.indexOf('ad_name');

    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(rowIndex, statusColIndex + 1).setValue('반려');
    if (rejectionDateColIndex > -1) sheet.getRange(rowIndex, rejectionDateColIndex + 1).setValue(timestamp);
    if (rejectionReasonColIndex > -1) sheet.getRange(rowIndex, rejectionReasonColIndex + 1).setValue(reason);

    const registrantEmail = rowData[registrantColIndex];
    const threadId = rowData[threadIdColIndex];
    const adName = String(rowData[adNameIndex] || modId).split('\n')[0];

    if (registrantEmail) {
      const subject = `[광고 등록 시스템] 요청하신 캐시슬라이드 수정(ID: ${modId})이 반려되었습니다.`;
      let emailBody = `<p>안녕하세요, ${registrantEmail.split('@')[0]}님.</p>
                       <p>요청하신 캐시슬라이드 수정(ID: <b>${modId}</b>)이 아래 사유로 반려되었습니다.</p>`;
      if (reason) {
        emailBody += `<p style="margin-top:20px;"><b>반려 사유:</b></p>
                      <div style="padding: 12px; border: 1px solid #ddd; background-color: #f9f9f9;">${reason.replace(/\n/g, '<br>')}</div>`;
      }
      emailBody += `<p style="margin-top:20px;">수정 후 재요청하시거나 담당자(${rejectorEmail})에게 문의해주세요.</p>
                    <p><a href="${SYSTEM_URL}">광고 등록 요청 시스템 바로가기</a></p>`;

      let thread = null;

      // 1. 저장된 ID로 스레드 조회 시도
      if (threadId) {
        try {
          thread = GmailApp.getThreadById(threadId);
        } catch (e) {
          console.warn(`저장된 ID(${threadId})는 스레드 ID가 아닙니다. 검색으로 전환합니다.`);
        }
      }

      // 2. 실패 시 제목(ID)으로 검색 시도
      if (!thread) {
        try {
          // 수정 ID가 포함된 보낸 편지함 검색
          const threads = GmailApp.search(`"${modId}"`, 0, 1);
          if (threads.length > 0) {
            thread = threads[0];
          }
        } catch (e) {
          console.warn(`제목 검색 실패: ${e.toString()}`);
        }
      }

      // 3. 발송 (답장 또는 새 메일)
      if (thread) {
        try {
          thread.replyAll('', { htmlBody: emailBody, cc: registrantEmail });
        } catch (e) {
          console.error(`답장 실패, 새 메일 발송: ${e.toString()}`);
          GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
        }
      } else {
        console.log('스레드 찾기 실패, 새 메일 발송');
        GmailApp.sendEmail(registrantEmail, subject, '', { htmlBody: emailBody });
      }
    }

    try {
      const slackMessage = { 'text': `[캐시슬라이드 수정 반려] ${adName} (ID: ${modId})` };
      UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(slackMessage) });
    } catch (e) {
      console.error(`캐시슬라이드 수정 반려 슬랙 발송 실패 (ID: ${modId}): ${e.toString()}`);
    }

    logUserAction(rejectorEmail, '캐시슬라이드 수정 반려', { targetId: modId, message: `사유: ${reason}` });
    return { success: true, message: `캐시슬라이드 수정 ID(${modId})가 성공적으로 반려 처리되었습니다.` };
  } catch (e) {
    console.error(`Error in processCashslideModificationRejection: ${e.toString()}`);
    return { success: false, message: '처리 중 오류가 발생했습니다: ' + e.toString() };
  }
}

function getCashslideFields() {
  const adTypeOptions = ['CPC', 'CPMC', 'CPM'];
  const webviewOptions = ['웹뷰 미진행', '웹뷰 진행'];
  const demoTarget1Options = ['전체', '남성', '여성'];
  const tagOptions = ['FashionMF', 'FashionM', 'FashionF', 'Beauty/Cosmetics', 'Luxury goods', 'Cosmetic surgery', 'Shopping', 'Furniture/Interior', 'Grocery/Consumables', 'Electronics', 'Pet Care', 'Living Info', 'Cooking/Recipes', 'Coffee/Dessert/Drinks', 'Liquor', 'Delivery/Takeout', 'Restaurant', 'Baby (0 ~ 3 yrs.)', 'Kids (4 ~ 13 yrs.)', 'Language', 'Middle & High School/SAT', 'License/Transfer', 'Job Search', 'Media & Entertainment', 'Movies', 'Music', 'Webcomic/webfiction', 'Books', 'Tickets', 'Other hobbies', 'Action', 'Arcade/Shooting/Racing', 'Board/Gostop/Poker', 'PC/Online Game', 'Puzzle', 'RPG', 'Simulation/Casual', 'Sports', 'Strategy/TCG', 'Smartphone', 'High Tech Product', 'Flights', 'Hotels', 'Car Rental', 'Travel Info', 'Credit Card', 'Finance/Stock/Loan/Insurance', 'Real Estate', 'Medical', 'Fitness/Diet/Healthcare', 'Sports General', 'Sports Apparel', 'Golf', 'Leisure', 'Auto', 'Date/Love', 'Event', 'Wedding', 'Current Affairs', 'Donation', 'Snack Culture'];
  const frequencyTypeOptions = ['노출', '클릭', '노출 & 클릭'];
  const trackerOptions = ['Tune', 'Appsflyer', 'Ad brix', 'Adjust', 'Airbridge', 'Singular', 'Ad brix Remaster', 'Branch with google_ad_id'];
  
  const basicExposureFields = [
    { name: '요청사항', type: 'textarea', rows: 3, required: true, serverKey: 'request_details' },
    { name: '광고주 어드민 ID (신규 or 기존 ID)', type: 'text', required: true, serverKey: 'advertiser_admin_id' },
    { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
    { name: '캠페인', type: 'text', placeholder: '예 : 디지털 페스티벌', required: false, serverKey: 'campaign_name' },
    { name: '광고 타입', type: 'select', options: adTypeOptions, required: true, serverKey: 'ad_type_option' },
    { name: '웹뷰 진행여부', type: 'select', options: webviewOptions, required: true, serverKey: 'webview_enabled' },
    { name: '웹뷰 URL', type: 'text', required: true, serverKey: 'webview_url', dependency: { field: '웹뷰 진행여부', showsOn: '웹뷰 진행' } },
    { name: '상단 오버레이 영역 높이', type: 'text', required: true, defaultValue: '100dp', serverKey: 'webview_top_overlay_height', dependency: { field: '웹뷰 진행여부', showsOn: '웹뷰 진행' } },
    { name: '상단 오버레이 색', type: 'text', required: true, defaultValue: '#000000', serverKey: 'webview_top_overlay_color', dependency: { field: '웹뷰 진행여부', showsOn: '웹뷰 진행' } },
    { name: '하단 오버레이 영역 높이', type: 'text', required: true, defaultValue: '100dp', serverKey: 'webview_bottom_overlay_height', dependency: { field: '웹뷰 진행여부', showsOn: '웹뷰 진행' } },
    { name: '하단 오버레이 색', type: 'text', required: true, defaultValue: '#000000', serverKey: 'webview_bottom_overlay_color', dependency: { field: '웹뷰 진행여부', showsOn: '웹뷰 진행' } },
    { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
    { name: '데모타겟1', type: 'select', options: demoTarget1Options, required: false, serverKey: 'demo_target_1' },
    { name: '데모타겟2', type: 'text', placeholder: '연령 (예 : 20-39)', required: false, serverKey: 'demo_target_2' },
    { name: '리타겟 클러스터', type: 'textarea', rows: 3, required: false, serverKey: 'retarget_cluster' },
    { name: '앱 패키지명 - 리타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_retargeting' },
    { name: '앱 패키지명 - 디타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_detargeting' },
    { name: 'LBS 타겟팅 - 위도', type: 'text', required: false, serverKey: 'lbs_latitude' },
    { name: 'LBS 타겟팅 - 경도', type: 'text', required: false, serverKey: 'lbs_longitude' },
    { name: 'LBS 타겟팅 - 범위', type: 'text', required: false, serverKey: 'lbs_radius' },
    { name: '태그', type: 'multiselect_checkbox', options: tagOptions, required: true, serverKey: 'tag' },
    { name: '프리퀀시 기준', type: 'select', options: ['캠페인별', '소재별'], required: false, serverKey: 'frequency_criteria' },
    { name: '프리퀀시 타입', type: 'select', options: frequencyTypeOptions, required: true, serverKey: 'frequency_type', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
    { name: '일일 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_daily_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
    { name: '최대 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_max_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
    { name: '트래커', type: 'select', options: trackerOptions, required: false, serverKey: 'tracker' },
    { name: '노출 우선순위', type: 'text', required: true, defaultValue: '10', serverKey: 'priority' },
    { name: 'Seg Filter', type: 'text', required: false, serverKey: 'seg_filter' },
    { name: '기본 노출형', type: 'dynamic_table', serverKey: 'exposure_settings' }
  ];

  const maxviewUniqueFields = [
    { name: '커버 템플릿 설정', type: 'select', options: ['Teaser 형', 'Native 형'], required: true, serverKey: 'cover_template_setting' },
    // -- Native 형 전용 필드 --
    { name: '커버) 광고 타이틀 1번째 줄', type: 'textarea_limited', rows: 2, required: true, maxLength: 13, placeholder: '최대 13자 입력 가능, 띄어쓰기 포함. 특수문자 입력 불가', serverKey: 'cover_title_1', dependency: { field: '커버 템플릿 설정', showsOn: 'Native 형' } },
    { name: '커버) 광고 타이틀 2번째 줄', type: 'textarea_limited', rows: 2, required: true, maxLength: 13, placeholder: '최대 13자 입력 가능, 띄어쓰기 포함. 특수문자 입력 불가', serverKey: 'cover_title_2', dependency: { field: '커버 템플릿 설정', showsOn: 'Native 형' } },
    { name: '커버) 광고주 타이틀', type: 'textarea_limited', rows: 2, required: true, maxLength: 16, placeholder: '최대 16자 입력 가능, 띄어쓰기 포함', serverKey: 'cover_advertiser_title', dependency: { field: '커버 템플릿 설정', showsOn: 'Native 형' } },
    // -- 공통 하위 필드 --
    { name: '선택형 화면비율', type: 'select', options: ['Full Screen(영상 비율에 상관 없이 Full Screen)', '16:9 (가로영상)', '9:16 (세로영상)', '4:3 (가로영상)', '1:1 (정방형 영상)'], required: true, defaultValue: 'Full Screen(영상 비율에 상관 없이 Full Screen)', serverKey: 'aspect_ratio', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
    { name: '동적소재 반복재생 여부', type: 'select', options: ['NO', 'YES'], required: true, defaultValue: 'NO', serverKey: 'dynamic_creative_loop', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
    { name: '첫 슬라이드시 자동 재생 여부', type: 'select', options: ['YES', 'NO'], required: true, defaultValue: 'YES', serverKey: 'autoplay_on_first_slide', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
    { name: '동영상 스트리밍 여부', type: 'select', options: ['NO', 'YES'], required: true, defaultValue: 'NO', serverKey: 'video_streaming', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
    { name: '여백허용', type: 'select', options: ['NO', 'YES'], required: true, defaultValue: 'NO', serverKey: 'allow_margin', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
    { name: '동영상 재생중간 적립 설정', type: 'select', options: ['미설정', '1회 설정', '2회 설정', '3회 설정', '4회 설정', '5회 설정'], required: true, defaultValue: '미설정', serverKey: 'mid_roll_reward', dependency: { field: '커버 템플릿 설정', showsOn: ['Teaser 형', 'Native 형'] } },
  ];

  const fields = {
    'Cashslide': {
      '캐슬_노출형(기본)': { fields: basicExposureFields },
      '캐슬_노출형(라방패키지)': {
        fields: [
          { name: '요청사항', type: 'textarea', rows: 3, required: true, serverKey: 'request_details' },
          { name: '광고주 어드민 ID (신규 or 기존 ID)', type: 'text', required: true, serverKey: 'advertiser_admin_id' },
          { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
          { name: '캠페인', type: 'text', placeholder: '디지털 페스티벌', required: false, serverKey: 'campaign_name' },
          { name: '라이브 시작일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '00:00', serverKey: 'live_start_date' },
          { name: '라이브 종료일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '23:59', serverKey: 'live_end_date' },
          { name: '태그', type: 'multiselect_checkbox', options: tagOptions, required: true, serverKey: 'tag' },
          { name: '랜딩 URL', type: 'textarea', rows: 5, required: true, serverKey: 'landing_url' },
          { name: '웹뷰 URL', type: 'textarea', rows: 5, required: true, serverKey: 'webview_url' }
        ]
      },
      '캐슬_노출형(라방패키지)_유튜브': {
        fields: [
          { name: '요청사항', type: 'textarea', rows: 3, required: true, serverKey: 'request_details' },
          { name: '광고주 어드민 ID (신규 or 기존 ID)', type: 'text', required: true, serverKey: 'advertiser_admin_id' },
          { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
          { name: '캠페인', type: 'text', placeholder: '디지털 페스티벌', required: false, serverKey: 'campaign_name' },
          { name: '라이브 시작일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '00:00', serverKey: 'live_start_date' },
          { name: '라이브 종료일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '23:59', serverKey: 'live_end_date' },
          { name: '태그', type: 'multiselect_checkbox', options: tagOptions, required: true, serverKey: 'tag' },
          { name: '랜딩 URL', type: 'textarea', rows: 5, required: true, serverKey: 'landing_url' },
          { name: '웹뷰 URL', type: 'textarea', rows: 5, required: true, serverKey: 'webview_url' }
        ]
      },
      '캐슬_라이브적립': { 
          fields: [
            { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
            { name: '라이브 시작일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '00:00', serverKey: 'live_start_date' },
            { name: '소재', type: 'file_upload', required: true, maxSize: 200, serverKey: 'creative_file' },
            { name: '타이틀', type: 'textarea', rows: 2, required: true, placeholder: '한 줄당 14자, 최대 2줄까지 입력 가능', serverKey: 'title' },
            { name: '상품정보', type: 'textarea', rows: 5, required: true, placeholder: '각 줄은 공백 포함 20자까지 입력 가능합니다.', serverKey: 'product_info' },
            { name: '랜딩 URL', type: 'text', required: true, serverKey: 'landing_url' },
            { name: '상세 페이지 URL', type: 'text', required: false, serverKey: 'detail_page_url' }
        ]
      },
      '캐슬_노출형(홈앤쇼핑)': { 
        fields: [
          { name: '홈앤쇼핑 광고 타입', type: 'sub_type_select', options: ['딥링크', '기획전', '프로모션', '휴면'], required: true }
        ],
        sub_forms: {
          '딥링크': [
            { name: '요청사항', type: 'textarea', rows: 4, required: true, defaultValue: '1. 리포팅 그룹 꼭 셋팅 부탁드립니다.\n2. 1차 URL에 캐슬 브릿지 링크 필수셋팅 부탁드립니다.\n3. 신규 소재만 탭분리 O / 데일리 셋팅\n4. 광고 등록하신 후 프리뷰 게재면 공유 부탁드립니다. (랜딩 지면 포함)', serverKey: 'request_details' },
            { name: '슬랏 / 우선순위', type: 'text', required: true, defaultValue: '3번 / 10', serverKey: 'slot_priority' },
            { name: '프리퀀시', type: 'text', required: true, defaultValue: '클릭/캠페인기준 3회', serverKey: 'frequency' },
            { name: '타겟팅', type: 'text', required: true, defaultValue: '20세 이상 + 홈앤쇼핑 앱 설치자', serverKey: 'targeting' },
            { name: '앱타겟팅', type: 'text', required: true, defaultValue: 'com.hnsmall', serverKey: 'app_targeting' },
            { name: '클러스터', type: 'textarea', rows: 2, required: false, serverKey: 'cluster' },
            { name: '일물량', type: 'text', required: true, defaultValue: '노출 120,000 / 클릭 60', serverKey: 'daily_volume' },
            { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
            { name: '세부 정보 (복사 붙여넣기)', type: 'pasteable_table', required: true, headers: ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'], serverKey: 'deeplink_info_json' }
          ],
          '기획전': [
            { name: '요청사항', type: 'textarea', rows: 4, required: true, defaultValue: '1. 리포팅 그룹 꼭 셋팅 부탁드립니다.\n2. 1차 URL에 캐슬 브릿지 링크 필수셋팅 부탁드립니다.\n3. 신규 소재만 탭분리 O / 데일리 셋팅\n4. 광고 등록하신 후 프리뷰 게재면 공유 부탁드립니다. (랜딩 지면 포함)', serverKey: 'request_details' },
            { name: '슬랏 / 우선순위', type: 'text', required: true, defaultValue: '3번 / 10', serverKey: 'slot_priority' },
            { name: '프리퀀시', type: 'text', required: true, defaultValue: '클릭/캠페인기준 2회', serverKey: 'frequency' },
            { name: '타겟팅', type: 'text', required: true, defaultValue: '20세 이상 + 홈앤쇼핑 앱 설치자', serverKey: 'targeting' },
            { name: '앱타겟팅', type: 'text', required: true, defaultValue: 'com.hnsmall', serverKey: 'app_targeting' },
            { name: '클러스터', type: 'textarea', rows: 2, required: false, serverKey: 'cluster' },
            { name: '일물량', type: 'text', required: true, defaultValue: '노출 100,000 / 클릭 40', serverKey: 'daily_volume' },
            { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
            { name: '세부 정보 (복사 붙여넣기)', type: 'pasteable_table', required: true, headers: ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'], serverKey: 'deeplink_info_json' }
          ],
          '프로모션': [
            { name: '요청사항', type: 'textarea', rows: 4, required: true, defaultValue: '1. 리포팅 그룹 꼭 셋팅 부탁드립니다.\n2. 1차 URL에 캐슬 브릿지 링크 필수셋팅 부탁드립니다.\n3. 신규 소재만 탭분리 O / 데일리 셋팅\n4. 광고 등록하신 후 프리뷰 게재면 공유 부탁드립니다. (랜딩 지면 포함)', serverKey: 'request_details' },
            { name: '슬랏 / 우선순위', type: 'text', required: true, defaultValue: '3번 / 10', serverKey: 'slot_priority' },
            { name: '프리퀀시', type: 'text', required: true, defaultValue: '클릭/캠페인기준 2회', serverKey: 'frequency' },
            { name: '타겟팅', type: 'text', required: true, defaultValue: '20세 이상 + 홈앤쇼핑 앱 설치자', serverKey: 'targeting' },
            { name: '앱타겟팅', type: 'text', required: true, defaultValue: 'com.hnsmall', serverKey: 'app_targeting' },
            { name: '클러스터', type: 'textarea', rows: 2, required: false, serverKey: 'cluster' },
            { name: '일물량', type: 'text', required: true, defaultValue: '노출 120,000 / 클릭 60', serverKey: 'daily_volume' },
            { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
            { name: '세부 정보 (복사 붙여넣기)', type: 'pasteable_table', required: true, headers: ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'], serverKey: 'deeplink_info_json' }
          ],
          '휴면': [
            { name: '요청사항', type: 'textarea', rows: 4, required: true, defaultValue: '1. 리포팅 그룹 꼭 셋팅 부탁드립니다.\n2. 1차 URL에 캐슬 브릿지 링크 필수셋팅 부탁드립니다.\n3. 신규 소재만 탭분리 O / 데일리 셋팅\n4. 광고 등록하신 후 프리뷰 게재면 공유 부탁드립니다. (랜딩 지면 포함)', serverKey: 'request_details' },
            { name: '슬랏 / 우선순위', type: 'text', required: true, defaultValue: '3번 / 10', serverKey: 'slot_priority' },
            { name: '프리퀀시', type: 'text', required: true, defaultValue: '클릭/캠페인기준 3회', serverKey: 'frequency' },
            { name: '타겟팅', type: 'text', required: false, serverKey: 'targeting' },
            { name: '앱타겟팅', type: 'text', required: false, serverKey: 'app_targeting' },
            { name: '클러스터', type: 'textarea', rows: 2, required: false, serverKey: 'cluster' },
            { name: '일물량', type: 'text', required: true, defaultValue: '노출 160,000 / 타겟 80', serverKey: 'daily_volume' },
            { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
            { name: '세부 정보 (복사 붙여넣기)', type: 'pasteable_table', required: true, headers: ['라이브타이틀', '어드민', '2차 랜딩 URL', 'BgImg(W,H) 사이즈 확인', '시작일자', '종료일자', '라이브 시간', '소재명'], serverKey: 'deeplink_info_json' }
          ]
        }
      },
      '캐슬_노출형(오토뷰)': {
        fields: [
          { name: '요청사항', type: 'textarea', rows: 3, required: true, serverKey: 'request_details' },
          { name: '광고주 어드민 ID (신규 or 기존 ID)', type: 'text', required: true, serverKey: 'advertiser_admin_id' },
          { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
          { name: '캠페인', type: 'text', placeholder: '디지털 페스티벌', required: false, serverKey: 'campaign_name' },
          { name: '이미지 소재경로', type: 'textarea', rows: 3, required: true, serverKey: 'creative_path_image' },
          { name: '영상 소재경로', type: 'textarea', rows: 3, required: true, serverKey: 'creative_path_video' },
          { name: '첫 슬라이드시 자동 재생 여부', type: 'select', options: ['YES', 'NO'], required: true, defaultValue: 'YES', serverKey: 'autoplay_on_first_slide' },
          { name: '동적소재 반복재생 여부', type: 'select', options: ['NO', 'YES'], required: true, defaultValue: 'NO', serverKey: 'dynamic_creative_loop' },
          { name: '데모타겟1', type: 'select', options: demoTarget1Options, required: false, serverKey: 'demo_target_1' },
          { name: '데모타겟2', type: 'text', placeholder: '연령 (예 : 20-39)', required: false, serverKey: 'demo_target_2' },
          { name: '리타겟 클러스터', type: 'textarea', rows: 3, required: false, serverKey: 'retarget_cluster' },
          { name: '앱 패키지명 - 리타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_retargeting' },
          { name: '앱 패키지명 - 디타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_detargeting' },
          { name: 'LBS 타겟팅 - 위도', type: 'text', required: false, serverKey: 'lbs_latitude' },
          { name: 'LBS 타겟팅 - 경도', type: 'text', required: false, serverKey: 'lbs_longitude' },
          { name: 'LBS 타겟팅 - 범위', type: 'text', required: false, serverKey: 'lbs_radius' },
          { name: '태그', type: 'multiselect_checkbox', options: tagOptions, required: true, serverKey: 'tag' },
          { name: '프리퀀시 기준', type: 'select', options: ['캠페인별', '소재별'], required: false, serverKey: 'frequency_criteria' },
          { name: '프리퀀시 타입', type: 'select', options: frequencyTypeOptions, required: true, serverKey: 'frequency_type', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '일일 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_daily_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '최대 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_max_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '트래커', type: 'select', options: trackerOptions, required: false, serverKey: 'tracker' },
          { name: '노출 우선순위', type: 'text', required: true, defaultValue: '10', serverKey: 'priority' },
          { name: 'Seg Filter', type: 'text', required: false, serverKey: 'seg_filter' },
          { name: '기본 노출형', type: 'dynamic_table', serverKey: 'exposure_settings' }
        ]
      },
     '캐슬_노출형(맥스뷰)': {
        fields: [
          { name: '요청사항', type: 'textarea', rows: 3, required: true, serverKey: 'request_details' },
          { name: '광고주 어드민 ID (신규 or 기존 ID)', type: 'text', required: true, serverKey: 'advertiser_admin_id' },
          { name: '광고주', type: 'text', required: true, serverKey: 'advertiser_name' },
          { name: '캠페인', type: 'text', placeholder: '디지털 페스티벌', required: false, serverKey: 'campaign_name' },
          { name: '소재경로', type: 'textarea', rows: 5, required: true, serverKey: 'creative_path' },
          ...maxviewUniqueFields, // 위에서 정의한 맥스뷰 전용 필드들 삽입
          { name: '데모타겟1', type: 'select', options: demoTarget1Options, required: false, serverKey: 'demo_target_1' },
          { name: '데모타겟2', type: 'text', placeholder: '연령 (예 : 20-39)', required: false, serverKey: 'demo_target_2' },
          { name: '리타겟 클러스터', type: 'textarea', rows: 3, required: false, serverKey: 'retarget_cluster' },
          { name: '앱 패키지명 - 리타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_retargeting' },
          { name: '앱 패키지명 - 디타겟팅', type: 'textarea', rows: 3, required: false, serverKey: 'app_package_detargeting' },
          { name: 'LBS 타겟팅 - 위도', type: 'text', required: false, serverKey: 'lbs_latitude' },
          { name: 'LBS 타겟팅 - 경도', type: 'text', required: false, serverKey: 'lbs_longitude' },
          { name: 'LBS 타겟팅 - 범위', type: 'text', required: false, serverKey: 'lbs_radius' },
          { name: '태그', type: 'multiselect_checkbox', options: tagOptions, required: true, serverKey: 'tag' },
          { name: '프리퀀시 기준', type: 'select', options: ['캠페인별', '소재별'], required: false, serverKey: 'frequency_criteria' },
          { name: '프리퀀시 타입', type: 'select', options: frequencyTypeOptions, required: true, serverKey: 'frequency_type', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '일일 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_daily_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '최대 프리퀀시 제한 횟수', type: 'text', required: false, serverKey: 'frequency_max_limit', dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
          { name: '트래커', type: 'select', options: trackerOptions, required: false, serverKey: 'tracker' },
          { name: '노출 우선순위', type: 'text', required: true, defaultValue: '10', serverKey: 'priority' },
          { name: 'Seg Filter', type: 'text', required: false, serverKey: 'seg_filter' },
          { name: '기본 노출형', type: 'dynamic_table', serverKey: 'exposure_settings' }
        ]
      }
    },
    'dropdowns': {
      advertisers: []
    }
  };
  return fields;
}

/**
 * 캐시슬라이드 광고 수정 타입별 폼 필드 정의를 반환합니다.
 * @returns {object} 캐시슬라이드 수정 폼 필드 설정 객체
 */
function getCashslideModificationFields() {
  const demoTarget1Options = ['전체', '남성', '여성'];
  const frequencyTypeOptions = ['노출', '클릭', '노출 & 클릭'];
  const trackerOptions = ['Tune', 'Appsflyer', 'Ad brix', 'Adjust', 'Airbridge', 'Singular', 'Ad brix Remaster', 'Branch with google_ad_id'];

  const fields = {
    'fields': [
      // 필수 항목
      { name: '광고ID', type: 'textarea', rows: 5, required: true, placeholder: '복수의 경우, "," 로 구분' },
      { name: '광고명', type: 'textarea', rows: 5, required: true },
      { name: '요청사항', type: 'textarea', rows: 5, required: true },
      
      // 선택 항목
      { name: '광고주 계정 ID', type: 'text', required: false },
      { name: '라이브 시작일시', type: 'datetime_picker', required: false },
      { name: '라이브 종료일시', type: 'datetime_picker', required: false },
      { name: '데일리 라이브 시작 시간', type: 'time_picker', required: false },
      { name: '데일리 라이브 종료 시간', type: 'time_picker', required: false },
      { name: '소재경로', type: 'textarea', rows: 5, required: false },
      { name: '일물량', type: 'textarea', rows: 5, required: false },
      { name: '랜딩 URL', type: 'textarea', rows: 5, required: false },
      { name: '데모타겟1', type: 'select', options: demoTarget1Options, required: false },
      { name: '데모타겟2', type: 'text', required: false, placeholder: '연령 (예 : 20-39)' },
      { name: '리타겟 클러스터', type: 'text', required: false },
      { name: '앱 패키지명 - 리타겟팅', type: 'text', required: false },
      { name: '앱 패키지명 - 디타겟팅', type: 'text', required: false },
      
      // LBS 타겟팅 (3개의 개별 필드로 구현)
      { name: 'LBS 타겟팅 - 위도', type: 'text', required: false },
      { name: 'LBS 타겟팅 - 경도', type: 'text', required: false },
      { name: 'LBS 타겟팅 - 범위', type: 'text', required: false },
      
      // 기존 광고 등록과 동일한 필드들
      { name: '프리퀀시 기준', type: 'select', options: ['캠페인별', '소재별'], required: false },
      { name: '프리퀀시 타입', type: 'select', options: frequencyTypeOptions, required: true, dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
      { name: '일일 프리퀀시 제한 횟수', type: 'text', required: false, dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
      { name: '최대 프리퀀시 제한 횟수', type: 'text', required: false, dependency: { field: '프리퀀시 기준', showsOn: ['캠페인별', '소재별'] } },
      { name: '트래커', type: 'select', options: trackerOptions, required: false },
      { name: '노출 우선순위', type: 'text', required: false }, // defaultValue 제거
      { name: 'Seg Filter', type: 'text', required: false }
    ]
  };

  return fields;
}

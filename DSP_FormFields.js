/**
 * DSP 광고 타입별 폼 필드 정의를 반환합니다.
 * @returns {object} DSP 폼 필드 설정 객체
 */
function getDspFields() {
  const trackerOptions = ['ApplsFlyer', 'Adjust', 'Adbrix Remaster', 'Branch', 'Airbridge', 'Singular'];
  const eventConditionOptions = ['Value', 'Range', 'not_allowed_value'];
  const subscribeTargetOptions = ['기본이벤트', '유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우'];

  const dspFields = {
    'DSP': {
      'DSP_광고주연동': {
        fields: [
          { name: '요청사항', label: '요청사항', type: 'textarea', rows: 5, required: true },
          { name: '광고주', label: '광고주', type: 'searchable_dropdown', optionsKey: 'advertisers', required: true },
          { name: '캠페인명', label: '캠페인명', type: 'text', required: true },
          { name: '광고명', label: '광고명', type: 'text', required: true },
          { name: '광고 타입', label: '광고 타입', type: 'select', options: ['CPI', 'CPA S2S', 'CPA TRACKER'], required: true },
          { name: '재참여 타입', label: '재참여 타입', type: 'select', options: ['단일 참여', '무한 참여'], required: true },
          { name: '광고주 연동 토큰', label: '광고주 연동 토큰', type: 'text', required: false },
          { name: '트래커', label: '트래커', type: 'select', options: trackerOptions, required: false },
          { name: '완료 이벤트 이름', label: '완료 이벤트 이름', type: 'text', required: false },
          { name: '이벤트 토큰', label: '이벤트 토큰', type: 'text', placeholder: 'Adjust 전용', required: false },
          {
            name: '완료 이벤트 조건(event_parameters JSON 파라미터)',
            label: '완료 이벤트 조건(event_parameters JSON 파라미터)',
            type: 'group_checkbox',
            // ▼▼▼ [수정] 하위 필드의 name 속성을 고유하게 변경 ▼▼▼
            subFields: [
              { name: '완료 이벤트 조건(event_parameters JSON 파라미터) 이벤트 타입', label: '이벤트 타입', type: 'select', options: eventConditionOptions, required: true },
              { name: '완료 이벤트 조건(event_parameters JSON 파라미터) 이벤트 이름', label: '이벤트 이름', type: 'text' },
              { name: '완료 이벤트 조건(event_parameters JSON 파라미터) value', label: 'value', type: 'text' },
              { name: '완료 이벤트 조건(event_parameters JSON 파라미터) from', label: 'from', type: 'text' },
              { name: '완료 이벤트 조건(event_parameters JSON 파라미터) to', label: 'to', type: 'text' }
            ]
            // ▲▲▲ [수정] 여기까지 ▲▲▲
          },
          {
            name: '완료 이벤트 조건 (개별 파라미터)',
            label: '완료 이벤트 조건 (개별 파라미터)',
            type: 'group_checkbox',
            subFields: [
              { name: '완료 이벤트 조건 (개별 파라미터) 파라미터 타입', label: '파라미터 타입', type: 'select', options: eventConditionOptions, required: true },
              { name: '완료 이벤트 조건 (개별 파라미터) 파라미터 이름', label: '파라미터 이름', type: 'text' },
              { name: '완료 이벤트 조건 (개별 파라미터) value', label: 'value', type: 'text' },
              { name: '완료 이벤트 조건 (개별 파라미터) from', label: 'from', type: 'text' },
              { name: '완료 이벤트 조건 (개별 파라미터) to', label: 'to', type: 'text' }
            ]
          },
          { name: '랜딩 URL', label: '랜딩 URL', type: 'text', required: true },
          { name: '완료 인정 유효기간 (일단위)', label: '완료 인정 유효기간 (일단위)', type: 'text', required: true, defaultValue: '7일' },
          { name: '총물량', label: '총물량', type: 'number_unlimited', required: true },
          { name: '일물량', label: '일물량', type: 'number_unlimited', required: true },
          { name: '광고 단가', label: '광고 단가', type: 'number', allowFloat: true, required: true, placeholder: '소수점 입력 가능' },
          { name: '총예산', label: '총예산', type: 'text', readonly: true, placeholder: '총물량과 광고단가 입력 시 자동계산' },
          { name: '광고 시작일시', label: '광고 시작일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '00:00' },
          { name: '광고 종료일시', label: '광고 종료일시', type: 'datetime_picker', required: true, hasUnlimited: true, defaultTime: '23:59' },
          { type: 'heading', label: '매체별 상세 설정' },
          { name: 'media_pivot_table', type: 'media_pivot_table' }
        ]
      },
'DSP_구독형(EVENT)': {
        fields: [
          { name: '요청사항', label: '요청사항', type: 'textarea', rows: 5, required: true },
          { name: '광고주', label: '광고주', type: 'searchable_dropdown', optionsKey: 'advertisers', required: true },
          { name: '캠페인명', label: '캠페인명', type: 'text', required: true },
          { name: '광고명', label: '광고명', type: 'text', required: true },
          { name: '재참여 타입', label: '재참여 타입', type: 'select', options: ['단일 참여', '무한 참여'], required: true },
          { name: '구독 대상 이름', label: '구독 대상 이름', type: 'select', options: subscribeTargetOptions, required: true },
          { 
            name: '이미지 인식에 사용할 식별자', label: '이미지 인식에 사용할 식별자', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: '기본이벤트' }
          },
          { 
            name: '광고주 계정 식별자 1', label: '광고주 계정 식별자 1', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우'] }
          },
          { 
            name: '광고주 계정 식별자 2', label: '광고주 계정 식별자 2', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우'] }
          },
          { 
            name: '광고주 계정 식별자 3', label: '광고주 계정 식별자 3', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: ['유튜브 구독(채널메인)', '유튜브 구독(특정영상)', '팔로우'] }
          },
          { 
            name: '가이드 메세지', label: '가이드 메세지', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: '기본이벤트' }
          },
          { 
            name: '버튼 메세지', label: '버튼 메세지', type: 'text', required: true, 
            dependency: { field: '구독 대상 이름', showsOn: '기본이벤트' }
          },
          { 
            name: '구독 페이지 랜딩 URL', label: '구독 페이지 랜딩 URL', type: 'text', required: true, placeholder: '인스타그램 사용자 ID만 입력', 
            dependency: { field: '구독 대상 이름', showsOn: '팔로우' }
          },
          { name: '이미지 인식 예시 이미지', label: '이미지 인식 예시 이미지', type: 'text', required: true },
          
          { 
            name: '랜딩 URL', label: '랜딩 URL', type: 'text', required: true
          },
          { name: '완료 인정 유효기간 (일단위)', label: '완료 인정 유효기간 (일단위)', type: 'text', required: true, defaultValue: '7일' },
          { type: 'heading', label: '예산 및 일정' },
          { name: '총물량', label: '총물량', type: 'number_unlimited', required: true },
          { name: '일물량', label: '일물량', type: 'number_unlimited', required: true },
          { name: '광고 단가', label: '광고 단가', type: 'number', allowFloat: true, required: true, placeholder: '소수점 입력 가능' },
          { name: '총예산', label: '총예산', type: 'text', readonly: true, placeholder: '총물량과 광고단가 입력 시 자동계산' },
          { name: '광고 시작일시', label: '광고 시작일시', type: 'datetime_picker', required: true, hasUnlimited: false, defaultTime: '00:00' },
          { name: '광고 종료일시', label: '광고 종료일시', type: 'datetime_picker', required: true, hasUnlimited: true, defaultTime: '23:59' },
          { type: 'heading', label: '매체별 상세 설정' },
          { name: 'media_pivot_table', type: 'media_pivot_table' }
        ]
      }
    },
    'dropdowns': {
      // 드롭다운 목록은 외부에서 채워줍니다.
      advertisers: []
    }
  };

  return dspFields;
}
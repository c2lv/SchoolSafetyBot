# 시스템 오류
ERR_NO_REQUESTS  = "⚠️ requests 라이브러리가 없어요.\n관리자에게 문의해 주세요."
ERR_NO_COOKIE    = "⚠️ 쿠키가 설정되지 않았어요.\n관리자에게 문의해 주세요."
ERR_NDRIMS_CONN  = "⚠️ NDRIMS 연결 실패\n관리자에게 문의해 주세요."
ERR_MENU_NOT_FOUND  = "⚠️ 메뉴를 찾을 수 없어요."
ERR_UNKNOWN_ATYPE   = "⚠️ 알 수 없는 액션 타입이에요."
ERR_UNKNOWN_DATA   = "⚠️ 알 수 없는 데이터 형식이에요.\n관리자에게 문의해 주세요."

# 쿠키 갱신 플로우
MSG_COOKIE_START   = "🔐 쿠키 갱신\n\n먼저 관리자 비밀번호를 입력해 주세요.\n\n비밀번호 확인 후 쿠키 입력 단계로 넘어갑니다."
MSG_NO_ADMIN_PW    = "⚠️ 관리자 비밀번호가 설정되지 않았어요."
MSG_PASSWORD_OK    = "✅ 비밀번호가 확인됐어요.\n\n이제 아래 형식으로 쿠키를 입력해 주세요:\nWMONID=값 JSESSIONID=값"
MSG_WRONG_PASSWORD = "❌ 잘못된 비밀번호예요."
MSG_WRONG_FORMAT   = "❌ 형식이 맞지 않아요. 다시 시도해 주세요.\n\n올바른 형식:\nWMONID=값 JSESSIONID=값"

TEXT_RESPONSES = {
    "budget_rule": (
        "💰 연구비 사용 한도(테스트입니다)\n\n"
        "• 50만 원 이하 → 팀장 승인\n"
        "• 50만 원 초과 → 소장 결재\n"
        "• 100만 원 초과 → 사전 품의서 필수\n\n"
        "📌 근거: 연구비관리규정 제7조"
    ),
    "travel_domestic": (
        "🇰🇷 국내 출장 규정(테스트입니다)\n\n"
        "• 일비: 연구원 2만 원 / 책임 3만 원\n"
        "• 숙박: 실비 (상한 7만 원/박)\n"
        "• 서울 출장: 당일 복귀 원칙\n"
        "• 신청: 출발 3일 전까지\n\n"
        "📌 근거: 출장여비규정 제3·5조"
    ),
    "travel_overseas": (
        "✈️ 해외 출장 규정(테스트입니다)\n\n"
        "• 신청: 출발 7일 전까지\n"
        "• 항공: 이코노미 원칙\n"
        "  (14시간 초과 시 비즈니스 허용)\n"
        "• 일비: 국가별 기준표 적용\n\n"
        "📌 근거: 출장여비규정 제8조"
    ),
    "equipment_rule": (
        "🔧 공용 장비 예약(테스트입니다)\n\n"
        "• 사용 2일 전까지 시스템 예약 필수\n"
        "• 예약 없이 사용 불가\n"
        "• 미사용 취소: 24시간 전까지\n"
        "• 고장 발견 시 즉시 내선 234 신고\n\n"
        "📌 근거: 장비관리규정 제4·10조"
    ),
    "help": (
        "📖 사용법\n\n"
        "버튼을 눌러 연구소 관련 문제 해결\n"
        "기타 문제는 채널 관리자에게 문의"
    ),
    "work_link": (
        "🔗 링크\n\n"
        "행정 업무에 도움이 되는 사이트 링크 모음입니다.\n\n"
        "한국연구재단\nhttps://www.nrf.re.kr/index\n\n"
        "통합이지바로(GAIA)\nhttps://www.gaia.go.kr\n\n"
        "동국대학교 포털\nhttps://www.dongguk.edu\n\n"
        "동국대 그룹웨어\nhttps://gw.dongguk.edu\n\n"
        "범부처통합연구지원시스템(IRIS)\nhttps://www.iris.go.kr\n\n"
        "산학협력단\nhttps://rnd.dongguk.edu\n\n"
        "구글드라이브\nhttps://drive.google.com\n\n"
        "바른한글\nhttps://www.nara-speller.co.kr\n\n"
        "신한카드 법인\nhttps://www.shinhancard.com/cconts/html/main.html\n"
    ),
}

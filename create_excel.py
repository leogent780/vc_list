import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# 색상 정의
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILLS = {
    "초기 (Pre-Seed/Seed)": PatternFill("solid", fgColor="D6E4F0"),
    "초기~중기 (Series A/B)": PatternFill("solid", fgColor="D5F5E3"),
    "성장·후기 (Series B/C/Pre-IPO)": PatternFill("solid", fgColor="FEF9E7"),
    "글로벌·크로스보더": PatternFill("solid", fgColor="FADBD8"),
    "정책 기관": PatternFill("solid", fgColor="E8DAEF"),
}
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(bold=True, size=10)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

data = [
    # (단계, VC명, 투자단계, 특징, 웹사이트)
    ("초기 (Pre-Seed/Seed)", "프라이머 (Primer)", "Pre-Seed / Seed", "국내 최초 시드 전문 액셀러레이터, 배기홍·권도균 설립", "primer.kr"),
    ("초기 (Pre-Seed/Seed)", "스파크랩 (SparkLabs)", "Seed / Series A", "글로벌 네트워크 보유 액셀러레이터", "sparklabs.co.kr"),
    ("초기 (Pre-Seed/Seed)", "매쉬업엔젤스 (Mashup Angels)", "Seed", "이택경 설립, 초기 스타트업 전문", "mashupangels.com"),
    ("초기 (Pre-Seed/Seed)", "본엔젤스 (BonAngels)", "Seed / Series A", "장병규 설립, 국내 대표 시드 VC", "bonangels.net"),
    ("초기 (Pre-Seed/Seed)", "퓨처플레이 (FuturePlay)", "Seed / Series A", "류중희 설립, 딥테크 특화", "futureplay.co"),
    ("초기 (Pre-Seed/Seed)", "패스트벤처스 (Fast Ventures)", "Seed / Series A", "빠른 의사결정, IT/플랫폼 집중", "fast.vc"),
    ("초기 (Pre-Seed/Seed)", "카카오벤처스 (Kakao Ventures)", "Seed / Series A", "카카오 계열 초기 투자 CVC", "kakaoventures.com"),
    ("초기 (Pre-Seed/Seed)", "네이버 D2SF", "Seed / Series A", "딥테크·개발자 도구 특화, 네이버 CVC", "d2startup.com"),
    ("초기 (Pre-Seed/Seed)", "디캠프 (D.CAMP)", "Seed", "은행권청년창업재단, 사무공간+투자", "dcamp.kr"),
    ("초기 (Pre-Seed/Seed)", "현대기아 ZER01NE", "Seed / Series A", "현대기아차 모빌리티 CVC", "zer01ne.zone"),
    ("초기 (Pre-Seed/Seed)", "롯데액셀러레이터", "Seed / Series A", "롯데그룹 계열 액셀러레이터", ""),
    ("초기 (Pre-Seed/Seed)", "에이벤처스 (A Ventures)", "Seed", "이커머스·소비자 서비스 특화", ""),
    ("초기 (Pre-Seed/Seed)", "컴퍼니케이파트너스", "Seed / Series A", "초기 소비재·서비스 특화", ""),
    ("초기 (Pre-Seed/Seed)", "엔슬파트너스", "Seed / Series A", "초기 기술 스타트업 투자", ""),

    # 추가: 액셀러레이터 & 초기 VC
    ("초기 (Pre-Seed/Seed)", "블루포인트파트너스 (Bluepoint Partners)", "Pre-Seed / Seed", "이용관 설립, 딥테크 특화, 창업자 선호 1위 AC, 200개+ 투자", "bluepoint.ac"),
    ("초기 (Pre-Seed/Seed)", "소풍벤처스 (Sopoong Ventures)", "Pre-Seed / Seed", "임팩트·소셜벤처 특화, 2024년 VC 라이선스 취득", "sopoong.net"),
    ("초기 (Pre-Seed/Seed)", "해시드 (Hashed)", "Seed / Series A", "블록체인·Web3 전문 글로벌 VC, 김서준 설립", "hashed.com"),
    ("초기 (Pre-Seed/Seed)", "캡스톤파트너스 (Capstone Partners)", "Seed / Series A", "독립계 초기 VC, 500억+ 펀드 운용", "capstonepartners.co.kr"),
    ("초기 (Pre-Seed/Seed)", "스트롱벤처스 (Strong Ventures)", "Seed / Series A", "미국·한국 크로스보더 초기 투자, 2025H1 11건 투자", "strongvc.com"),
    ("초기 (Pre-Seed/Seed)", "씨엔티테크 (CNT Tech)", "Pre-Seed / Seed", "2003년 설립, 스포츠·엔터프라이즈 특화 AC, 상장 추진 중", "cntt.co.kr"),
    ("초기 (Pre-Seed/Seed)", "더인벤션랩 (The Invention Lab)", "Pre-Seed / Seed", "2014년 설립, 누적 투자 260억, 25개+ 펀드 운용, AI·딥테크", "theinventionlab.co.kr"),
    ("초기 (Pre-Seed/Seed)", "빅뱅엔젤스 (BigBang Angels)", "Pre-Seed / Seed", "2012년 설립, 엔터프라이즈·B2B SaaS 특화 AC", "bigbangangels.com"),
    ("초기 (Pre-Seed/Seed)", "슈미트 (Schmidt)", "Pre-Seed / Seed", "2025H1 10건+ 투자, 초기 스타트업 특화 AC", ""),
    ("초기 (Pre-Seed/Seed)", "엠와이소셜컴퍼니 (MY Social Company)", "Pre-Seed / Seed", "2025H1 초기 라운드 투자 최다 참여 AC 중 하나", ""),
    ("초기 (Pre-Seed/Seed)", "한국투자액셀러레이터", "Pre-Seed / Seed", "한국투자증권 계열 AC, 검증된 금융 네트워크 보유", "ac.koreainvestment.com"),
    ("초기 (Pre-Seed/Seed)", "테크스타스 코리아 (Techstars Korea)", "Pre-Seed / Seed", "글로벌 액셀러레이터 Techstars 한국 프로그램", "techstars.com"),
    ("초기 (Pre-Seed/Seed)", "K-Startup (정부 프로그램)", "Pre-Seed / Seed", "정부 지원 글로벌 스타트업 육성 프로그램, 중기부 운영", "k-startup.go.kr"),
    ("초기 (Pre-Seed/Seed)", "로아인벤션랩 (ROA Invention Lab)", "Pre-Seed / Seed", "지능정보·AI 특화 AC, K-Global 액셀러레이터 육성사업 참여", ""),
    ("초기 (Pre-Seed/Seed)", "한국엔젤투자협회 (KBAN)", "Pre-Seed / Seed", "2012년 설립, 엔젤투자 지원 및 매칭 플랫폼 운영", "kban.or.kr"),

    # 2차 추가
    ("초기 (Pre-Seed/Seed)", "N파트너스 (N Partners)", "Pre-Seed / Seed / Series A", "2022년 설립, 카카오·크래프톤·토스·당근 투자 경험 파트너진, AI·블록체인·핀테크·헬스케어 집중", "n.partners"),
    ("초기 (Pre-Seed/Seed)", "ZDVC (지디벤처스)", "Pre-Seed / Seed", "2023년 설립, 20대 파트너 4인 운영, K팝·콘텐츠·뷰티·패션·Web3 집중, 5천만~1.5억 소액 초기 투자", ""),
    ("초기 (Pre-Seed/Seed)", "스프링캠프 (SpringCamp)", "Pre-Seed / Seed", "2015년 설립, 서울대·스탠포드 인근 거점, 10년 532개사 투자 32% IRR, 미국 진출 지원 강점", "springcamp.co"),
    ("초기 (Pre-Seed/Seed)", "프라이머사제파트너스", "Seed / Series A", "프라이머 출신 설립, 업스테이지 등 초기 시드 투자, 딥테크·AI 집중", ""),

    ("초기~중기 (Series A/B)", "알토스벤처스 (Altos Ventures)", "Series A~C", "쿠팡·크래프톤 등 대형 투자, 미국계 독립 VC", "altosvc.com"),
    ("초기~중기 (Series A/B)", "소프트뱅크벤처스아시아", "Series A~C", "소프트뱅크 계열, 아시아 특화", "softbank.vc"),
    ("초기~중기 (Series A/B)", "IMM인베스트먼트", "Series A~C", "대형 독립계 VC", "imminvestment.com"),
    ("초기~중기 (Series A/B)", "한국투자파트너스", "Series A~C", "한국투자증권 계열", "kipvc.com"),
    ("초기~중기 (Series A/B)", "LB인베스트먼트", "Series A~C", "LG그룹 계열 VC", "lbinvestment.co.kr"),
    ("초기~중기 (Series A/B)", "미래에셋벤처투자", "Series A~C", "미래에셋 계열", "miraeassetventure.com"),
    ("초기~중기 (Series A/B)", "KB인베스트먼트", "Series A~C", "KB금융 계열", "kbinvestment.co.kr"),
    ("초기~중기 (Series A/B)", "신한벤처투자", "Series A~C", "신한금융 계열", "shinhanventure.com"),
    ("초기~중기 (Series A/B)", "하나벤처스", "Series A~C", "하나금융 계열", "hanaventures.co.kr"),
    ("초기~중기 (Series A/B)", "인터베스트 (InterVest)", "Series A~C", "독립계, IT·바이오 특화", "intervest.co.kr"),
    ("초기~중기 (Series A/B)", "스마일게이트인베스트먼트", "Series A~C", "스마일게이트 CVC", ""),
    ("초기~중기 (Series A/B)", "카카오인베스트먼트", "Series A~C", "카카오 계열 성장 투자", ""),
    ("초기~중기 (Series A/B)", "KDB산업은행 캐피탈", "Series A~C", "정책금융 기관", "kdbcapital.co.kr"),
    ("초기~중기 (Series A/B)", "DSC인베스트먼트", "Series A~C", "상장 독립계 VC", "dscinvestment.com"),
    ("초기~중기 (Series A/B)", "SV인베스트먼트", "Series A~C", "독립계 VC", "svinvestment.co.kr"),
    ("초기~중기 (Series A/B)", "세마인베스트먼트", "Series A~C", "독립계 VC", "semavc.com"),

    ("성장·후기 (Series B/C/Pre-IPO)", "스틱인베스트먼트 (STIC)", "Series B~D", "PE·VC 겸업, 대형 펀드 운용", "sticvc.com"),
    ("성장·후기 (Series B/C/Pre-IPO)", "에이티넘인베스트먼트", "Series B~D", "대형 독립계 VC", "atinum.co.kr"),
    ("성장·후기 (Series B/C/Pre-IPO)", "UTC인베스트먼트", "Series B~D", "독립계, 제조·기술 특화", "utcis.co.kr"),
    ("성장·후기 (Series B/C/Pre-IPO)", "큐캐피탈파트너스", "Series B~D", "PE·VC 겸업", "qcapital.co.kr"),
    ("성장·후기 (Series B/C/Pre-IPO)", "TS인베스트먼트", "Series A~C", "상장사, 바이오·IT 특화", "ts-investment.co.kr"),
    ("성장·후기 (Series B/C/Pre-IPO)", "한화인베스트먼트", "Series B~D", "한화그룹 CVC", ""),
    ("성장·후기 (Series B/C/Pre-IPO)", "신세계 시그나이트파트너스", "Series A~C", "신세계그룹 CVC", "signiteventures.com"),
    ("성장·후기 (Series B/C/Pre-IPO)", "GS벤처스", "Series B~C", "GS그룹 CVC", ""),
    ("성장·후기 (Series B/C/Pre-IPO)", "CJ인베스트먼트", "Series A~C", "CJ그룹 CVC", ""),
    ("성장·후기 (Series B/C/Pre-IPO)", "롯데벤처스", "Series A~C", "롯데그룹 CVC", ""),
    ("성장·후기 (Series B/C/Pre-IPO)", "크레비스파트너스", "Series A~C", "임팩트·사회적 가치 투자", "crevisse.com"),
    ("성장·후기 (Series B/C/Pre-IPO)", "위벤처스 (We Ventures)", "Series A~C", "여성 창업자·다양성 강조", "we.vc"),
    ("성장·후기 (Series B/C/Pre-IPO)", "스퀘어원파트너스", "Series A~C", "핀테크·SaaS 특화", ""),
    ("성장·후기 (Series B/C/Pre-IPO)", "NH농협PE", "Series B~D", "NH농협금융 계열", ""),

    ("글로벌·크로스보더", "알토스벤처스 (Altos Ventures)", "Series A~C+", "미국 본사, 한국 스타트업 집중 투자", "altosvc.com"),
    ("글로벌·크로스보더", "소프트뱅크벤처스아시아", "Series B~C", "일본 소프트뱅크 계열, 아시아 투자", "softbank.vc"),
    ("글로벌·크로스보더", "500 스타트업스 코리아", "Seed~Series A", "글로벌 액셀러레이터 한국 프로그램", ""),
    ("글로벌·크로스보더", "스파크랩 글로벌", "Seed~Series A", "미국·한국 연결 액셀러레이터", "sparklabs.co.kr"),
    ("글로벌·크로스보더", "블루런벤처스 (BlueRun)", "Series A~C", "미국계, 모바일·인터넷 특화", "brv.com"),

    ("정책 기관", "한국벤처투자 (KVIC)", "모태펀드 출자", "모태펀드 운용, 전체 VC 출자 기관", "kvic.or.kr"),
    ("정책 기관", "기술보증기금 (KIBO)", "Seed~Series A", "기술 창업 보증·투자", "kibo.or.kr"),
    ("정책 기관", "신용보증기금", "Seed~Series A", "스타트업 보증 지원", "kodit.co.kr"),
    ("정책 기관", "TIPS 프로그램 (중기부)", "Seed~Series A", "민간 투자 주도형 기술 창업 지원", "jointips.or.kr"),
]

ws = wb.active
ws.title = "한국 VC 리스트"

# 헤더
headers = ["투자 단계", "VC명", "주요 투자 라운드", "특징", "웹사이트"]
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = BORDER

# 데이터 입력
for row_idx, (stage, name, rounds, desc, url) in enumerate(data, 2):
    fill = SECTION_FILLS.get(stage, PatternFill("solid", fgColor="FFFFFF"))
    values = [stage, name, rounds, desc, url]
    for col_idx, val in enumerate(values, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.fill = fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        if col_idx == 2:
            cell.font = Font(bold=True, size=10)

# 열 너비 조정
col_widths = [28, 30, 22, 50, 25]
for i, width in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = width

# 행 높이
ws.row_dimensions[1].height = 22
for i in range(2, len(data) + 2):
    ws.row_dimensions[i].height = 18

# 틀 고정 (헤더 고정)
ws.freeze_panes = "A2"

# 자동 필터
ws.auto_filter.ref = f"A1:E{len(data) + 1}"

wb.save("/home/user/vc_list/korea_vc_list.xlsx")
print("Done")

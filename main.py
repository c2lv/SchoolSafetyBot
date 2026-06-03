"""
NDRIMS 카카오 챗봇 스킬 서버
- menu.xlsx 를 읽어 카카오 버튼 메뉴 자동 생성
- 버튼 클릭 → NDRIMS API 호출 → 텍스트/카드 응답
- 카톡에서 쿠키값 직접 갱신 가능
"""

import json, os, threading, urllib.request, re, sqlite3
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from datetime import datetime, timezone, timedelta
from text_responses import (
    TEXT_RESPONSES,
    ERR_NO_REQUESTS, ERR_NO_COOKIE, ERR_NDRIMS_CONN,
    ERR_MENU_NOT_FOUND, ERR_UNKNOWN_ATYPE, ERR_UNKNOWN_DATA,
    MSG_COOKIE_START, MSG_NO_ADMIN_PW, MSG_PASSWORD_OK,
    MSG_WRONG_PASSWORD, MSG_WRONG_FORMAT,
)

# openpyxl 은 pip install openpyxl 로 설치
try:
    import openpyxl
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False
    print("[경고] openpyxl 없음 — pip install openpyxl")

try:
    from dotenv import load_dotenv
    DOTENV_OK = True
except ImportError:
    DOTENV_OK = False
    print("[경고] python-dotenv 없음 — pip install python-dotenv")

try:
    import requests as req_lib
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False
    print("[경고] requests 없음 — pip install requests")

BASE_DIR = Path(__file__).parent

if DOTENV_OK:
    load_dotenv(BASE_DIR / ".env")

ADMIN_PASSWORD = os.environ.get("DGUSSB_PASSWORD", "")
_base_url = ""  # 첫 요청 시 Host 헤더에서 자동 감지

# ── Utils ────────────────────────────────────────────
def get_kst() -> str:
    kst = datetime.now(timezone(timedelta(hours=9)))
    return kst.strftime("%Y-%m-%d %H:%M:%S")

# ── SQLite 캐시 ──────────────────────────────────────
DB_PATH = BASE_DIR / "cache.db"

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS cache (
                action TEXT PRIMARY KEY,
                text   TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS cookies (
                key   TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

def cache_save(action: str, text: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO cache (action, text, updated_at) VALUES (?, ?, ?)",
            (action, text, get_kst())
        )

def cache_load(action: str) -> tuple[str | None, str | None]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT text, updated_at FROM cache WHERE action = ?", (action,)
        ).fetchone()
    return (row[0], row[1]) if row else (None, None)

def save_cookies():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("INSERT OR REPLACE INTO cookies VALUES ('WMONID', ?)",     (cookie_store["WMONID"],))
        conn.execute("INSERT OR REPLACE INTO cookies VALUES ('JSESSIONID', ?)", (cookie_store["JSESSIONID"],))

def load_cookies():
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute("SELECT key, value FROM cookies").fetchall()
            return {k: v for k, v in rows}
    except Exception:
        return {}

init_db()
_saved = load_cookies()
# ── 쿠키 상태 (DB → 환경변수 순으로 로드) ────────────────────
cookie_store = {
    "WMONID":     _saved.get("WMONID")     or os.environ.get("WMONID", ""),
    "JSESSIONID": _saved.get("JSESSIONID") or os.environ.get("JSESSIONID", ""),
}

# ── 엑셀 메뉴 로드 ────────────────────────────────────────────
def load_menu() -> dict:
    """
    menu.xlsx 를 읽어 트리 구조로 변환.
    반환 형태:
    {
      "A": {
        "label": "📂 과제 조회",
        "children": {
          "A1": {"label": "과제 목록", "action": "project_list", "action_type": "NDRIMS", "children": {}},
          "A2": {"label": "지출 현황", "action": None, "children": {
            "A2a": {"label": "📊 과제별 지출", "action": "expense_by_project", ...},
            ...
          }},
        }
      }, ...
    }
    """
    path = BASE_DIR / "menu.xlsx"
    if not EXCEL_OK or not path.exists():
        print("[경고] menu.xlsx 없음 — 기본 메뉴 사용")
        return {}

    wb = openpyxl.load_workbook(path)
    ws = wb["menu"]

    tree = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        id1, lbl1, id2, lbl2, id3, lbl3, atype, aval, _ = row[:9]
        if not id1:
            continue

        # 1단계
        if id1 not in tree:
            tree[id1] = {"label": lbl1 or id1, "action_type": None, "action": None, "children": {}}

        if not id2:
            # 1단계에서 바로 액션
            tree[id1]["action_type"] = atype
            tree[id1]["action"] = aval
            continue

        # 2단계
        if id2 not in tree[id1]["children"]:
            tree[id1]["children"][id2] = {
                "label": lbl2 or id2, "action_type": None, "action": None, "children": {}
            }

        if not id3:
            # 2단계에서 바로 액션
            tree[id1]["children"][id2]["action_type"] = atype
            tree[id1]["children"][id2]["action"] = aval
            continue

        # 3단계
        tree[id1]["children"][id2]["children"][id3] = {
            "label": lbl3 or id3,
            "action_type": atype,
            "action": aval,
            "children": {}
        }

    return tree

MENU = load_menu()


def build_menu_label_index(tree: dict) -> dict[str, str]:
    index: dict[str, str] = {}

    def walk(nodes: dict):
        for menu_id, node in nodes.items():
            label = (node.get("label") or "").strip()
            if label and label not in index:
                index[label] = menu_id
            children = node.get("children", {})
            if children:
                walk(children)

    walk(tree)
    return index


MENU_LABEL_INDEX = build_menu_label_index(MENU)

# ── nDRIMS API 호출 ──────────────────────────────────────────
def make_ndrims_session():
    if not REQUESTS_OK:
        return None
    s = req_lib.Session()
    s.cookies.set("WMONID",     cookie_store["WMONID"],     domain="ndrims.dongguk.edu")
    s.cookies.set("JSESSIONID", cookie_store["JSESSIONID"], domain="ndrims.dongguk.edu")
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": "https://ndrims.dongguk.edu/",
    })
    return s


def get_running_values(session) -> dict:
    res = session.post(
        "https://ndrims.dongguk.edu/unis/main/view/login/doUserLoginAccs.do",
        data={}, timeout=10
    )
    data = res.json()
    nana = data["dmLoginAccsResult"]["NANA"]

    res2 = session.post(
        "https://ndrims.dongguk.edu/unis/main/view/menu/doListUserMenuListTop.do",
        data={
            "_runningNana": nana,
            "@d1#MOBILE_YN": "N",
            "@d#": "@d1#", "@d1#": "dmSearch", "@d1#tp": "dm"
        }, timeout=10
    )
    data2 = res2.json()
    user_info = data2["dmUserInfo"]
    return nana, user_info["LOGIN_IDEN_NO"], user_info["MAIN_OPEN_KEY"]


def call_ndrims(action: str) -> str:
    """액션값에 따라 NDRIMS API 호출 → 텍스트 결과 반환."""
    if not REQUESTS_OK:
        text, _ = cache_load(action)
        if text:
            return text
        return ERR_NO_REQUESTS
    if not cookie_store["JSESSIONID"]:
        text, _ = cache_load(action)
        if text:
            return text
        return ERR_NO_COOKIE

    try:
        session = make_ndrims_session()
        print(f"[NDRIMS] 세션 생성 완료")

        try:
            rv = get_running_values(session)
            menu_payload = {
                "_runningNana": rv[0],
                "_runningLoginIdenNo": rv[1],
                "_runningMainOpenKey": rv[2],
                "@d1#MENU_NO": "10497",
                "@d1#MOBILE_YN": "N",
                "@d#": "@d1#",
                "@d1#": "dmSearch",
                "@d1#tp": "dm",
            }
            
            menu_res = session.post(
                "https://ndrims.dongguk.edu/unis/main/view/menu/doListUserMenuListLeft.do",
                data=menu_payload, timeout=10
            )
            menu_data = menu_res.json()
            print(f"[NDRIMS] doListUserMenuListLeft 호출 성공")
            # dsUserMenuListLeft에서 PGM_URL이 app/rs/rsb/prjfi/RsbPrjfi100인 항목 찾기
            RsbPrjfi100_nana = None
            menu_items = menu_data.get("dsUserMenuListLeft", [])
            for item in menu_items:
                pgm_url = item.get("PGM_URL") or ""
                if "app/rs/rsb/prjfi/RsbPrjfi100" in pgm_url:
                    RsbPrjfi100_nana = item.get("NANA")
                    print(f"[NDRIMS] RsbPrjfi100_nana 추출: {RsbPrjfi100_nana[:8] if RsbPrjfi100_nana else 'EMPTY'}...")
                    break
            if not RsbPrjfi100_nana:
                raise ValueError("RsbPrjfi100 메뉴 항목을 찾을 수 없어요.")
        except Exception as e:
            print(f"[NDRIMS] ⚠️ 에러: {e}")
            text, _ = cache_load(action)
            if text:
                return text
            return ERR_NDRIMS_CONN

        print(f"[NDRIMS] 요청 액션: {action}")

        # ── 과제 목록 ─────────────────────────────────────
        if action == "project_list":
            # build payload as requested
            payload = {
                "_runningNana": RsbPrjfi100_nana,
                "_runningLoginIdenNo": rv[1],
                "_runningMainOpenKey": rv[2],
                "@d1#CHIFR_MEMB_NO": "SE260018",
                "@d1#CHIFR_MEMB_NM": "박현준",
                "@d1#INCLD_YN": "N",
                "@d#": "@d1#",
                "@d1#": "dmSearch",
                "@d1#tp": "dm",
            }

            # call the list endpoint (same endpoint as before)
            res = session.post(
                "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doList.do",
                data=payload, timeout=15
            )
            data = res.json()
            print(f"[NDRIMS] 응답 상태: {res.status_code}")
            
            # response structure: dsMain is directly a list
            items = []
            if isinstance(data.get("dsMain"), list):
                items = data.get("dsMain", [])
                print(f"[NDRIMS] dsMain 리스트에서 {len(items)}개 항목 발견")
            elif isinstance(data.get("dsMain"), dict):
                # fallback: dsMain.dsMain 형태
                items = data.get("dsMain", {}).get("dsMain", []) or []
                print(f"[NDRIMS] dsMain.dsMain에서 {len(items)}개 항목 발견")
            
            if not items:
                # try dsSearchDscMain as last resort
                items = data.get("dsSearchDscMain", {}).get("dsSearchDscMain", []) or []
                if items:
                    print(f"[NDRIMS] dsSearchDscMain에서 {len(items)}개 항목 발견 (fallback)")

            if not items:
                print(f"[NDRIMS] ⚠️ 데이터 없음 - 응답 키: {list(data.keys())}")
                return ERR_UNKNOWN_DATA

            lines = ["📂 확인 가능한 과제 목록\n(" + get_kst() + " 기준)\n"]
            for it in items:
                proj_no = it.get("PROJ_NO", "")
                proj_nm = it.get("PROJ_NM", "")
                # 총연구기간: MYY_STR_DT ~ MYY_END_DT
                myy_str = it.get("MYY_STR_DT", "")
                myy_end = it.get("MYY_END_DT", "")
                total_period = f"{myy_str[:4]}-{myy_str[4:6]}-{myy_str[6:]}" if myy_str else ""
                total_period += f" ~ {myy_end[:4]}-{myy_end[4:6]}-{myy_end[6:]}" if myy_end else ""
                total_period = total_period.strip()
                
                # 당해과제기간: STR_DT ~ END_DT
                str_dt = it.get("STR_DT", "")
                stt_close = it.get("END_DT", "")
                this_period = f"{str_dt[:4]}-{str_dt[4:6]}-{str_dt[6:]}" if str_dt else ""
                this_period += f" ~ {stt_close[:4]}-{stt_close[4:6]}-{stt_close[6:]}" if stt_close else ""
                this_period = this_period.strip()
                
                chief = it.get("MAIN_OFFI_MEMB_NM", "")
                this_amount = it.get("ALL_RES_AMT", "")
                # 금액 포맷
                if this_amount:
                    this_amount = f"{int(this_amount):,}원"

                lines.append(f"• 과제번호: {proj_no}")
                lines.append(f"  과제명: {proj_nm}")
                lines.append(f"  총연구기간: {total_period}")
                lines.append(f"  당해과제기간: {this_period}")
                lines.append(f"  과제담당자: {chief}")
                lines.append(f"  당해연구비: {this_amount}")
                lines.append("")

            result = "\n".join(lines)
            cache_save(action, result)
            return result

        # ── 과제 상세 정보 ──────────────────────────────
        elif action in ("unsettled_details", "budget_balance", "research_details"):
            if action == "unsettled_details":
                payload = {
                "_runningNana": RsbPrjfi100_nana,
                "_runningLoginIdenNo": rv[1],
                "_runningMainOpenKey": rv[2],
                "@d1#PROJ_NO": "S2026A043400091", # [국고] [2단계] 지속가능한 학교안전 생태계 조성을 위한 연구[1/4]
                "@d1#CARD_DIV_CD": "RE488.10",
                "@d1#SELF_PROC_YN": "N",
                "@d#": "@d1#", "@d1#": "dmSearchCardMain", "@d1#tp": "dm",
                }
                res = session.post(
                    "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListCardMain.do",
                    data=payload, timeout=10
                )
                data = res.json()
                if isinstance(data.get("dsProjCardMain"), list):
                    items = data.get("dsProjCardMain", [])
                    print(f"[NDRIMS] dsProjCardMain 리스트에서 {len(items)}개 항목 발견")
                elif isinstance(data.get("dsProjCardMain"), dict):
                    # fallback: dsProjCardMain.dsProjCardMain 형태
                    items = data.get("dsProjCardMain", {}).get("dsProjCardMain", []) or []
                    print(f"[NDRIMS] dsProjCardMain.dsProjCardMain에서 {len(items)}개 항목 발견")
                lines = ["📊 미정산 내역\n(" + get_kst() + " 기준)\n"]
                lines.append("• 연구비카드\n")
                for it in items:
                    # 카드사용자명, 승인일시, 금액
                    card_user_nm = it.get("CARD_USER_NM", "(소유자)")
                    card_aprv_dttms = it.get("CARD_APRV_DTTMS", "(승인일시)")
                    aprv_amt = it.get("APRV_AMT", "(승인금액)")
                    lines.append(f"- {card_aprv_dttms}\n{card_user_nm} 카드로 {aprv_amt}원 사용")
                
                payload = {
                "_runningNana": RsbPrjfi100_nana,
                "_runningLoginIdenNo": rv[1],
                "_runningMainOpenKey": rv[2],
                "@d1#PROJ_NO": "S2026G999900003", # [대응] [2단계] 지속가능한 학교안전 생태계 조성을 위한 연구[1/4]
                "@d1#CARD_DIV_CD": "RE488.20",
                "@d1#STR_DT": "20260301",
                "@d1#END_DT": "20270228",
                "@d1#CAMPUS_CD": "CM030.10",
                "@d1#SELF_PROC_YN": "N",
                "@d#": "@d1#", "@d1#": "dmSearchCorpCardMain", "@d1#tp": "dm",
                }
                res = session.post(
                    "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListCorpCardMain.do",
                    data=payload, timeout=10
                )
                data = res.json()
                if isinstance(data.get("dsCorpCardMain"), list):
                    items = data.get("dsCorpCardMain", [])
                    print(f"[NDRIMS] dsCorpCardMain 리스트에서 {len(items)}개 항목 발견")
                elif isinstance(data.get("dsCorpCardMain"), dict):
                    # fallback: dsCorpCardMain.dsCorpCardMain 형태
                    items = data.get("dsCorpCardMain", {}).get("dsCorpCardMain", []) or []
                    print(f"[NDRIMS] dsCorpCardMain.dsCorpCardMain에서 {len(items)}개 항목 발견")
                lines.append(f"\n• 법인카드\n")
                for it in items:
                    # 카드사용자명, 승인일시, 금액
                    card_user_nm = it.get("CARD_USER_NM", "(소유자)")
                    card_aprv_dttms = it.get("CARD_APRV_DTTMS", "(승인일시)")
                    aprv_amt = it.get("APRV_AMT", "(승인금액)")
                    lines.append(f"- {card_aprv_dttms}\n{card_user_nm} 카드로 {aprv_amt}원 사용")
                result = "\n".join(lines)
                cache_save(action, result)
                return result

            elif action == "budget_balance":
                payload = {
                    "_runningNana": RsbPrjfi100_nana,
                    "_runningLoginIdenNo": rv[1],
                    "_runningMainOpenKey": rv[2],
                    "@d1#PROJ_NO": "S2026A043400091", # 국고
                    "@d1#SEARCH_GB": "1",
                    "@d1#GRD_GB": "2",
                    "@d1#SUPP_DIV_CD": "SB115.120",
                    "@d#": "@d1#", "@d1#": "dmSearchBdgt", "@d1#tp": "dm",
                }
                res = session.post(
                    "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListBdgt.do",
                    data=payload, timeout=10
                )
                data = res.json()
                if isinstance(data.get("dsBdgt"), list):
                    items = data.get("dsBdgt", [])
                elif isinstance(data.get("dsBdgt"), dict):
                    items = data.get("dsBdgt", {}).get("dsBdgt", []) or []
                else:
                    items = []

                lines = ["💰 실행예산잔액\n(" + get_kst() + " 기준)\n"]
                totals = {k: 0 for k in ("CARY_AMT", "BDGT_AMT", "NOW_BDGT_AMT", "USE_AMT", "USE_BALC_AMT", "NOT_USE_AMT", "NOT_USE_BALC_AMT")}
                lines.append("• 연구비카드\n")
                for it in items:
                    def a(key): return int(it.get(key, 0) or 0)
                    lines.append(f"비목명: {it.get('PROJ_ITEM_NM', '')}")
                    lines.append(f"이월액: {a('CARY_AMT'):,}원")
                    lines.append(f"예산액: {a('BDGT_AMT'):,}원")
                    lines.append(f"예산총액: {a('NOW_BDGT_AMT'):,}원")
                    lines.append(f"집행액: {a('USE_AMT'):,}원")
                    lines.append(f"잔액: {a('USE_BALC_AMT'):,}원")
                    lines.append(f"진행중: {a('NOT_USE_AMT'):,}원")
                    lines.append(f"진행포함잔액: {a('NOT_USE_BALC_AMT'):,}원\n")
                    for k in totals:
                        totals[k] += int(it.get(k, 0) or 0)
                lines.append("↓\n")
                lines.append(f"총비목명: 합계 ({len(items)}건)")
                lines.append(f"총이월액: {totals['CARY_AMT']:,}원")
                lines.append(f"총예산액: {totals['BDGT_AMT']:,}원")
                lines.append(f"총예산총액: {totals['NOW_BDGT_AMT']:,}원")
                lines.append(f"총집행액: {totals['USE_AMT']:,}원")
                lines.append(f"총잔액: {totals['USE_BALC_AMT']:,}원")
                lines.append(f"총진행중: {totals['NOT_USE_AMT']:,}원")
                lines.append(f"총진행포함잔액: {totals['NOT_USE_BALC_AMT']:,}원")
                payload = {
                    "_runningNana": RsbPrjfi100_nana,
                    "_runningLoginIdenNo": rv[1],
                    "_runningMainOpenKey": rv[2],
                    "@d1#PROJ_NO": "S2026G999900003", # 대응
                    "@d1#SEARCH_GB": "1",
                    "@d1#GRD_GB": "2",
                    "@d1#SUPP_DIV_CD": "SB115.60",
                    "@d#": "@d1#", "@d1#": "dmSearchBdgt", "@d1#tp": "dm",
                }
                res = session.post(
                    "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListBdgt.do",
                    data=payload, timeout=10
                )
                data = res.json()
                if isinstance(data.get("dsBdgt"), list):
                    items = data.get("dsBdgt", [])
                elif isinstance(data.get("dsBdgt"), dict):
                    items = data.get("dsBdgt", {}).get("dsBdgt", []) or []
                else:
                    items = []
                totals = {k: 0 for k in ("CARY_AMT", "BDGT_AMT", "NOW_BDGT_AMT", "USE_AMT", "USE_BALC_AMT", "NOT_USE_AMT", "NOT_USE_BALC_AMT")}
                lines.append("\n• 법인카드\n")
                for it in items:
                    def a(key): return int(it.get(key, 0) or 0)
                    lines.append(f"비목명: {it.get('PROJ_ITEM_NM', '')}")
                    lines.append(f"이월액: {a('CARY_AMT'):,}원")
                    lines.append(f"예산액: {a('BDGT_AMT'):,}원")
                    lines.append(f"예산총액: {a('NOW_BDGT_AMT'):,}원")
                    lines.append(f"집행액: {a('USE_AMT'):,}원")
                    lines.append(f"잔액: {a('USE_BALC_AMT'):,}원")
                    lines.append(f"진행중: {a('NOT_USE_AMT'):,}원")
                    lines.append(f"진행포함잔액: {a('NOT_USE_BALC_AMT'):,}원\n")
                    for k in totals:
                        totals[k] += int(it.get(k, 0) or 0)
                lines.append("↓\n")
                lines.append(f"총비목명: 합계 ({len(items)}건)")
                lines.append(f"총이월액: {totals['CARY_AMT']:,}원")
                lines.append(f"총예산액: {totals['BDGT_AMT']:,}원")
                lines.append(f"총예산총액: {totals['NOW_BDGT_AMT']:,}원")
                lines.append(f"총집행액: {totals['USE_AMT']:,}원")
                lines.append(f"총잔액: {totals['USE_BALC_AMT']:,}원")
                lines.append(f"총진행중: {totals['NOT_USE_AMT']:,}원")
                lines.append(f"총진행포함잔액: {totals['NOT_USE_BALC_AMT']:,}원")
                result = "\n".join(lines)
                cache_save(action, result)
                return result

            elif action == "research_details":
                def fetch_groups(proj_no: str) -> dict:
                    p = {
                        "_runningNana": RsbPrjfi100_nana,
                        "_runningLoginIdenNo": rv[1],
                        "_runningMainOpenKey": rv[2],
                        "@d1#PROJ_NO": proj_no,
                        "@d1#APRV_INCLD_YN": "Y",
                        "@d#": "@d1#", "@d1#": "dmSearchDscMain", "@d1#tp": "dm",
                    }
                    r = session.post(
                        "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListDscMain.do",
                        data=p, timeout=10
                    )
                    d = r.json()
                    if isinstance(d.get("dsDscMain"), list):
                        raw = d.get("dsDscMain", [])
                    elif isinstance(d.get("dsDscMain"), dict):
                        raw = d.get("dsDscMain", {}).get("dsDscMain", []) or []
                    else:
                        raw = []
                    grps: dict[str, dict] = {}
                    for it in raw:
                        if it.get("PROJ_ITEM_NM") != "연구활동비":
                            continue
                        typ = it.get("EXPS_TYP_CD_NM") or "기타"
                        if typ not in grps:
                            grps[typ] = {"count": 0, "amount": 0}
                        grps[typ]["count"] += 1
                        grps[typ]["amount"] += int(it.get("USE_AMT", 0) or 0)
                    return grps

                def append_section(lines: list, label: str, grps: dict) -> tuple[int, int]:
                    lines.append(f"• {label}\n")
                    sc = sa = 0
                    for typ, g in grps.items():
                        lines.append(f"{typ}: 총 {g['count']}건 {g['amount']:,}원")
                        sc += g["count"]; sa += g["amount"]
                    lines.append(f"총사용금액: 총 {sc}건 {sa:,}원\n")
                    return sc, sa

                lines = ["📋 연구활동비 상세\n(" + get_kst() + " 기준)\n"]
                c1, a1 = append_section(lines, "연구비카드", fetch_groups("S2026A043400091"))
                c2, a2 = append_section(lines, "법인카드",   fetch_groups("S2026G999900003"))
                lines.append(f"합계: 총 {c1+c2}건 {a1+a2:,}원")
                result = "\n".join(lines)
                cache_save(action, result)
                return result

        else:
            return f"⚠️ '{action}' 기능은 아직 구현되지 않았어요."

    except Exception as e:
        text, _ = cache_load(action)
        if text:
            return text
        return f"⚠️ NDRIMS 조회 중 오류가 발생했어요.\n{type(e).__name__}: {e}"


# ── 쿠키 갱신 처리 ───────────────────────────────────────────
# 세션별로 쿠키 입력 대기 상태를 기록
cookie_pending: dict[str, str] = {}   # user_id -> "password" | "cookie"

def looks_like_cookie_input(utterance: str) -> bool:
    text = utterance.strip().upper()
    return "WMONID" in text or "JSESSIONID" in text


def handle_cookie_input(user_id: str, utterance: str) -> str | None:
    """
    쿠키 갱신 대화 흐름:
    1) 버튼 클릭 → 비밀번호 입력 대기
    2) 비밀번호 확인 후 쿠키 입력 대기
    3) 다음 발화에서 'WMONID=xxx JSESSIONID=yyy' 파싱 → 저장
    """
    stage = cookie_pending.get(user_id)
    if stage:
        if not utterance.strip():
            return None

        if utterance.strip() in ("처음으로", "시작", "메인", "홈"):
            cookie_pending.pop(user_id, None)
            return None

        if stage == "password":
            if not ADMIN_PASSWORD:
                cookie_pending.pop(user_id, None)
                return MSG_NO_ADMIN_PW

            if utterance.strip() == ADMIN_PASSWORD:
                cookie_pending[user_id] = "cookie"
                return MSG_PASSWORD_OK

            if not looks_like_cookie_input(utterance):
                cookie_pending.pop(user_id, None)
                return None

            cookie_pending.pop(user_id, None)
            return MSG_WRONG_PASSWORD

        if not looks_like_cookie_input(utterance):
            cookie_pending.pop(user_id, None)
            return None

        # 입력 파싱 시도
        m_wmon = re.search(r"WMONID\s*=\s*(\S+)", utterance, re.I)
        m_sess = re.search(r"JSESSIONID\s*=\s*(\S+)", utterance, re.I)
        if m_wmon and m_sess:
            cookie_store["WMONID"]     = m_wmon.group(1)
            cookie_store["JSESSIONID"] = m_sess.group(1)
            save_cookies()
            cookie_pending.pop(user_id, None)
            return (
                "✅ 쿠키가 갱신됐어요!\n\n"
                f"WMONID: {cookie_store['WMONID'][:8]}…\n"
                f"JSESSIONID: {cookie_store['JSESSIONID'][:8]}…"
            )
        else:
            return MSG_WRONG_FORMAT
    return None   # 쿠키 대기 상태 아님


# ── 카카오 응답 빌더 ─────────────────────────────────────────
MAX_QUICK = 10   # 카카오 퀵리플라이 최대 10개

def quick(label: str, msg: str) -> dict:
    return {"label": label[:14], "action": "message", "messageText": msg}

def basic_card(title: str, description: str) -> dict:
    card = {"title": title, "description": description}
    if _base_url and (BASE_DIR / "thumbnail.jpg").exists():
        card["thumbnail"] = {"imageUrl": f"{_base_url}/thumbnail.jpg"}
    return card

def make_main_menu() -> dict:
    """메인 메뉴: 1단계 버튼들."""
    replies = [quick(v["label"], v["label"]) for k, v in MENU.items()]
    replies = replies[:MAX_QUICK]
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"basicCard": basic_card("학교종합안전연구소", "무엇을 도와드릴까요?")}],
            "quickReplies": replies,
        }
    }

def make_submenu(node: dict, title: str) -> dict:
    children = list(node["children"].values())
    card_title = title.split(" > ")[-1]

    if len(children) < 6:
        replies = [quick(c["label"], c["label"]) for c in children]
        replies.append(quick("🏠 처음으로", "처음으로"))
        return {
            "version": "2.0",
            "template": {
                "outputs": [{"basicCard": basic_card(card_title, "항목을 선택해 주세요.")}],
                "quickReplies": replies[:MAX_QUICK],
            }
        }

    btns = [
        {"label": c["label"][:14], "action": "message", "messageText": c["label"]}
        for c in children
    ]
    btns.append({"label": "🏠 처음으로", "action": "message", "messageText": "처음으로"})
    cards = [
        {**basic_card(card_title, "항목을 선택해 주세요."), "buttons": btns[i:i+3]}
        for i in range(0, len(btns), 3)
    ]
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"carousel": {"type": "basicCard", "items": cards}}],
        }
    }

def make_text_response(text: str, back_label: str = "🏠 처음으로") -> dict:
    btn = {"action": "message", "label": back_label[:14], "messageText": "처음으로"}
    if len(text) <= 400:
        return {
            "version": "2.0",
            "template": {
                "outputs": [{"textCard": {"text": text, "buttons": [btn]}}],
            }
        }
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"simpleText": {"text": text}}],
            "quickReplies": [quick(back_label, "처음으로")],
        }
    }

def make_waiting() -> dict:
    return {"version": "2.0", "useCallback": True, "data": {}}


# ── 메뉴 ID 찾기 ─────────────────────────────────────────────
def find_node(menu_id: str) -> tuple[dict | None, list[str]]:
    """menu_id (예: 'A2a') 로 노드와 부모 경로를 찾아 반환."""
    # 1단계
    if menu_id in MENU:
        return MENU[menu_id], [MENU[menu_id]["label"]]
    # 2단계
    for _, v1 in MENU.items():
        if menu_id in v1["children"]:
            node = v1["children"][menu_id]
            return node, [v1["label"], node["label"]]
        # 3단계
        for k2, v2 in v1["children"].items():
            if menu_id in v2["children"]:
                node = v2["children"][menu_id]
                return node, [v1["label"], v2["label"], node["label"]]
    return None, []


# ── 핵심 라우터 ──────────────────────────────────────────────
def route(utterance: str, user_id: str) -> dict:
    utt = utterance.strip()

    # 빈 발화는 메인 메뉴 표시
    if not utt:
        return make_main_menu()

    if utt in ("처음으로", "시작", "메인", "홈"):
        return make_main_menu()

    # 쿠키 대기 중이면 먼저 처리
    cookie_resp = handle_cookie_input(user_id, utt)
    if cookie_resp is not None:
        return make_text_response(cookie_resp)

    # 버튼 클릭 — "__MENU__ID" 형태
    if utt.startswith("__MENU__"):
        menu_id = utt[len("__MENU__"):]
        node, path = find_node(menu_id)

    # 버튼 텍스트가 label로 들어오는 경우
    elif utt in MENU_LABEL_INDEX:
        menu_id = MENU_LABEL_INDEX[utt]
        node, path = find_node(menu_id)
    else:
        node, path = None, []

    if node is None:
        return make_text_response(ERR_MENU_NOT_FOUND)

    atype  = node.get("action_type")
    action = node.get("action")

    # 하위 메뉴가 있으면 서브메뉴 표시
    if node["children"]:
        return make_submenu(node, " > ".join(path))

    # TEXT 응답
    if atype == "TEXT":
        text = TEXT_RESPONSES.get(action, ERR_UNKNOWN_ATYPE)
        return make_text_response(text)

    # COOKIE 갱신
    if atype == "COOKIE":
        cookie_pending[user_id] = "password"
        return make_text_response(MSG_COOKIE_START)

    # NDRIMS API — 콜백으로 처리 (시간이 걸릴 수 있으므로)
    if atype == "NDRIMS":
        return None  # 콜백 필요 신호

    return make_text_response(ERR_UNKNOWN_ATYPE)


# ── HTTP 서버 ────────────────────────────────────────────────
def send_callback(url: str, payload: dict):
    data = json.dumps(payload, ensure_ascii=False).encode()
    req = urllib.request.Request(url, data=data,
                                  headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            print(f"[CB] {r.status}")
    except Exception as e:
        print(f"[CB ERR] {e}")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[{self.address_string()}] {fmt % args}")

    def do_GET(self):
        self._detect_base_url()
        if self.path == "/thumbnail.jpg":
            img_path = BASE_DIR / "thumbnail.jpg"
            if img_path.exists():
                data = img_path.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            self._json(404, {"error": "not found"})
            return
        self._json(200, {"status": "ok", "message": "nDRIMS 봇 실행 중"})

    def _detect_base_url(self):
        global _base_url
        if not _base_url:
            host = self.headers.get("Host", "")
            if host:
                _base_url = f"https://{host}"

    def do_POST(self):
        self._detect_base_url()
        if self.path != "/skill":
            self._json(404, {"error": "not found"}); return

        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length))

        utterance    = body.get("userRequest", {}).get("utterance", "").strip()
        user_id      = body.get("userRequest", {}).get("user", {}).get("id", "anonymous")
        callback_url = body.get("userRequest", {}).get("callbackUrl", "")
        # print(f"[CB URL] {callback_url[:40] if callback_url else '없음'}")

        print(f"[USER] {user_id[:8]}… | {utterance[:60]}")

        response = route(utterance, user_id)

        # NDRIMS 호출이 필요한 경우 (route가 None 반환)
        if response is None:
            # resolve menu id from either internal token or visible label
            if utterance.startswith("__MENU__"):
                menu_id = utterance[len("__MENU__"):]
            elif utterance in MENU_LABEL_INDEX:
                menu_id = MENU_LABEL_INDEX[utterance]
            else:
                self._json(200, make_text_response("⚠️ 메뉴를 찾을 수 없어요."))
                return

            node, _ = find_node(menu_id)
            if not node:
                self._json(200, make_text_response("⚠️ 메뉴를 찾을 수 없어요."))
                return

            action = node.get("action", "")

            if callback_url:
                self._json(200, make_waiting())
                _action, _url = action, callback_url
                def bg():
                    try:
                        text = call_ndrims(_action)
                    except Exception as e:
                        cached, _ = cache_load(_action)
                        text = cached or f"⚠️ 오류가 발생했어요.\n{type(e).__name__}: {e}"
                    send_callback(_url, make_text_response(text))
                threading.Thread(target=bg, daemon=True).start()
            else:
                text = call_ndrims(action)
                self._json(200, make_text_response(text))
            return

        self._json(200, response)

    def _json(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    print(f"✅ nDRIMS 챗봇 서버 시작 — http://0.0.0.0:{port}/skill")
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()

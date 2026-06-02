"""
NDRIMS 카카오 챗봇 스킬 서버
- 메뉴구조.xlsx 를 읽어 카카오 버튼 메뉴 자동 생성
- 버튼 클릭 → NDRIMS API 호출 → 텍스트/카드 응답
- 카톡에서 쿠키값 직접 갱신 가능
"""

import json, os, threading, urllib.request, re
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

# openpyxl 은 pip install openpyxl 로 설치
try:
    import openpyxl
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False
    print("[경고] openpyxl 없음 — pip install openpyxl")

try:
    import requests as req_lib
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False
    print("[경고] requests 없음 — pip install requests")

BASE_DIR = Path(__file__).parent

# ── 쿠키 상태 (메모리 보관, 서버 재시작 시 초기화) ─────────────
cookie_store = {
    "WMONID":     os.environ.get("WMONID", ""),
    "JSESSIONID": os.environ.get("JSESSIONID", ""),
}

# ── 규정 텍스트 (간단하게 딕셔너리로 관리) ─────────────────────
TEXT_RESPONSES = {
    "budget_rule": (
        "💰 연구비 사용 한도\n\n"
        "• 50만 원 이하 → 팀장 승인\n"
        "• 50만 원 초과 → 소장 결재\n"
        "• 100만 원 초과 → 사전 품의서 필수\n\n"
        "📌 근거: 연구비관리규정 제7조"
    ),
    "travel_domestic": (
        "🇰🇷 국내 출장 규정\n\n"
        "• 일비: 연구원 2만 원 / 책임 3만 원\n"
        "• 숙박: 실비 (상한 7만 원/박)\n"
        "• 서울 출장: 당일 복귀 원칙\n"
        "• 신청: 출발 3일 전까지\n\n"
        "📌 근거: 출장여비규정 제3·5조"
    ),
    "travel_overseas": (
        "✈️ 해외 출장 규정\n\n"
        "• 신청: 출발 7일 전까지\n"
        "• 항공: 이코노미 원칙\n"
        "  (14시간 초과 시 비즈니스 허용)\n"
        "• 일비: 국가별 기준표 적용\n\n"
        "📌 근거: 출장여비규정 제8조"
    ),
    "equipment_rule": (
        "🔧 공용 장비 예약\n\n"
        "• 사용 2일 전까지 시스템 예약 필수\n"
        "• 예약 없이 사용 불가\n"
        "• 미사용 취소: 24시간 전까지\n"
        "• 고장 발견 시 즉시 내선 234 신고\n\n"
        "📌 근거: 장비관리규정 제4·10조"
    ),
    "help": (
        "📖 사용법\n\n"
        "1️⃣ 메인 메뉴 버튼을 눌러 탐색\n"
        "2️⃣ 버튼이 없으면 자유롭게 질문\n\n"
        "⚙️ 쿠키 갱신 방법\n"
        "설정 → 쿠키 갱신 버튼 후\n"
        "아래 형식으로 입력하세요:\n"
        "WMONID=값 JSESSIONID=값"
    ),
    "update_cookie": "__COOKIE_INPUT__",  # 특수 처리
}

# ── 엑셀 메뉴 로드 ────────────────────────────────────────────
def load_menu() -> dict:
    """
    메뉴구조.xlsx 를 읽어 트리 구조로 변환.
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
    path = BASE_DIR / "메뉴구조.xlsx"
    if not EXCEL_OK or not path.exists():
        print("[경고] 메뉴구조.xlsx 없음 — 기본 메뉴 사용")
        return {}

    wb = openpyxl.load_workbook(path)
    ws = wb["메뉴구조"]

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

# ── NDRIMS API 호출 ──────────────────────────────────────────
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
    return {
        "_runningNana": nana,
        "_runningLoginIdenNo": user_info["LOGIN_IDEN_NO"],
        "_runningMainOpenKey": user_info["MAIN_OPEN_KEY"],
    }


def call_ndrims(action: str) -> str:
    """액션값에 따라 NDRIMS API 호출 → 텍스트 결과 반환."""
    if not REQUESTS_OK:
        return "⚠️ requests 라이브러리가 없어요.\npip install requests 후 재시작하세요."
    if not cookie_store["JSESSIONID"]:
        return "⚠️ 쿠키가 설정되지 않았어요.\n설정 → 쿠키 갱신에서 먼저 쿠키를 입력해 주세요."

    try:
        session = make_ndrims_session()
        rv = get_running_values(session)

        # ── 과제 목록 ─────────────────────────────────────
        if action == "project_list":
            payload = {
                **rv,
                "@d1#PROJ_NO": "",
                "@d1#USE_YY_MM": "",
                "@d1#APRV_INCLD_YN": "Y",
                "@d#": "@d1#", "@d1#": "dmSearchDscMain", "@d1#tp": "dm",
            }
            res = session.post(
                "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListDscMain.do",
                data=payload, timeout=10
            )
            data = res.json()
            items = data.get("dsSearchDscMain", {}).get("dsSearchDscMain", [])
            if not items:
                return "조회된 과제가 없어요."
            lines = ["📂 진행 중인 과제 목록\n"]
            for it in items[:10]:
                lines.append(f"• [{it.get('PROJ_NO','')}] {it.get('PROJ_NM','')}")
            return "\n".join(lines)

        # ── 과제별 지출 현황 ──────────────────────────────
        elif action in ("expense_by_project", "expense_by_item", "budget_remain"):
            payload = {
                **rv,
                "@d1#PROJ_NO": "",
                "@d1#USE_YY_MM": "",
                "@d1#APRV_INCLD_YN": "Y",
                "@d#": "@d1#", "@d1#": "dmSearchDscMain", "@d1#tp": "dm",
            }
            res = session.post(
                "https://ndrims.dongguk.edu/rs/rsb/prjfi/RsbPrjfi100/doListDscMain.do",
                data=payload, timeout=10
            )
            data = res.json()
            items = data.get("dsSearchDscMain", {}).get("dsSearchDscMain", [])
            if not items:
                return "조회된 데이터가 없어요."

            if action == "expense_by_project":
                lines = ["📊 과제별 지출 현황\n"]
                for it in items[:8]:
                    used  = int(it.get("USE_AMT", 0) or 0)
                    total = int(it.get("TOT_AMT", 0) or 0)
                    pct   = f"{used/total*100:.1f}%" if total else "—"
                    lines.append(
                        f"• {it.get('PROJ_NM','')}\n"
                        f"  지출 {used:,}원 / 예산 {total:,}원 ({pct})"
                    )
                return "\n".join(lines)

            elif action == "budget_remain":
                lines = ["💰 예산 잔액 현황\n"]
                for it in items[:8]:
                    total  = int(it.get("TOT_AMT", 0) or 0)
                    used   = int(it.get("USE_AMT", 0) or 0)
                    remain = total - used
                    lines.append(
                        f"• {it.get('PROJ_NM','')}\n"
                        f"  잔액 {remain:,}원"
                    )
                return "\n".join(lines)

            else:  # expense_by_item
                return "📋 항목별 지출 조회는 현재 개발 중이에요."

        else:
            return f"⚠️ '{action}' 액션은 아직 구현되지 않았어요."

    except Exception as e:
        return f"⚠️ NDRIMS 조회 중 오류가 발생했어요.\n{type(e).__name__}: {e}"


# ── 쿠키 갱신 처리 ───────────────────────────────────────────
# 세션별로 쿠키 입력 대기 상태를 기록
cookie_pending: set = set()   # user_id 집합

def handle_cookie_input(user_id: str, utterance: str) -> str | None:
    """
    쿠키 갱신 대화 흐름:
    1) 버튼 클릭 → "COOKIE_WAIT" 상태 진입 → 입력 안내 반환
    2) 다음 발화에서 'WMONID=xxx JSESSIONID=yyy' 파싱 → 저장
    """
    if user_id in cookie_pending:
        # 입력 파싱 시도
        m_wmon = re.search(r"WMONID\s*=\s*(\S+)", utterance, re.I)
        m_sess = re.search(r"JSESSIONID\s*=\s*(\S+)", utterance, re.I)
        if m_wmon and m_sess:
            cookie_store["WMONID"]     = m_wmon.group(1)
            cookie_store["JSESSIONID"] = m_sess.group(1)
            cookie_pending.discard(user_id)
            return (
                "✅ 쿠키가 갱신됐어요!\n\n"
                f"WMONID: {cookie_store['WMONID'][:8]}…\n"
                f"JSESSIONID: {cookie_store['JSESSIONID'][:8]}…"
            )
        else:
            cookie_pending.discard(user_id)
            return (
                "❌ 형식이 맞지 않아요. 다시 시도해 주세요.\n\n"
                "올바른 형식:\n"
                "WMONID=값 JSESSIONID=값"
            )
    return None   # 쿠키 대기 상태 아님


# ── 카카오 응답 빌더 ─────────────────────────────────────────
MAX_QUICK = 10   # 카카오 퀵리플라이 최대 10개

def quick(label: str, msg: str) -> dict:
    return {"label": label[:14], "action": "message", "messageText": msg}

def make_main_menu() -> dict:
    """메인 메뉴: 1단계 버튼들."""
    replies = [quick(v["label"], f"__MENU__{k}") for k, v in MENU.items()]
    replies = replies[:MAX_QUICK]
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"simpleText": {"text": "🏠 메인 메뉴\n무엇을 도와드릴까요?"}}],
            "quickReplies": replies,
        }
    }

def make_submenu(node: dict, title: str) -> dict:
    """2단계 버튼 목록."""
    replies = []
    for cid, child in node["children"].items():
        replies.append(quick(child["label"], f"__MENU__{cid}"))
    replies.append(quick("🏠 처음으로", "처음으로"))
    replies = replies[:MAX_QUICK]
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"simpleText": {"text": f"{title}\n항목을 선택해 주세요."}}],
            "quickReplies": replies,
        }
    }

def make_text_response(text: str, back_label: str = "🏠 처음으로") -> dict:
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
    for k1, v1 in MENU.items():
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

    # 처음으로
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

        if node is None:
            return make_text_response("⚠️ 메뉴를 찾을 수 없어요.")

        atype  = node.get("action_type")
        action = node.get("action")

        # 하위 메뉴가 있으면 서브메뉴 표시
        if node["children"]:
            return make_submenu(node, " > ".join(path))

        # TEXT 응답
        if atype == "TEXT":
            text = TEXT_RESPONSES.get(action, f"'{action}' 응답이 아직 없어요.")
            return make_text_response(text)

        # COOKIE 갱신
        if atype == "COOKIE":
            cookie_pending.add(user_id)
            return make_text_response(
                "🔐 쿠키 갱신\n\n"
                "아래 형식으로 입력해 주세요:\n\n"
                "WMONID=값 JSESSIONID=값\n\n"
                "브라우저 개발자 도구(F12) → Application\n"
                "→ Cookies 에서 값을 복사하세요."
            )

        # NDRIMS API — 콜백으로 처리 (시간이 걸릴 수 있으므로)
        if atype == "NDRIMS":
            return None  # 콜백 필요 신호

        return make_text_response("⚠️ 알 수 없는 액션 타입이에요.")

    # 자유 발화 — 메인 메뉴로 안내
    return make_main_menu()


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
        self._json(200, {"status": "ok", "message": "NDRIMS 규정봇 실행 중"})

    def do_POST(self):
        if self.path != "/skill":
            self._json(404, {"error": "not found"}); return

        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length))

        utterance    = body.get("userRequest", {}).get("utterance", "").strip()
        user_id      = body.get("userRequest", {}).get("user", {}).get("id", "anonymous")
        callback_url = body.get("callback_url", "")

        print(f"[Q] {user_id[:8]}… | {utterance[:60]}")

        response = route(utterance, user_id)

        # NDRIMS 호출이 필요한 경우 (route가 None 반환)
        if response is None:
            menu_id = utterance[len("__MENU__"):]
            node, _ = find_node(menu_id)
            action  = node["action"] if node else ""

            if callback_url:
                self._json(200, make_waiting())
                def bg():
                    text = call_ndrims(action)
                    send_callback(callback_url, make_text_response(text))
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
    print(f"✅ NDRIMS 규정봇 서버 시작 — http://0.0.0.0:{port}/skill")
    print(f"   쿠키 상태: WMONID={'설정됨' if cookie_store['WMONID'] else '미설정'}")
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()

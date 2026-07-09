# =============================================================================
# 🐾 펫패스(PetPass) - 반려동물 동반 출입 사전 확인 웹앱
# =============================================================================
#
# [이 앱이 하는 일]
# 사용자가 자기 반려동물의 무게/크기를 입력하면, 한국관광공사의
# '반려동물 동반여행 정보(KorPetTourService2)' 공공데이터를 불러와서
# 가려는 장소의 동반 조건과 대조한 뒤,
#   🐾 가능 / 🐾 조건부 / 🐾 불가  를 신호등처럼 판정해 줍니다.
#
# -----------------------------------------------------------------------------
# [설치 방법] 아래 명령을 터미널(명령 프롬프트)에 붙여넣어 실행하세요.
#   pip install streamlit requests pandas
#
# [실행 방법] 이 파일이 있는 폴더에서 아래 명령을 실행하세요.
#   streamlit run app.py
# -----------------------------------------------------------------------------


# --- 1. 필요한 도구(라이브러리) 불러오기 ------------------------------------
import re                      # 정규식: 텍스트에서 숫자(무게 등)를 뽑아낼 때 사용
import requests                # 인터넷으로 공공데이터 API를 호출할 때 사용
import pandas as pd            # 표(엑셀 같은 데이터) 형태로 정리/필터/정렬할 때 사용
import streamlit as st         # 웹앱 화면을 만드는 핵심 라이브러리


# --- 2. 기본 설정값(상수) ----------------------------------------------------

# 🔑 인증키(serviceKey) 불러오기
#   - 실제 키는 코드가 아니라 '.streamlit/secrets.toml'(로컬) 또는
#     Streamlit Cloud의 [Secrets] 설정(배포)에 보관합니다.
#     → 이렇게 하면 GitHub에 키가 노출되지 않아 안전합니다.
#   - 아래는 secrets 에서 먼저 찾고, 없으면 기본값(placeholder)을 쓰는 구조입니다.
#_DEFAULT_KEY = "a259c8267267163027ba8e9a1beec3a2973d0f1ad8e7a885d2dc8b9f46be0fde"
try:
    # secrets.toml 이나 Cloud Secrets 에 SERVICE_KEY 가 있으면 그 값을 사용
    SERVICE_KEY = st.secrets["SERVICE_KEY"]
except Exception:
    # secrets 가 없을 때만 기본값 사용 (배포 전 임시 확인용)
    SERVICE_KEY = _DEFAULT_KEY

# 공공데이터 API의 기본 주소(Base URL)
# ※ 구버전 KorPetTourService 는 폐지되었고, 현재는 v2(KorPetTourService2) 를 씁니다.
BASE_URL = "http://apis.data.go.kr/B551011/KorPetTourService2"

# 우리 앱이 공공데이터포털에 스스로를 소개하는 이름(아무 값이나 가능)
MOBILE_OS = "ETC"          # 운영체제 구분(대부분 ETC로 둠)
MOBILE_APP = "PetPass"     # 앱 이름

# 지역 선택용 코드표 (한국관광공사 areaCode 기준의 대표적인 시/도 코드)
# key = 화면에 보여줄 이름, value = API에 넘길 지역 코드 숫자
AREA_CODES = {
    "서울": 1, "인천": 2, "대전": 3, "대구": 4, "광주": 5,
    "부산": 6, "울산": 7, "세종": 8, "경기": 31, "강원": 32,
    "충북": 33, "충남": 34, "경북": 35, "경남": 36,
    "전북": 37, "전남": 38, "제주": 39,
}

# 견종/종류 드롭다운 목록 (자주 찾는 종류 약 10가지 + 직접 입력)
# 목록에 없는 종류는 마지막 '기타 (직접 입력)'을 골라 타이핑할 수 있습니다.
BREED_OPTIONS = [
    "말티즈",
    "푸들 / 토이푸들",
    "포메라니안",
    "시츄",
    "치와와",
    "비숑 프리제",
    "웰시코기",
    "진돗개",
    "골든/래브라도 리트리버",
    "고양이 (코리안숏헤어 등)",
    "기타 (직접 입력)",
]

# 관광 유형 코드표 (한국관광공사 contentTypeId 기준)
# key = 화면에 보여줄 이름, value = API에 넘길 유형 코드
CONTENT_TYPES = {
    "전체": None,
    "관광지": 12,
    "문화시설": 14,
    "축제/행사": 15,
    "여행코스": 25,
    "레포츠": 28,
    "숙소": 32,
    "쇼핑": 38,
    "음식점": 39,
}

# detailPetTour2 응답의 '실제' 필드명 모음 (한국관광공사 KorPetTourService2 기준).
# key = API 필드명, value = 화면에 보여줄 한글 이름.
# 값이 바뀌면 여기만 고치면 앱 전체(판정/상세카드)에 반영됩니다.
PET_FIELDS = {
    "acmpyTypeCd":      "동반 유형",              # 예: '전구역 동반가능'
    "acmpyPsblCpam":    "동반 가능 반려동물",       # 예: '반려견, 반려묘'
    "acmpyNeedMtr":     "동반 시 필요사항(준비물)",  # 예: '목줄 착용'
    "etcAcmpyInfo":     "기타 동반 정보",
    "relaAcdntRiskMtr": "관련 사고 위험 사항",
    "relaPosesFclty":   "관련 구비 시설",
    "relaFrnshPrdlst":  "관련 비치 품목",
    "relaPurcPrdlst":   "관련 구매 품목",
    "relaRntlPrdlst":   "관련 대여 품목",
}


# --- 3. 신호등(색상) 관련 정의 ----------------------------------------------

# 판정 결과별로 (표시 문구, 색상코드)를 미리 정해 둡니다.
# 색상코드(green/orange/red)는 상세 안내 박스 등에 사용됩니다.
SIGNAL_STYLE = {
    "가능":   {"emoji": "🐾 가능",   "color": "green",  "hex": "#2ecc71"},
    "조건부": {"emoji": "🐾 조건부", "color": "orange", "hex": "#f39c12"},
    "불가":   {"emoji": "🐾 불가",   "color": "red",    "hex": "#e74c3c"},
}


# --- 4-0. 작은 도우미 함수들 -------------------------------------------------

def _to_float(value):
    """문자열 좌표를 실수로 바꿔 주는 작은 도우미 함수(변환 실패 시 None)."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def pick(d, *keys, default=""):
    """
    딕셔너리(d)에서 여러 '후보 키(keys)' 중 값이 실제로 들어 있는
    첫 번째 값을 돌려줍니다. 하나도 없으면 default 반환.

    → API 필드명이 문서와 조금 달라도 앱이 견디도록 하는 안전장치입니다.
    """
    if not isinstance(d, dict):
        return default
    for k in keys:
        v = d.get(k)
        if v not in (None, "", " "):
            return v
    return default


def find_by_keyword(d, *substrings):
    """
    딕셔너리(d)에서 '키 이름'에 특정 문자열이 포함된 첫 값을 돌려줍니다.
    (필드명이 유형마다 조금씩 달라서 부분일치로 찾습니다.)
    예) 주차: parking / parkinglodging / parkingfood ... → 'parking'으로 한 번에
    """
    if not isinstance(d, dict):
        return ""
    for k, v in d.items():
        if any(s in k.lower() for s in substrings) and v not in (None, "", " "):
            return v
    return ""


def strip_html(text):
    """홈페이지 등에 섞여 오는 <a href=...>같은 HTML 태그를 제거해 순수 텍스트만 남깁니다."""
    if not text:
        return ""
    # <...> 형태의 태그를 모두 제거
    return re.sub(r"<[^>]+>", " ", str(text)).strip()


# =============================================================================
# 4. 공공데이터 API 호출 함수들
# =============================================================================
#
# @st.cache_data 를 붙이면, 똑같은 검색을 다시 할 때 인터넷에 또 물어보지 않고
# 이전에 받아 둔 결과를 재사용합니다. => API 호출 횟수 절약 + 속도 향상
# (ttl=3600 : 1시간 동안 캐시 유지)

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_area_based_list(area_code, content_type_id, sigungu_code=None, num_rows=100):
    """
    [지역기반 목록 조회] areaBasedList2
    선택한 지역(area_code)·구군(sigungu_code)·유형(content_type_id)에 해당하는
    반려동물 동반 가능 장소 '목록'을 가져옵니다.

    반환값: (장소 리스트, 원본 JSON)  ← 원본은 디버그(필드 확인)용
    """
    # API에 넘길 파라미터(질문 항목)들을 딕셔너리로 준비
    params = {
        "serviceKey": SERVICE_KEY,   # 인증키
        "MobileOS": MOBILE_OS,
        "MobileApp": MOBILE_APP,
        "areaCode": area_code,       # 지역 코드
        "numOfRows": num_rows,       # 한 번에 몇 개까지 받아올지
        "pageNo": 1,                 # 페이지 번호
        "arrange": "A",              # 정렬(A=제목순). 목록이 안정적으로 오도록 지정
        "_type": "json",             # ⭐ 응답을 JSON 형식으로 달라는 요청
    }
    # 유형(관광지/음식점 등)을 선택했을 때만 파라미터에 추가
    if content_type_id is not None:
        params["contentTypeId"] = content_type_id
    # 구/군을 선택했을 때만 파라미터에 추가('전체'면 None 이라 생략)
    if sigungu_code is not None:
        params["sigunguCode"] = sigungu_code

    # 실제 호출 주소 (v2 오퍼레이션 이름은 끝에 '2'가 붙습니다)
    url = f"{BASE_URL}/areaBasedList2"

    # requests 호출은 반드시 try/except 로 감싸서, 실패해도 앱이 죽지 않게 합니다.
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()      # 200이 아니면 여기서 에러 발생
        data = resp.json()           # 응답 본문을 JSON(딕셔너리)으로 변환
    except requests.exceptions.RequestException as e:
        # 네트워크/서버 오류 → 화면에 친절한 메시지, 빈 결과 반환
        st.error(f"😢 목록을 불러오지 못했어요. 잠시 후 다시 시도해 주세요.\n\n(상세: {e})")
        return [], {}
    except ValueError:
        # JSON 변환 실패(응답이 JSON이 아닐 때) → 인증키 문제일 가능성 큼
        st.error("😢 응답 형식이 올바르지 않아요. 인증키가 맞는지 확인해 주세요.")
        return [], {}

    # JSON 안에서 실제 목록(items)이 있는 곳까지 '안전하게' 파고들기
    # 키가 없을 수도 있으니 .get(..., 기본값) 을 계속 사용합니다.
    items = (
        data.get("response", {})
            .get("body", {})
            .get("items", {})
    )
    # items 가 빈 문자열("")로 오는 경우도 있어 방어 (결과 0건일 때)
    if not isinstance(items, dict):
        return [], data

    item_list = items.get("item", [])
    # 결과가 1개면 딕셔너리, 여러 개면 리스트로 오므로 항상 리스트로 통일
    if isinstance(item_list, dict):
        item_list = [item_list]

    return item_list, data


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_sigungu_list(area_code):
    """
    [시군구 코드 조회] areaCode2
    선택한 시/도(area_code)에 속한 '구/군' 목록을 가져옵니다.
    (예: 서울 → 강남구, 강동구, ...)

    반환값: {구이름: 구코드} 딕셔너리 (실패하면 빈 딕셔너리)
    ※ 구/군 목록은 잘 바뀌지 않으므로 하루(86400초) 동안 캐시합니다.
    """
    params = {
        "serviceKey": SERVICE_KEY,
        "MobileOS": MOBILE_OS,
        "MobileApp": MOBILE_APP,
        "areaCode": area_code,   # 이 시/도에 속한 구/군을 달라는 뜻
        "numOfRows": 50,         # 구/군은 많아야 수십 개라 50이면 충분
        "pageNo": 1,
        "_type": "json",
    }
    url = f"{BASE_URL}/areaCode2"

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
    except (requests.exceptions.RequestException, ValueError):
        # 구/군 목록을 못 받아도 앱이 멈추지 않게, 빈 딕셔너리 반환
        return {}

    items = (
        data.get("response", {})
            .get("body", {})
            .get("items", {})
    )
    if not isinstance(items, dict):
        return {}

    item_list = items.get("item", [])
    if isinstance(item_list, dict):
        item_list = [item_list]

    # {구이름: 코드} 형태로 정리 (이름/코드가 모두 있는 항목만)
    result = {}
    for it in item_list:
        name = it.get("name")
        code = it.get("code")
        if name and code is not None:
            result[name] = code
    return result


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_detail_pet_tour(content_id):
    """
    [반려동물 상세조건 조회] detailPetTour2
    특정 장소(content_id)의 '반려동물 동반 상세 조건'을 가져옵니다.
    (예: 동반 유형, 동반 가능 반려동물, 필수 준비물, 유의사항 등)

    반환값: (상세 딕셔너리, 원본 JSON)
    """
    params = {
        "serviceKey": SERVICE_KEY,
        "MobileOS": MOBILE_OS,
        "MobileApp": MOBILE_APP,
        "contentId": content_id,     # 어떤 장소인지 지정하는 고유 번호
        "_type": "json",
    }
    url = f"{BASE_URL}/detailPetTour2"

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
    except requests.exceptions.RequestException as e:
        st.warning(f"상세 조건을 불러오지 못했어요. (상세: {e})")
        return {}, {}
    except ValueError:
        st.warning("상세 조건 응답 형식이 올바르지 않아요.")
        return {}, {}

    items = (
        data.get("response", {})
            .get("body", {})
            .get("items", {})
    )
    # 이 장소에 반려동물 상세정보가 없으면 items 가 ""(빈 문자열)로 옵니다.
    if not isinstance(items, dict):
        return {}, data

    item = items.get("item", {})
    if isinstance(item, list):
        # 여러 개면 첫 번째만 사용
        item = item[0] if item else {}

    return item, data


def _fetch_single_item(operation, extra_params):
    """
    공통정보/소개정보처럼 '항목 1개'를 돌려주는 API를 호출하는 도우미 함수.
    실패하거나 내용이 없으면 빈 딕셔너리를 돌려줍니다.
    """
    params = {
        "serviceKey": SERVICE_KEY,
        "MobileOS": MOBILE_OS,
        "MobileApp": MOBILE_APP,
        "_type": "json",
    }
    params.update(extra_params)
    url = f"{BASE_URL}/{operation}"
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
    except (requests.exceptions.RequestException, ValueError):
        return {}

    items = (
        data.get("response", {})
            .get("body", {})
            .get("items", {})
    )
    if not isinstance(items, dict):
        return {}
    item = items.get("item", {})
    if isinstance(item, list):
        item = item[0] if item else {}
    return item if isinstance(item, dict) else {}


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_detail_common(content_id):
    """
    [공통정보 조회] detailCommon2
    전화번호(tel), 홈페이지(homepage), 소개글(overview) 등 공통 정보를 가져옵니다.
    반환값: 상세 딕셔너리 (실패 시 빈 딕셔너리)
    """
    return _fetch_single_item("detailCommon2", {"contentId": content_id})


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_detail_intro(content_id, content_type_id):
    """
    [소개정보 조회] detailIntro2
    주차 가능 여부, 이용/체크인 시간, 문의처 등 '유형별' 소개 정보를 가져옵니다.
    ⚠️ 이 API는 반드시 '정확한' contentTypeId(유형코드)를 함께 넘겨야 값이 나옵니다.
    반환값: 상세 딕셔너리 (실패 시 빈 딕셔너리)
    """
    if not content_type_id:
        return {}
    return _fetch_single_item(
        "detailIntro2",
        {"contentId": content_id, "contentTypeId": content_type_id},
    )


# =============================================================================
# 5. 신호등 판정 로직 (이 앱의 핵심!)
# =============================================================================

def extract_weight_limit(text):
    """
    조건 텍스트에서 '허용 무게 상한'을 숫자로 뽑아냅니다.
    예) '20kg 이하 반려견 동반 가능' → 20.0

    - 정규식으로 '숫자 + kg' 패턴을 찾습니다.
    - 못 찾으면 None 을 돌려줍니다(= 무게 제한 정보 없음/불명확).
    """
    if not text:
        return None

    # 소문자로 바꿔서 'KG', 'Kg', 'kg' 모두 잡히게 함
    lowered = str(text).lower()

    # 정규식 설명:
    #   (\d+(?:\.\d+)?)  → 정수 또는 소수(예: 10, 7.5)
    #   \s*              → 공백이 있어도/없어도 됨
    #   (?:kg|킬로그램|킬로) → kg, 킬로그램, 킬로 중 하나
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:kg|킬로그램|킬로)", lowered)
    if match:
        return float(match.group(1))
    return None


def judge_pet(user, place):
    """
    사용자 입력(user)과 장소 조건(place)을 대조해 신호등을 판정합니다.

    [입력]
      user  : {"breed": 견종, "weight": 무게(kg), "size": 크기} 딕셔너리
      place : detailPetTour2 상세 조건 딕셔너리

    [반환]
      (신호 문구, 색상코드, 사유리스트)
      예) ("🐾 조건부", "orange", ["동반 시 준비물 필요: 목줄 착용"])
    """
    reasons = []   # 판정 사유를 하나씩 모아 둘 리스트

    # 실제 필드명으로 값 읽기 (pick 으로 안전하게)
    acmpy_type = str(pick(place, "acmpyTypeCd"))    # 동반 유형(가장 중요) 예:'전구역 동반가능'
    allow_pet  = str(pick(place, "acmpyPsblCpam"))  # 동반 가능 반려동물
    need_items = str(pick(place, "acmpyNeedMtr"))   # 동반 시 필요사항(준비물)

    # 무게/키워드 검색용: 반려동물 관련 필드를 전부 하나로 합침
    full_text = " ".join(str(place.get(k, "")) for k in PET_FIELDS)

    # --- (A) 조건 정보 자체가 거의 없는 경우 → 조건부 + 현장 확인 권장 --------
    if not full_text.strip():
        reasons.append("동반 조건 정보가 등록되어 있지 않아요. 방문 전 현장 확인을 권장합니다.")
        return SIGNAL_STYLE["조건부"]["emoji"], SIGNAL_STYLE["조건부"]["color"], reasons

    # --- (B) '전체 동반 불가'인지 판정 → 불가 ---------------------------------
    # ⚠️ 중요: 동반 유형(acmpyTypeCd)이 판정의 '기준'입니다.
    #   - acmpyTypeCd 에 '동반가능'이 들어 있으면(전구역/일부구역 동반가능),
    #     본문(etcAcmpyInfo)의 '~는 동반 불가'는 '일부 구역' 제한일 뿐이므로
    #     전체 불가로 보지 않습니다. (D단계에서 조건부 사유로 처리)
    #   - acmpyTypeCd 자체가 '동반불가/금지'이거나, 유형 정보가 아예 없는데
    #     본문에 불가 표현이 있으면 → 전체 불가.
    ban_keywords = ["동반불가", "동반 불가", "출입불가", "출입 불가", "동반 금지", "입장불가"]
    type_says_ok = "동반가능" in acmpy_type            # 예: '전구역/일부구역 동반가능'
    type_says_ban = any(word in acmpy_type for word in ban_keywords) and not type_says_ok
    no_type_info = not acmpy_type.strip()
    text_has_ban = any(word in full_text for word in ban_keywords)

    if type_says_ban or (no_type_info and text_has_ban):
        reasons.append("장소 안내에 '동반 불가/금지' 표현이 포함되어 있어요.")
        return SIGNAL_STYLE["불가"]["emoji"], SIGNAL_STYLE["불가"]["color"], reasons

    # --- (C) 무게 상한 검사 ---------------------------------------------------
    weight_limit = extract_weight_limit(full_text)  # 조건에서 상한 숫자 추출
    user_weight = user.get("weight")                # 사용자 반려동물 무게

    if weight_limit is not None and user_weight is not None:
        if user_weight > weight_limit:
            # 무게 초과 → 불가
            reasons.append(
                f"허용 무게 상한({weight_limit:.0f}kg)을 초과했어요 "
                f"(내 반려동물: {user_weight:.0f}kg)."
            )
            return SIGNAL_STYLE["불가"]["emoji"], SIGNAL_STYLE["불가"]["color"], reasons
        else:
            # 무게는 통과했다는 사실을 사유에 기록
            reasons.append(f"무게 조건 충족 (상한 {weight_limit:.0f}kg 이내).")

    # --- (D) '조건부'로 만드는 부가 조건들 수집 -------------------------------
    # (D-1) 동반 유형이 '일부/제한'이면 조건부 사유로 명시
    if any(word in acmpy_type for word in ["일부", "제한", "부분"]):
        reasons.append(f"일부 구역만 동반 가능해요 (동반 유형: {acmpy_type}).")

    # (D-1-b) 유형은 동반가능이지만 본문에 '~는 동반 불가' 같은 부분 제한이 있으면 안내
    if type_says_ok and text_has_ban:
        reasons.append("실내 등 일부 시설·구역은 동반이 제한돼요. 상세 조건을 확인하세요.")

    # (D-2) 준비물이 적혀 있으면 조건부 사유로 추가
    if need_items.strip():
        reasons.append(f"동반 시 준비물 필요: {need_items.strip()}")

    # (D-3) 그 밖의 조건성 키워드(추가요금/야외제한 등) 탐지
    conditional_rules = {
        "이동장(케이지) 필요": ["이동장", "케이지", "이동가방", "켄넬"],
        "추가 요금 발생 가능": ["추가요금", "추가 요금", "별도요금", "별도 요금"],
        "야외/특정 공간만 가능": ["야외", "테라스", "실외"],
        "입마개 등 추가 준비 필요": ["입마개"],
    }
    for reason_text, keywords in conditional_rules.items():
        if any(word in full_text for word in keywords):
            reasons.append(reason_text)

    # --- (E) 최종 판정 --------------------------------------------------------
    # '무게 충족' 은 제약이 아니므로 제외한, 실제 '조건부 사유' 목록
    restriction_reasons = [r for r in reasons if not r.startswith("무게 조건 충족")]

    # 동반 유형에 '동반가능' 이 명시돼 있으면 긍정 신호로 봅니다.
    is_positive = ("동반가능" in acmpy_type) or ("가능" in acmpy_type) or bool(allow_pet.strip())

    if restriction_reasons:
        # 조건이 하나라도 있으면 조건부
        return SIGNAL_STYLE["조건부"]["emoji"], SIGNAL_STYLE["조건부"]["color"], reasons
    if is_positive:
        # 제약이 없고 '동반가능'이 명시 → 가능
        if not reasons:
            reasons.append("별도 제약 조건이 없어요.")
        return SIGNAL_STYLE["가능"]["emoji"], SIGNAL_STYLE["가능"]["color"], reasons

    # 그 외(정보가 애매) → 조건부 + 현장 확인 권장
    reasons.append("동반 조건이 명확하지 않아요. 방문 전 현장 확인을 권장합니다.")
    return SIGNAL_STYLE["조건부"]["emoji"], SIGNAL_STYLE["조건부"]["color"], reasons


# =============================================================================
# 6. 목록 데이터를 표(DataFrame)로 정리하는 함수
# =============================================================================

def build_dataframe(item_list):
    """
    API가 준 장소 목록(item_list)을 pandas DataFrame(표)으로 정리합니다.
    필터링/정렬을 쉽게 하기 위해서예요.

    아래 키들은 한국관광공사 areaBasedList2 의 '실제' 필드명입니다.
      - title      : 장소명
      - addr1      : 주소
      - contentid  : 장소 고유번호(상세조회에 사용)
      - contenttypeid : 유형 코드
      - firstimage : 대표 사진 URL
      - mapx       : 경도(x)   ← 지도 표시에 필요
      - mapy       : 위도(y)   ← 지도 표시에 필요
    """
    rows = []
    for it in item_list:
        rows.append({
            "장소명": it.get("title", "이름 없음"),
            "주소": it.get("addr1", ""),
            "content_id": it.get("contentid", ""),
            "유형코드": it.get("contenttypeid", ""),
            "사진": it.get("firstimage", ""),
            # 좌표는 문자열로 올 수 있으니 숫자로 안전 변환(실패 시 None)
            "경도": _to_float(it.get("mapx")),
            "위도": _to_float(it.get("mapy")),
        })
    return pd.DataFrame(rows)


def content_type_name(type_code):
    """유형 코드(숫자)를 사람이 읽는 이름으로 바꿔 줍니다."""
    for name, code in CONTENT_TYPES.items():
        if str(code) == str(type_code):
            return name
    return "기타"


# =============================================================================
# 7. Streamlit 화면 구성 시작
# =============================================================================

# 페이지 기본 설정 (탭 제목, 넓은 레이아웃)
st.set_page_config(page_title="펫패스(PetPass)", page_icon="🐾", layout="wide")

# 앱 제목과 소개
st.title("🐾 펫패스(PetPass)")
st.caption("반려동물 동반 출입 조건을 미리 확인하고, 헛걸음을 줄여요!")


# -----------------------------------------------------------------------------
# 7-1. 사이드바: '내 반려동물' 입력
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("🐕 내 반려동물")

    # 견종/종류 입력(드롭다운 목록 + 직접 입력)
    # 목록에서 고르되, 원하는 종류가 없으면 '기타(직접 입력)'를 선택해 타이핑합니다.
    breed_choice = st.selectbox("견종 / 종류", options=BREED_OPTIONS)
    if breed_choice == "기타 (직접 입력)":
        # '기타'를 고른 경우에만 자유 입력창을 보여 줍니다.
        breed = st.text_input("견종/종류 직접 입력",
                              value="", placeholder="예: 코숏 고양이, 웰시코기")
    else:
        breed = breed_choice

    # 무게 입력(숫자). min/max/step 으로 입력 범위를 제한
    weight = st.number_input("무게 (kg)", min_value=0.0, max_value=100.0,
                             value=5.0, step=0.5)

    # 크기 선택(라디오 버튼)
    size = st.radio("크기", options=["소형", "중형", "대형"], horizontal=True)

    # ⭐ 입력값을 st.session_state 에 저장 ⭐
    # session_state 는 화면이 새로고침돼도 값이 유지되는 '앱의 기억 공간'입니다.
    st.session_state["user_pet"] = {
        "breed": breed,
        "weight": weight,
        "size": size,
    }

    st.divider()

    # -------------------------------------------------------------------------
    # 7-2. 사이드바: 검색 필터
    # -------------------------------------------------------------------------
    st.header("🔎 검색 필터")

    # 지역(시/도) 선택
    selected_area = st.selectbox("지역 (시/도)", options=list(AREA_CODES.keys()))
    area_code = AREA_CODES[selected_area]   # 선택한 시/도의 코드

    # 구/군 선택 — 위에서 고른 시/도의 구/군 목록을 API로 받아와 보여 줍니다.
    # (시/도를 바꾸면 이 목록도 자동으로 바뀝니다.)
    sigungu_map = fetch_sigungu_list(area_code)              # {구이름: 코드}
    sigungu_options = ["전체"] + list(sigungu_map.keys())     # 맨 앞에 '전체' 추가
    selected_sigungu = st.selectbox("구 / 군", options=sigungu_options)
    # '전체'면 None(=구 필터 없음), 아니면 해당 구 코드
    sigungu_code = None if selected_sigungu == "전체" else sigungu_map.get(selected_sigungu)

    # 유형(관광지/음식점/숙소 등) 선택
    selected_type = st.selectbox("유형", options=list(CONTENT_TYPES.keys()))

    # '🐾만 보기' 토글: 켜면 '불가'로 판정된 곳은 목록에서 숨깁니다.
    only_ok = st.toggle("🐾만 보기 (동반 가능/조건부만)", value=False)

    # 디버그 옵션: 켜면 API 응답 원본(JSON)을 화면에 그대로 보여 줍니다.
    # → 실제 필드명을 확인할 때 매우 유용합니다!
    debug_mode = st.checkbox("🔧 디버그: API 응답 원본 보기", value=False)

    # 검색 버튼
    search_clicked = st.button("검색하기", type="primary", use_container_width=True)


# -----------------------------------------------------------------------------
# 7-3. 검색 실행: 버튼을 눌렀을 때만 API를 새로 호출하고 결과를 저장
# -----------------------------------------------------------------------------
if search_clicked:
    # area_code / sigungu_code 는 위 사이드바에서 이미 계산해 두었습니다.
    type_code = CONTENT_TYPES[selected_type]

    with st.spinner("공공데이터를 불러오는 중이에요... 🐾"):
        item_list, raw_json = fetch_area_based_list(area_code, type_code, sigungu_code)

    # 검색 결과를 앱의 기억 공간에 저장 → 다른 조작을 해도 유지됨
    st.session_state["search_df"] = build_dataframe(item_list)
    st.session_state["search_raw"] = raw_json
    # 새 검색을 하면 이전에 선택한 상세 장소는 초기화
    st.session_state.pop("selected_content_id", None)

# 디버그 모드면 목록 원본 JSON을 펼쳐 보여 줌
if debug_mode and "search_raw" in st.session_state:
    with st.expander("🔧 [디버그] areaBasedList2 응답 원본"):
        st.write(st.session_state["search_raw"])


# -----------------------------------------------------------------------------
# 7-4. 메인 영역: 왼쪽(목록) + 오른쪽(지도)
# -----------------------------------------------------------------------------

# 아직 검색을 한 번도 안 했으면 안내 문구만 보여 주고 종료
if "search_df" not in st.session_state:
    st.info("👈 왼쪽 사이드바에서 반려동물 정보와 검색 조건을 입력한 뒤 '검색하기'를 눌러 주세요.")
else:
    df = st.session_state["search_df"]           # 검색 결과 표
    user_pet = st.session_state["user_pet"]      # 내 반려동물 정보

    if df.empty:
        st.warning("검색 결과가 없어요. 다른 지역이나 유형으로 다시 검색해 보세요.")
    else:
        # -------------------------------------------------------------
        # 각 장소마다 상세조건을 불러와 신호등 판정을 미리 계산해 둡니다.
        # (표에 '신호등'/'색상'/'사유' 컬럼을 추가)
        # -------------------------------------------------------------
        with st.spinner("동반 조건을 분석하는 중이에요... 🐾"):
            signals, colors, reason_lists = [], [], []
            for _, row in df.iterrows():
                # 상세 조건 조회(반려동물 상세 API). 캐시 덕분에 부담이 적습니다.
                detail, _ = fetch_detail_pet_tour(row["content_id"])
                sig, col, reasons = judge_pet(user_pet, detail)
                signals.append(sig)
                colors.append(col)
                reason_lists.append(reasons)

        df = df.copy()
        df["신호등"] = signals
        df["색상"] = colors
        df["사유"] = reason_lists

        # '🐾만 보기'가 켜져 있으면 '불가'인 행을 제거
        if only_ok:
            df = df[df["신호등"] != SIGNAL_STYLE["불가"]["emoji"]]

        # 신호등 순서대로 정렬(가능 → 조건부 → 불가)
        order = {SIGNAL_STYLE["가능"]["emoji"]: 0,
                 SIGNAL_STYLE["조건부"]["emoji"]: 1,
                 SIGNAL_STYLE["불가"]["emoji"]: 2}
        df["정렬키"] = df["신호등"].map(order).fillna(3)
        df = df.sort_values("정렬키").reset_index(drop=True)

        # 화면을 두 칸(왼쪽 목록 / 오른쪽 상세정보)으로 나눔
        # 상세 정보가 내용이 많아 오른쪽 칸을 조금 더 넓게 잡습니다.
        left_col, right_col = st.columns([1, 1.4])

        # ============ 왼쪽: 검색 결과 목록 (전체 표시) ============
        with left_col:
            # 서브헤더 (기본 글자 크기)
            st.subheader(f"📋 검색 결과 ({len(df)}곳)")

            if df.empty:
                st.info("조건에 맞는 결과가 없어요.")
            else:
                # 검색된 모든 장소를 카드로 표시 (페이지 나눔 없이 전체)
                for idx, row in df.iterrows():
                    with st.container(border=True):
                        # 목록은 '신호등 + 상호명'만 간결하게 표시.
                        # 제목(h3)이 아니라 일반 크기(약 1rem, 굵게)로 표시해 리스트 느낌을 줍니다.
                        # (유형/주소 등 자세한 정보는 '상세 보기'에서 보여 줍니다.)
                        st.markdown(
                            f"<span style='font-size:1rem; font-weight:600;'>"
                            f"{row['신호등']}  {row['장소명']}</span>",
                            unsafe_allow_html=True,
                        )

                        # '상세 보기' 버튼 → 누르면 선택한 장소를 기억
                        if st.button("상세 보기", key=f"detail_{idx}"):
                            st.session_state["selected_content_id"] = row["content_id"]
                            st.session_state["selected_row"] = row.to_dict()

        # ============ 오른쪽: 선택한 장소의 상세 정보 ============
        with right_col:
            st.subheader("🐾 상세 정보")

            if "selected_content_id" not in st.session_state:
                # 아직 아무것도 선택하지 않았을 때 안내
                st.info("👈 왼쪽 목록에서 '상세 보기'를 누르면 여기에 자세한 정보가 나와요.")
            else:
                sel_id = st.session_state["selected_content_id"]
                sel_row = st.session_state.get("selected_row", {})

                # 반려동물 동반 상세 조건 조회 + 신호등 판정
                detail, detail_raw = fetch_detail_pet_tour(sel_id)
                sig, col, reasons = judge_pet(st.session_state["user_pet"], detail)

                if debug_mode:
                    with st.expander("🔧 [디버그] detailPetTour2 응답 원본"):
                        st.write(detail_raw)

                # 상세 카드 그리기
                with st.container(border=True):
                    st.markdown(f"## {sig}  {sel_row.get('장소명', '선택한 장소')}")

                    # 신호등 색상에 맞는 안내 박스
                    if col == "green":
                        st.success("동반 방문에 큰 제약이 없어 보여요!")
                    elif col == "orange":
                        st.warning("조건부예요. 아래 사유와 준비물을 꼭 확인하세요.")
                    else:
                        st.error("동반이 어려워 보여요. 방문 전 재확인을 권장합니다.")

                    # 판정 사유 목록
                    if reasons:
                        st.markdown("**판정 사유**")
                        for r in reasons:
                            st.write(f"- {r}")

                    # 🖼️ 가게 그림(대표 이미지)이 있으면 표시
                    photo_url = sel_row.get("사진", "")
                    if photo_url:
                        st.markdown("### 🖼️ 가게 그림")
                        st.image(photo_url, caption=sel_row.get("장소명", ""),
                                 use_container_width=True)

                    # ---------------------------------------------------------
                    # 기본 정보: 유형 / 주소 / 연락처 / 주차 / 시간 / 홈페이지 / 소개
                    #   - 유형/주소는 목록 데이터(sel_row)에서
                    #   - 연락처/홈페이지/소개는 detailCommon2 에서
                    #   - 주차/시간/문의/메뉴는 detailIntro2 에서 (정확한 유형코드 필요)
                    # ---------------------------------------------------------
                    common = fetch_detail_common(sel_id)
                    intro = fetch_detail_intro(sel_id, sel_row.get("유형코드"))

                    if debug_mode:
                        with st.expander("🔧 [디버그] detailCommon2 / detailIntro2 원본"):
                            st.write({"detailCommon2": common, "detailIntro2": intro})

                    # 각 정보를 안전하게 뽑아내기 (없으면 빈 값)
                    place_type = content_type_name(sel_row.get("유형코드"))          # 유형
                    addr = sel_row.get("주소", "")                                  # 주소
                    tel = pick(common, "tel") or find_by_keyword(intro, "infocenter")  # 연락처
                    parking = find_by_keyword(intro, "parking")                      # 주차 가능 여부
                    usetime = find_by_keyword(intro, "usetime", "opentime", "checkintime")  # 시간
                    restdate = find_by_keyword(intro, "restdate")                    # 휴무일
                    menu = strip_html(find_by_keyword(intro, "menu"))                # 메뉴(음식점 등)
                    homepage = strip_html(pick(common, "homepage"))                  # 홈페이지
                    overview = strip_html(pick(common, "overview"))                  # 소개글

                    st.markdown("### 📌 기본 정보")
                    st.write(f"- **유형:** {place_type}")
                    st.write(f"- **주소:** {addr or '정보 없음'}")
                    st.write(f"- **연락처:** {tel or '정보 없음'}")
                    st.write(f"- **주차 가능 여부:** {parking or '정보 없음'}")
                    if usetime:
                        st.write(f"- **이용/영업/체크인 시간:** {usetime}")
                    if restdate:
                        st.write(f"- **휴무일:** {restdate}")
                    if homepage:
                        st.write(f"- **홈페이지:** {homepage}")

                    # 🍽️ 메뉴 (음식점/카페 등 메뉴 정보가 있을 때만 표시)
                    if menu:
                        st.markdown("### 🍽️ 메뉴")
                        st.write(menu)

                    if overview:
                        with st.expander("📖 장소 소개 보기"):
                            st.write(overview)

                    # 📍 위치: 지도를 없앤 대신 주소·좌표 + 외부 지도 링크로 안내
                    st.markdown("### 📍 위치")
                    lat = sel_row.get("위도")
                    lon = sel_row.get("경도")
                    if lat is not None and lon is not None:
                        map_url = ("https://www.google.com/maps/search/?api=1"
                                   f"&query={lat},{lon}")
                        st.write(f"- 좌표: {lat:.5f}, {lon:.5f}")
                        st.markdown(f"[🗺️ 지도로 열기(새 창)]({map_url})")
                    else:
                        st.write("위치(좌표) 정보가 없어요.")

                    # ---------------------------------------------------------
                    # 동반 조건 원문 표시 (PET_FIELDS에 정의된 실제 필드들을 순회)
                    # ---------------------------------------------------------
                    st.markdown("### 동반 조건 원문")
                    shown_any = False
                    for key, label in PET_FIELDS.items():
                        value = detail.get(key, "")
                        if value not in (None, "", " "):
                            st.markdown(f"**{label}**")
                            st.write(value)
                            shown_any = True
                    if not shown_any:
                        st.write("등록된 동반 조건 정보가 없어요. 방문 전 현장 확인을 권장합니다.")

                    # ---------------------------------------------------------
                    # 준비물 체크리스트: '동반 시 필요사항(acmpyNeedMtr)'을
                    # 쉼표/줄바꿈 등으로 쪼개 각 항목을 체크박스로 보여 줍니다.
                    # ---------------------------------------------------------
                    st.markdown("**✅ 준비물 체크리스트**")
                    need_items = pick(detail, "acmpyNeedMtr")  # 준비물 원문
                    raw_items = re.split(r"[,/\n·•]", str(need_items)) if need_items else []
                    checklist = [x.strip() for x in raw_items if x.strip()]
                    if not checklist:
                        # 필수 준비물이 비어 있으면 기본 체크리스트 제공
                        checklist = ["목줄/하네스", "배변봉투", "이동장(필요 시)", "예방접종 확인"]
                    for i, item in enumerate(checklist):
                        st.checkbox(item, key=f"check_{sel_id}_{i}")


# =============================================================================
# 8. 푸터(맨 아래 출처/제작자 표기) — 공모전 요구사항이라 반드시 표시
# =============================================================================
st.divider()
st.caption(
    "데이터 출처: 한국관광공사 반려동물 동반여행 정보(공공데이터포털) · "
    "제작: 숭실대학교 노은희"
)

# 단독/다가구 실거래가(마스킹된 번지)에, 건축물대장 데이터를 대조해 확정번지를 채워 넣는 스크립트.
# data/raw 의 모든 실거래가 xlsx를 전국 대상으로 한번에 처리한다.
# 매칭된 건물의 대장pk는 각 실거래 xlsx에 컬럼 하나로 붙이고, 건물 상세 정보는 대장pk 기준으로
# 중복 없이 모아 data/processed/건물정보.xlsx 마스터 테이블 하나로 별도 저장한다.
# 현재 표제부로 매칭되지 않은 행은 폐쇄말소 표제부(거래 후 멸실된 건물 등)로 한 번 더 보완 매칭한다.
import glob
import os
import re
import sys
import warnings

import pandas as pd

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

# 실거래가 xlsx에 기본 셀 서식 정보가 없어 openpyxl이 띄우는 경고. 데이터 읽기에는 영향이 없으므로 숨긴다.
warnings.filterwarnings("ignore", message="Workbook contains no default style")

INPUT_DIR = r"D:\dev\data\processed\traded_house_processed\overall"  # 이 디렉터리의 모든 xlsx를 전국 대상으로 한번에 처리한다.
# 전국 건축물대장 CSV(대장구분 미필터링, 809만행 규모).
# 청크 단위로 읽으며 '일반' 대장구분 + 필요한 동만 추려낸다.
BASIS_CSV = r"D:\dev\data\basis\building_ledger\표제부_2026_09.csv"
BASIS_CSV_ENCODING = "utf-8-sig"  # CSV가 BOM 포함 UTF-8이라, BOM이 첫 컬럼명에 붙지 않도록 utf-8-sig로 읽는다.
# 폐쇄말소 표제부 CSV. 현재 표제부로 매칭 실패한 행만 이걸로 보완 매칭한다.
# (과거 거래일수록 거래 후 철거(멸실)돼 현재 표제부에서 빠진 건물이 많아, 이 경로로 상당수가 매칭된다.)
CLOSED_BASIS_CSV = r"D:\dev\data\basis\building_ledger\closed\폐쇄말소_표제부_2026_09.csv"
# 매칭출처 컬럼 값: 어떤 대장에서 확정번지를 찾았는지 구분한다.
SOURCE_OVERALL = "총괄표제부"  # find_overall_ledger_house.py 단계에서 이미 확정번지를 찾은 행
SOURCE_CURRENT = "표제부"
SOURCE_CLOSED = "폐쇄말소"
# 허용오차 매칭: 정확 매칭(표제부 -> 폐쇄말소)으로 못 찾은 행만, 건축년도/연면적 차이를 아래 범위까지 허용해 한 번 더 매칭한다.
# (실거래 신고 건축년도가 사용승인연도와 1년 어긋나거나, 연면적이 소수점/반올림 차이로 조금 다른 경우를 잡기 위함)
# 정확 매칭보다 오매칭 위험이 있으므로 매칭출처에 "(허용오차)"를 붙여 구분한다.
YEAR_TOLERANCE = 1  # 건축년도 ±1년
AREA_TOLERANCE = 1.0  # 연면적 ±1㎡
FLOAT_EPSILON = 1e-9  # 51.04 - 50.04 = 1.0000000000000002 같은 부동소수점 오차로 경계값이 빠지지 않도록 더하는 값
TOLERANCE_SUFFIX = "(허용오차)"
OUTPUT_DIR = r"D:\dev\data\processed\traded_house_processed"
SUMMARY_FILE_NAME = "표제부매핑_단독다가구_요약통계.xlsx"  # 파일별 매칭 통계를 모아 OUTPUT_DIR에 저장한다.
SUMMARY_SHEET_NAME = "요약통계"

HEADER_ROW = 2  # find_house_multi_buillding.py가 저장한 xlsx의 실제 표 헤더 행(0-index)
OUTPUT_SHEET_NAME = "실거래가"
DATA_START_ROW = 2  # 0-index. 1행에 통계 요약, 2행은 빈 줄, 3행부터 표 작성
EMPTY_VALUE = "-"  # 실제 거래를 나타내는 값(해제된 경우는 날짜가 들어감)
CSV_CHUNK_SIZE = 200_000  # 대용량 전국 건축물대장 CSV를 청크 단위로 읽는다.

# 실거래 주택유형 -> 건축물대장 주용도 중 허용되는 값들.
# 건축물대장에는 '다가구주택' 라벨이 거의 없고(전체 중 2건뿐) 대부분 '단독주택'으로 기록되므로,
# 단독/다가구 모두 '단독주택'을 허용하고, 다가구는 혹시 존재하는 '다가구주택'도 함께 허용한다.
HOUSE_TYPE_TO_JUYONGDO = {
    "단독": {"단독주택"},
    "다가구": {"단독주택", "다가구주택"},
}

# 매칭에 성공한 행에 건축물대장에서 함께 붙여줄 컬럼들(건축물대장 기준 원본 컬럼명). PK는 항상 맨 앞에 둔다.
BASIS_EXTRA_COLUMNS = [
    "PK", "건물명", "동명", "대지면적(㎡)", "건축면적(㎡)", "연면적(㎡)", "건폐율(%)", "용적률(%)",
    "용적률산정연면적(㎡)", "주구조", "주용도", "주지붕", "지상층수", "지하층수", "사용승인일",
]
# PK는 결과에서 '대장pk'로, 대지면적(㎡)/연면적(㎡)은 실거래가에도 같은 이름의 컬럼(신고된 값)이 이미
# 있으므로 덮어쓰지 않고 건축물대장 값임을 구분할 수 있도록 별도 컬럼명으로 붙인다.
BASIS_EXTRA_COLUMN_RENAME = {
    "PK": "대장pk",
    "대지면적(㎡)": "대장상 대지면적(㎡)",
    "연면적(㎡)": "대장상 연면적(㎡)",
}
# 위 rename이 적용된 최종 컬럼명들(대장pk 포함 14개). deals에는 이제 이 중 '대장pk' 하나만 직접 붙고,
# 나머지 13개 상세 컬럼은 건물정보.xlsx 마스터 테이블(대장pk 기준, 중복 제거)의 컬럼 순서로 쓰인다.
OUTPUT_EXTRA_COLUMNS = [BASIS_EXTRA_COLUMN_RENAME.get(col, col) for col in BASIS_EXTRA_COLUMNS]
# 숫자로 변환해두는 게 자연스러운 붙임 컬럼들. (건물명/동명/주구조/주용도/주지붕/사용승인일은 문자열 그대로 둔다.)
BASIS_NUMERIC_EXTRA_COLUMNS = [
    "대지면적(㎡)", "건축면적(㎡)", "연면적(㎡)", "건폐율(%)", "용적률(%)", "용적률산정연면적(㎡)",
    "지상층수", "지하층수",
]
# 건축물대장 CSV의 실제 헤더명 -> 이 스크립트 내부에서 쓰는 컬럼명. CSV 헤더 네이밍이 바뀌어도 내부 로직과
# 결과 컬럼명은 그대로 두고, 읽은 직후 이 매핑으로 이름을 바꿔 맞춘다.
BASIS_CSV_COLUMN_RENAME = {
    "관리_건축물대장_PK": "PK",
    "대장_구분_코드_명": "대장구분",
    "대지_위치": "대지위치",
    "번": "번",
    "지": "지",
    "대지_구분_코드": "대지구분",
    "사용승인_일": "사용승인일",
    "건물_명": "건물명",
    "동_명": "동명",
    "대지_면적(㎡)": "대지면적(㎡)",
    "건축_면적(㎡)": "건축면적(㎡)",
    "연면적(㎡)": "연면적(㎡)",
    "건폐_율(%)": "건폐율(%)",
    "용적_률(%)": "용적률(%)",
    "용적_률_산정_연면적(㎡)": "용적률산정연면적(㎡)",
    "구조_코드_명": "주구조",
    "주_용도_코드_명": "주용도",
    "지붕_코드_명": "주지붕",
    "지상_층_수": "지상층수",
    "지하_층_수": "지하층수",
}
# 매칭(대장구분/대지위치/번/지/사용승인일)에 필요한 컬럼 + 붙임 컬럼만 읽어 CSV 파싱 비용을 줄인다.
# 대지구분은 '총괄표제 존재여부'가 O인 지번주소의 표제부 상세 조회에서 산 지번을 구분하는 데 쓴다.
# (읽을 컬럼은 각 CSV의 rename 매핑 키로 정하므로 _load_basis_candidates에서 rename 매핑으로부터 만든다.)
# 폐쇄말소 표제부 CSV는 PK 컬럼명만 다르고 나머지 헤더는 표제부와 같다. 폐쇄말소일은 건물정보.xlsx에 함께 남긴다.
CLOSED_BASIS_CSV_COLUMN_RENAME = {
    **{k: v for k, v in BASIS_CSV_COLUMN_RENAME.items() if k != "관리_건축물대장_PK"},
    "관리_폐쇄말소대장_PK": "PK",
    "폐쇄말소_일": "폐쇄말소일",
}
# 건축물대장 대지_구분_코드 중 산 지번을 뜻하는 값(0: 대지, 1: 산, 2: 블록).
SAN_DAEJI_GUBUN_CODE = "1"
# "서울특별시 성북구 돈암동 552-3번지" / "...산11-1번지" 형태의 대지위치 끝에서 번지 부분을 떼어내기 위한 패턴.
_DAEJI_WICHI_BEONJI_RE = re.compile(r"\s+산?\d+(?:-\d+)?번지$")


"""마스킹된 번지 문자열을 (자리수 접두어, 자리수)로 분리한다.

    '1**' -> ('1', 3), '*' -> ('', 1)
"""
def parse_beonji(raw: str):
    body = raw[1:] if raw.startswith("산") else raw
    prefix = body.split("*", 1)[0]
    return prefix, len(body)


"""시군구 문자열의 연속 공백을 1개로 줄인다. 두 데이터 출처의 공백 표기가 달라도 같은 키로 매칭되게 한다.

    법정동명을 비교용으로 정규화한다. (실거래 시군구, 건축물대장 대지위치에서 뗀 시군구에 공통 적용)
    - 앞뒤 공백 제거
    - 텍스트 사이 연속 공백(2칸 이상, 탭/전각 공백 포함)을 한 칸으로
    문자열이 아니면(NaN 등) 그대로 반환한다.
    예: "  경기도 화성시   새솔동 " -> "경기도 화성시 새솔동"
"""
def _normalize_dong_name(name):
    if not isinstance(name, str):
        return name
    return re.sub(r"\s+", " ", name.strip())


"""실거래가 xlsx를 읽어 해제(취소) 거래를 제외한 뒤, 매칭에 쓸 파생 컬럼(시군구, 번지 접두어/자리수)을 추가해 반환한다."""
def _load_deals(xlsx_path: str):
    # 확정번지/번지는 셀에 문자열로 저장돼 있어도, 값이 전부 숫자 형태(예: "1556", '-' 없음)면 pandas가 float로 추론해
    # "1556.0"이 되므로 문자열로 고정해서 읽는다.
    deals = pd.read_excel(xlsx_path, header=HEADER_ROW, dtype={"확정번지": str, "번지": str}, engine="openpyxl")
    # 확정번지가 전부 비어있으면(NaN) float64로 읽혀, 이후 문자열 대입 시 타입 충돌이 나므로 object로 맞춰둔다.
    deals["확정번지"] = deals["확정번지"].astype(object)

    cancelled_mask = deals["해제사유발생일"].notna() & (deals["해제사유발생일"] != EMPTY_VALUE)
    deals = deals[~cancelled_mask].reset_index(drop=True)
    total_count = len(deals)

    # 시군구 컬럼을 매칭키로 쓴다. 연속 공백을 1개로 줄여 건축물대장 쪽 키 표기와 맞춘다.
    deals["_시군구"] = deals["시군구"].map(_normalize_dong_name)

    parsed = deals["번지"].map(parse_beonji)
    deals["_번지_접두어"] = [p[0] for p in parsed]
    deals["_번지_자리수"] = [p[1] for p in parsed]

    return deals, total_count


"""건축물대장의 '대지위치' 문자열에서 끝의 번지 부분을 떼어내, 시군구("시도 구 법정동") 문자열을 반환한다.

    번지로 끝나지 않는 값(예: 번지 없이 동까지만 있는 값)은 그대로 공백만 정리해 반환하고, 문자열이 아니면 None을 반환한다.
"""
def parse_daeji_wichi(raw):
    if not isinstance(raw, str):
        return None
    return _normalize_dong_name(_DAEJI_WICHI_BEONJI_RE.sub("", raw.strip()))


"""전국 건축물대장 CSV를 청크 단위로 읽어, '일반' 대장구분만 남기고 시군구(대지위치에서 번지를 뗀 "시도 구 법정동" 문자열) 별로 묶은 후보 딕셔너리를 만든다.

    전국 xlsx를 모두 처리하므로 동 단위 필터링 없이 전체를 대상으로 한다.
    column_rename은 CSV 헤더 -> 내부 컬럼명 매핑으로, 읽을 컬럼(usecols)도 이 매핑의 키로 정한다.
    (현재 표제부는 BASIS_CSV_COLUMN_RENAME, 폐쇄말소 표제부는 CLOSED_BASIS_CSV_COLUMN_RENAME)
"""
def _load_basis_candidates(basis_csv_path: str, column_rename: dict) -> dict:
    if not os.path.exists(basis_csv_path):
        raise FileNotFoundError(f"{basis_csv_path} 가 없습니다.")

    print(f"{os.path.basename(basis_csv_path)} 로딩 중...")
    chunks = []
    reader = pd.read_csv(
        basis_csv_path, encoding=BASIS_CSV_ENCODING, dtype=str, usecols=list(column_rename.keys()), chunksize=CSV_CHUNK_SIZE
    )
    for chunk in reader:
        chunk = chunk.rename(columns=column_rename)
        chunk = chunk[chunk["대장구분"] == "일반"]
        if chunk.empty:
            continue

        chunk = chunk.copy()
        # CSV에 시도/시군구/법정동 컬럼이 따로 없으므로, 대지위치에서 번지를 떼어낸 문자열을 시군구 키로 쓴다.
        chunk["_시군구"] = chunk["대지위치"].map(parse_daeji_wichi)
        chunk = chunk[chunk["_시군구"].notna()]
        chunk["_번_정수"] = pd.to_numeric(chunk["번"], errors="coerce")
        chunk["_지_정수"] = pd.to_numeric(chunk["지"], errors="coerce").fillna(0)
        chunk = chunk[chunk["_번_정수"].notna()].copy()
        chunk["_번_정수"] = chunk["_번_정수"].astype("int64")
        chunk["_지_정수"] = chunk["_지_정수"].astype("int64")
        chunk["_번_문자열"] = chunk["_번_정수"].astype(str)
        chunk["_번_자리수"] = chunk["_번_문자열"].str.len()
        chunk["_사용승인연도"] = pd.to_numeric(chunk["사용승인일"].astype(str).str[:4], errors="coerce")
        for col in BASIS_NUMERIC_EXTRA_COLUMNS:
            chunk[col] = pd.to_numeric(chunk[col], errors="coerce")

        chunks.append(chunk)

    if not chunks:
        return {}

    basis = pd.concat(chunks, ignore_index=True)
    return {key: df for key, df in basis.groupby("_시군구")}


"""실거래 한 건(딕셔너리 형태의 행)에 대해 고유 매핑키 -> 서브 매핑키 순으로 후보를 좁혀 확정번지를 찾는다.

    확정되면 {"확정번지": ..., "지번주소": ..., OUTPUT_EXTRA_COLUMNS...(대장pk 포함)} 딕셔너리를,
    실패하면 None을 반환한다. 지번주소는 산 여부를 원본 '번지'(마스킹된 값, 예: "산1**")로 판단해
    시군구+확정번지로 조합한 것으로, 건물정보.xlsx에 저장할 때 쓰인다.
    year_tolerance/area_tolerance가 0이면 건축년도/연면적 정확 일치, 0보다 크면 그 차이까지 허용한다.
"""
def _match_deal(deal: dict, candidates_by_dong: dict, year_tolerance: int = 0, area_tolerance: float = 0):
    candidates = candidates_by_dong.get(deal["_시군구"])
    if candidates is None:
        return None

    # 고유 매핑키: 번지(자리수+접두어, 지는 제외) + 연면적 + 건축년도가 모두 정확히 일치해야 한다.
    # (허용오차 매칭이면 연면적/건축년도는 허용 범위 안이면 일치로 본다.)
    if year_tolerance or area_tolerance:
        area_mask = (candidates["연면적(㎡)"] - deal["연면적(㎡)"]).abs() <= area_tolerance + FLOAT_EPSILON
        year_mask = (candidates["_사용승인연도"] - deal["건축년도"]).abs() <= year_tolerance
    else:
        area_mask = candidates["연면적(㎡)"] == deal["연면적(㎡)"]
        year_mask = candidates["_사용승인연도"] == deal["건축년도"]
    mask = (
        (candidates["_번_자리수"] == deal["_번지_자리수"])
        & candidates["_번_문자열"].str.startswith(deal["_번지_접두어"])
        & area_mask
        & year_mask
    )
    hits = candidates[mask]

    # 서브 매핑키: 주택유형(주용도) -> 대지면적 순으로, 후보가 2개 이상일 때만 한 단계씩 더 좁힌다.
    # 어떤 서브 매핑키를 적용한 결과가 0개면 그 키는 건너뛰고, 직전 후보를 유지한 채 다음 키로 넘어간다.
    allowed_juyongdo = HOUSE_TYPE_TO_JUYONGDO.get(deal["주택유형"], set())
    sub_key_filters = [
        lambda df: df["주용도"].isin(allowed_juyongdo),
        lambda df: df["대지면적(㎡)"] == deal["대지면적(㎡)"],
    ]
    for sub_key_filter in sub_key_filters:
        if len(hits) < 2:
            break
        narrowed = hits[sub_key_filter(hits)]
        if not narrowed.empty:
            hits = narrowed

    if len(hits) != 1:
        return None

    matched = hits.iloc[0]
    beonji = str(matched["_번_정수"]) if matched["_지_정수"] == 0 else f"{matched['_번_정수']}-{matched['_지_정수']}"
    산여부 = str(deal["번지"]).startswith("산")
    지번주소 = deal["시군구"] + (" 산 " if 산여부 else " ") + beonji
    extra = {BASIS_EXTRA_COLUMN_RENAME.get(col, col): matched[col] for col in BASIS_EXTRA_COLUMNS}
    # 폐쇄말소 표제부에서 찾은 경우에만 폐쇄말소일이 있다(현재 표제부면 None).
    return {"확정번지": beonji, "지번주소": 지번주소, "폐쇄말소일": matched.get("폐쇄말소일"), **extra}


"""전체 실거래 행 중 '총괄표제 존재여부'가 비어있는(확정번지가 아직 없는) 행은 _match_deal로 확정번지와
대장pk를 채우고, '총괄표제 존재여부'가 'O'인 행(확정번지가 이미 채워져 있던 행)은 지번주소를 만들어
표제부에서 대장pk만 채운다(둘 다 실패한 행은 대장pk가 비어있는 채로 남는다).

    두 경로에서 매칭된 건물의 상세 정보(OUTPUT_EXTRA_COLUMNS)는 이 xlsx에는 더 이상 붙이지 않고,
    buildings_by_pk(대장pk -> 상세 컬럼 딕셔너리, main()에서 전체 실행 동안 누적)에 모아
    나중에 건물정보.xlsx 하나로 저장한다. address_pk_cache는 총괄표제 O-case 지번주소 ->
    대장pk 문자열 캐시로, main()에서 파일 간에 재사용되어 같은 건물을 여러 번 표제부 조회하지 않게 한다.
    통계를 계산해 결과 xlsx로 저장한다.

    현재 표제부(candidates_by_dong)로 매칭 실패한 행은 폐쇄말소 표제부(closed_candidates_by_dong)로 다시 매칭한다.
    어느 대장에서 찾았는지는 '매칭출처' 컬럼(총괄표제부/표제부/폐쇄말소)에 남긴다.
    buildings_by_pk의 키는 (대장출처, 대장pk)다. 현재 표제부와 폐쇄말소 표제부의 PK는 서로 다른 체계라 출처를 함께 키로 쓴다.
    정확 매칭으로도 못 찾은 행은 건축년도 ±YEAR_TOLERANCE, 연면적 ±AREA_TOLERANCE를 허용해
    표제부 -> 폐쇄말소 순으로 한 번 더 매칭하고, 매칭출처에 "(허용오차)"를 붙인다.

    요약 통계 파일 작성을 위해 (확정번지 있는 건수, 대장pk가 채워진(표제부에 건물정보 있는) 건수, 폐쇄말소 매칭 건수, 허용오차 매칭 건수)를 반환한다."""
def _match_and_write(
    xlsx_path: str, deals, total_count: int, candidates_by_dong: dict, closed_candidates_by_dong: dict,
    address_pk_cache: dict, buildings_by_pk: dict
):
    needs_match_mask = deals["확정번지"].isna()
    needs_match_index = deals.index[needs_match_mask]
    needs_match_records = deals.loc[needs_match_mask].to_dict("records")

    # 매칭 단계: (후보, 대장출처, 건축년도 허용오차, 연면적 허용오차, 매칭출처). 앞 단계에서 못 찾은 행만 다음 단계로 넘어간다.
    # 정확 매칭(표제부 -> 폐쇄말소)을 모두 시도한 뒤에야 허용오차 매칭을 하므로, 정확히 맞는 건물이 있으면 항상 그쪽이 우선이다.
    match_stages = [
        (candidates_by_dong, SOURCE_CURRENT, 0, 0, SOURCE_CURRENT),
        (closed_candidates_by_dong, SOURCE_CLOSED, 0, 0, SOURCE_CLOSED),
        (candidates_by_dong, SOURCE_CURRENT, YEAR_TOLERANCE, AREA_TOLERANCE, SOURCE_CURRENT + TOLERANCE_SUFFIX),
        (closed_candidates_by_dong, SOURCE_CLOSED, YEAR_TOLERANCE, AREA_TOLERANCE, SOURCE_CLOSED + TOLERANCE_SUFFIX),
    ]
    # 행별 (매칭 결과, 대장출처, 매칭출처). 아직 못 찾은 행은 None.
    stage_results = [None] * len(needs_match_records)
    for stage_candidates, ledger, year_tolerance, area_tolerance, source in match_stages:
        for i, deal in enumerate(needs_match_records):
            if stage_results[i] is not None:
                continue
            result = _match_deal(deal, stage_candidates, year_tolerance, area_tolerance)
            if result is not None:
                stage_results[i] = (result, ledger, source)

    # (행 index, 매칭 결과, 대장출처, 매칭출처) 목록
    found = [(idx, *stage_result) for idx, stage_result in zip(needs_match_index, stage_results) if stage_result is not None]
    found_index = [idx for idx, _, _, _ in found]

    if "대장pk" not in deals.columns:
        deals["대장pk"] = pd.NA
    # 총괄표제부 단계에서 이미 확정번지를 찾은 행은 출처를 총괄표제부로 둔다.
    deals["매칭출처"] = deals["총괄표제 존재여부"].map({"O": SOURCE_OVERALL})
    if found_index:
        deals.loc[found_index, "확정번지"] = [result["확정번지"] for _, result, _, _ in found]
        deals.loc[found_index, "대장pk"] = [result["대장pk"] for _, result, _, _ in found]
        deals.loc[found_index, "매칭출처"] = [source for _, _, _, source in found]

    # 마스킹 번지 경로(_match_deal)에서 매칭된 건물 상세를 (대장출처, 대장pk) 기준으로 누적(건물정보.xlsx용).
    # 대장출처는 허용오차 여부와 상관없이 실제 대장(표제부/폐쇄말소)이다.
    for _, result, ledger, _ in found:
        buildings_by_pk[(ledger, result["대장pk"])] = {
            "지번주소": result["지번주소"],
            "대장출처": ledger,
            "폐쇄말소일": result["폐쇄말소일"],
            **{col: result[col] for col in OUTPUT_EXTRA_COLUMNS},
        }

    o_case_pk_count = _fill_o_case_pk(deals, candidates_by_dong, address_pk_cache, buildings_by_pk)

    deals = deals.drop(columns=["_시군구", "_번지_접두어", "_번지_자리수"])

    # 매칭출처별 건수(표제부, 폐쇄말소, 표제부(허용오차), 폐쇄말소(허용오차))
    source_counts = {source: 0 for *_, source in match_stages}
    for _, _, _, source in found:
        source_counts[source] += 1
    closed_success_count = sum(1 for _, _, ledger, _ in found if ledger == SOURCE_CLOSED)
    tolerance_success_count = sum(1 for _, _, _, source in found if source.endswith(TOLERANCE_SUFFIX))
    new_success_count = len(found_index)
    success_count = deals["확정번지"].notna().sum()
    success_rate = success_count / total_count if total_count else 0
    pk_filled_count = deals["대장pk"].notna().sum()
    source_breakdown = " + ".join(f"{source} {count}" for source, count in source_counts.items())
    print(f"{os.path.basename(xlsx_path)} 전체 건수: {total_count}")
    print(f"{os.path.basename(xlsx_path)} 신규 매칭 성공 건수(총괄표제 미존재 대상): {new_success_count} / {needs_match_mask.sum()} ({source_breakdown})")
    print(f"{os.path.basename(xlsx_path)} 매칭 성공 건수(확정번지): {success_count}")
    print(f"{os.path.basename(xlsx_path)} 매칭 성공률: {success_rate:.1%}")
    print(f"{os.path.basename(xlsx_path)} 대장pk 채워진 건수: {pk_filled_count} (신규 매칭 {new_success_count} + 총괄표제 O-case {o_case_pk_count})")

    _xlsx_name, _xlsx_ext = os.path.splitext(os.path.basename(xlsx_path))
    output_path = os.path.join(OUTPUT_DIR, f"{_xlsx_name}_매핑{_xlsx_ext}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        deals.to_excel(writer, sheet_name=OUTPUT_SHEET_NAME, index=False, startrow=DATA_START_ROW)
        sheet = writer.sheets[OUTPUT_SHEET_NAME]
        sheet["A1"] = "total_info"
        sheet["B1"] = f"전체 건수: {total_count}"
        sheet["G1"] = f"매칭 성공 건수(확정번지): {success_count}"
        sheet["L1"] = f"매칭 성공률: {success_rate:.1%}"
        sheet["Q1"] = f"대장pk 채워진 건수: {pk_filled_count}"
        sheet["V1"] = f"폐쇄말소 매칭 건수: {closed_success_count}"
        sheet["AA1"] = f"허용오차 매칭 건수: {tolerance_success_count}"

    print(f"완료: {output_path}")
    return int(success_count), int(pk_filled_count), closed_success_count, tolerance_success_count


"""'총괄표제 존재여부'가 'O'인 행(이미 총괄표제부로 확정번지를 찾은 행)에 대해 지번주소를 만들고,
해당 주소로 매칭되는 건축물대장 PK를 deals의 '대장pk' 컬럼에 채운다. 한 지번주소가 여러 PK(여러
건물/동)에 매칭될 수 있으므로, 매칭된 PK들을 콤마로 이어붙인 문자열로 채운다(예: "12345,67890").

    확정번지는 "번-지"(지가 0이면 "번") 형식이라 산 여부가 드러나지 않으므로, 원본 '번지'
    (마스킹된 값, 예: "산1**")로 산 여부를 판단해 지번주소에 반영한다.

    address_pk_cache는 지번주소 -> 대장pk 문자열(매칭 실패 시 None) 캐시로, main()의 파일 루프
    전체에서 재사용된다. 같은 지번주소가 여러 xlsx 파일에 걸쳐 등장해도 표제부 조회는 한 번만
    수행하고, 매칭된 건물 상세는 buildings_by_pk(대장pk -> 상세 컬럼 딕셔너리)에 함께 누적해
    건물정보.xlsx 마스터 테이블 데이터로 쓴다.

    반환값은 이번 파일에서 대장pk가 채워진 O-case 행 수다.
"""
def _fill_o_case_pk(deals, candidates_by_dong: dict, address_pk_cache: dict, buildings_by_pk: dict) -> int:
    o_case_mask = deals["총괄표제 존재여부"] == "O"
    if not o_case_mask.any():
        return 0

    o_case = deals.loc[o_case_mask].copy()
    o_case["_산여부"] = o_case["번지"].astype(str).str.startswith("산")

    beonji_parts = o_case["확정번지"].astype(str).str.split("-", n=1)
    o_case["_확정번"] = beonji_parts.str[0].astype("int64")
    o_case["_확정지"] = beonji_parts.str.get(1).fillna("0").astype("int64")

    o_case["지번주소"] = (
        o_case["시군구"] + " " + o_case["_산여부"].map({True: "산 ", False: ""}) + o_case["확정번지"]
    )

    # 이 파일 안에서 유니크한 지번주소만 골라, 아직 캐시에 없는 것만 표제부 조회한다.
    unique_addresses = o_case.drop_duplicates(subset="지번주소")
    for address in unique_addresses.to_dict("records"):
        addr_key = address["지번주소"]
        if addr_key in address_pk_cache:
            continue
        basis_rows = _match_address_basis_rows(address, candidates_by_dong)
        if basis_rows[0]["대장pk"] is None:
            address_pk_cache[addr_key] = None
            continue
        for basis_row in basis_rows:
            buildings_by_pk[(SOURCE_CURRENT, basis_row["대장pk"])] = {"지번주소": addr_key, "대장출처": SOURCE_CURRENT, **basis_row}
        address_pk_cache[addr_key] = ",".join(str(basis_row["대장pk"]) for basis_row in basis_rows)

    pk_series = o_case["지번주소"].map(address_pk_cache)
    deals.loc[o_case_mask, "대장pk"] = pk_series.values
    return int(pk_series.notna().sum())


"""지번주소 하나(딕셔너리 형태의 행)에 대해 매핑되는 표제부 row(들)를 찾는다.

    산 지번이면 대지구분 == SAN_DAEJI_GUBUN_CODE('1', 산) 조건을 추가로 건다. 매칭된 row마다 {OUTPUT_EXTRA_COLUMNS...(대장pk 포함)}
    딕셔너리를 만들어 리스트로 반환하고, 하나도 못 찾으면 값이 모두 비어있는 딕셔너리 1개짜리 리스트를 반환한다.
"""
def _match_address_basis_rows(address: dict, candidates_by_dong: dict) -> list:
    candidates = candidates_by_dong.get(address["_시군구"])
    if candidates is not None:
        mask = (candidates["_번_정수"] == address["_확정번"]) & (candidates["_지_정수"] == address["_확정지"])
        if address["_산여부"]:
            mask &= candidates["대지구분"] == SAN_DAEJI_GUBUN_CODE
        hits = candidates[mask]
    else:
        hits = None

    if hits is None or hits.empty:
        return [{col: None for col in OUTPUT_EXTRA_COLUMNS}]

    return [
        {BASIS_EXTRA_COLUMN_RENAME.get(col, col): matched[col] for col in BASIS_EXTRA_COLUMNS}
        for matched in hits.to_dict("records")
    ]


"""buildings_by_pk(건축물대장 PK -> 건물 상세 컬럼 딕셔너리)에 누적된, 전체 실행(모든 입력 xlsx)에서
매칭된 건물들을 대장pk 기준으로 중복 없이 하나의 마스터 엑셀로 저장한다.

    마스킹된 번지 매칭 경로(_match_deal)와 총괄표제 O-case 경로(_match_address_basis_rows) 양쪽에서
    나온 건물이 섞여 들어오며, 동일 대장pk는 buildings_by_pk 단계에서 이미 하나로 합쳐져 있다.
    기존 단독다가구_총괄있는_표제부.xlsx(총괄표제 O-case 지번주소 전용, 지번주소 컬럼 포함)를 대체한다.
    두 매칭 경로 모두 지번주소를 계산해 buildings_by_pk에 함께 담아두므로, 대장pk 옆에 지번주소
    컬럼도 포함해 저장한다.
    대장출처(표제부/폐쇄말소)와 폐쇄말소일(폐쇄말소 표제부에서 찾은 건물만)도 함께 저장한다."""
def _write_building_info(buildings_by_pk: dict):
    result_columns = ["대장pk", "대장출처", "지번주소", "폐쇄말소일"] + [col for col in OUTPUT_EXTRA_COLUMNS if col != "대장pk"]
    result = pd.DataFrame(buildings_by_pk.values(), columns=result_columns)

    building_count = len(result)
    print(f"매칭된 고유 건물(대장pk) 수: {building_count}")

    output_path = os.path.join(OUTPUT_DIR, "단독다가구_건물정보.xlsx")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name=OUTPUT_SHEET_NAME, index=False, startrow=DATA_START_ROW)
        sheet = writer.sheets[OUTPUT_SHEET_NAME]
        sheet["A1"] = "total_info"
        sheet["B1"] = f"매칭된 고유 건물(대장pk) 수: {building_count}"

    print(f"완료: {output_path}")


"""파일별 매칭 통계를 하나의 요약 xlsx로 저장한다.

    성공률은 표제부에 건물정보 있는(대장pk가 채워진) 개수 / 전체 개수다.
    마지막에 합계 행(개수 합계)과 성공률 평균 행(파일별 성공률의 단순 평균)을 덧붙인다.
"""
def _write_summary(summary_rows: list):
    summary = pd.DataFrame(
        summary_rows,
        columns=[
            "파일명", "전체 개수", "확정번지 있는 개수", "표제부에 건물정보 있는 개수", "폐쇄말소 매칭 개수",
            "허용오차 매칭 개수", "성공률",
        ],
    )

    total_count = summary["전체 개수"].sum()
    total_beonji = summary["확정번지 있는 개수"].sum()
    total_building = summary["표제부에 건물정보 있는 개수"].sum()
    total_closed = summary["폐쇄말소 매칭 개수"].sum()
    total_tolerance = summary["허용오차 매칭 개수"].sum()
    mean_rate = summary["성공률"].mean() if not summary.empty else 0
    summary.loc[len(summary)] = ["합계", total_count, total_beonji, total_building, total_closed, total_tolerance, None]
    summary.loc[len(summary)] = ["성공률 평균", None, None, None, None, None, mean_rate]

    output_path = os.path.join(OUTPUT_DIR, SUMMARY_FILE_NAME)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name=SUMMARY_SHEET_NAME, index=False)
        sheet = writer.sheets[SUMMARY_SHEET_NAME]
        # 성공률 컬럼(G열)을 퍼센트 형식으로 표시한다. 1행은 헤더.
        for row in range(2, len(summary) + 2):
            sheet.cell(row=row, column=7).number_format = "0.0%"

    print(f"요약 통계 완료: {output_path} (성공률 평균: {mean_rate:.1%})")


def main():
    xlsx_paths = sorted(glob.glob(os.path.join(INPUT_DIR, "*.xlsx")))
    if not xlsx_paths:
        raise FileNotFoundError(f"{INPUT_DIR} 에 xlsx 파일이 없습니다.")

    candidates_by_dong = _load_basis_candidates(BASIS_CSV, BASIS_CSV_COLUMN_RENAME)
    closed_candidates_by_dong = _load_basis_candidates(CLOSED_BASIS_CSV, CLOSED_BASIS_CSV_COLUMN_RENAME)

    print("매칭 중...")
    address_pk_cache = {}  # 지번주소 -> 대장pk 문자열(또는 None) 캐시, 파일 간 재사용
    buildings_by_pk = {}  # (대장출처, 대장pk) -> 건물 상세 딕셔너리, 전체 실행 동안 누적 -> 건물정보.xlsx
    summary_rows = []
    for xlsx_path in xlsx_paths:
        print(f"{os.path.basename(xlsx_path)} 로딩 중...")
        deals, total_count = _load_deals(xlsx_path)
        success_count, pk_filled_count, closed_count, tolerance_count = _match_and_write(
            xlsx_path, deals, total_count, candidates_by_dong, closed_candidates_by_dong, address_pk_cache, buildings_by_pk
        )
        building_rate = pk_filled_count / total_count if total_count else 0
        summary_rows.append(
            [os.path.basename(xlsx_path), total_count, success_count, pk_filled_count, closed_count, tolerance_count,
             building_rate]
        )

    _write_building_info(buildings_by_pk)
    _write_summary(summary_rows)


if __name__ == "__main__":
    main()

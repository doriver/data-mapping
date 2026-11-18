# 토지 실거래가(마스킹된 번지)에, 토지 특성정보 CSV를 대조해 실제 번지와 PNU등을 채워 넣는 스크립트.
# 지역(서울/경기/... 등)별로 실거래가 xlsx와 토지특성정보 csv가 따로 존재할 수 있으므로,
# data/raw, data/basis를 훑어 파일명 접두어(지역명)로 짝을 찾아 지역별로 순회 처리한다.
import gc
import glob
import os
import re
import sys

import numpy as np
import pandas as pd

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

RAW_DIR = r"D:\dev\data\raw\traded_land_raw\25"
BASIS_DIR = r"D:\dev\data\basis\land_characteristics\26"
# 파일명이 이 이름들 중 하나로 시작하면 해당 지역으로 판별한다.
# (예: "서울25년_토지(매매)_실거래가.xlsx", "서울토지특성정보_20260519.csv" -> "서울")
# 실제 파일명 규칙이 다르면 이 목록만 조정하면 된다.
REGION_NAMES = [
    "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
    "경기", "강원", "충청북", "충북", "충청남", "충남", "전북", "전라남", "전남", "경상북", "경북", "경상남", "경남", "제주",
]
# 법정동 변경 이력. 실거래는 최신 법정동명, 토지특성정보는 해당 연도 시점 법정동명을 쓰므로
# csv에 실거래 법정동이 없으면 이 이력을 거꾸로 따라가 과거 법정동명으로 다시 조회한다.
BJD_CHANGE_DIR = r"D:\dev\data\basis\bjd_change"
BJD_CHANGE_PATTERN = "tbl_bjd_code_mapping_history_*.csv"
BJD_CHANGE_ENCODING = "utf-8-sig"
CSV_ENCODING = "cp949"
CSV_CHUNK_SIZE = 200_000
CSV_USECOLS = ["고유번호", "법정동명", "지번", "지목명", "용도지역명1", "토지면적", "대장구분명", "공시지가", "도로접면코드"]
HEADER_ROW = 12  # 실제 표 헤더가 있는 엑셀 행(0-index)
OUTPUT_DIR = r"D:\dev\data\processed\traded_land_processed\25"
DATA_START_ROW = 5  # 0-index. 상단 4행(전체/제외/처리대상/성공 건수) + 빈 줄을 남기고 그 아래부터 표 작성
AREA_TOLERANCE = 0.01
# 실거래 용도지역 -> 토지특성정보 용도지역명1 별칭. 같은 용도지역인데 표기가 다른 경우 비교 전에 맞춘다.
# (실거래는 "자연환경보전", 토지특성정보는 "자연환경보전지역"으로 표기)
USE_ZONE_ALIASES = {
    "자연환경보전": "자연환경보전지역",
}
EMPTY_VALUE = "-"  # 실제 거래된거를 나타내는 값( 해제된경우는 날짜 들어가있음 )
ROAD_CONDITION_TO_ROAD_SIDE_CODES = {
    "25m이상": {"01", "02", "03"},
    "25m미만": {"04", "05"},
    "12m미만": {"06", "07"},
    "8m미만": {"08", "09", "10", "11"},
    "-": {"12", "00"},
}
# 도로접면코드 -> 도로조건 역매핑. 후보를 도로조건 기준으로 미리 그룹핑하는 데 사용한다.
ROAD_SIDE_CODE_TO_CONDITION = {
    code: condition for condition, codes in ROAD_CONDITION_TO_ROAD_SIDE_CODES.items() for code in codes
}

"""마스킹된 번지 문자열을 (자리수 접두어, 자리수) 로 분리한다.

    '산4*' -> ('4', 2), '산*' -> ('', 1), '1***' -> ('1', 4)
"""
def parse_beonji(raw: str):    
    body = raw[1:] if raw.startswith("산") else raw
    prefix = body.split("*", 1)[0]
    return prefix, len(body)

"""파일명이 REGION_NAMES 중 하나로 시작하면 그 지역명을 반환한다."""
def _match_region(filename: str):
    
    for region in REGION_NAMES:
        if filename.startswith(region):
            return region
    return None

"""법정동명을 비교용으로 정규화한다. (실거래 시군구, 토지특성 법정동명, 법정동 변경 이력 이름에 공통 적용)

    - 앞뒤 공백 제거
    - 텍스트 사이 연속 공백(2칸 이상, 탭/전각 공백 포함)을 한 칸으로
    문자열이 아니면(NaN 등) 그대로 반환한다.
    예: "  경기도 화성시   새솔동 " -> "경기도 화성시 새솔동"
"""
def _normalize_dong_name(name):
    if not isinstance(name, str):
        return name
    return re.sub(r"\s+", " ", name.strip())

"""법정동 변경 이력 csv를 읽어 {최신 법정동명: {과거 법정동명, ...}} 역방향 dict를 만든다.

    이름은 "시도 시군구 읍면동 [리]"를 공백으로 이은 전체 이름이다(실거래 시군구, csv 법정동명과 같은 형식).
    (예: "경기도 화성시 만세구 새솔동" -> {"경기도 화성시 새솔동"})
    읍면동이 없는 시도/시군구 단위 행은 실거래 시군구와 비교할 수 없으므로 제외한다.
    합병처럼 새 이름 하나에 과거 이름이 여럿일 수 있으므로 값은 집합이다.
"""
def _load_bjd_reverse_map() -> dict:
    paths = sorted(glob.glob(os.path.join(BJD_CHANGE_DIR, BJD_CHANGE_PATTERN)))
    if not paths:
        print(f"경고: 법정동 변경 이력 파일이 없어 과거 법정동 회귀 조회를 하지 않음 ({BJD_CHANGE_DIR})")
        return {}
    path = paths[-1]  # 파일명 끝의 생성시각 기준 최신 파일
    history = pd.read_csv(path, encoding=BJD_CHANGE_ENCODING, dtype=str)
    history = history[history["prev_emd_nm"].notna() & history["next_emd_nm"].notna()]

    def full_name(row, side):
        parts = [row[f"{side}_{level}_nm"] for level in ("sido", "sgg", "emd", "ri")]
        return _normalize_dong_name(" ".join(p.strip() for p in parts if isinstance(p, str) and p.strip()))

    reverse_map = {}
    for _, row in history.iterrows():
        prev_name, next_name = full_name(row, "prev"), full_name(row, "next")
        if prev_name != next_name:
            reverse_map.setdefault(next_name, set()).add(prev_name)
    print(f"법정동 변경 이력 로드: {os.path.basename(path)} (읍면동 단위 {len(reverse_map)}건)")
    return reverse_map

"""변경 이력을 거꾸로 따라가(BFS) dong의 과거 법정동명들을 가까운 시점부터 순서대로 반환한다. (dong 자신은 제외)"""
def _ancestors(dong: str, reverse_map: dict) -> list:
    result = []
    visited = {dong}
    frontier = [dong]
    while frontier:
        next_frontier = []
        for name in frontier:
            for prev_name in sorted(reverse_map.get(name, ())):
                if prev_name not in visited:
                    visited.add(prev_name)
                    result.append(prev_name)
                    next_frontier.append(prev_name)
        frontier = next_frontier
    return result

"""RAW_DIR의 xlsx와 BASIS_DIR의 csv를 파일명 접두어(지역명) 기준으로 짝짓는다.

    같은 지역의 실거래가 xlsx는 (기간 등에 따라)여러 개 있을 수 있으므로 개수 제한 없이 모두 모으고,
    해당 지역의 토지특성정보 csv 1개를 공통으로 매칭해 각 xlsx마다 개별 처리·출력한다.
"""
def discover_region_files():    
    xlsx_by_region = {}
    for path in glob.glob(os.path.join(RAW_DIR, "*.xlsx")):
        region = _match_region(os.path.basename(path))
        if region is None:
            print(f"건너뜀: 지역명을 알 수 없는 실거래가 파일 - {path}")
            continue
        xlsx_by_region.setdefault(region, []).append(path)

    csv_by_region = {}
    for path in glob.glob(os.path.join(BASIS_DIR, "*.csv")):
        region = _match_region(os.path.basename(path))
        if region is None:
            print(f"건너뜀: 지역명을 알 수 없는 토지특성정보 파일 - {path}")
            continue
        if region in csv_by_region:
            raise ValueError(f"'{region}' 지역의 토지특성정보 파일이 2개 이상 발견됨: {csv_by_region[region]}, {path}")
        csv_by_region[region] = path

    region_files = {}
    for region in xlsx_by_region.keys() | csv_by_region.keys():
        xlsx_paths = xlsx_by_region.get(region)
        csv_path = csv_by_region.get(region)
        if not xlsx_paths or csv_path is None:
            print(f"건너뜀: '{region}' 지역은 xlsx/csv 짝이 없음 (xlsx={xlsx_paths}, csv={csv_path})")
            continue
        region_files[region] = (xlsx_paths, csv_path)
    return region_files


def _load_deals(xlsx_path: str):
    deals = pd.read_excel(xlsx_path, header=HEADER_ROW)
    total_count = len(deals)

    # 계약면적 30 이하인 소규모 거래는 분석 대상에서 제외한다.
    small_area_mask = deals["계약면적"] <= 30
    small_area_count = small_area_mask.sum()
    deals = deals[~small_area_mask].reset_index(drop=True)

    # 해제(취소) 신고가 등록되면, 취소되기 전 원본 행(해제사유발생일 "-")이 지워지지 않고, 취소 행(해제사유발생일에 날짜)과 별개로 하나 더 남아 중복되는 경우가 있다.
    # 이런 원본 중복 행은 잘못된 데이터이므로 집계 전에 제거한다.
    dup_key_cols = ["시군구", "번지", "지목", "용도지역", "도로조건", "계약면적", "거래금액(만원)", "지분구분", "거래유형"]
    has_cancel_date = deals["해제사유발생일"].notna() & (deals["해제사유발생일"] != EMPTY_VALUE)
    group_ids = deals.groupby(dup_key_cols, dropna=False).ngroup()
    cancelled_group_ids = set(group_ids[has_cancel_date])
    duplicate_original_mask = (~has_cancel_date) & group_ids.isin(cancelled_group_ids)
    if duplicate_original_mask.any():
        print(f"{os.path.basename(xlsx_path)} 해제 중복 원본 행 제거: {duplicate_original_mask.sum()}건")
    deals = deals[~duplicate_original_mask].reset_index(drop=True)

    # 거래 취소된경우 제외
    cancelled_mask = deals["해제사유발생일"].notna() & (deals["해제사유발생일"] != EMPTY_VALUE)
    cancelled_count = cancelled_mask.sum()
    deals = deals[~cancelled_mask].reset_index(drop=True)

    # 제외할꺼 제외하고, 처리대상이 되는 전체 데이터 수
    actual_target_count = len(deals)

    # 지분구분이 "지분"인 거래(지분 거래) 건수
    share_mask = deals["지분구분"] == "지분"
    share_count = share_mask.sum()

    return deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count


"""csv_path를 청크 단위로 한 번만 읽어, needed_dongs에 해당하는 행만 묶어 반환한다.

    반환 딕셔너리의 키는 (법정동명, 산여부) 튜플이다.
    (지목/용도지역은 "둘 중 하나 일치" 조건이라 groupby 키로 미리 나눌 수 없어 키에서 제외했다.)
    _match_and_write에서 row마다 반복되는 등호 조건(법정동명/산여부)을
    미리 groupby로 나눠둠으로써, row 순회 시에는 훨씬 작아진 후보군에서
    번지(본번 자리수/접두어)와 면적만 비교하면 되도록 한다.
    반환: (candidates_by_group, present_dongs)
    present_dongs는 needed_dongs 중 이 csv에 실제로 존재하는 법정동명 집합이다.
    ("csv에 해당 법정동이 있는지" 판단에 쓴다. 없으면 과거 법정동으로 회귀한다.)
"""
def _load_candidates_by_dong(region: str, csv_path: str, needed_dongs: set) -> tuple:
    candidate_chunks = []
    for chunk in pd.read_csv(
        csv_path, encoding=CSV_ENCODING, dtype=str, usecols=CSV_USECOLS, chunksize=CSV_CHUNK_SIZE
    ):
        # 실거래 시군구와 같은 기준으로 비교하도록 법정동명을 정규화한다.
        # (청크 안 고유 법정동명은 수백 개 수준이므로, 고유값만 정규화해 map으로 펼친다.)
        dong_names = chunk["법정동명"]
        normalized = {name: _normalize_dong_name(name) for name in dong_names.dropna().unique()}
        chunk["법정동명"] = dong_names.map(normalized)
        sub = chunk[chunk["법정동명"].isin(needed_dongs)]
        if len(sub):
            candidate_chunks.append(sub)
    candidates = pd.concat(candidate_chunks, ignore_index=True)
    candidates = candidates.dropna(subset=["지번"]).copy()
    present_dongs = set(candidates["법정동명"].unique())
    # 지번은 "본번-부번" 형태이므로, 자리수/접두어 비교에 쓸 본번만 분리해둔다.
    candidates["본번"] = candidates["지번"].str.split("-", n=1).str[0]
    # row마다 다시 계산하지 않도록 본번 자리수를 미리 구해둔다.
    candidates["본번_자리수"] = candidates["본번"].str.len()
    candidates["토지면적"] = candidates["토지면적"].astype(float)
    candidates["공시지가"] = candidates["공시지가"].astype(float)
    candidates["산여부"] = candidates["대장구분명"] == "산"
    candidates["도로조건"] = candidates["도로접면코드"].map(ROAD_SIDE_CODE_TO_CONDITION)
    print(f"[{region}] 토지 특성정보 후보 건수(대상 법정동 내): {len(candidates)}")

    group_keys = ["법정동명", "산여부"]
    return {key: df for key, df in candidates.groupby(group_keys)}, present_dongs

""" 실질적인 매핑 로직
    실거래 법정동(최신)이 csv에 없으면 변경 이력을 거꾸로 따라가 csv에 있는 과거 법정동으로 조회한다.
    (PNU/확정번지는 csv에서 찾은 과거 값을 그대로 쓰고, 실제 조회한 법정동명은 조회법정동 컬럼에 남긴다.)
"""
def _match_and_write(region: str, xlsx_path: str, deals, total_count: int, actual_target_count: int, cancelled_count: int, small_area_count: int, share_count: int, candidates_by_group: dict, present_dongs: set, reverse_map: dict) -> dict:
    land_total_count = actual_target_count - share_count  # 특정토지 전체 거래 건수

    # 실거래 법정동 -> 이 csv에서 조회할 법정동 (같은 법정동이 반복되므로 캐시)
    lookup_dong_cache = {}

    def resolve_lookup_dong(raw_dong):
        if raw_dong not in lookup_dong_cache:
            lookup = None
            # csv 법정동명(present_dongs)도 정규화돼 있으므로 실거래 시군구도 정규화한 뒤 비교한다.
            dong = _normalize_dong_name(raw_dong)
            for name in [dong] + _ancestors(dong, reverse_map):
                if name in present_dongs:
                    lookup = name
                    break
            lookup_dong_cache[raw_dong] = lookup
        return lookup_dong_cache[raw_dong]

    # (조회법정동, 산여부, 자리수) -> 자리수까지 맞는 후보 DataFrame
    # (조회법정동, 산여부, 자리수, 접두어) -> 3번 조건까지 통과한 후보의 numpy 배열 묶음
    # 같은 법정동/자리수/접두어 조합이 여러 행에서 반복되므로, 후보 전체 필터링은 조합마다 한 번만 하고
    # 행마다는 작아진 numpy 배열에서 면적/지목/용도지역/도로조건만 비교한다.
    digit_cache = {}
    prefix_cache = {}

    def get_candidate_arrays(lookup_dong, is_san, digit_count, prefix):
        key = (lookup_dong, is_san, digit_count, prefix)
        if key not in prefix_cache:
            digit_key = (lookup_dong, is_san, digit_count)
            if digit_key not in digit_cache:
                group = candidates_by_group.get((lookup_dong, is_san))
                digit_cache[digit_key] = None if group is None else group[group["본번_자리수"] == digit_count]
            by_digit = digit_cache[digit_key]
            if by_digit is None:
                prefix_cache[key] = None
            else:
                sub = by_digit[by_digit["본번"].str.startswith(prefix)]
                prefix_cache[key] = {
                    "고유번호": sub["고유번호"].to_numpy(dtype=object),
                    "지번": sub["지번"].to_numpy(dtype=object),
                    "토지면적": sub["토지면적"].to_numpy(dtype=float),
                    "공시지가": sub["공시지가"].to_numpy(dtype=float),
                    "지목명": sub["지목명"].to_numpy(dtype=object),
                    "용도지역명1": sub["용도지역명1"].to_numpy(dtype=object),
                    "도로조건": sub["도로조건"].to_numpy(dtype=object),
                }
        return prefix_cache[key]

    matched_lookup_dong = []
    matched_pnu = []
    matched_beonji = []
    matched_area = []
    matched_official_price = []
    for row in deals.itertuples(index=False):
        # 0. 번지가 비어있는(NaN) 행은 매칭 불가능하므로 바로 실패 처리한다.
        #    (예: 경상북도 데이터의 포항시 북구 흥해읍 곡강리 건처럼 원본에 번지 누락이 있는 경우)
        if pd.isna(row.번지):
            matched_lookup_dong.append(None)
            matched_pnu.append(None)
            matched_beonji.append(None)
            matched_area.append(None)
            matched_official_price.append(None)
            continue

        # 1. 마스킹된 번지("1***", "산4*" 등)를 "산 여부 / 앞부분 접두어 / 전체 자리수"로 분해한다.
        #    산으로 시작하면 "산"을 뗀 뒷부분이 실제 번지 문자열이다.
        is_san = row.번지.startswith("산")
        prefix, digit_count = parse_beonji(row.번지)

        # 2. [중요조건] 시군구(법정동명)/산여부가 모두 일치하는 후보군으로 범위를 좁힌다.
        #    (도로조건은 여기서 제외하고, 5번 단계에서 후보가 2건 이상일 때만 서브조건으로 적용한다.)
        #    (이 조건들은 row마다 값이 그대로 반복되므로, candidates_by_group을 만들 때 미리
        #    groupby 해뒀다. 매칭에 실패할 조합이면 딕셔너리에 아예 없으므로 후보군이 없다.)
        #    실거래 법정동이 csv에 없으면(법정동 개편 등 시점 차이) 과거 법정동으로 회귀해 조회한다.
        lookup_dong = resolve_lookup_dong(row.시군구)

        # 3. [중요조건] 토지 특성정보의 본번이 접두어로 시작하고, 마스킹 전 번지와 자리수가 같아야 한다.
        #    (자리수까지 맞춰야 "1**"가 "12" 같은 다른 자리수 번지와 잘못 매칭되지 않는다.)
        #    (이 조건도 row마다 값이 반복되므로 get_candidate_arrays에서 조합별로 한 번만 계산해 캐시한다.)
        candidates = None if lookup_dong is None else get_candidate_arrays(lookup_dong, is_san, digit_count, prefix)
        if candidates is None or not len(candidates["고유번호"]):
            matched_lookup_dong.append(None)
            matched_pnu.append(None)
            matched_beonji.append(None)
            matched_area.append(None)
            matched_official_price.append(None)
            continue

        # 4. [중요조건] 면적 조건: 지분구분이 없는 단독 거래는 계약면적과 오차범위 내로 일치해야 하고,
        #    지분 거래는 토지 전체 면적 중 일부만 거래된 것이므로 토지면적이 계약면적 이상이어야 한다.
        area = candidates["토지면적"]
        if pd.isna(row.지분구분):
            mask = (area > row.계약면적 - AREA_TOLERANCE) & (area < row.계약면적 + AREA_TOLERANCE)
        else:
            mask = area >= row.계약면적

        # 5. [중요조건] 용도지역 또는 지목 중 하나라도 일치해야 한다.
        #    (실거래 용도지역은 USE_ZONE_ALIASES로 토지특성정보 표기에 맞춘 뒤 비교한다. 결과 엑셀의 원본 값은 그대로 둔다.)
        use_zone = USE_ZONE_ALIASES.get(row.용도지역, row.용도지역)
        mask &= (candidates["지목명"] == row.지목) | (candidates["용도지역명1"] == use_zone)
        # hits: 조건을 통과한 후보의 위치(index) 배열
        hits = np.flatnonzero(mask)

        # 6. [서브조건] 중요조건을 만족하는 후보가 1건이면 바로 매칭 확정.
        #    2건 이상이면(중요조건만으로 특정 불가) 아래 서브조건을 순서대로 하나씩 적용해 다시 좁힌다.
        #    (1) 도로조건 일치 -> (2) 용도지역, 지목 모두 일치
        #    각 단계에서 1건이 되면 확정, 0건이면 실패, 2건 이상이면 다음 서브조건으로 넘어간다.
        #    (모든 서브조건을 거친 뒤에도 2건 이상이면 특정할 수 없으므로 매칭하지 않는다.)
        sub_filters = [
            lambda idx: candidates["도로조건"][idx] == row.도로조건,
            lambda idx: (candidates["지목명"][idx] == row.지목) & (candidates["용도지역명1"][idx] == use_zone),
        ]
        for sub_filter in sub_filters:
            if len(hits) < 2: # 서브조건을 적용하기 전에 먼저 건수를 확인
                break
            hits = hits[sub_filter(hits)]

        # 7. 최종 후보가 정확히 1건일 때만 매칭 확정(PNU/실제 번지 확정)한다.
        if len(hits) == 1:
            hit = hits[0]
            matched_lookup_dong.append(lookup_dong)
            matched_pnu.append(candidates["고유번호"][hit])
            beonji = candidates["지번"][hit]
            if is_san:
                beonji = f"산 {beonji}"
            matched_beonji.append(beonji)
            matched_area.append(candidates["토지면적"][hit])
            matched_official_price.append(candidates["공시지가"][hit])
        else:
            matched_lookup_dong.append(None)
            matched_pnu.append(None)
            matched_beonji.append(None)
            matched_area.append(None)
            matched_official_price.append(None)

    deals["PNU"] = matched_pnu
    deals["확정번지"] = matched_beonji
    deals["토지면적"] = matched_area
    deals["공시지가"] = matched_official_price
    # 조회법정동: 실제로 토지특성정보를 조회한 법정동명(과거 법정동으로 회귀했다면 과거 이름)
    deals["조회법정동"] = matched_lookup_dong

    success_count = deals["PNU"].notna().sum()
    # 공백 차이만 있는 경우를 회귀로 세지 않도록 시군구도 정규화해서 비교한다.
    regressed_count = (deals["조회법정동"].notna() & (deals["조회법정동"] != deals["시군구"].map(_normalize_dong_name))).sum()
    share_deal_mask = deals["지분구분"] == "지분"
    share_success_count = (share_deal_mask & deals["PNU"].notna()).sum()  # 지분거래 매칭 성공 건수
    land_success_count = (~share_deal_mask & deals["PNU"].notna()).sum()
    print(f"[{region}] {os.path.basename(xlsx_path)} 매칭 성공 건수: {success_count} / {actual_target_count} (그중 과거 법정동 회귀 매칭 {regressed_count}건)")
    print(f"[{region}] {os.path.basename(xlsx_path)} 지분거래 매칭 성공 건수: {share_success_count} / {share_count}")
    print(f"[{region}] {os.path.basename(xlsx_path)} 특정토지 전체 거래 매칭 성공 건수: {land_success_count} / {land_total_count}")

    _xlsx_name, _xlsx_ext = os.path.splitext(os.path.basename(xlsx_path))
    output_path = os.path.join(OUTPUT_DIR, f"{_xlsx_name}_매핑{_xlsx_ext}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        deals.to_excel(writer, sheet_name="실거래가", index=False, startrow=DATA_START_ROW)
        sheet = writer.sheets["실거래가"]
        sheet["A1"] = f"total_info"
        sheet["B1"] = f"전체 건수: {total_count}"
        sheet["G1"] = f"실제 처리대상 건수: {actual_target_count}"
        sheet["L1"] = f"매칭 성공 건수: {success_count}"
        overall_success_rate = success_count / actual_target_count if actual_target_count else 0
        sheet["Q1"] = f"매칭 성공률: {overall_success_rate:.1%}"

        sheet["A2"] = f"제외 case"
        sheet["B2"] = f"거래취소로 제외된 건수: {cancelled_count}"
        sheet["G2"] = f"면적30이하로 제외된 건수 :  {small_area_count}"

        sheet["A3"] = f"지분거래"
        sheet["B3"] = f"지분 거래 건수: {share_count}"
        share_ratio = share_count / actual_target_count if actual_target_count else 0
        sheet["G3"] = f"지분거래 비율: {share_ratio:.1%}"
        sheet["L3"] = f"지분거래 매칭 성공 건수: {share_success_count}"
        share_success_rate = share_success_count / share_count if share_count else 0
        sheet["Q3"] = f"지분거래 성공률: {share_success_rate:.1%}"
        
        sheet["A4"] = f"전체거래"
        sheet["B4"] = f"특정토지 전체거래 건수: {land_total_count}"
        sheet["L4"] = f"특정토지 전체거래 매칭 성공 건수: {land_success_count}"
        land_success_rate = land_success_count / land_total_count if land_total_count else 0    
        sheet["Q4"] = f"전체거래 성공률: {land_success_rate:.1%}"
        
    print(f"[{region}] 완료: {output_path}")

    return {
        "region": region,
        "file": os.path.basename(xlsx_path),
        "actual_target_count": actual_target_count,
        "cancelled": cancelled_count,
        "success": success_count,
        "share_ratio": share_ratio,
        "land_success_rate": land_success_rate,
        "share_success_rate": share_success_rate,
        "output": output_path,
    }

"""같은 지역의 xlsx 여러 개가 csv 1개를 공유하므로, csv는 지역당 한 번만 읽고 xlsx만 반복 처리한다.

    실거래 법정동이 csv에 없으면 reverse_map(법정동 변경 이력)으로 과거 법정동에 회귀해 조회하므로,
    csv에서 걸러 둘 법정동(needed_dongs)에 과거 법정동명도 함께 넣는다.
"""
def process_region_group(region: str, xlsx_paths, csv_path: str, reverse_map: dict) -> list:
    loaded = []
    needed_dongs = set()
    for xlsx_path in xlsx_paths:
        try:
            deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count = _load_deals(xlsx_path)
        except Exception as e:
            print(f"[{region}] {xlsx_path} 실거래가 로드 실패로 건너뜀: {e}")
            continue
        print(f"[{region}] {os.path.basename(xlsx_path)} 실거래가 전체 건수: {total_count}")
        print(f"[{region}] {os.path.basename(xlsx_path)} 해제사유발생일 있어 제외된 건수: {cancelled_count}")
        print(f"[{region}] {os.path.basename(xlsx_path)} 계약면적 30이하로 제외된 건수: {small_area_count}")
        print(f"[{region}] {os.path.basename(xlsx_path)} 지분 거래 건수: {share_count}")
        for raw_dong in deals["시군구"].dropna().unique():
            # csv 법정동명과 같은 기준으로 비교하도록 정규화한 이름으로 키를 만든다.
            dong = _normalize_dong_name(raw_dong)
            needed_dongs.add(dong)
            needed_dongs.update(_ancestors(dong, reverse_map))
        loaded.append((xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count))

    if not loaded:
        return []

    # 실거래가에 등장하는 시군구(법정동명)만 후보로 남기면 되므로, CSV 전체를 메모리에 올리지 않고
    # 청크 단위로 읽으면서 법정동명이 일치하는 행만 추려낸다. 이 지역의 xlsx 전체에서 필요한 동을
    # 먼저 모아뒀기 때문에 CSV는 이 한 번만 읽으면 된다.
    candidates_by_group, present_dongs = _load_candidates_by_dong(region, csv_path, needed_dongs)

    summaries = []
    # loaded에서 하나씩 꺼내 처리 즉시 참조를 없애, 처리 끝난 xlsx의 데이터가 다음 xlsx 처리 중에도
    # 계속 메모리에 남아있지 않도록 한다.
    while loaded:
        xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count = loaded.pop(0)
        try:
            summaries.append(_match_and_write(region, xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count, candidates_by_group, present_dongs, reverse_map))
        except Exception as e:
            print(f"[{region}] {xlsx_path} 매칭 처리 실패로 건너뜀: {e}")
        finally:
            del deals
            gc.collect()

    # 이 지역의 CSV 후보(candidates_by_group)는 다음 지역 처리 전에 해제한다.
    del candidates_by_group, present_dongs
    gc.collect()
    return summaries


def main():
    region_files = discover_region_files()
    if not region_files:
        print("처리할 지역 파일 쌍을 찾지 못했습니다.")
        return

    # 법정동 변경 이력(최신 -> 과거 역방향)은 전 지역 공통이므로 한 번만 읽는다.
    reverse_map = _load_bjd_reverse_map()

    # 실질적인 처리및 매핑로직은 메서드 process_region_group에 다 들어있음
    summaries = []
    for region, (xlsx_paths, csv_path) in region_files.items():
        try:
            summaries.extend(process_region_group(region, xlsx_paths, csv_path, reverse_map))
        except Exception as e:
            print(f"[{region}] 처리 중 오류로 건너뜀: {e}")

    print("\n=== 전국 처리 요약 ===")
    for s in summaries:
        print(f"[{s['region']}] 매칭 성공 {s['success']} / {s['actual_target_count']} -> {s['output']}")
    region_count = len(set(s["region"] for s in summaries))
    print(f"전체 처리 파일 수: {len(summaries)} (지역 수: {region_count}), 총 매칭 성공: {sum(s['success'] for s in summaries)}")

    # 원본 파일명별 핵심 통계(실제 처리대상 건수, 지분거래 비율, 전체거래 성공률, 지분거래 성공률)를
    # 개별 매핑 xlsx를 일일이 열어보지 않고 한눈에 비교할 수 있도록 별도 요약 xlsx로 모아둔다.
    summary_rows = [
        {
            "원본파일명": s["file"],
            "실제 처리대상 건수": s["actual_target_count"],
            "지분거래 비율": f"{s['share_ratio']:.1%}",
            "전체거래 성공률": f"{s['land_success_rate']:.1%}",
            "지분거래 성공률": f"{s['share_success_rate']:.1%}",
        }
        for s in summaries
    ]
    if summary_rows:
        avg_share_ratio = sum(s["share_ratio"] for s in summaries) / len(summaries)
        avg_land_success_rate = sum(s["land_success_rate"] for s in summaries) / len(summaries)
        summary_rows.append({
            "원본파일명": "평균",
            "실제 처리대상 건수": "",
            "지분거래 비율": f"{avg_share_ratio:.1%}",
            "전체거래 성공률": f"{avg_land_success_rate:.1%}",
            "지분거래 성공률": "",
        })

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        summary_path = os.path.join(OUTPUT_DIR, "통계_요약_토지매핑.xlsx")
        pd.DataFrame(summary_rows).to_excel(summary_path, sheet_name="통계", index=False)
        print(f"통계 요약 저장: {summary_path}")


if __name__ == "__main__":
    main()

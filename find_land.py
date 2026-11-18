# 토지 실거래가(마스킹된 번지)에, 토지 특성정보 CSV를 대조해 실제 번지와 PNU등을 채워 넣는 스크립트.
# 지역(서울/경기/... 등)별로 실거래가 xlsx와 토지특성정보 csv가 따로 존재할 수 있으므로,
# data/raw, data/basis를 훑어 파일명 접두어(지역명)로 짝을 찾아 지역별로 순회 처리한다.
# 토지특성정보는 여러 연도(BASIS_DIRS)를 우선순위대로 사용한다.
# (1차: 26년 특성정보로 전체 매칭 -> 2차: 1차에서 매칭 안 된 행만 25년 특성정보로 재매칭)
import gc
import glob
import os
import sys
import time

import pandas as pd

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

RAW_DIR = r"D:\dev\data\raw\traded_land_raw\25"
# (연도, 경로) 목록. 리스트 순서가 곧 매칭 우선순위다. 앞 연도에서 매칭 실패한 행만 다음 연도로 재매칭한다.
BASIS_DIRS = [
    ("26", r"D:\dev\data\basis\land_characteristics\26"),
    ("25", r"D:\dev\data\basis\land_characteristics\25"),
]
# 파일명이 이 별칭들 중 하나로 시작하면 해당 표준 지역명으로 판별한다.
# (예: "경남_25_토지(매매)_실거래가.xlsx", "경상남도_토지특성정보_2025-01-23.csv", "경남_26_토지특성정보.csv" -> "경남")
# 파일마다 정식명칭/줄임말이 섞여 있으므로 별칭을 표준명으로 정규화해 짝을 맞춘다.
# 실제 파일명 규칙이 다르면 이 목록만 조정하면 된다.
REGION_ALIASES = {
    "서울특별시": "서울", "서울": "서울",
    "부산광역시": "부산", "부산": "부산",
    "대구광역시": "대구", "대구": "대구",
    "인천광역시": "인천", "인천": "인천",
    "광주광역시": "광주", "광주": "광주",
    "대전광역시": "대전", "대전": "대전",
    "울산광역시": "울산", "울산": "울산",
    "세종특별자치시": "세종", "세종": "세종",
    "경기도": "경기", "경기": "경기",
    "강원특별자치도": "강원", "강원도": "강원", "강원": "강원",
    "충청북도": "충북", "충북": "충북",
    "충청남도": "충남", "충남": "충남",
    "전북특별자치도": "전북", "전라북도": "전북", "전북": "전북",
    "전라남도": "전남", "전남": "전남",
    "경상북도": "경북", "경북": "경북",
    "경상남도": "경남", "경남": "경남",
    "제주특별자치도": "제주", "제주도": "제주", "제주": "제주",
}
# 긴 별칭부터 비교해야 "경기도"가 "경기"보다 먼저 걸린다(결과는 같지만 의도를 명확히).
_REGION_ALIAS_ORDER = sorted(REGION_ALIASES, key=len, reverse=True)
# 법정동 변경 이력. 실거래는 최신 법정동명, 토지특성정보는 해당 연도 시점 법정동명을 쓰므로
# csv에 실거래 법정동이 없으면 이 이력을 거꾸로 따라가 과거 법정동명으로 다시 조회한다.
BJD_CHANGE_DIR = r"D:\dev\data\basis\bjd_change"
BJD_CHANGE_PATTERN = "tbl_bjd_code_mapping_history_*.csv"
BJD_CHANGE_ENCODING = "utf-8-sig"
CSV_ENCODING = "cp949"
CSV_CHUNK_SIZE = 200_000
CSV_USECOLS = ["고유번호", "법정동명", "지번", "지목명", "용도지역명1", "토지면적", "대장구분명", "공시지가", "도로접면코드"]
# 후보 메모리를 줄이기 위해 반복값이 많은 문자열 컬럼은 category로 변환한다.
CANDIDATE_CATEGORY_COLS = ["법정동명", "지목명", "용도지역명1", "도로조건"]
HEADER_ROW = 12  # 실제 표 헤더가 있는 엑셀 행(0-index)
OUTPUT_DIR = r"D:\dev\data\traded_land_processed\25"
DATA_START_ROW = 6  # 0-index. 상단 5행(전체/제외/지분/전체거래/연도별 통계) + 빈 줄을 남기고 그 아래부터 표 작성
AREA_TOLERANCE = 0.01
EMPTY_VALUE = "-"  # 실제 거래된거를 나타내는 값( 해제된경우는 날짜 들어가있음 )
# 조회법정동: 실제로 토지특성정보를 조회한 법정동명(과거 법정동으로 회귀했다면 과거 이름)
RESULT_COLS = ["PNU", "확정번지", "토지면적", "공시지가", "매칭연도", "조회법정동"]
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

"""파일명이 REGION_ALIASES의 별칭 중 하나로 시작하면 그 표준 지역명을 반환한다."""
def _match_region(filename: str):

    for alias in _REGION_ALIAS_ORDER:
        if filename.startswith(alias):
            return REGION_ALIASES[alias]
    return None

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
        return " ".join(p.strip() for p in parts if isinstance(p, str) and p.strip())

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

"""RAW_DIR의 xlsx와 BASIS_DIRS(연도별)의 csv를 파일명 접두어(표준 지역명) 기준으로 짝짓는다.

    같은 지역의 실거래가 xlsx는 (기간 등에 따라)여러 개 있을 수 있으므로 개수 제한 없이 모두 모으고,
    해당 지역의 토지특성정보 csv는 연도별로 1개씩, BASIS_DIRS 우선순위 순서로 묶는다.
    반환: {지역: (xlsx_paths, [(연도, csv_path), ...])}
"""
def discover_region_files():
    xlsx_by_region = {}
    for path in glob.glob(os.path.join(RAW_DIR, "*.xlsx")):
        region = _match_region(os.path.basename(path))
        if region is None:
            print(f"건너뜀: 지역명을 알 수 없는 실거래가 파일 - {path}")
            continue
        xlsx_by_region.setdefault(region, []).append(path)

    # 연도별로 {지역: csv_path}
    csv_by_year = {}
    for year, basis_dir in BASIS_DIRS:
        csv_by_region = {}
        for path in glob.glob(os.path.join(basis_dir, "*.csv")):
            region = _match_region(os.path.basename(path))
            if region is None:
                print(f"건너뜀: 지역명을 알 수 없는 토지특성정보 파일({year}) - {path}")
                continue
            if region in csv_by_region:
                raise ValueError(f"'{region}' 지역의 {year}년 토지특성정보 파일이 2개 이상 발견됨: {csv_by_region[region]}, {path}")
            csv_by_region[region] = path
        csv_by_year[year] = csv_by_region

    region_files = {}
    for region, xlsx_paths in xlsx_by_region.items():
        csv_list = []
        for year, _ in BASIS_DIRS:
            csv_path = csv_by_year[year].get(region)
            if csv_path is None:
                print(f"경고: '{region}' 지역은 {year}년 토지특성정보 csv가 없어 해당 연도는 건너뜀")
                continue
            csv_list.append((year, csv_path))
        if not csv_list:
            print(f"건너뜀: '{region}' 지역은 토지특성정보 csv가 하나도 없음 (xlsx={xlsx_paths})")
            continue
        region_files[region] = (xlsx_paths, csv_list)
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

    # 매칭 결과 컬럼을 미리 만들어 두고, 연도별 매칭에서 성공한 행만 채운다.
    for col in RESULT_COLS:
        deals[col] = pd.Series([None] * len(deals), dtype=object)

    return deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count


"""아직 매칭되지 않은 행들로부터, csv 청크를 걸러낼 매칭 키를 만든다.

    키는 "법정동명|산여부|본번자리수|접두어" 문자열이며, 접두어 길이별로 집합을 나눠 둔다.
    (매칭 3단계의 접두어/자리수 조건과 동일하므로, 이 키에 걸리지 않는 csv 행은 어차피 매칭될 수 없다.)
    csv에 최신 법정동이 없을 때 과거 법정동으로 회귀 조회할 수 있도록, 과거 법정동명으로도 키를 함께 만든다.
    반환: (needed_dongs, {접두어길이: 키 집합})
"""
def _build_match_keys(deals_list, reverse_map: dict) -> tuple:
    needed_dongs = set()
    keys_by_prefix_len = {}
    ancestors_cache = {}
    for deals in deals_list:
        targets = deals[deals["PNU"].isna() & deals["번지"].notna()]
        for dong, raw in zip(targets["시군구"], targets["번지"]):
            if dong not in ancestors_cache:
                ancestors_cache[dong] = [dong] + _ancestors(dong, reverse_map)
            is_san = raw.startswith("산")
            prefix, digit_count = parse_beonji(raw)
            for name in ancestors_cache[dong]:
                needed_dongs.add(name)
                keys_by_prefix_len.setdefault(len(prefix), set()).add(f"{name}|{int(is_san)}|{digit_count}|{prefix}")
    return needed_dongs, keys_by_prefix_len


"""csv_path를 청크 단위로 한 번만 읽어, 매칭 키(동/산여부/본번 자리수/접두어)에 걸리는 행만 묶어 반환한다.

    1GB 안팎의 csv도 있으므로 메모리 절감을 위해
    - 청크마다 동 필터 + 매칭 키 필터를 걸어 실제로 매칭될 수 있는 행만 남기고,
    - 파생 컬럼 계산/형변환도 청크 안에서 끝내 필요 없는 컬럼은 바로 버리며,
    - 최종 후보의 반복 문자열 컬럼은 category로 바꾼다.
    반환: (candidates_by_group, present_dongs)
    candidates_by_group의 키는 (법정동명, 산여부) 튜플이다.
    (지목/용도지역은 "둘 중 하나 일치" 조건이라 groupby 키로 미리 나눌 수 없어 키에서 제외했다.)
    _match_deals에서 row마다 반복되는 등호 조건(법정동명/산여부)을
    미리 groupby로 나눠둠으로써, row 순회 시에는 훨씬 작아진 후보군에서
    번지(본번 자리수/접두어)와 면적만 비교하면 되도록 한다.
    present_dongs는 needed_dongs 중 이 csv에 실제로 존재하는 법정동명 집합이다.
    (매칭 키 필터 전에 모으므로, "csv에 해당 법정동이 있는지" 판단에 쓴다. 없으면 과거 법정동으로 회귀한다.)
"""
def _load_candidates(region: str, year: str, csv_path: str, needed_dongs: set, keys_by_prefix_len: dict) -> tuple:
    started = time.time()
    candidate_chunks = []
    present_dongs = set()
    for chunk in pd.read_csv(
        csv_path, encoding=CSV_ENCODING, dtype=str, usecols=CSV_USECOLS, chunksize=CSV_CHUNK_SIZE
    ):
        sub = chunk[chunk["법정동명"].isin(needed_dongs) & chunk["지번"].notna()]
        if not len(sub):
            continue
        present_dongs.update(sub["법정동명"].unique())
        # 지번은 "본번-부번" 형태이므로, 자리수/접두어 비교에 쓸 본번만 분리해둔다.
        bonbeon = sub["지번"].str.split("-", n=1).str[0]
        # row마다 다시 계산하지 않도록 본번 자리수를 미리 구해둔다.
        bonbeon_len = bonbeon.str.len()
        is_san = sub["대장구분명"] == "산"
        key_base = sub["법정동명"] + "|" + is_san.astype(int).astype(str) + "|" + bonbeon_len.astype(str) + "|"
        keep = pd.Series(False, index=sub.index)
        for prefix_len, keys in keys_by_prefix_len.items():
            keep |= (key_base + bonbeon.str[:prefix_len]).isin(keys)
        if not keep.any():
            continue
        sub = sub[keep]
        candidate_chunks.append(pd.DataFrame({
            "고유번호": sub["고유번호"],
            "법정동명": sub["법정동명"],
            "지번": sub["지번"],
            "본번": bonbeon[keep],
            "본번_자리수": bonbeon_len[keep].astype("int16"),
            "지목명": sub["지목명"],
            "용도지역명1": sub["용도지역명1"],
            "토지면적": sub["토지면적"].astype(float),
            "공시지가": sub["공시지가"].astype(float),
            "산여부": is_san[keep],
            "도로조건": sub["도로접면코드"].map(ROAD_SIDE_CODE_TO_CONDITION),
        }))
    elapsed = time.time() - started
    if not candidate_chunks:
        print(f"[{region}] {year}년 토지 특성정보 후보 건수: 0 ({elapsed:.0f}초)")
        return {}, present_dongs
    candidates = pd.concat(candidate_chunks, ignore_index=True)
    del candidate_chunks
    for col in CANDIDATE_CATEGORY_COLS:
        candidates[col] = candidates[col].astype("category")
    print(f"[{region}] {year}년 토지 특성정보 후보 건수(매칭 키 해당): {len(candidates)} ({elapsed:.0f}초)")

    group_keys = ["법정동명", "산여부"]
    return {key: df for key, df in candidates.groupby(group_keys, observed=True)}, present_dongs

""" 실질적인 매핑 로직
    아직 PNU가 없는 행만 candidates_by_group으로 매칭하고, 성공한 행에 결과 컬럼과 매칭연도를 채운다.
    (이미 앞 연도에서 매칭된 행은 건드리지 않는다.)
    실거래 법정동(최신)이 이 csv에 없으면 변경 이력을 거꾸로 따라가 csv에 있는 과거 법정동으로 조회한다.
    반환값: (이번 연도에서 새로 매칭된 건수, 그중 과거 법정동으로 회귀해 매칭된 건수)
"""
def _match_deals(deals, candidates_by_group: dict, present_dongs: set, reverse_map: dict, year: str) -> tuple:
    target = deals[deals["PNU"].isna()]

    # 실거래 법정동 -> 이 csv에서 조회할 법정동 (같은 법정동이 반복되므로 캐시)
    lookup_dong_cache = {}

    def resolve_lookup_dong(dong):
        if dong not in lookup_dong_cache:
            lookup = None
            for name in [dong] + _ancestors(dong, reverse_map):
                if name in present_dongs:
                    lookup = name
                    break
            lookup_dong_cache[dong] = lookup
        return lookup_dong_cache[dong]

    matched_idx = []
    matched_lookup_dong = []
    matched_pnu = []
    matched_beonji = []
    matched_area = []
    matched_official_price = []
    for row in target.itertuples(index=True):
        # 0. 번지가 비어있는(NaN) 행은 매칭 불가능하므로 바로 실패 처리한다.
        #    (예: 경상북도 데이터의 포항시 북구 흥해읍 곡강리 건처럼 원본에 번지 누락이 있는 경우)
        if pd.isna(row.번지):
            continue

        # 1. 마스킹된 번지("1***", "산4*" 등)를 "산 여부 / 앞부분 접두어 / 전체 자리수"로 분해한다.
        #    산으로 시작하면 "산"을 뗀 뒷부분이 실제 번지 문자열이다.
        is_san = row.번지.startswith("산")
        prefix, digit_count = parse_beonji(row.번지)

        # 2. [중요조건] 시군구(법정동명)/산여부가 모두 일치하는 후보군으로 범위를 좁힌다.
        #    (도로조건은 여기서 제외하고, 6번 단계에서 후보가 2건 이상일 때만 서브조건으로 적용한다.)
        #    (이 조건들은 row마다 값이 그대로 반복되므로, candidates_by_group을 만들 때 미리
        #    groupby 해뒀다. 매칭에 실패할 조합이면 딕셔너리에 아예 없으므로 후보군이 없다.)
        #    실거래 법정동이 csv에 없으면(법정동 개편 등 시점 차이) 과거 법정동으로 회귀해 조회한다.
        lookup_dong = resolve_lookup_dong(row.시군구)
        if lookup_dong is None:
            continue
        candidates = candidates_by_group.get((lookup_dong, is_san))
        if candidates is None:
            continue

        # 3. [중요조건] 토지 특성정보의 본번이 접두어로 시작하고, 마스킹 전 번지와 자리수가 같아야 한다.
        #    (자리수까지 맞춰야 "1**"가 "12" 같은 다른 자리수 번지와 잘못 매칭되지 않는다.)
        mask = (candidates["본번_자리수"] == digit_count) & candidates["본번"].str.startswith(prefix)

        # 4. [중요조건] 면적 조건: 지분구분이 없는 단독 거래는 계약면적과 오차범위 내로 일치해야 하고,
        #    지분 거래는 토지 전체 면적 중 일부만 거래된 것이므로 토지면적이 계약면적 이상이어야 한다.
        if pd.isna(row.지분구분):
            mask &= (candidates["토지면적"] > row.계약면적 - AREA_TOLERANCE) & (
                candidates["토지면적"] < row.계약면적 + AREA_TOLERANCE
            )
        else:
            mask &= candidates["토지면적"] >= row.계약면적

        # 5. [중요조건] 용도지역 또는 지목 중 하나라도 일치해야 한다.
        mask &= (candidates["지목명"] == row.지목) | (candidates["용도지역명1"] == row.용도지역)
        hits = candidates[mask]

        # 6. [서브조건] 중요조건을 만족하는 후보가 1건이면 바로 매칭 확정.
        #    2건 이상이면(중요조건만으로 특정 불가) 아래 서브조건을 순서대로 하나씩 적용해 다시 좁힌다.
        #    (1) 도로조건 일치 -> (2) 용도지역, 지목 모두 일치
        #    각 단계에서 1건이 되면 확정, 0건이면 실패, 2건 이상이면 다음 서브조건으로 넘어간다.
        #    (모든 서브조건을 거친 뒤에도 2건 이상이면 특정할 수 없으므로 매칭하지 않는다.)
        sub_filters = [
            lambda df: df["도로조건"] == row.도로조건,
            lambda df: (df["지목명"] == row.지목) & (df["용도지역명1"] == row.용도지역),
        ]
        for sub_filter in sub_filters:
            if len(hits) < 2: # 서브조건을 적용하기 전에 먼저 건수를 확인
                break
            hits = hits[sub_filter(hits)]

        # 7. 최종 후보가 정확히 1건일 때만 매칭 확정(PNU/실제 번지 확정)한다.
        if len(hits) == 1:
            matched_idx.append(row.Index)
            matched_lookup_dong.append(lookup_dong)
            matched_pnu.append(hits["고유번호"].iloc[0])
            beonji = hits["지번"].iloc[0]
            if is_san:
                beonji = f"산 {beonji}"
            matched_beonji.append(beonji)
            matched_area.append(hits["토지면적"].iloc[0])
            matched_official_price.append(hits["공시지가"].iloc[0])

    if matched_idx:
        deals.loc[matched_idx, "PNU"] = matched_pnu
        deals.loc[matched_idx, "확정번지"] = matched_beonji
        deals.loc[matched_idx, "토지면적"] = matched_area
        deals.loc[matched_idx, "공시지가"] = matched_official_price
        deals.loc[matched_idx, "매칭연도"] = year
        deals.loc[matched_idx, "조회법정동"] = matched_lookup_dong
    regressed_count = sum(
        1 for idx, lookup in zip(matched_idx, matched_lookup_dong) if deals.at[idx, "시군구"] != lookup
    )
    return len(matched_idx), regressed_count

"""모든 연도 매칭이 끝난 최종 deals로 통계 수치를 계산한다."""
def _compute_stats(region: str, xlsx_path: str, deals, total_count: int, actual_target_count: int, cancelled_count: int, small_area_count: int, share_count: int) -> dict:
    land_total_count = actual_target_count - share_count  # 특정토지 전체 거래 건수

    success_count = deals["PNU"].notna().sum()
    share_deal_mask = deals["지분구분"] == "지분"
    share_success_count = (share_deal_mask & deals["PNU"].notna()).sum()  # 지분거래 매칭 성공 건수
    land_success_count = (~share_deal_mask & deals["PNU"].notna()).sum()
    # 연도별 매칭 기여 건수 (BASIS_DIRS 우선순위 순서)
    year_counts = {year: int((deals["매칭연도"] == year).sum()) for year, _ in BASIS_DIRS}

    return {
        "region": region,
        "file": os.path.basename(xlsx_path),
        "total_count": total_count,
        "actual_target_count": actual_target_count,
        "cancelled": cancelled_count,
        "small_area": small_area_count,
        "share_count": share_count,
        "land_total_count": land_total_count,
        "success": success_count,
        "share_success": share_success_count,
        "land_success": land_success_count,
        "overall_success_rate": success_count / actual_target_count if actual_target_count else 0,
        "share_ratio": share_count / actual_target_count if actual_target_count else 0,
        "land_success_rate": land_success_count / land_total_count if land_total_count else 0,
        "share_success_rate": share_success_count / share_count if share_count else 0,
        "year_counts": year_counts,
    }

"""연도 라벨: 첫 우선순위 연도는 "n년 매칭", 이후 연도는 "n년 보완 매칭"."""
def _year_label(year: str) -> str:
    return f"{year}년 매칭" if year == BASIS_DIRS[0][0] else f"{year}년 보완 매칭"

"""매핑 결과 deals와 통계를 xlsx로 저장하고 출력 경로를 반환한다."""
def _write_output(deals, stats: dict) -> str:
    region = stats["region"]
    file_name = stats["file"]
    print(f"[{region}] {file_name} 매칭 성공 건수: {stats['success']} / {stats['actual_target_count']}")
    print(f"[{region}] {file_name} 지분거래 매칭 성공 건수: {stats['share_success']} / {stats['share_count']}")
    print(f"[{region}] {file_name} 특정토지 전체 거래 매칭 성공 건수: {stats['land_success']} / {stats['land_total_count']}")
    print(f"[{region}] {file_name} 연도별 매칭: " + ", ".join(f"{_year_label(y)} {n}건" for y, n in stats["year_counts"].items()))

    _xlsx_name, _xlsx_ext = os.path.splitext(file_name)
    output_path = os.path.join(OUTPUT_DIR, f"{_xlsx_name}_매핑{_xlsx_ext}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        deals.to_excel(writer, sheet_name="실거래가", index=False, startrow=DATA_START_ROW)
        sheet = writer.sheets["실거래가"]
        sheet["A1"] = f"total_info"
        sheet["B1"] = f"전체 건수: {stats['total_count']}"
        sheet["G1"] = f"실제 처리대상 건수: {stats['actual_target_count']}"
        sheet["L1"] = f"매칭 성공 건수: {stats['success']}"
        sheet["Q1"] = f"매칭 성공률: {stats['overall_success_rate']:.1%}"

        sheet["A2"] = f"제외 case"
        sheet["B2"] = f"거래취소로 제외된 건수: {stats['cancelled']}"
        sheet["G2"] = f"면적30이하로 제외된 건수 :  {stats['small_area']}"

        sheet["A3"] = f"지분거래"
        sheet["B3"] = f"지분 거래 건수: {stats['share_count']}"
        sheet["G3"] = f"지분거래 비율: {stats['share_ratio']:.1%}"
        sheet["L3"] = f"지분거래 매칭 성공 건수: {stats['share_success']}"
        sheet["Q3"] = f"지분거래 성공률: {stats['share_success_rate']:.1%}"

        sheet["A4"] = f"전체거래"
        sheet["B4"] = f"특정토지 전체거래 건수: {stats['land_total_count']}"
        sheet["L4"] = f"특정토지 전체거래 매칭 성공 건수: {stats['land_success']}"
        sheet["Q4"] = f"전체거래 성공률: {stats['land_success_rate']:.1%}"

        sheet["A5"] = f"연도별 매칭"
        for col, (year, count) in zip(["B5", "G5", "L5", "Q5"], stats["year_counts"].items()):
            sheet[col] = f"{_year_label(year)} 건수: {count}"

    print(f"[{region}] 완료: {output_path}")
    return output_path

"""같은 지역의 xlsx 여러 개가 연도별 csv를 공유하므로, csv는 연도마다 한 번만 읽고 xlsx만 반복 처리한다.

    1차: 우선순위 첫 연도(26) csv로 전체 행 매칭 -> 메모리 해제
    2차: 1차에서 매칭 안 된 행만 키로 삼아 다음 연도(25) csv를 읽어 재매칭 -> 메모리 해제
    모든 연도 매칭이 끝난 뒤 통계를 산출하고 결과를 쓴다.
    각 연도에서 실거래 법정동이 csv에 없으면 reverse_map(법정동 변경 이력)으로 과거 법정동에 회귀해 조회한다.
"""
def process_region_group(region: str, xlsx_paths, csv_list, reverse_map: dict) -> list:
    loaded = []
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
        loaded.append((xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count))

    if not loaded:
        return []

    # 연도 우선순위대로 csv를 하나씩만 메모리에 올려 매칭한다.
    # 매칭 키는 매 연도마다 "아직 매칭 안 된 행"으로 다시 만들기 때문에, 뒤 연도일수록 읽어 둘 후보가 줄어든다.
    for year, csv_path in csv_list:
        needed_dongs, keys_by_prefix_len = _build_match_keys([item[1] for item in loaded], reverse_map)
        if not needed_dongs:
            print(f"[{region}] 남은 미매칭 행이 없어 {year}년 매칭은 생략")
            break
        candidates_by_group, present_dongs = _load_candidates(region, year, csv_path, needed_dongs, keys_by_prefix_len)
        for xlsx_path, deals, *_ in loaded:
            try:
                newly_matched, regressed = _match_deals(deals, candidates_by_group, present_dongs, reverse_map, year)
                print(f"[{region}] {os.path.basename(xlsx_path)} {_year_label(year)}: {newly_matched}건 (그중 과거 법정동 회귀 매칭 {regressed}건)")
            except Exception as e:
                print(f"[{region}] {xlsx_path} {year}년 매칭 처리 실패: {e}")
        # 이 연도의 CSV 후보(candidates_by_group)는 다음 연도 처리 전에 해제한다.
        del candidates_by_group, present_dongs
        gc.collect()

    summaries = []
    # loaded에서 하나씩 꺼내 처리 즉시 참조를 없애, 처리 끝난 xlsx의 데이터가 다음 xlsx 처리 중에도
    # 계속 메모리에 남아있지 않도록 한다.
    while loaded:
        xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count = loaded.pop(0)
        try:
            stats = _compute_stats(region, xlsx_path, deals, total_count, actual_target_count, cancelled_count, small_area_count, share_count)
            stats["output"] = _write_output(deals, stats)
            summaries.append(stats)
        except Exception as e:
            print(f"[{region}] {xlsx_path} 결과 저장 실패로 건너뜀: {e}")
        finally:
            del deals
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
    for region, (xlsx_paths, csv_list) in region_files.items():
        print(f"\n=== [{region}] xlsx {len(xlsx_paths)}개, csv: {[year for year, _ in csv_list]} ===")
        try:
            summaries.extend(process_region_group(region, xlsx_paths, csv_list, reverse_map))
        except Exception as e:
            print(f"[{region}] 처리 중 오류로 건너뜀: {e}")

    print("\n=== 전국 처리 요약 ===")
    for s in summaries:
        print(f"[{s['region']}] 매칭 성공 {s['success']} / {s['actual_target_count']} -> {s['output']}")
    region_count = len(set(s["region"] for s in summaries))
    print(f"전체 처리 파일 수: {len(summaries)} (지역 수: {region_count}), 총 매칭 성공: {sum(s['success'] for s in summaries)}")

    # 원본 파일명별 핵심 통계(실제 처리대상 건수, 지분거래 비율, 전체거래 성공률, 지분거래 성공률, 연도별 매칭 건수)를
    # 개별 매핑 xlsx를 일일이 열어보지 않고 한눈에 비교할 수 있도록 별도 요약 xlsx로 모아둔다.
    summary_rows = [
        {
            "원본파일명": s["file"],
            "실제 처리대상 건수": s["actual_target_count"],
            "지분거래 비율": f"{s['share_ratio']:.1%}",
            "전체거래 성공률": f"{s['land_success_rate']:.1%}",
            "지분거래 성공률": f"{s['share_success_rate']:.1%}",
            **{f"{_year_label(year)} 건수": s["year_counts"].get(year, 0) for year, _ in BASIS_DIRS},
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
            **{f"{_year_label(year)} 건수": "" for year, _ in BASIS_DIRS},
        })

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        summary_path = os.path.join(OUTPUT_DIR, "통계_요약_토지매핑.xlsx")
        pd.DataFrame(summary_rows).to_excel(summary_path, sheet_name="통계", index=False)
        print(f"통계 요약 저장: {summary_path}")


if __name__ == "__main__":
    main()

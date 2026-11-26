# 건물 실거래가 데이터(확정번지가 채워진 매핑 결과)에, 토지특성정보 데이터를 대조해
# 해당 건물이 위치한 토지의 PNU(고유번호)/토지면적/공시지가를 채워 넣는 스크립트.
# 지역(서울/경기/... 등)별로 건물 실거래가 xlsx와 토지특성정보 csv가 따로 존재하므로,
# RAW_DIR, BASIS_DIRS(연도별)를 훑어 파일명 접두어(지역명)로 짝을 찾아 지역별로 순회 처리한다.
# (실거래-토지특성 짝짓기/연도별 보완 매칭 방식은 find_traded_land.py와 같다.)
import gc
import glob
import os
import re
import sys

import pandas as pd

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

RAW_DIR = r"D:\dev\data\processed\traded_house_processed\building" # 바로 아래에 있는 .xlsx 파일만 대상
# (연도, 경로) 목록. 리스트 순서가 곧 매칭 우선순위다. 각 경로 바로 아래에 있는 .csv 파일만 대상
# 1차로 첫 연도(26) 토지특성정보로 전체 실거래를 매칭하고, 이후 연도(25 -> 24 -> 23)는 앞 연도까지 매칭 안 된 행만 재매칭한다.
# 지역마다 한 번에 1개 연도 csv만 메모리에 올리고, 그 연도 매칭이 끝나면 해제한 뒤 다음 연도로 넘어간다.
BASIS_DIRS = [
    ("26", r"D:\dev\data\basis\land_characteristics\26"),
    ("25", r"D:\dev\data\basis\land_characteristics\25"),
    ("24", r"D:\dev\data\basis\land_characteristics\24"),
    ("23", r"D:\dev\data\basis\land_characteristics\23"),
]
BASIS_CSV_ENCODING = "cp949"
BASIS_USECOLS = ["고유번호", "법정동명", "지번", "대장구분명", "토지면적", "공시지가"]
CSV_CHUNK_SIZE = 200_000
OUTPUT_DIR = r"D:\dev\data\processed\traded_land_processed\land"

HEADER_ROW = 2  # find_house_buillding.py/find_general_building.py가 저장한 xlsx의 실제 표 헤더 행(0-index)
OUTPUT_SHEET_NAME = "실거래가"
DATA_START_ROW = 2  # 0-index. 1행에 통계 요약, 2행은 빈 줄, 3행부터 표 작성
SUMMARY_FILE_NAME = "건물토지매핑_요약통계.xlsx"  # 파일별 매칭 통계를 모아 OUTPUT_DIR에 저장한다.
SUMMARY_SHEET_NAME = "요약통계"

# 파일명이 이 이름들 중 하나로 시작하면 해당 지역으로 판별한다. (find_traded_land.py와 동일)
# 현재 파일명은 실거래 xlsx/토지특성정보 csv 모두 시도 정식명칭으로 시작한다.
# (예: "서울특별시_2019_단독다가구_실거래가_총괄매핑_매핑.xlsx", "서울특별시_토지특성정보_2025-01-23.csv" -> "서울특별시")
# 실거래가는 전남/광주가 "전남광주통합특별시" 1개 파일이고, 토지특성정보는 "전라남도"/"광주광역시"로 나뉘어 있다.
REGION_NAMES = [
    "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시", "대전광역시", "울산광역시", "세종특별자치시",
    "경기도", "강원특별자치도", "충청북도", "충청남도", "전북특별자치도", "전라남도", "전남광주통합특별시",
    "경상북도", "경상남도", "제주특별자치도",
]
# 실거래 지역 1개에 토지특성정보 csv 여러 개를 짝지어야 하는 예외 지역.
# 실거래 "전남광주통합특별시"는 토지특성정보 "전라남도", "광주광역시" csv 2개를 함께 읽어 모집단으로 쓴다.
# 여기 없는 지역은 실거래 지역명과 같은 이름의 csv 1개와 짝짓는다.
DEAL_REGION_TO_CSV_REGIONS = {
    "전남광주통합특별시": ["전라남도", "광주광역시"],
}
# 법정동명 비교 시 맨 앞 '시도' 단어를 떼고 비교하는 실거래 지역.
# 실거래는 "전남광주통합특별시 ...", 토지특성정보는 "전라남도 ..."/"광주광역시 ..."로 시도명이 달라
# 시도를 뺀 "시군구 읍면동 [리]"만으로 비교한다.
# 전북특별자치도/강원특별자치도도 시도 명칭이 바뀐 지역이라(전라북도 -> 전북특별자치도, 강원도 -> 강원특별자치도)
# 연도별 토지특성정보의 시도명이 실거래와 다를 수 있으므로 같은 방식으로 시도를 떼고 비교한다.
SIDO_STRIP_REGIONS = {"전남광주통합특별시", "전북특별자치도", "강원특별자치도"}


"""파일명이 REGION_NAMES 중 하나로 시작하면 그 지역명을 반환한다."""
def _match_region(filename: str):
    for region in REGION_NAMES:
        if filename.startswith(region):
            return region
    return None


"""법정동명/시군구를 비교용으로 정규화한다. 앞뒤 공백을 제거하고 연속 공백(탭/전각 공백 포함)을 한 칸으로 줄인다.

    실거래가 쪽 시군구 표기와 토지특성정보 쪽 법정동명 표기의 공백이 달라도 같은 키로 매칭되게 하기 위함이다.
    (기존에 import하던 find_overall_ledger_building.normalize_sigungu가 프로젝트에 없어, 같은 역할을 여기서 정의한다.)
    문자열이 아니면(NaN 등) 그대로 반환한다.
"""
def normalize_sigungu(name):
    if not isinstance(name, str):
        return name
    return re.sub(r"\s+", " ", name.strip())


"""법정동명을 비교 키로 만든다. 정규화한 뒤, strip_sido면 맨 앞 '시도' 단어를 뗀다.

    SIDO_STRIP_REGIONS 지역(전남광주통합특별시, 전북특별자치도, 강원특별자치도)에서만 strip_sido=True로 쓴다.
    예: "전남광주통합특별시 광산구 도산동" -> "광산구 도산동", "광주광역시 광산구 도산동" -> "광산구 도산동"
    문자열이 아니면(NaN 등) 그대로 반환한다.
"""
def _dong_compare_key(name, strip_sido: bool):
    name = normalize_sigungu(name)
    if not strip_sido or not isinstance(name, str):
        return name
    parts = name.split(" ", 1)
    return parts[1] if len(parts) == 2 else name


"""RAW_DIR의 xlsx와 BASIS_DIRS(연도별)의 csv를 파일명 접두어(지역명) 기준으로 짝짓는다.

    같은 지역의 건물 실거래가(확정번지 매핑 결과) xlsx는 여러 개 있을 수 있으므로 개수 제한 없이
    모두 모으고, 해당 지역의 토지특성정보 csv를 연도별로 공통 매칭해 각 xlsx마다 개별 처리·출력한다.
    단, DEAL_REGION_TO_CSV_REGIONS에 있는 지역(전남광주통합특별시)은 연도마다 csv 여러 개를 모두 짝짓는다.
    어떤 연도에 그 지역 csv가 (하나라도) 없으면 그 연도만 건너뛰고, 쓸 수 있는 연도가 하나도 없으면 지역을 건너뛴다.
    반환: {실거래 지역명: (xlsx 경로 리스트, [(연도, csv 경로 리스트), ...])}  (연도는 BASIS_DIRS 우선순위 순서)
"""
def discover_region_files():
    xlsx_by_region = {}
    for path in glob.glob(os.path.join(RAW_DIR, "*.xlsx")):
        region = _match_region(os.path.basename(path))
        if region is None:
            print(f"건너뜀: 지역명을 알 수 없는 실거래가 파일 - {path}")
            continue
        xlsx_by_region.setdefault(region, []).append(path)

    # {연도: {csv 지역명: csv 경로}}
    csv_by_year = {}
    for year, basis_dir in BASIS_DIRS:
        csv_by_region = {}
        for path in glob.glob(os.path.join(basis_dir, "*.csv")):
            region = _match_region(os.path.basename(path))
            if region is None:
                print(f"건너뜀: 지역명을 알 수 없는 토지특성정보 파일 - {path}")
                continue
            if region in csv_by_region:
                raise ValueError(f"{year}년 '{region}' 지역의 토지특성정보 파일이 2개 이상 발견됨: {csv_by_region[region]}, {path}")
            csv_by_region[region] = path
        csv_by_year[year] = csv_by_region

    # 다른 실거래 지역에 묶여 들어가는 csv 지역(전라남도, 광주광역시)은 단독 지역으로 짝을 찾지 않는다.
    merged_csv_regions = {r for csv_regions in DEAL_REGION_TO_CSV_REGIONS.values() for r in csv_regions}

    all_csv_regions = set().union(*(m.keys() for m in csv_by_year.values()))
    region_files = {}
    for region in xlsx_by_region.keys() | (all_csv_regions - merged_csv_regions):
        xlsx_paths = xlsx_by_region.get(region)
        if not xlsx_paths:
            print(f"건너뜀: '{region}' 지역은 실거래가 xlsx가 없음")
            continue
        csv_regions = DEAL_REGION_TO_CSV_REGIONS.get(region, [region])
        csv_list = []
        for year, _ in BASIS_DIRS:
            csv_by_region = csv_by_year[year]
            csv_paths = [csv_by_region[r] for r in csv_regions if r in csv_by_region]
            # 묶어야 하는 csv 중 하나라도 없으면 그 연도 모집단이 불완전하므로 그 연도는 쓰지 않는다.
            if len(csv_paths) != len(csv_regions):
                print(f"경고: '{region}' 지역은 {year}년 csv 짝이 없어 {year}년 매칭은 건너뜀 (csv={csv_paths}, 필요 csv 지역={csv_regions})")
                continue
            csv_list.append((year, csv_paths))
        if not csv_list:
            print(f"건너뜀: '{region}' 지역은 어느 연도에도 csv 짝이 없음 (xlsx={xlsx_paths}, 필요 csv 지역={csv_regions})")
            continue
        region_files[region] = (xlsx_paths, csv_list)
    return region_files


"""건물 실거래가(확정번지 매핑 결과) xlsx를 읽는다.

    확정번지 컬럼이 없으면(아직 건물-건축물대장 매핑을 거치지 않은 파일) None을 반환해 건너뛴다.
    확정번지가 없는 행(매칭 실패한 건물 거래)은 처리 대상에서 제외하되, 결과물에서 삭제하지는 않는다.
"""
def _load_deals(xlsx_path: str):
    # 확정번지는 셀에 문자열로 저장돼 있어도, 값이 전부 숫자 형태(예: "1556", '-' 없음)면 pandas가 float로 추론해
    # "1556.0"이 되므로 문자열로 고정해서 읽는다. (확정번지 컬럼이 없는 파일이어도 dtype 지정은 에러 없이 무시된다)
    deals = pd.read_excel(xlsx_path, header=HEADER_ROW, dtype={"확정번지": str})
    if "확정번지" not in deals.columns:
        return None, 0

    target_count = int(deals["확정번지"].notna().sum())
    return deals, target_count


"""csv_paths(지역 토지특성정보, 한 연도분)를 청크 단위로 한 번만 읽어, needed_dongs에 해당하는 행만 법정동명별로 묶어 반환한다.

    법정동명은 _dong_compare_key(normalize_sigungu + strip_sido면 시도 제거)로 비교 키로 바꿔둔다. 실거래가 쪽 시군구 표기와 토지특성정보
    쪽 법정동명 표기의 공백이 달라도 같은 키로 매칭되게 하기 위함이다(find_overall_ledger_building.py에서
    시군구/대지_위치 매칭에 쓰는 것과 동일한 정규화). 법정동명이 비어있는(NaN) 행은 normalize_sigungu가
    문자열만 받으므로 정규화 전에 미리 제외한다(어차피 그런 행은 실거래 시군구와 매칭될 수 없어, 기존에도
    결과에 포함되지 않던 행이다).
    반환 딕셔너리의 키는 정규화된 법정동명(비교 키)이다. 실거래 행마다 반복되는 동 단위 필터링을 미리 끝내두어,
    매칭 시에는 훨씬 작아진 후보군에서 지번/대장구분명만 비교하면 되도록 한다.
    csv 전체를 올리지 않고, 아직 매칭 안 된 실거래에 등장하는 법정동(needed_dongs)의 행만 남긴다.
    csv_paths가 여러 개(전남광주통합특별시 -> 전라남도 + 광주광역시)면 모두 읽어 하나의 후보군(모집단)으로 합친다.
"""
def _load_land_candidates(region: str, csv_paths: list, needed_dongs: set, strip_sido: bool, year: str) -> dict:
    candidate_chunks = []
    for csv_path in csv_paths:
        for chunk in pd.read_csv(
            csv_path, encoding=BASIS_CSV_ENCODING, usecols=BASIS_USECOLS,
            dtype={"고유번호": str, "지번": str}, chunksize=CSV_CHUNK_SIZE,
        ):
            chunk = chunk[chunk["법정동명"].notna()]
            # 청크 안 고유 법정동명은 수백 개 수준이므로, 고유값만 정규화해 map으로 펼친다.
            normalized = {name: _dong_compare_key(name, strip_sido) for name in chunk["법정동명"].unique()}
            chunk = chunk.assign(법정동명=chunk["법정동명"].map(normalized))
            chunk = chunk[chunk["법정동명"].isin(needed_dongs)]
            if len(chunk):
                candidate_chunks.append(chunk)

    # 필요한 법정동이 이 연도 csv에 하나도 없으면 후보가 없다. (빈 리스트를 concat하면 오류)
    if not candidate_chunks:
        print(f"[{region}] {year}년 토지 특성정보 후보 건수(대상 법정동 내): 0")
        return {}

    candidates = pd.concat(candidate_chunks, ignore_index=True)
    candidates = candidates.dropna(subset=["지번"]).copy()
    candidates["토지면적"] = candidates["토지면적"].astype(float)
    candidates["공시지가"] = candidates["공시지가"].astype(float)
    print(f"[{region}] {year}년 토지 특성정보 후보 건수(대상 법정동 내): {len(candidates)}")

    return {dong: df for dong, df in candidates.groupby("법정동명")}


"""실거래 한 건(딕셔너리 형태의 행)에 대해 시군구+확정번지로 토지 후보를 찾는다.

    시군구는 토지특성정보 법정동명과 같은 비교 키(_dong_compare_key)로 바꿔 조회한다.
    확정번지가 '산'으로 시작하면 '산'을 뗀 나머지를 지번과 비교하고, 대장구분명이 '산'인 후보만 허용한다.
    후보가 정확히 1개일 때만 확정하며, (고유번호, 토지면적, 공시지가)를 반환한다. 실패하면 (None, None, None).
"""
def _match_deal(deal: dict, candidates_by_dong: dict, strip_sido: bool):
    candidates = candidates_by_dong.get(_dong_compare_key(deal["시군구"], strip_sido))
    if candidates is None:
        return None, None, None

    beonji = deal["확정번지"]
    is_san = beonji.startswith("산")
    beonji_key = beonji[1:].lstrip() if is_san else beonji

    mask = candidates["지번"] == beonji_key
    if is_san:
        mask &= candidates["대장구분명"] == "산"
    hits = candidates[mask]

    if len(hits) != 1:
        return None, None, None

    matched = hits.iloc[0]
    return matched["고유번호"], matched["토지면적"], matched["공시지가"]


"""연도별 보완 매칭: 확정번지가 있고 아직 PNU가 없는(앞 연도까지 매칭 안 된) 행만 이 연도 후보로 매칭하고,
    성공한 행에만 PNU/토지면적/공시지가/매칭연도(year)를 채운다. (이미 매칭된 행은 건드리지 않는다.)
    반환: 이번 연도 신규 매칭 건수
"""
def _match_deals(deals, candidates_by_dong: dict, strip_sido: bool, year: str) -> int:
    unmatched = deals[deals["확정번지"].notna() & deals["PNU"].isna()]

    success_idx, matched_pnu, matched_area, matched_price = [], [], [], []
    for idx, deal in zip(unmatched.index, unmatched.to_dict("records")):
        pnu, area, price = _match_deal(deal, candidates_by_dong, strip_sido)
        if pnu is None:
            continue
        success_idx.append(idx)
        matched_pnu.append(pnu)
        matched_area.append(area)
        matched_price.append(price)

    deals.loc[success_idx, "PNU"] = matched_pnu
    deals.loc[success_idx, "토지면적"] = matched_area
    deals.loc[success_idx, "공시지가"] = matched_price
    deals.loc[success_idx, "매칭연도"] = year
    return len(success_idx)


"""연도 라벨: 첫 우선순위 연도는 "n년 매칭", 이후 연도는 "n년 보완 매칭"."""
def _year_label(year: str) -> str:
    return f"{year}년 매칭" if year == BASIS_DIRS[0][0] else f"{year}년 보완 매칭"


"""모든 연도 매칭이 끝난 deals의 통계를 계산해 결과 xlsx로 저장한다. 로그에는 어느 지역 처리 중인지 알 수 있도록 [region]을 붙인다.
요약 통계 파일에 쓰기 위해 (PNU가 채워진(매칭 성공) 건수, 연도별 매칭 건수)를 반환한다."""
def _write_output(region: str, xlsx_path: str, deals, target_count: int):
    success_count = deals["PNU"].notna().sum()
    success_rate = success_count / target_count if target_count else 0
    # 연도별 매칭 기여 건수 (BASIS_DIRS 우선순위 순서)
    year_counts = {year: int((deals["매칭연도"] == year).sum()) for year, _ in BASIS_DIRS}
    print(f"[{region}] {os.path.basename(xlsx_path)} 처리대상 개수: {target_count}")
    print(f"[{region}] {os.path.basename(xlsx_path)} 매칭 성공 건수: {success_count}")
    print(f"[{region}] {os.path.basename(xlsx_path)} 매칭 성공률: {success_rate:.1%}")
    print(f"[{region}] {os.path.basename(xlsx_path)} 연도별 매칭: " + ", ".join(f"{_year_label(y)} {n}건" for y, n in year_counts.items()))

    _xlsx_name, _xlsx_ext = os.path.splitext(os.path.basename(xlsx_path))
    output_path = os.path.join(OUTPUT_DIR, f"{_xlsx_name}_토지매핑{_xlsx_ext}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        deals.to_excel(writer, sheet_name=OUTPUT_SHEET_NAME, index=False, startrow=DATA_START_ROW)
        sheet = writer.sheets[OUTPUT_SHEET_NAME]
        sheet["A1"] = "total_info"
        sheet["B1"] = f"처리대상 개수: {target_count}"
        sheet["G1"] = f"매칭 성공 건수: {success_count}"
        sheet["L1"] = f"매칭 성공률: {success_rate:.1%}"
        # 연도별 매칭 건수도 1행 오른쪽에 이어서 쓴다.
        for col, (year, count) in zip(["Q1", "V1", "AA1", "AF1"], year_counts.items()):
            sheet[col] = f"{_year_label(year)} 건수: {count}"

    print(f"[{region}] 완료: {output_path}")
    return int(success_count), year_counts


"""파일별 토지 매칭 통계를 하나의 요약 xlsx로 저장한다.

    토지 매핑 성공률은 PNU 개수 / 확정번지 개수, 전체 매핑 성공률은 PNU 개수 / 전체 개수다.
    성공률 뒤에는 연도별 매칭 건수 컬럼(BASIS_DIRS 우선순위 순서)이 붙는다.
    마지막에 합계 행(개수 합계)과 성공률 평균 행(파일별 성공률의 단순 평균)을 덧붙인다.
"""
def _write_summary(summary_rows: list):
    year_columns = [f"{_year_label(year)} 건수" for year, _ in BASIS_DIRS]
    summary = pd.DataFrame(
        summary_rows,
        columns=["파일명", "전체 개수", "확정번지 개수", "PNU 개수", "토지 매핑 성공률", "전체 매핑 성공률"] + year_columns,
    )

    total_count = summary["전체 개수"].sum()
    total_beonji = summary["확정번지 개수"].sum()
    total_pnu = summary["PNU 개수"].sum()
    total_year_counts = [summary[col].sum() for col in year_columns]
    mean_land_rate = summary["토지 매핑 성공률"].mean() if not summary.empty else 0
    mean_overall_rate = summary["전체 매핑 성공률"].mean() if not summary.empty else 0
    summary.loc[len(summary)] = ["합계", total_count, total_beonji, total_pnu, None, None] + total_year_counts
    summary.loc[len(summary)] = ["성공률 평균", None, None, None, mean_land_rate, mean_overall_rate] + [None] * len(year_columns)

    output_path = os.path.join(OUTPUT_DIR, SUMMARY_FILE_NAME)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name=SUMMARY_SHEET_NAME, index=False)
        sheet = writer.sheets[SUMMARY_SHEET_NAME]
        # 성공률 컬럼(E, F열)을 퍼센트 형식으로 표시한다. 1행은 헤더.
        for row in range(2, len(summary) + 2):
            sheet.cell(row=row, column=5).number_format = "0.0%"
            sheet.cell(row=row, column=6).number_format = "0.0%"

    print(
        f"요약 통계 완료: {output_path} "
        f"(토지 매핑 성공률 평균: {mean_land_rate:.1%}, 전체 매핑 성공률 평균: {mean_overall_rate:.1%})"
    )


"""같은 지역의 xlsx 여러 개가 연도별 csv를 공유하므로, csv는 연도마다 한 번만 읽고 xlsx만 반복 처리한다.
    (전남광주통합특별시는 연도마다 전라남도/광주광역시 csv 2개가 들어오며, 둘을 합친 후보군을 공유한다.)

    0: 이 지역 xlsx를 모두 로드한다. (연도별로 미매칭 행을 이어서 채워야 하므로 지역 처리 동안 유지)
    1차: 우선순위 첫 연도(26) csv로 확정번지가 있는 전체 행 매칭 -> 메모리 해제
    2차~: 앞 연도까지 매칭 안 된 행만으로 needed_dongs를 다시 만들어 다음 연도(25 -> 24 -> 23) csv를 읽어 재매칭 -> 메모리 해제
    모든 연도 매칭이 끝난 뒤 xlsx를 하나씩 저장하고 즉시 해제한다.
    파일별 통계(파일명, 전체 개수, 확정번지 개수, PNU 개수, 성공률 2종, 연도별 매칭 건수)는 summary_rows에 추가한다.
"""
def process_region_group(region: str, xlsx_paths, csv_list: list, summary_rows: list):
    # 전남광주통합특별시/전북/강원은 실거래와 토지특성 시도명이 다를 수 있어 시도를 떼고 법정동명을 비교한다.
    strip_sido = region in SIDO_STRIP_REGIONS
    loaded = []
    for xlsx_path in xlsx_paths:
        print(f"[{region}] {os.path.basename(xlsx_path)} 로딩 중...")
        deals, target_count = _load_deals(xlsx_path)
        if deals is None:
            print(f"[{region}] 건너뜀: 확정번지 컬럼이 없는 파일 - {xlsx_path}")
            continue
        # 매칭 결과 컬럼. 연도별 매칭에서 성공한 행만 채워 나가므로 빈 값으로 시작한다.
        for col in ["PNU", "토지면적", "공시지가", "매칭연도"]:
            deals[col] = None
        loaded.append((xlsx_path, deals, target_count))

    if not loaded:
        return

    # 연도 우선순위대로 csv를 한 연도씩만 메모리에 올려 매칭한다.
    for year, csv_paths in csv_list:
        # 아직 매칭 안 된(확정번지 있음 + PNU 없음) 행의 시군구만 후보로 남기면 되므로, 이 지역 xlsx 전체에서
        # 필요한 동을 먼저 모은다. 뒤 연도일수록 미매칭 행이 줄어 읽어 둘 후보도 줄어든다.
        needed_dongs = set()
        for _, deals, _ in loaded:
            pending = deals["확정번지"].notna() & deals["PNU"].isna()
            for raw_dong in deals.loc[pending, "시군구"].dropna().unique():
                needed_dongs.add(_dong_compare_key(raw_dong, strip_sido))
        if not needed_dongs:
            print(f"[{region}] 남은 미매칭 행이 없어 {year}년 이후 매칭은 생략")
            break

        print(f"[{region}] {year}년 {[os.path.basename(p) for p in csv_paths]} 로딩 중...")
        candidates_by_dong = _load_land_candidates(region, csv_paths, needed_dongs, strip_sido, year)
        for xlsx_path, deals, _ in loaded:
            newly_matched = _match_deals(deals, candidates_by_dong, strip_sido, year)
            print(f"[{region}] {os.path.basename(xlsx_path)} {_year_label(year)}: {newly_matched}건")
        # 이 연도의 CSV 후보는 다음 연도 처리 전에 해제한다.
        del candidates_by_dong
        gc.collect()

    # loaded에서 하나씩 꺼내 저장 즉시 참조를 없애, 처리 끝난 xlsx가 계속 메모리에 남아있지 않도록 한다.
    while loaded:
        xlsx_path, deals, target_count = loaded.pop(0)
        total_count = len(deals)
        success_count, year_counts = _write_output(region, xlsx_path, deals, target_count)
        land_rate = success_count / target_count if target_count else 0
        overall_rate = success_count / total_count if total_count else 0
        summary_rows.append(
            [os.path.basename(xlsx_path), total_count, target_count, success_count, land_rate, overall_rate]
            + [year_counts[year] for year, _ in BASIS_DIRS]
        )
        del deals
        gc.collect()


def main():
    region_files = discover_region_files()
    if not region_files:
        print("처리할 지역 파일 쌍을 찾지 못했습니다.")
        return

    summary_rows = []
    for region, (xlsx_paths, csv_list) in region_files.items():
        print(f"\n=== [{region}] xlsx {len(xlsx_paths)}개, csv 연도: {[year for year, _ in csv_list]} ===")
        process_region_group(region, xlsx_paths, csv_list, summary_rows)

    _write_summary(summary_rows)


if __name__ == "__main__":
    main()

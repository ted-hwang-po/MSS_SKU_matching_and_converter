import pandas as pd
from rapidfuzz import fuzz, process

from .utils import normalize_for_matching, normalize_strict, split_product_option


def _find_column(df: pd.DataFrame, candidates: list[str], label: str) -> str:
    """여러 후보 컬럼명 중 실제 존재하는 것을 찾음"""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"'{label}' 컬럼을 찾을 수 없습니다. 후보: {candidates}, 실제 컬럼: {list(df.columns)}")


def match_barcode_to_uid(order_df: pd.DataFrame, matching_df: pd.DataFrame) -> pd.DataFrame:
    # 파일2 컬럼명 유연 매핑
    barcode_col = _find_column(matching_df, ['바코드'], '바코드')
    uid_col = _find_column(matching_df, ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], '상품코드(UID)')
    style_col = _find_column(matching_df, ['스타일번호'], '스타일번호')

    match_subset = matching_df[[barcode_col, uid_col, style_col]].copy()
    match_subset = match_subset.rename(columns={
        barcode_col: '바코드',
        uid_col: '상품코드',
        style_col: '스타일번호',
    })
    match_subset['바코드'] = match_subset['바코드'].astype(str).str.strip()

    order_df = order_df.copy()
    order_df['88코드'] = order_df['88코드'].astype(str).str.strip()

    merged = order_df.merge(
        match_subset,
        left_on='88코드',
        right_on='바코드',
        how='left',
    )

    unmatched = merged[merged['상품코드'].isna()]
    if len(unmatched) > 0:
        unmatched_barcodes = unmatched['88코드'].tolist()
    else:
        unmatched_barcodes = []

    return merged, unmatched_barcodes


def detect_option_products(merged_df: pd.DataFrame, matching_df: pd.DataFrame) -> dict:
    """UID별 바코드 개수로 옵션 유무 판별. 파일2 전체를 기준으로 판단."""
    barcode_col = _find_column(matching_df, ['바코드'], '바코드')
    uid_col = _find_column(matching_df, ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], '상품코드(UID)')

    match_subset = matching_df[[barcode_col, uid_col]].copy()
    match_subset = match_subset.rename(columns={barcode_col: '바코드', uid_col: '상품코드'})
    match_subset['바코드'] = match_subset['바코드'].astype(str).str.strip()
    match_subset['상품코드'] = match_subset['상품코드'].astype(str).str.strip()

    barcode_count = match_subset.groupby('상품코드')['바코드'].nunique()

    # True = 옵션 상품 (바코드 2개 이상)
    has_option = {uid: (count >= 2) for uid, count in barcode_count.items()}
    return has_option


def match_option_info(
    merged_df: pd.DataFrame,
    option_df: pd.DataFrame,
    has_option: dict,
) -> pd.DataFrame:
    """옵션 매핑: 사이즈유형, 옵션 슬롯 결정"""

    # 파일3의 사이즈유형명 목록 준비
    size_type_names = []
    for _, row in option_df.iterrows():
        size_type_names.append({
            'raw': str(row.get('사이즈유형명', '')),
            'normalized': normalize_for_matching(str(row.get('사이즈유형명', ''))),
            'strict': normalize_strict(str(row.get('사이즈유형명', ''))),
            '사이즈유형': str(row.get('사이즈유형', '')),
            'row': row,
        })

    # Size 컬럼 목록 추출
    size_cols = [c for c in option_df.columns if c.lower().startswith('size') and c[4:].isdigit()]
    size_cols = sorted(size_cols, key=lambda x: int(x[4:].lstrip('0') or '0'))

    results = []
    warnings = []

    for idx, row in merged_df.iterrows():
        uid = str(row.get('상품코드', ''))
        product_name = str(row.get('상품명', ''))
        barcode = str(row.get('88코드', ''))

        is_option = has_option.get(uid, False)

        if not is_option:
            results.append({
                'idx': idx,
                '사이즈유형': 'MF3',
                '옵션슬롯': 1,  # Size01
                '옵션값': 'FREE',
            })
            continue

        # 옵션 상품 처리
        base_name, option_name = split_product_option(product_name)

        # Step 1: 사이즈유형 매칭
        matched_type = _match_size_type(base_name, size_type_names)

        if matched_type is None:
            warnings.append({
                '88코드': barcode,
                '상품명': product_name,
                '원인': '사이즈유형 매칭 실패',
            })
            results.append({
                'idx': idx,
                '사이즈유형': None,
                '옵션슬롯': None,
                '옵션값': option_name,
            })
            continue

        # Step 2: 옵션 슬롯 매칭
        slot_num, slot_value = _match_option_slot(
            option_name, matched_type['row'], size_cols
        )

        if slot_num is None:
            warnings.append({
                '88코드': barcode,
                '상품명': product_name,
                '원인': f"옵션 슬롯 매칭 실패 (사이즈유형: {matched_type['사이즈유형']})",
            })

        results.append({
            'idx': idx,
            '사이즈유형': matched_type['사이즈유형'],
            '옵션슬롯': slot_num,
            '옵션값': slot_value or option_name,
        })

    result_df = pd.DataFrame(results).set_index('idx')

    for col in ['사이즈유형', '옵션슬롯', '옵션값']:
        merged_df[col] = result_df[col]

    return merged_df, warnings


def _match_size_type(base_name: str, size_type_names: list[dict]) -> dict | None:
    if not base_name:
        return None

    base_strict = normalize_strict(base_name)

    # Level 1: 정규화 후 정확 매칭
    for entry in size_type_names:
        if base_strict == entry['strict']:
            return entry

    # Level 2: 포함관계 매칭
    for entry in size_type_names:
        type_strict = entry['strict']
        if base_strict and type_strict:
            if base_strict in type_strict or type_strict in base_strict:
                return entry

    # Level 3: RapidFuzz 유사도 매칭
    base_norm = normalize_for_matching(base_name)
    choices = {i: entry['normalized'] for i, entry in enumerate(size_type_names)}

    if not choices:
        return None

    result = process.extractOne(
        query=base_norm,
        choices=choices,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=80,
    )

    if result:
        _, score, matched_idx = result
        return size_type_names[matched_idx]

    return None


def _match_option_slot(
    option_name: str | None,
    option_row: pd.Series,
    size_cols: list[str],
) -> tuple[int | None, str | None]:
    if not option_name:
        return (None, None)

    option_strict = normalize_strict(option_name)

    # Level 1: 정규화 후 정확 매칭
    for i, col in enumerate(size_cols, start=1):
        val = str(option_row.get(col, ''))
        if not val or val == 'nan':
            continue
        if normalize_strict(val) == option_strict:
            return (i, val)

    # Level 2: RapidFuzz
    option_norm = normalize_for_matching(option_name)
    slot_values = {}
    for i, col in enumerate(size_cols, start=1):
        val = str(option_row.get(col, ''))
        if val and val != 'nan' and val.strip():
            slot_values[i] = normalize_for_matching(val)

    if not slot_values:
        return (None, None)

    result = process.extractOne(
        query=option_norm,
        choices=slot_values,
        scorer=fuzz.ratio,
        score_cutoff=75,
    )

    if result:
        _, score, matched_slot = result
        col = size_cols[matched_slot - 1]
        return (matched_slot, str(option_row.get(col, '')))

    return (None, None)

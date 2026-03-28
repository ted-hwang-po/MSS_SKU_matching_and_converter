import math

import pandas as pd
from rapidfuzz import fuzz, process

from .utils import normalize_for_matching, normalize_strict, split_product_option


def _find_column(df: pd.DataFrame, candidates: list[str], label: str) -> str:
    """여러 후보 컬럼명 중 실제 존재하는 것을 찾음"""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"'{label}' 컬럼을 찾을 수 없습니다. 후보: {candidates}, 실제 컬럼: {list(df.columns)}")


def _find_column_optional(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _safe_str(val) -> str:
    """숫자(float/int)를 안전하게 문자열로 변환."""
    if val is None:
        return ''
    if isinstance(val, float):
        if math.isnan(val):
            return ''
        if val == int(val):
            return str(int(val))
    return str(val).strip()


def match_barcode_to_uid(order_df: pd.DataFrame, matching_df: pd.DataFrame) -> pd.DataFrame:
    barcode_col = _find_column(matching_df, ['바코드'], '바코드')
    uid_col = _find_column(matching_df, ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], '상품코드(UID)')
    style_col = _find_column(matching_df, ['스타일번호'], '스타일번호')
    option_col = _find_column_optional(matching_df, ['옵션명'])
    product_col = _find_column_optional(matching_df, ['상품명'])

    # 가져올 컬럼 목록
    cols = [barcode_col, uid_col, style_col]
    rename_map = {barcode_col: '바코드', uid_col: '상품코드', style_col: '스타일번호'}
    if option_col:
        cols.append(option_col)
        rename_map[option_col] = '_파일2_옵션명'
    if product_col:
        cols.append(product_col)
        rename_map[product_col] = '_파일2_상품명'

    match_subset = matching_df[cols].copy()
    match_subset = match_subset.rename(columns=rename_map)
    match_subset['바코드'] = match_subset['바코드'].apply(_safe_str)

    order_df = order_df.copy()
    order_df['88코드'] = order_df['88코드'].apply(_safe_str)

    merged = order_df.merge(
        match_subset,
        left_on='88코드',
        right_on='바코드',
        how='left',
    )

    unmatched = merged[merged['상품코드'].isna()]
    unmatched_barcodes = unmatched['88코드'].tolist() if len(unmatched) > 0 else []

    return merged, unmatched_barcodes


def detect_option_products(merged_df: pd.DataFrame, matching_df: pd.DataFrame) -> dict:
    """UID별 바코드 개수로 옵션 유무 판별."""
    barcode_col = _find_column(matching_df, ['바코드'], '바코드')
    uid_col = _find_column(matching_df, ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], '상품코드(UID)')

    match_subset = matching_df[[barcode_col, uid_col]].copy()
    match_subset = match_subset.rename(columns={barcode_col: '바코드', uid_col: '상품코드'})
    match_subset['바코드'] = match_subset['바코드'].apply(_safe_str)
    match_subset['상품코드'] = match_subset['상품코드'].apply(_safe_str)

    barcode_count = match_subset.groupby('상품코드')['바코드'].nunique()
    has_option = {uid: (count >= 2) for uid, count in barcode_count.items()}
    return has_option


def _build_uid_option_order(matching_df: pd.DataFrame) -> dict:
    """파일2에서 UID별 옵션 순서 맵 생성.
    {uid_str: {barcode_str: (slot_num, option_name)}}
    """
    barcode_col = _find_column(matching_df, ['바코드'], '바코드')
    uid_col = _find_column(matching_df, ['상품코드', '1P 상품코드', '기존 위탁 상품코드 UID'], '상품코드(UID)')
    option_col = _find_column_optional(matching_df, ['옵션명'])

    uid_option_map = {}

    for _, row in matching_df.iterrows():
        uid = _safe_str(row[uid_col])
        barcode = _safe_str(row[barcode_col])
        option_name = str(row[option_col]).strip() if option_col and pd.notna(row.get(option_col)) else ''

        if not uid or not barcode:
            continue

        if uid not in uid_option_map:
            uid_option_map[uid] = {}

        if barcode not in uid_option_map[uid]:
            slot_num = len(uid_option_map[uid]) + 1
            uid_option_map[uid][barcode] = (slot_num, option_name)

    return uid_option_map


def match_option_info(
    merged_df: pd.DataFrame,
    option_df: pd.DataFrame,
    has_option: dict,
    matching_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """옵션 매핑: 사이즈유형, 옵션 슬롯 결정

    핵심 전략:
    1. 파일2의 옵션명/상품명을 우선 활용 (# 구분자 의존 제거)
    2. 사이즈유형은 파일2 상품명 → 파일3 사이즈유형명 매칭
    3. 슬롯 번호는 파일2 내 UID별 순서 기반
    """

    # 파일3 사이즈유형 정보
    size_type_entries = []
    for _, row in option_df.iterrows():
        size_type_entries.append({
            'raw': str(row.get('사이즈유형명', '')),
            'normalized': normalize_for_matching(str(row.get('사이즈유형명', ''))),
            'strict': normalize_strict(str(row.get('사이즈유형명', ''))),
            '사이즈유형': str(row.get('사이즈유형', '')),
            'row': row,
        })

    size_cols = [c for c in option_df.columns if c.lower().startswith('size') and c[4:].isdigit()]
    size_cols = sorted(size_cols, key=lambda x: int(x[4:].lstrip('0') or '0'))

    # 파일2 기반 UID별 옵션 순서 맵
    uid_option_map = {}
    if matching_df is not None:
        uid_option_map = _build_uid_option_order(matching_df)

    # UID별 사이즈유형 캐시 (동일 UID는 같은 사이즈유형)
    uid_size_type_cache = {}

    results = []
    warnings = []

    for idx, row in merged_df.iterrows():
        uid = _safe_str(row.get('상품코드', ''))
        barcode = _safe_str(row.get('88코드', ''))
        product_name = str(row.get('상품명', ''))
        file2_option = str(row.get('_파일2_옵션명', '')).strip() if '_파일2_옵션명' in row.index else ''
        file2_product = str(row.get('_파일2_상품명', '')).strip() if '_파일2_상품명' in row.index else ''

        is_option = has_option.get(uid, False)

        if not is_option:
            results.append({
                'idx': idx,
                '사이즈유형': 'MF3',
                '옵션슬롯': 1,
                '옵션값': 'FREE',
            })
            continue

        # ── 옵션 상품 처리 ──

        # Step 1: 사이즈유형 매칭 (캐시 활용)
        if uid in uid_size_type_cache:
            matched_type = uid_size_type_cache[uid]
        else:
            # 파일2 상품명 → 파일3 사이즈유형명 매칭 시도
            matched_type = _match_size_type(file2_product, size_type_entries)
            if matched_type is None:
                # fallback: 파일1 상품명의 기본 부분으로 시도
                base_name, _ = split_product_option(product_name)
                matched_type = _match_size_type(base_name, size_type_entries)
            if matched_type is None:
                # fallback: 파일1 상품명 전체로 시도
                matched_type = _match_size_type(product_name, size_type_entries)
            uid_size_type_cache[uid] = matched_type

        if matched_type is None:
            warnings.append({
                '88코드': barcode,
                '상품명': product_name,
                '원인': f'사이즈유형 매칭 실패 (파일2 상품명: {file2_product})',
            })
            results.append({
                'idx': idx,
                '사이즈유형': None,
                '옵션슬롯': None,
                '옵션값': file2_option or None,
            })
            continue

        size_type_code = matched_type['사이즈유형']

        # Step 2: 옵션 슬롯 결정
        option_name = file2_option  # 파일2 옵션명 우선
        slot_num = None
        slot_value = option_name

        # 방법 A: 파일2의 UID별 순서 맵 사용
        if uid in uid_option_map and barcode in uid_option_map[uid]:
            slot_num, map_option = uid_option_map[uid][barcode]
            if not option_name and map_option:
                option_name = map_option
                slot_value = map_option

        # 방법 B: 옵션명으로 파일3 Size 슬롯 매칭 (검증/보완)
        if option_name and slot_num is None:
            matched_slot, matched_val = _match_option_slot(
                option_name, matched_type['row'], size_cols
            )
            if matched_slot is not None:
                slot_num = matched_slot
                slot_value = matched_val

        # 방법 C: 파일1 상품명에서 # 기반 옵션 추출 (레거시 호환)
        if slot_num is None and not option_name:
            _, extracted_option = split_product_option(product_name)
            if extracted_option:
                option_name = extracted_option
                matched_slot, matched_val = _match_option_slot(
                    extracted_option, matched_type['row'], size_cols
                )
                if matched_slot is not None:
                    slot_num = matched_slot
                    slot_value = matched_val

        if slot_num is None:
            warnings.append({
                '88코드': barcode,
                '상품명': product_name,
                '원인': f"옵션 슬롯 매칭 실패 (사이즈유형: {size_type_code}, 옵션명: {option_name})",
            })

        results.append({
            'idx': idx,
            '사이즈유형': size_type_code,
            '옵션슬롯': slot_num,
            '옵션값': slot_value or option_name or None,
        })

    result_df = pd.DataFrame(results).set_index('idx')

    for col in ['사이즈유형', '옵션슬롯', '옵션값']:
        merged_df[col] = result_df[col]

    return merged_df, warnings


def _match_size_type(name: str, size_type_entries: list[dict]) -> dict | None:
    if not name or name == 'nan':
        return None

    name_strict = normalize_strict(name)

    # Level 1: 정규화 후 정확 매칭
    for entry in size_type_entries:
        if name_strict == entry['strict']:
            return entry

    # Level 2: 포함관계 매칭
    for entry in size_type_entries:
        type_strict = entry['strict']
        if name_strict and type_strict:
            if name_strict in type_strict or type_strict in name_strict:
                return entry

    # Level 3: RapidFuzz 유사도 매칭
    name_norm = normalize_for_matching(name)
    choices = {i: entry['normalized'] for i, entry in enumerate(size_type_entries)}

    if not choices:
        return None

    result = process.extractOne(
        query=name_norm,
        choices=choices,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=80,
    )

    if result:
        _, score, matched_idx = result
        return size_type_entries[matched_idx]

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

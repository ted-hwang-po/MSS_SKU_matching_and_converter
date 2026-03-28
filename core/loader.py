import re

import pandas as pd
from pathlib import Path


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명의 줄바꿈, 연속 공백을 정리"""
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    return df


def _find_header_row(filepath: Path, markers: list[str] = None) -> int | None:
    """엑셀 파일에서 헤더 행 번호를 탐색 (0-indexed).
    markers 중 하나라도 포함된 셀이 있는 첫 번째 행을 반환.
    """
    if markers is None:
        markers = ['브랜드명', '88코드', '바코드']

    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active
    found_row = None
    for row_idx in range(1, 15):  # 최대 14행까지 탐색
        for col_idx in range(1, 20):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val and any(m in str(val) for m in markers):
                found_row = row_idx - 1  # 0-indexed
                break
        if found_row is not None:
            break
    wb.close()
    return found_row


def load_excel_or_csv(filepath: str | Path, header: int = 0) -> pd.DataFrame:
    filepath = Path(filepath)
    if filepath.suffix.lower() == '.csv':
        return pd.read_csv(filepath, header=header)
    else:
        return pd.read_excel(filepath, engine='openpyxl', header=header)


def load_order_data(filepath: str | Path) -> pd.DataFrame:
    filepath = Path(filepath)

    # 먼저 기본(header=0)으로 시도
    df = load_excel_or_csv(filepath, header=0)
    df = _clean_columns(df)

    if '브랜드명' in df.columns:
        return df

    # 브랜드명을 못 찾으면 헤더 행 자동 탐색
    if filepath.suffix.lower() != '.csv':
        header_row = _find_header_row(filepath, ['브랜드명', '88코드'])
        if header_row is not None and header_row > 0:
            df = load_excel_or_csv(filepath, header=header_row)
            df = _clean_columns(df)
            if '브랜드명' in df.columns:
                return df

    return df


def load_matching_data(filepath: str | Path) -> pd.DataFrame:
    filepath = Path(filepath)
    if filepath.suffix.lower() == '.csv':
        df = pd.read_csv(filepath)
    else:
        df = pd.read_excel(filepath, engine='openpyxl')

    # 첫 행이 빈 행이거나, 컬럼명이 Unnamed으로 시작하면 헤더 재탐색
    unnamed_cols = [c for c in df.columns if str(c).startswith('Unnamed')]
    if len(unnamed_cols) > len(df.columns) // 2:
        # 헤더 행 자동 탐색
        if filepath.suffix.lower() != '.csv':
            header_row = _find_header_row(Path(filepath), ['바코드', '스타일번호', '상품코드'])
            if header_row is not None:
                df = pd.read_excel(filepath, engine='openpyxl', header=header_row)
        else:
            for i in range(len(df)):
                row = df.iloc[i]
                non_null = row.dropna()
                if len(non_null) > len(df.columns) // 2:
                    df = pd.read_csv(filepath, header=i + 1)
                    break

    df = _clean_columns(df)
    return df


def load_option_data(filepath: str | Path) -> pd.DataFrame:
    df = load_excel_or_csv(filepath)
    df = _clean_columns(df)
    return df


def get_brand_list(order_df: pd.DataFrame) -> list[str]:
    if '브랜드명' not in order_df.columns:
        raise ValueError("총 발주 수량 데이터셋에 '브랜드명' 컬럼이 없습니다.")
    brands = order_df['브랜드명'].dropna().unique().tolist()
    return sorted(brands)


def filter_by_brand(order_df: pd.DataFrame, brand: str) -> pd.DataFrame:
    return order_df[order_df['브랜드명'] == brand].copy()

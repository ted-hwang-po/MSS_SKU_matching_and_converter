import re

import pandas as pd
from pathlib import Path


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명의 줄바꿈, 연속 공백을 정리"""
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    return df


def load_excel_or_csv(filepath: str | Path) -> pd.DataFrame:
    filepath = Path(filepath)
    if filepath.suffix.lower() == '.csv':
        return pd.read_csv(filepath)
    else:
        return pd.read_excel(filepath, engine='openpyxl')


def load_order_data(filepath: str | Path) -> pd.DataFrame:
    df = load_excel_or_csv(filepath)
    df = _clean_columns(df)
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
        # 첫 번째 비어있지 않은 행을 헤더로 사용
        for i in range(len(df)):
            row = df.iloc[i]
            non_null = row.dropna()
            if len(non_null) > len(df.columns) // 2:
                if filepath.suffix.lower() == '.csv':
                    df = pd.read_csv(filepath, header=i + 1)
                else:
                    df = pd.read_excel(filepath, engine='openpyxl', header=i + 1)
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

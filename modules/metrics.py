"""
財務指標の計算を担当するモジュール

主な機能：
- 各種財務比率の計算
- 外れ値の判定
"""

from typing import Tuple, Optional

import numpy as np
import pandas as pd


def add_financial_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """
    財務指標を追加する
    
    以下の指標を計算して列として追加：
    - 営業利益率 = 営業利益 / 売上高
    - 当期純利益率 = 当期純利益 / 売上高
    - 自己資本比率 = 自己資本 / 総資産
    - 総資産回転率 = 売上高 / 総資産
    - ROE_ROA_gap = ROE - ROA
    
    Parameters
    ----------
    df : pd.DataFrame
        財務データを含むデータフレーム
    
    Returns
    -------
    pd.DataFrame
        指標が追加されたデータフレーム
    """
    df_result = df.copy()
    
    # 営業利益率（%）
    if "営業利益" in df.columns and "売上高" in df.columns:
        df_result["営業利益率"] = (df_result["営業利益"] / df_result["売上高"]) * 100
    
    # 当期純利益率（%）
    if "当期純利益" in df.columns and "売上高" in df.columns:
        df_result["当期純利益率"] = (df_result["当期純利益"] / df_result["売上高"]) * 100
    
    # 自己資本比率（%）
    if "自己資本" in df.columns and "総資産" in df.columns:
        df_result["自己資本比率"] = (df_result["自己資本"] / df_result["総資産"]) * 100
    
    # 総資産回転率（回）
    if "売上高" in df.columns and "総資産" in df.columns:
        df_result["総資産回転率"] = df_result["売上高"] / df_result["総資産"]
    
    # ROE-ROA差分
    if "ROE" in df.columns and "ROA" in df.columns:
        df_result["ROE_ROA_gap"] = df_result["ROE"] - df_result["ROA"]
    
    # 利益率差分（営業利益率 - 当期純利益率）
    if "営業利益率" in df_result.columns and "当期純利益率" in df_result.columns:
        df_result["利益率差分"] = df_result["営業利益率"] - df_result["当期純利益率"]
    
    return df_result


def normalize_percentage_columns(
    df: pd.DataFrame,
    cols: list[str]
) -> pd.DataFrame:
    """
    パーセント表記の列を0-100の範囲に正規化する
    
    例：0.05（5%）→ 5.0、または既に5.0の場合はそのまま
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    cols : list[str]
        正規化する列名のリスト
    
    Returns
    -------
    pd.DataFrame
        正規化後のデータフレーム
    """
    df_result = df.copy()
    
    for col in cols:
        if col not in df_result.columns:
            continue
        
        # 最大値が1以下の場合は0-1スケールと判断して100倍
        max_val = df_result[col].max()
        if max_val <= 1.0 and max_val > 0:
            print(f"列 '{col}' を100倍してパーセント表記に変換しました")
            df_result[col] = df_result[col] * 100
    
    return df_result


def detect_outliers_iqr(
    df: pd.DataFrame,
    col: str,
    factor: float = 1.5
) -> Tuple[pd.Series, float, float]:
    """
    IQR法で外れ値を検出する
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    col : str
        外れ値を検出する列名
    factor : float, default 1.5
        IQRに掛ける係数（通常は1.5）
    
    Returns
    -------
    outliers : pd.Series
        外れ値かどうかのブール値（Trueが外れ値）
    lower_bound : float
        下限値
    upper_bound : float
        上限値
    """
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    
    lower_bound = Q1 - factor * IQR
    upper_bound = Q3 + factor * IQR
    
    outliers = (df[col] < lower_bound) | (df[col] > upper_bound)
    
    return outliers, lower_bound, upper_bound


def detect_outliers_zscore(
    df: pd.DataFrame,
    col: str,
    threshold: float = 3.0
) -> pd.Series:
    """
    Z-score法で外れ値を検出する
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    col : str
        外れ値を検出する列名
    threshold : float, default 3.0
        外れ値とみなすZ-scoreの閾値
    
    Returns
    -------
    pd.Series
        外れ値かどうかのブール値（Trueが外れ値）
    """
    mean = df[col].mean()
    std = df[col].std()
    
    z_scores = np.abs((df[col] - mean) / std)
    outliers = z_scores > threshold
    
    return outliers


def get_top_bottom_companies(
    df: pd.DataFrame,
    col: str,
    n: int = 3,
    company_col: str = "企業名"
) -> dict:
    """
    指定列の上位・下位企業を取得する
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    col : str
        ソート対象の列名
    n : int, default 3
        取得する企業数
    company_col : str, default "企業名"
        企業名を含む列名
    
    Returns
    -------
    dict
        'top'と'bottom'をキーとする辞書
        各値は企業名と値のタプルのリスト
    """
    # 欠損値を除外
    df_valid = df[[company_col, col]].dropna()
    
    # 上位n社
    top_companies = df_valid.nlargest(n, col)
    top_list = [
        (row[company_col], row[col]) 
        for _, row in top_companies.iterrows()
    ]
    
    # 下位n社
    bottom_companies = df_valid.nsmallest(n, col)
    bottom_list = [
        (row[company_col], row[col]) 
        for _, row in bottom_companies.iterrows()
    ]
    
    return {
        "top": top_list,
        "bottom": bottom_list
    }

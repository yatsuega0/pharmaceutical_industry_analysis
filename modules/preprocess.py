"""
データの前処理を担当するモジュール

主な機能：
- 数値列の型変換
- 必須列の存在確認
- 欠損値の処理
"""

from typing import List, Optional, Literal

import numpy as np
import pandas as pd


def coerce_numeric(
    df: pd.DataFrame,
    cols: List[str],
    inplace: bool = False
) -> pd.DataFrame:
    """
    指定された列を数値型に変換する
    
    カンマや全角数字などを除去してから変換を行う
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    cols : List[str]
        数値型に変換する列名のリスト
    inplace : bool, default False
        元のデータフレームを変更するかどうか
    
    Returns
    -------
    pd.DataFrame
        変換後のデータフレーム
    """
    if not inplace:
        df = df.copy()
    
    for col in cols:
        if col not in df.columns:
            print(f"警告: 列 '{col}' が見つかりません")
            continue
        
        # 文字列型の場合はカンマなどを除去
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.replace(",", "").str.replace("，", "")
            df[col] = df[col].str.strip()
        
        # 数値に変換（変換できない場合はNaNになる）
        df[col] = pd.to_numeric(df[col], errors="coerce")
    
    return df


def validate_columns(
    df: pd.DataFrame,
    required_cols: List[str]
) -> None:
    """
    必須列の存在をチェックする
    
    Parameters
    ----------
    df : pd.DataFrame
        チェック対象のデータフレーム
    required_cols : List[str]
        必須列名のリスト
    
    Raises
    ------
    ValueError
        必須列が存在しない場合
    """
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        raise ValueError(f"必須列が見つかりません: {missing_cols}")
    
    print(f"必須列の確認完了: {len(required_cols)}列")


def handle_missing(
    df: pd.DataFrame,
    strategy: Literal["drop", "warn", "fill_zero"] = "warn",
    subset: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    欠損値の処理を行う
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    strategy : {"drop", "warn", "fill_zero"}, default "warn"
        欠損値の処理方法
        - "drop": 欠損を含む行を削除
        - "warn": 警告のみ表示（削除しない）
        - "fill_zero": 欠損値を0で埋める
    subset : List[str], optional
        処理対象の列（Noneの場合は全列）
    
    Returns
    -------
    pd.DataFrame
        処理後のデータフレーム
    """
    df_result = df.copy()
    
    # 対象列の設定
    cols_to_check = subset if subset is not None else df.columns.tolist()
    
    # 欠損値の確認
    missing_counts = df_result[cols_to_check].isnull().sum()
    has_missing = missing_counts > 0
    
    if has_missing.any():
        print("欠損値が検出されました:")
        for col in missing_counts[has_missing].index:
            print(f"  {col}: {missing_counts[col]}件")
        
        if strategy == "drop":
            df_result = df_result.dropna(subset=cols_to_check)
            print(f"欠損を含む行を削除しました（残り{len(df_result)}行）")
        elif strategy == "fill_zero":
            df_result[cols_to_check] = df_result[cols_to_check].fillna(0)
            print("欠損値を0で埋めました")
        elif strategy == "warn":
            print("警告: 欠損値が存在します（処理は行いません）")
    else:
        print("欠損値は検出されませんでした")
    
    return df_result

"""
データの入出力を担当するモジュール

主な機能：
- Excel形式の財務データの読み込み
- 列名の正規化
- テーブルデータの保存
- 出力ディレクトリの確認・作成
"""

from pathlib import Path
from typing import Optional

import pandas as pd


def load_financial_xlsx(path: str) -> pd.DataFrame:
    """
    Excel形式の財務データを読み込む
    
    列名の正規化（全角/半角統一、空白除去など）を行う
    
    Parameters
    ----------
    path : str
        読み込むExcelファイルのパス
    
    Returns
    -------
    pd.DataFrame
        読み込んだデータフレーム（列名正規化済み）
    
    Raises
    ------
    FileNotFoundError
        指定されたファイルが見つからない場合
    """
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {path}")
    
    # Excelファイル読み込み
    df = pd.read_excel(path)
    
    # 列名の正規化
    df.columns = df.columns.str.strip()  # 前後の空白除去
    df.columns = df.columns.str.replace("　", " ")  # 全角スペースを半角に
    
    return df


def save_table(
    df: pd.DataFrame,
    path: str,
    index: bool = False,
    **kwargs
) -> None:
    """
    データフレームをファイルに保存する
    
    拡張子に応じて自動的に形式を判定（.csv, .xlsx）
    
    Parameters
    ----------
    df : pd.DataFrame
        保存するデータフレーム
    path : str
        保存先のファイルパス
    index : bool, default False
        インデックスを含めるかどうか
    **kwargs
        その他のpandas保存関数に渡す引数
    """
    file_path = Path(path)
    
    # 親ディレクトリを作成
    ensure_output_dir(str(file_path.parent))
    
    # 拡張子に応じて保存
    if file_path.suffix == ".csv":
        df.to_csv(path, index=index, encoding="utf-8-sig", **kwargs)
    elif file_path.suffix in [".xlsx", ".xls"]:
        df.to_excel(path, index=index, **kwargs)
    else:
        raise ValueError(f"サポートされていないファイル形式です: {file_path.suffix}")
    
    print(f"ファイルを保存しました: {path}")


def ensure_output_dir(path: str) -> None:
    """
    指定されたディレクトリが存在しない場合は作成する
    
    Parameters
    ----------
    path : str
        確認・作成するディレクトリのパス
    """
    dir_path = Path(path)
    if not dir_path.exists():
        dir_path.mkdir(parents=True, exist_ok=True)
        print(f"ディレクトリを作成しました: {path}")

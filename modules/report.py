"""
レポート生成を担当するモジュール

主な機能：
- 集計サマリーの作成
- Markdownレポートの生成
"""

from typing import List, Optional, Dict, Any

import pandas as pd
import numpy as np


def generate_summary_stats(
    df: pd.DataFrame,
    cols: List[str]
) -> pd.DataFrame:
    """
    指定された列の基本統計量を計算する
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    cols : List[str]
        統計量を計算する列名のリスト
    
    Returns
    -------
    pd.DataFrame
        統計量をまとめたデータフレーム
        （行：統計量の種類、列：指定した列）
    """
    stats = []
    
    for col in cols:
        if col not in df.columns:
            continue
        
        col_stats = {
            "列名": col,
            "件数": df[col].count(),
            "平均": df[col].mean(),
            "中央値": df[col].median(),
            "標準偏差": df[col].std(),
            "最小値": df[col].min(),
            "25%点": df[col].quantile(0.25),
            "75%点": df[col].quantile(0.75),
            "最大値": df[col].max()
        }
        stats.append(col_stats)
    
    return pd.DataFrame(stats)


def create_markdown_summary(
    output_path: str,
    title: str,
    sections: List[Dict[str, Any]]
) -> None:
    """
    Markdown形式のサマリーレポートを生成する
    
    Parameters
    ----------
    output_path : str
        保存先のファイルパス
    title : str
        レポートのタイトル
    sections : List[Dict[str, Any]]
        セクションのリスト
        各セクションは以下のキーを持つ辞書：
        - "heading": セクションの見出し
        - "content": セクションの内容（文字列またはリスト）
    """
    with open(output_path, "w", encoding="utf-8") as f:
        # タイトル
        f.write(f"# {title}\n\n")
        
        # 各セクション
        for section in sections:
            heading = section.get("heading", "")
            content = section.get("content", "")
            
            # 見出し
            if heading:
                f.write(f"## {heading}\n\n")
            
            # 内容
            if isinstance(content, list):
                for item in content:
                    # 見出し（###で始まる）や空文字列の場合は箇条書き記号を付けない
                    if item.strip().startswith("###") or item.strip() == "":
                        f.write(f"{item}\n")
                    else:
                        f.write(f"- {item}\n")
                f.write("\n")
            else:
                f.write(f"{content}\n\n")
    
    print(f"Markdownサマリーを保存しました: {output_path}")


def format_top_bottom_summary(
    top_bottom_dict: Dict[str, List[tuple]],
    metric_name: str
) -> List[str]:
    """
    上位・下位企業の情報をフォーマットする
    
    Parameters
    ----------
    top_bottom_dict : Dict[str, List[tuple]]
        modules.metrics.get_top_bottom_companiesの戻り値
    metric_name : str
        指標名
    
    Returns
    -------
    List[str]
        フォーマット済みの文字列リスト
    """
    lines = []
    
    # 上位
    lines.append(f"### {metric_name} 上位企業")
    for i, (company, value) in enumerate(top_bottom_dict["top"], 1):
        lines.append(f"{i}. {company}: {value:.2f}")
    
    lines.append("")
    
    # 下位
    lines.append(f"### {metric_name} 下位企業")
    for i, (company, value) in enumerate(top_bottom_dict["bottom"], 1):
        lines.append(f"{i}. {company}: {value:.2f}")
    
    return lines


def create_data_quality_report(
    df: pd.DataFrame,
    required_cols: List[str]
) -> List[str]:
    """
    データ品質に関するレポートを生成する
    
    Parameters
    ----------
    df : pd.DataFrame
        対象のデータフレーム
    required_cols : List[str]
        必須列のリスト
    
    Returns
    -------
    List[str]
        レポートの内容（行のリスト）
    """
    report = []
    
    # 基本情報
    report.append(f"総行数: {len(df)}")
    report.append(f"総列数: {len(df.columns)}")
    report.append("")
    
    # 欠損値の確認
    missing = df[required_cols].isnull().sum()
    if missing.sum() > 0:
        report.append("欠損値あり:")
        for col, count in missing[missing > 0].items():
            report.append(f"{col}: {count}件 ({count/len(df)*100:.1f}%)")
    else:
        report.append("欠損値: なし")
    
    report.append("")
    
    # データ型の確認
    report.append("#### 列のデータ型:")
    for col in required_cols:
        if col in df.columns:
            report.append(f"{col}: {df[col].dtype}")
    
    return report


def create_analysis_metadata(
    data_path: str,
    n_companies: int,
    target_year: str,
    used_columns: List[str],
    created_metrics: List[str],
    missing_strategy: str
) -> List[str]:
    """
    分析のメタデータ情報を生成する
    
    Parameters
    ----------
    data_path : str
        元データのパス
    n_companies : int
        分析対象企業数
    target_year : str
        対象年度
    used_columns : List[str]
        使用した列名のリスト
    created_metrics : List[str]
        作成した指標のリスト
    missing_strategy : str
        欠損値の処理方針
    
    Returns
    -------
    List[str]
        メタデータの内容（行のリスト）
    """
    metadata = [
        f"#### データソース: {data_path}",
        f"#### 分析対象: {n_companies}社",
        f"#### 対象年度: {target_year}",
        "",
        "#### 使用した列:",
    ]
    
    for col in used_columns:
        metadata.append(f"{col}")
    
    metadata.append("")
    metadata.append("#### 作成した指標:")
    
    for metric in created_metrics:
        metadata.append(f"{metric}")
    
    metadata.append("")
    metadata.append(f"#### 欠損値の処理: {missing_strategy}")
    
    return metadata


def format_table_for_markdown(
    df: pd.DataFrame,
    large_number_cols: Optional[List[str]] = None,
    decimal_places: int = 2
) -> str:
    """
    DataFrameをMarkdown形式でフォーマットする
    
    大きな数値列は科学的記数法を避けてカンマ区切り表記にする。
    その他の数値列は指定された小数点桁数で丸める。
    
    Parameters
    ----------
    df : pd.DataFrame
        フォーマット対象のデータフレーム
    large_number_cols : List[str], optional
        カンマ区切り表記にする列名のリスト（例: ["売上高", "総資産"]）
        指定しない場合は、全ての数値列を小数点表記にする
    decimal_places : int, default=2
        小数点以下の桁数
    
    Returns
    -------
    str
        Markdown形式のテーブル
    
    Examples
    --------
    >>> df = pd.DataFrame({
    ...     'クラスター': [1, 2],
    ...     '売上高': [1142544.0, 267070.8],
    ...     'ROA': [6.54, 9.76]
    ... })
    >>> print(format_table_for_markdown(df, large_number_cols=['売上高']))
    |   クラスター | 売上高       |   ROA |
    |------------:|:-------------|------:|
    |           1 | 1,142,544    |  6.54 |
    |           2 | 267,071      |  9.76 |
    """
    # コピーを作成して元のデータを変更しない
    df_formatted = df.copy()
    
    # large_number_colsが指定されている場合、それらの列をカンマ区切り整数表記に変換
    if large_number_cols:
        for col in large_number_cols:
            if col in df_formatted.columns:
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else ""
                )
    
    # その他の数値列は小数点桁数で丸める
    for col in df_formatted.columns:
        if large_number_cols and col in large_number_cols:
            # 既に文字列に変換済みなのでスキップ
            continue
        
        # 数値型の列のみ処理
        if pd.api.types.is_numeric_dtype(df_formatted[col]):
            df_formatted[col] = df_formatted[col].apply(
                lambda x: f"{x:.{decimal_places}f}" if pd.notna(x) else ""
            )
    
    return df_formatted.to_markdown()

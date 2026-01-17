"""
可視化を担当するモジュール

主な機能：
- 散布図・バブル図の作成
- 企業名ラベルの付与
- 画像保存の統一処理
"""

from typing import List, Optional, Tuple

import japanize_matplotlib  # 日本語フォント対応 # noqa: F401
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

# デフォルトのスタイル設定
plt.style.use("seaborn-v0_8-darkgrid")
sns.set_palette("husl")

# スタイル設定後に日本語フォントを再設定（スタイルでリセットされるため）
plt.rcParams["font.sans-serif"] = [
    "Hiragino Sans",
    "Yu Gothic",
    "Meirio",
    "Takao",
    "IPAexGothic",
    "IPAPGothic",
    "Noto Sans CJK JP",
]
plt.rcParams["axes.unicode_minus"] = False  # マイナス記号の文字化け対策


def create_scatter_plot(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    output_path: str,
    title: str = "",
    x_label: Optional[str] = None,
    y_label: Optional[str] = None,
    annotate: bool = False,
    company_col: str = "企業名",
    figsize: Tuple[float, float] = (10, 6),
    log_x: bool = False,
    log_y: bool = False,
    dpi: int = 150,
    show: bool = True,
) -> None:
    """
    散布図を作成して保存する

    Parameters
    ----------
    df : pd.DataFrame
        プロットするデータ
    x_col : str
        X軸の列名
    y_col : str
        Y軸の列名
    output_path : str
        保存先のファイルパス
    title : str, default ""
        グラフのタイトル
    x_label : str, optional
        X軸のラベル（Noneの場合は列名を使用）
    y_label : str, optional
        Y軸のラベル（Noneの場合は列名を使用）
    annotate : bool, default False
        企業名ラベルを付けるかどうか
    company_col : str, default "企業名"
        企業名の列名
    figsize : Tuple[float, float], default (10, 6)
        図のサイズ
    log_x : bool, default False
        X軸を対数スケールにするかどうか
    log_y : bool, default False
        Y軸を対数スケールにするかどうか
    dpi : int, default 150
        保存時の解像度
    show : bool, default True
        グラフを表示するかどうか
    """
    fig, ax = plt.subplots(figsize=figsize)

    # 欠損値を除外
    df_plot = df[[x_col, y_col, company_col]].dropna()

    # 散布図を描画
    ax.scatter(df_plot[x_col], df_plot[y_col], alpha=0.6, s=100)

    # ラベルの設定
    ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
    ax.set_ylabel(y_label if y_label else y_col, fontsize=12)
    ax.set_title(title, fontsize=14, fontweight="bold")

    # 対数スケール
    if log_x:
        ax.set_xscale("log")
    if log_y:
        ax.set_yscale("log")

    # 企業名のラベル付け
    if annotate:
        for _, row in df_plot.iterrows():
            ax.annotate(
                row[company_col],
                (row[x_col], row[y_col]),
                xytext=(5, 5),
                textcoords="offset points",
                fontsize=8,
                alpha=0.7,
            )

    # グリッド
    ax.grid(True, alpha=0.3)

    # 保存
    plt.tight_layout()
    plt.savefig(output_path, dpi=dpi, bbox_inches="tight")
    print(f"散布図を保存しました: {output_path}")

    # 表示
    if show:
        plt.show()
    else:
        plt.close()


def create_bubble_chart(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    size_col: str,
    output_path: str,
    title: str = "",
    x_label: Optional[str] = None,
    y_label: Optional[str] = None,
    annotate: bool = False,
    company_col: str = "企業名",
    figsize: Tuple[float, float] = (10, 6),
    size_scale: float = 1.0,
    dpi: int = 150,
    show: bool = True,
) -> None:
    """
    バブル図を作成して保存する

    Parameters
    ----------
    df : pd.DataFrame
        プロットするデータ
    x_col : str
        X軸の列名
    y_col : str
        Y軸の列名
    size_col : str
        バブルサイズを決める列名
    output_path : str
        保存先のファイルパス
    title : str, default ""
        グラフのタイトル
    x_label : str, optional
        X軸のラベル
    y_label : str, optional
        Y軸のラベル
    annotate : bool, default False
        企業名ラベルを付けるかどうか
    company_col : str, default "企業名"
        企業名の列名
    figsize : Tuple[float, float], default (10, 6)
        図のサイズ
    size_scale : float, default 1.0
        バブルサイズのスケーリング係数
    dpi : int, default 150
        保存時の解像度
    show : bool, default True
        グラフを表示するかどうか
    """
    fig, ax = plt.subplots(figsize=figsize)

    # 欠損値を除外
    df_plot = df[[x_col, y_col, size_col, company_col]].dropna()

    # サイズを正規化
    sizes = (df_plot[size_col] / df_plot[size_col].max()) * 1000 * size_scale

    # バブル図を描画
    scatter = ax.scatter(
        df_plot[x_col],
        df_plot[y_col],
        s=sizes,
        alpha=0.5,
        edgecolors="w",
        linewidth=0.5,
    )

    # ラベルの設定
    ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
    ax.set_ylabel(y_label if y_label else y_col, fontsize=12)
    ax.set_title(title, fontsize=14, fontweight="bold")

    # 企業名のラベル付け
    if annotate:
        for _, row in df_plot.iterrows():
            ax.annotate(
                row[company_col],
                (row[x_col], row[y_col]),
                xytext=(5, 5),
                textcoords="offset points",
                fontsize=8,
                alpha=0.7,
            )

    # グリッド
    ax.grid(True, alpha=0.3)

    # 保存
    plt.tight_layout()
    plt.savefig(output_path, dpi=dpi, bbox_inches="tight")
    print(f"バブル図を保存しました: {output_path}")

    # 表示
    if show:
        plt.show()
    else:
        plt.close()


def create_bar_chart(
    df: pd.DataFrame,
    x_col: str,
    y_cols: List[str],
    output_path: str,
    title: str = "",
    x_label: Optional[str] = None,
    y_label: Optional[str] = None,
    figsize: Tuple[float, float] = (12, 6),
    sort_by: Optional[str] = None,
    ascending: bool = False,
    dpi: int = 150,
    show: bool = True,
) -> None:
    """
    棒グラフを作成して保存する

    Parameters
    ----------
    df : pd.DataFrame
        プロットするデータ
    x_col : str
        X軸（カテゴリ）の列名
    y_cols : List[str]
        Y軸の列名のリスト（複数指定でグループ化棒グラフ）
    output_path : str
        保存先のファイルパス
    title : str, default ""
        グラフのタイトル
    x_label : str, optional
        X軸のラベル
    y_label : str, optional
        Y軸のラベル
    figsize : Tuple[float, float], default (12, 6)
        図のサイズ
    sort_by : str, optional
        ソートする列名
    ascending : bool, default False
        昇順でソートするかどうか
    dpi : int, default 150
        保存時の解像度
    show : bool, default True
        グラフを表示するかどうか
    """
    fig, ax = plt.subplots(figsize=figsize)

    # ソート
    df_plot = df.copy()
    if sort_by and sort_by in df_plot.columns:
        df_plot = df_plot.sort_values(by=sort_by, ascending=ascending)

    # 棒グラフを描画
    df_plot.plot(x=x_col, y=y_cols, kind="bar", ax=ax, width=0.8)

    # ラベルの設定
    ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
    ax.set_ylabel(y_label if y_label else "", fontsize=12)
    ax.set_title(title, fontsize=14, fontweight="bold")

    # X軸のラベルを回転
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha="right")

    # 凡例
    ax.legend(loc="best")

    # グリッド
    ax.grid(True, alpha=0.3, axis="y")

    # 保存
    plt.tight_layout()
    plt.savefig(output_path, dpi=dpi, bbox_inches="tight")
    print(f"棒グラフを保存しました: {output_path}")

    # 表示
    if show:
        plt.show()
    else:
        plt.close()


def annotate_points(
    ax: plt.Axes,
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    label_col: str,
    condition: Optional[pd.Series] = None,
) -> None:
    """
    指定した条件に合うポイントにラベルを付与する

    Parameters
    ----------
    ax : plt.Axes
        プロットするAxesオブジェクト
    df : pd.DataFrame
        データフレーム
    x_col : str
        X軸の列名
    y_col : str
        Y軸の列名
    label_col : str
        ラベルとして使用する列名
    condition : pd.Series, optional
        ラベル付けする条件（ブール値のSeries）
        Noneの場合は全てのポイントにラベルを付ける
    """
    df_to_label = df if condition is None else df[condition]

    for _, row in df_to_label.iterrows():
        ax.annotate(
            row[label_col],
            (row[x_col], row[y_col]),
            xytext=(5, 5),
            textcoords="offset points",
            fontsize=8,
            alpha=0.7,
        )

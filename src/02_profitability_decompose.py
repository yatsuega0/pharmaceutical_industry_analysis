# %%[markdown]
# # 分析②：収益性分解
# 
# ## 目的
# - 本業（営業利益）と最終利益（純利益）の差を比較し、利益の質を把握する
# 
# ### 仮説
# - **H2-1**: 営業利益率が高いが当期純利益率が低い企業は、特別損益・金融費用・税負担などの影響が相対的に大きい
# - **H2-2**: 営業利益率と当期純利益率が近い企業ほど、利益構造が安定的で説明しやすい
# 
# ### 実施内容
# 1. データの読み込みと前処理
# 2. 営業利益率・当期純利益率の計算
# 3. 利益率の差分分析
# 4. 棒グラフと散布図による可視化
# 5. 乖離が大きい企業の特定

# %%
import sys
from pathlib import Path

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import japanize_matplotlib

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from modules import io, preprocess, metrics, viz, report

# %%[markdown]
# ### データの読み込みと前処理
# - Excelファイルから財務データを読み込む
# - 数値列の型変換を行う

# %%
# データファイルのパス
data_path = project_root / "input" / "財務データ_製薬業界.xlsx"

# データ読み込み
df = io.load_financial_xlsx(str(data_path))

print(f"読み込んだデータ: {df.shape}")

# %%
# 必須列の確認
required_cols = [
    "証券コード", "企業名", "売上高", "営業利益", "当期純利益"
]

preprocess.validate_columns(df, required_cols)

# %%
# 数値列の型変換
print(df.shape)
numeric_cols = ["売上高", "営業利益", "当期純利益"]
df = preprocess.coerce_numeric(df, numeric_cols)
print(df.shape)
# %%
# 欠損値の確認
df = preprocess.handle_missing(df, strategy="warn", subset=required_cols)
print(df.isna().sum())
# %%[markdown]
# ### 利益率の計算
# - 営業利益率と当期純利益率を計算
# - 利益率差分（営業利益率 - 当期純利益率）を算出

# %%
# 財務指標を追加
print(df.shape)
df = metrics.add_financial_ratios(df)

# 利益率に関する列を確認
margin_cols = ["営業利益率", "当期純利益率", "利益率差分"]
print("\n利益率関連の指標:")
for col in margin_cols:
    if col in df.columns:
        print(f"  - {col}: {df[col].notna().sum()}件")
print(df.shape)
# %%
# 利益率テーブルを保存
output_cols = [
    "証券コード", "企業名", "売上高", "営業利益", "当期純利益",
    "営業利益率", "当期純利益率", "利益率差分"
]

df_output = df[output_cols].copy()

# 利益率差分で降順ソート（差が大きい順）
df_output = df_output.sort_values("利益率差分", ascending=False, na_position="last")

io.save_table(df_output, str(project_root / "output" / "02_margin_table.xlsx"))
# %%[markdown]
# ### 可視化①：企業別の営業利益率・当期純利益率（棒グラフ）
# - 各企業の営業利益率と当期純利益率を並べて比較

# %%
# 利益率差分で降順ソート
df_sorted = df.sort_values("利益率差分", ascending=False, na_position="last")

viz.create_bar_chart(
    df=df_sorted,
    x_col="企業名",
    y_cols=["営業利益率", "当期純利益率"],
    output_path=str(project_root / "output" / "fig_02_margins_by_company.png"),
    title="企業別の営業利益率・当期純利益率",
    y_label="利益率（%）",
    figsize=(14, 6)
)

# %%[markdown]
# ### 可視化②：営業利益率 vs 当期純利益率（散布図）
# - 45度線を引いて、乖離を視覚的に確認

# %%
# 散布図の作成
fig, ax = plt.subplots(figsize=(10, 8))

# データのプロット
df_plot = df[["企業名", "営業利益率", "当期純利益率"]].dropna()
ax.scatter(df_plot["営業利益率"], df_plot["当期純利益率"], alpha=0.6, s=100)

# 45度線を追加
max_val = max(df_plot["営業利益率"].max(), df_plot["当期純利益率"].max())
min_val = min(df_plot["営業利益率"].min(), df_plot["当期純利益率"].min())
ax.plot([min_val, max_val], [min_val, max_val], "r--", alpha=0.5, label="45度線")

# 企業名ラベル
for _, row in df_plot.iterrows():
    ax.annotate(
        row["企業名"],
        (row["営業利益率"], row["当期純利益率"]),
        xytext=(5, 5),
        textcoords="offset points",
        fontsize=8,
        alpha=0.7
    )

# ラベルとタイトル
ax.set_xlabel("営業利益率（%）", fontsize=12)
ax.set_ylabel("当期純利益率（%）", fontsize=12)
ax.set_title("営業利益率と当期純利益率の関係", fontsize=14, fontweight="bold")
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(project_root / "output" / "fig_02_opm_vs_npm.png", dpi=150, bbox_inches="tight")
plt.show()

print("散布図を保存しました: output/fig_02_opm_vs_npm.png")

# %%[markdown]
# ### 利益率差分の分析
# - 差分が大きい企業（上位・下位）を特定

# %%
# 利益率差分の上位・下位
margin_gap_top_bottom = metrics.get_top_bottom_companies(
    df, "利益率差分", n=3, company_col="企業名"
)

print("\n=== 利益率差分 上位3社（営業利益率 > 当期純利益率の差が大きい）===")
for company, value in margin_gap_top_bottom["top"]:
    print(f"  {company}: {value:.2f}")

print("\n=== 利益率差分 下位3社（営業利益率 < 当期純利益率 または差が小さい）===")
for company, value in margin_gap_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}")

# %%
# 営業利益率の上位・下位
opm_top_bottom = metrics.get_top_bottom_companies(df, "営業利益率", n=3)

print("\n=== 営業利益率 上位3社 ===")
for company, value in opm_top_bottom["top"]:
    print(f"  {company}: {value:.2f}%")

print("\n=== 営業利益率 下位3社 ===")
for company, value in opm_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}%")

# %%
# 純利益率の上位・下位
npm_top_bottom = metrics.get_top_bottom_companies(df, "当期純利益率", n=3)

print("\n=== 当期純利益率 上位3社 ===")
for company, value in npm_top_bottom["top"]:
    print(f"  {company}: {value:.2f}%")

print("\n=== 当期純利益率 下位3社 ===")
for company, value in npm_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}%")

# %%[markdown]
# ---
# ### 結果の確認と考察の記載
# 
# **ここで一旦、以下の結果を確認してください：**
# - 生成されたグラフ（営業利益率 vs 当期純利益率、棒グラフ）
# - 利益率差分、営業利益率、当期純利益率の上位・下位企業

# %%
# TODO: グラフと数値を確認した上で、観察された事実を記載してください
# 
# 記載方法の注意：
# - 各文字列の後に必ずカンマ(,)を付ける
# - セクション間に空行を入れたい場合は空文字列 "" を追加
# - Markdown形式で記載可能（**太字**、箇条書きなど）
#
# 記載例：
# observations = [
#     "### ① 営業利益率 vs 当期純利益率",
#     "観察事項１",
#     "観察事項２",
#     "### ② 利益率差分の大きい企業",
#     "観察事項３",
# ]
observations = [
    "#### ① 営業利益率 vs 当期純利益率",
    "高い営業利益率を示す企業であっても、当期純利益率は必ずしも同水準に達しておらず、本業の収益力と最終的な収益成果の間に乖離が生じるケースが確認される。",
    "特に小林製薬・武田薬品工業は乖離が大きく、非本業要因の影響が相対的に大きい構造が示唆される。",
    "#### ② 利益率差分の大きい企業",
    "45度線付近に位置する企業（日本新薬、第一三共など）は、営業利益と純利益の乖離が相対的に小さく、営業段階の収益性が最終利益に安定的に直結している可能性が高い。",
    "中外製薬は高い営業利益率を示す一方で45度線の下側に位置しており、営業利益に対して純利益が相対的に低く、最終利益段階での調整が生じていることが確認できる。",
    "#### ③ 利益率差分の分析",
    "差分上位（小林製薬・武田薬品工業・中外製薬）は、営業力の強さに対し最終利益の毀損が大きいグループ。",
    "差分下位（塩野義製薬・アステラス製薬・エーザイ）は、営業利益と当期純利益が近く、構造的に安定的。",
    "#### 仮説に対する評価",
    "**H2-1**：営業利益率が高い企業ほど差分が拡大するケースが確認され、仮説は概ね支持される。",
    "**H2-2**：両利益率が近い企業は分布上も安定しており、利益構造が説明しやすい点で仮説は妥当。"
]

# %%[markdown]
# ### サマリーレポートの作成
# - 分析結果をMarkdown形式でまとめる

# %%
# 基本統計量
summary_stats = report.generate_summary_stats(
    df,
    cols=["営業利益率", "当期純利益率", "利益率差分"]
)
print(summary_stats)

# メタデータ
metadata = report.create_analysis_metadata(
    data_path=data_path.name,  # ファイル名のみを使用
    n_companies=len(df),
    target_year="2024年度",
    used_columns=required_cols,
    created_metrics=["営業利益率", "当期純利益率", "利益率差分"],
    missing_strategy="警告のみ（除外なし）"
)

# データ品質レポート
quality_report = report.create_data_quality_report(df, required_cols)

# 各指標の上位下位のフォーマット
margin_gap_summary = report.format_top_bottom_summary(margin_gap_top_bottom, "利益率差分")
opm_summary = report.format_top_bottom_summary(opm_top_bottom, "営業利益率")
npm_summary = report.format_top_bottom_summary(npm_top_bottom, "当期純利益率")

# Markdownレポート作成
sections = [
    {
        "heading": "分析メタデータ",
        "content": metadata
    },
    {
        "heading": "データ品質",
        "content": quality_report
    },
    {
        "heading": "利益率の基本統計量",
        "content": summary_stats.to_markdown(index=False)
    },
    {
        "heading": "利益率差分（営業利益率 - 当期純利益率）",
        "content": margin_gap_summary
    },
    {
        "heading": "営業利益率",
        "content": opm_summary
    },
    {
        "heading": "当期純利益率",
        "content": npm_summary
    },
    {
        "heading": "観察事項",
        "content": observations
    }
]

report.create_markdown_summary(
    output_path=str(project_root / "output" / "02_summary.md"),
    title="分析②：収益性分解",
    sections=sections
)

# %%[markdown]
# ### 分析完了
# 
# 以下のファイルが生成されました：
# - `output/02_margin_table.xlsx`: 利益率のテーブル
# - `output/fig_02_margins_by_company.png`: 企業別利益率の棒グラフ
# - `output/fig_02_opm_vs_npm.png`: 営業利益率vs当期純利益率の散布図
# - `output/02_summary.md`: サマリーレポート

# %%
print("\n=== 分析②完了 ===")
print("生成されたファイル:")
print("output/02_margin_table.xlsx")
print("output/fig_02_margins_by_company.png")
print("output/fig_02_opm_vs_npm.png")
print("output/02_summary.md")

# %%

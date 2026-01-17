# %%[markdown]
# # 分析③：資本効率・財務体質（ROA/ROE/自己資本比率）
#
# ## 目的
# - ROAとROEの関係、および自己資本比率から、資本構造の違い（レバレッジ構造）を把握する
#
# ## 仮説
# - **H3-1**: ROEが高いがROAが相対的に低い企業は、自己資本比率が低めでレバレッジによりROEが押し上げられている
# - **H3-2**: ROAとROEが両方高い企業は、収益性と資産効率が両立しており「質の高い高効率」とみなせる
#
# ## 実施内容
# 1. データの読み込みと前処理
# 2. 自己資本比率、ROE-ROA差分の計算
# 3. ROA vs ROE、自己資本比率 vs ROEの散布図作成
# 4. レバレッジ効果の可視化
# 5. 象限別の企業分類

# %%
import sys
from pathlib import Path

import japanize_matplotlib  # noqa: F401
import matplotlib.pyplot as plt

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from modules import io, metrics, preprocess, report, viz

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
required_cols = ["証券コード", "企業名", "売上高", "総資産", "自己資本", "ROA", "ROE"]

preprocess.validate_columns(df, required_cols)

# %%
# 数値列の型変換
numeric_cols = ["売上高", "総資産", "自己資本", "ROA", "ROE"]
df = preprocess.coerce_numeric(df, numeric_cols)

# %%
# ROA/ROEがパーセント表記かどうかを確認して正規化
df = metrics.normalize_percentage_columns(df, ["ROA", "ROE"])

# %%
# 欠損値の確認
df = preprocess.handle_missing(df, strategy="warn", subset=required_cols)
print(df.isna().sum())

# %%[markdown]
# ### 財務指標の計算
# - 自己資本比率、総資産回転率、ROE-ROA差分を計算

# %%
# 財務指標を追加
print(df.shape)
df = metrics.add_financial_ratios(df)

# 資本効率関連の列を確認
capital_cols = ["自己資本比率", "総資産回転率", "ROE_ROA_gap"]
print("\n資本効率関連の指標:")
for col in capital_cols:
    if col in df.columns:
        print(f"  - {col}: {df[col].notna().sum()}件")
print(df.shape)
# %%
# 資本効率テーブルを保存
output_cols = [
    "証券コード",
    "企業名",
    "総資産",
    "自己資本",
    "ROA",
    "ROE",
    "自己資本比率",
    "総資産回転率",
    "ROE_ROA_gap",
]

df_output = df[output_cols].copy()

# ROEで降順ソート
df_output = df_output.sort_values("ROE", ascending=False, na_position="last")

io.save_table(df_output, str(project_root / "output" / "03_capital_table.xlsx"))

# %%[markdown]
# ### 可視化①：ROA vs ROE
# - 資産効率と株主資本効率の関係を可視化
# - 45度線を引いて、レバレッジ効果を確認

# %%
# 散布図の作成
fig, ax = plt.subplots(figsize=(10, 8))

# データのプロット
df_plot = df[["企業名", "ROA", "ROE"]].dropna()
ax.scatter(df_plot["ROA"], df_plot["ROE"], alpha=0.6, s=100)

# 45度線を追加（ROE = ROAの線）
max_val = max(df_plot["ROA"].max(), df_plot["ROE"].max())
min_val = min(df_plot["ROA"].min(), df_plot["ROE"].min())
ax.plot([min_val, max_val], [min_val, max_val], "r--", alpha=0.5, label="ROE = ROA")

# 企業名ラベル
for _, row in df_plot.iterrows():
    ax.annotate(
        row["企業名"],
        (row["ROA"], row["ROE"]),
        xytext=(5, 5),
        textcoords="offset points",
        fontsize=8,
        alpha=0.7,
    )

# ラベルとタイトル
ax.set_xlabel("ROA（%）", fontsize=12)
ax.set_ylabel("ROE（%）", fontsize=12)
ax.set_title("ROAとROEの関係", fontsize=14, fontweight="bold")
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(
    project_root / "output" / "fig_03_roa_vs_roe.png", dpi=150, bbox_inches="tight"
)
plt.show()

print("散布図を保存しました: output/fig_03_roa_vs_roe.png")

# %%[markdown]
# ### 可視化②：自己資本比率 vs ROE
# - 財務体質（自己資本比率）と株主資本効率の関係を可視化

# %%
viz.create_scatter_plot(
    df=df,
    x_col="自己資本比率",
    y_col="ROE",
    output_path=str(project_root / "output" / "fig_03_equity_ratio_vs_roe.png"),
    title="自己資本比率とROEの関係",
    x_label="自己資本比率（%）",
    y_label="ROE（%）",
    annotate=True,
    company_col="企業名",
)

# %%[markdown]
# ### 可視化③：自己資本比率 vs ROE-ROA差分（レバレッジ効果）
# - 財務レバレッジの影響を可視化

# %%
viz.create_scatter_plot(
    df=df,
    x_col="自己資本比率",
    y_col="ROE_ROA_gap",
    output_path=str(project_root / "output" / "fig_03_equity_vs_roe_roa_gap.png"),
    title="自己資本比率とROE-ROA差分の関係",
    x_label="自己資本比率（%）",
    y_label="ROE - ROA（ポイント）",
    annotate=True,
    company_col="企業名",
)

# %%[markdown]
# ### 可視化④：ROA分解（総資産回転率 vs 当期純利益率）
# - ROAを分解して、回転率と利益率の関係を確認
# - ROA ≈ 当期純利益率 × 総資産回転率

# %%
# 当期純利益率がない場合は計算
if "当期純利益率" not in df.columns:
    df = metrics.add_financial_ratios(df)

viz.create_scatter_plot(
    df=df,
    x_col="総資産回転率",
    y_col="当期純利益率",
    output_path=str(project_root / "output" / "fig_03_turnover_vs_margin.png"),
    title="総資産回転率と当期純利益率の関係（ROA分解）",
    x_label="総資産回転率（回）",
    y_label="当期純利益率（%）",
    annotate=True,
    company_col="企業名",
)

# %%[markdown]
# ### 象限別の企業分類
# - ROAとROEの中央値を基準に4象限に分類

# %%
# ROAとROEの中央値
roa_median = df["ROA"].median()
roe_median = df["ROE"].median()

print(f"\nROA中央値: {roa_median:.2f}%")
print(f"ROE中央値: {roe_median:.2f}%")

# 象限分類
df["象限"] = "その他"
df.loc[(df["ROA"] >= roa_median) & (df["ROE"] >= roe_median), "象限"] = "高ROA・高ROE"
df.loc[(df["ROA"] >= roa_median) & (df["ROE"] < roe_median), "象限"] = "高ROA・低ROE"
df.loc[(df["ROA"] < roa_median) & (df["ROE"] >= roe_median), "象限"] = "低ROA・高ROE"
df.loc[(df["ROA"] < roa_median) & (df["ROE"] < roe_median), "象限"] = "低ROA・低ROE"

# 象限別の企業数
quadrant_counts = df["象限"].value_counts()
print("\n象限別の企業数:")
print(quadrant_counts)

# 各象限の企業リスト
print("\n=== 象限別企業リスト ===")
for quadrant in ["高ROA・高ROE", "高ROA・低ROE", "低ROA・高ROE", "低ROA・低ROE"]:
    companies = df[df["象限"] == quadrant]["企業名"].tolist()
    if companies:
        print(f"\n{quadrant}:")
        for company in companies:
            print(f"  - {company}")

# %%[markdown]
# ### 各指標の上位・下位企業
# - ROA、ROE、自己資本比率、ROE-ROA差分の上位・下位を特定

# %%
# ROAの上位・下位
roa_top_bottom = metrics.get_top_bottom_companies(df, "ROA", n=3)
print("\n=== ROA 上位3社 ===")
for company, value in roa_top_bottom["top"]:
    print(f"  {company}: {value:.2f}%")

print("\n=== ROA 下位3社 ===")
for company, value in roa_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}%")

# %%
# ROEの上位・下位
roe_top_bottom = metrics.get_top_bottom_companies(df, "ROE", n=3)
print("\n=== ROE 上位3社 ===")
for company, value in roe_top_bottom["top"]:
    print(f"  {company}: {value:.2f}%")

print("\n=== ROE 下位3社 ===")
for company, value in roe_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}%")

# %%
# 自己資本比率の上位・下位
equity_top_bottom = metrics.get_top_bottom_companies(df, "自己資本比率", n=3)
print("\n=== 自己資本比率 上位3社 ===")
for company, value in equity_top_bottom["top"]:
    print(f"  {company}: {value:.2f}%")

print("\n=== 自己資本比率 下位3社 ===")
for company, value in equity_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}%")

# %%
# ROE-ROA差分の上位・下位
gap_top_bottom = metrics.get_top_bottom_companies(df, "ROE_ROA_gap", n=3)
print("\n=== ROE-ROA差分 上位3社（レバレッジ効果大）===")
for company, value in gap_top_bottom["top"]:
    print(f"  {company}: {value:.2f}")

print("\n=== ROE-ROA差分 下位3社 ===")
for company, value in gap_top_bottom["bottom"]:
    print(f"  {company}: {value:.2f}")

# %%[markdown]
# ---
# ### 結果の確認と考察の記載
#
# **ここで一旦、以下の結果を確認してください：**
# - 生成されたグラフ（ROA vs ROE、象限分析、自己資本比率、総資産回転率）
# - ROA、ROE、自己資本比率、ROE-ROA差分の上位・下位企業
# - 象限別の企業分類

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
#     "### ① ROA vs ROEの関係",
#     "観察事項１",
#     "観察事項２",
#     "### ② 象限分析",
#     "観察事項３",
# ]
observations = [
    "#### 可視化①：ROA vs ROE",
    "ROAとROEには概ね正の関係が見られ、資産効率の高い企業ほどROEも高い傾向が確認された",
    "45度線から大きく上方に乖離する企業では、ROAに比してROEが高く、レバレッジの影響が大きい",
    "中外製薬・第一三共はROA・ROEともに高く、事業収益力主導の高ROE構造",
    "住友ファーマはROAが低い一方でROEが高く、財務レバレッジ依存が強い",
    "#### 可視化②：自己資本比率 vs ROE",
    "自己資本比率とROEの間に単純な相関関係は見られない",
    "自己資本比率が低い企業ではROEが高く出やすい一方、自己資本比率が高い企業ではROEは相対的に抑制される",
    "ROEの水準は自己資本の厚みそのものより、ROAとレバレッジの組み合わせによって決まっている",
    "#### 可視化③：自己資本比率 vs ROE-ROA差分（レバレッジ効果）",
    "自己資本比率が低いほどROE-ROA差分が大きく、レバレッジ効果が強い傾向が確認された",
    "住友ファーマ、第一三共はROE-ROA差分が大きく、レバレッジによるROE押し上げが顕著",
    "塩野義製薬、日本新薬、中外製薬は自己資本が厚く、差分は相対的に小さい",
    "小林製薬はROE-ROA差分がマイナスで、レバレッジ効果が機能していない",
    "#### 象限別の企業分類（ROA×ROE）",
    "**高ROA・高ROE**：第一三共、中外製薬、塩野義製薬、日本新薬、参天製薬、ロート製薬",
    "**高ROA・低ROE**：小林製薬",
    "**低ROA・高ROE**：住友ファーマ",
    "**低ROA・低ROE**：武田薬品工業、アステラス製薬、エーザイ、小野薬品工業、ツムラ",
    "#### 各指標の上位・下位企業",
    "ROA上位は中外製薬、日本新薬、参天製薬で、事業収益力の高さが際立つ",
    "ROA下位はアステラス製薬、武田薬品工業、住友ファーマで、資産効率に課題が見られる",
    "ROE上位には中外製薬、第一三共、住友ファーマが含まれ、構造の違いが混在",
    "自己資本比率上位は塩野義製薬、日本新薬、中外製薬で、財務健全性が高い",
    "ROE-ROA差分上位は住友ファーマ、第一三共で、レバレッジ効果が大きい一方、小林製薬は差分がマイナス",
]

# %%[markdown]
# ### サマリーレポートの作成
# - 分析結果をMarkdown形式でまとめる

# %%
# 基本統計量
summary_stats = report.generate_summary_stats(
    df, cols=["ROA", "ROE", "自己資本比率", "総資産回転率", "ROE_ROA_gap"]
)

# メタデータ
metadata = report.create_analysis_metadata(
    data_path=data_path.name,  # ファイル名のみを使用
    n_companies=len(df),
    target_year="2024年度",
    used_columns=required_cols,
    created_metrics=["自己資本比率", "総資産回転率", "ROE_ROA_gap"],
    missing_strategy="警告のみ（除外なし）",
)

# データ品質レポート
quality_report = report.create_data_quality_report(df, required_cols)

# 各指標の上位下位のフォーマット
roa_summary = report.format_top_bottom_summary(roa_top_bottom, "ROA")
roe_summary = report.format_top_bottom_summary(roe_top_bottom, "ROE")
equity_summary = report.format_top_bottom_summary(equity_top_bottom, "自己資本比率")
gap_summary = report.format_top_bottom_summary(gap_top_bottom, "ROE-ROA差分")

# 象限別企業リスト
quadrant_summary = []
for quadrant in ["高ROA・高ROE", "高ROA・低ROE", "低ROA・高ROE", "低ROA・低ROE"]:
    companies = df[df["象限"] == quadrant]["企業名"].tolist()
    if companies:
        quadrant_summary.append(f"#### {quadrant}")
        for company in companies:
            quadrant_summary.append(f"{company}")
        quadrant_summary.append("")
print(quadrant_summary)

# Markdownレポート作成
sections = [
    {"heading": "分析メタデータ", "content": metadata},
    {"heading": "データ品質", "content": quality_report},
    {
        "heading": "資本効率指標の基本統計量",
        "content": summary_stats.to_markdown(index=False),
    },
    {"heading": "象限別企業分類", "content": quadrant_summary},
    {"heading": "ROA", "content": roa_summary},
    {"heading": "ROE", "content": roe_summary},
    {"heading": "自己資本比率", "content": equity_summary},
    {"heading": "ROE-ROA差分（レバレッジ効果）", "content": gap_summary},
    {"heading": "観察事項", "content": observations},
]

report.create_markdown_summary(
    output_path=str(project_root / "output" / "03_summary.md"),
    title="分析③：資本効率・財務体質（ROA/ROE/自己資本比率）",
    sections=sections,
)

# %%[markdown]
# ### 分析完了
#
# 以下のファイルが生成されました：
# - `output/03_capital_table.xlsx`: 資本効率のテーブル
# - `output/fig_03_roa_vs_roe.png`: ROA vs ROEの散布図
# - `output/fig_03_equity_ratio_vs_roe.png`: 自己資本比率 vs ROEの散布図
# - `output/fig_03_equity_vs_roe_roa_gap.png`: レバレッジ効果の可視化
# - `output/fig_03_turnover_vs_margin.png`: ROA分解図
# - `output/03_summary.md`: サマリーレポート

# %%
print("\n=== 分析③完了 ===")
print("生成されたファイル:")
print("output/03_capital_table.xlsx")
print("output/fig_03_roa_vs_roe.png")
print("output/fig_03_equity_ratio_vs_roe.png")
print("output/fig_03_equity_vs_roe_roa_gap.png")
print("output/fig_03_turnover_vs_margin.png")
print("output/03_summary.md")

# %%

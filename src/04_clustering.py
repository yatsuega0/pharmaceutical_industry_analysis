# %%[markdown]
# # 分析④：業界構造の類型化（クラスタリング）
#
# ## 目的
# - 13社を複数の型（ビジネス/財務特性）に分類し、業界構造を説明可能にする
#
# ## 仮説
# - **H4-1**: 製薬業界内に「大手安定（規模大・効率中）」「高効率中堅（効率高）」「高ROEレバレッジ型（ROE高・自己資本比率低）」のような複数クラスターが存在する
# - **H4-2**: クラスター間で、利益率・資産回転・自己資本比率の組合せが系統的に異なる
#
# ## 実施内容
# 1. データの読み込みと前処理
# 2. クラスタリング用の変数セット作成
# 3. k-meansクラスタリング（k=2〜4で試行）
# 4. PCAによる2次元可視化
# 5. クラスター別プロファイルの作成

# %%
import sys
from pathlib import Path

import japanize_matplotlib  # noqa: F401
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from modules import io, metrics, preprocess, report

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
    "証券コード",
    "企業名",
    "売上高",
    "営業利益",
    "当期純利益",
    "総資産",
    "自己資本",
    "ROA",
    "ROE",
]

preprocess.validate_columns(df, required_cols)

# %%
# 数値列の型変換
numeric_cols = ["売上高", "営業利益", "当期純利益", "総資産", "自己資本", "ROA", "ROE"]
df = preprocess.coerce_numeric(df, numeric_cols)

# %%
# ROA/ROEがパーセント表記かどうかを確認して正規化
df = metrics.normalize_percentage_columns(df, ["ROA", "ROE"])

# %%
print(df.shape)
# 財務指標を追加
df = metrics.add_financial_ratios(df)

print("\n追加された指標:")
new_cols = ["営業利益率", "当期純利益率", "自己資本比率", "総資産回転率"]
for col in new_cols:
    if col in df.columns:
        print(f"  - {col}")
print(df.shape)
df.head(2)
# %%[markdown]
# ### クラスタリング用の変数セット作成
# - 規模：売上高（log変換）
# - 収益性：営業利益率、当期純利益率
# - 効率：ROA、ROE
# - 体質：自己資本比率

# %%
# クラスタリングに使用する列
clustering_cols = [
    # 事業規模（対数変換）
    "売上高_log",
    # 収益性
    "営業利益率",
    "当期純利益率",
    # 効率
    "ROA",
    "ROE",
    # 財務体質
    "自己資本比率",
]

# 売上高の対数変換
df["売上高_log"] = np.log10(df["売上高"])

# 欠損値を含む行を除外（クラスタリング対象外）
df_cluster = df[["企業名"] + clustering_cols].dropna()

print(f"\nクラスタリング対象企業数: {len(df_cluster)}社")
print(f"除外された企業数: {len(df) - len(df_cluster)}社")
print(df.isna().sum())
# %%
# 標準化（z-score）
scaler = StandardScaler()
X = df_cluster[clustering_cols].values
X_scaled = scaler.fit_transform(X)

print(f"\n標準化後のデータ形状: {X_scaled.shape}")

# %%[markdown]
# ### 最適なクラスター数の決定
# - k=2〜4でシルエットスコアを計算

# %%
# k=2〜4でシルエットスコアを計算
silhouette_scores = {}

print("\n=== シルエットスコア ===")
for k in range(2, 5):
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_scaled)
    score = silhouette_score(X_scaled, labels)
    silhouette_scores[k] = score
    print(f"k={k}: {score:.4f}")

# 最適なクラスター数を選択（シルエットスコアが最大のもの）
optimal_k = max(silhouette_scores, key=silhouette_scores.get)
print(f"\n最適なクラスター数: k={optimal_k}")

# %%[markdown]
# ### k-meansクラスタリングの実行
# - 最適なクラスター数でクラスタリングを実行

# %%
# 最適なkでクラスタリング
kmeans = KMeans(n_clusters=optimal_k, random_state=42, n_init=10)
df_cluster["クラスター"] = kmeans.fit_predict(X_scaled)

# クラスター番号を1始まりに変更
df_cluster["クラスター"] = df_cluster["クラスター"] + 1

print(f"\nクラスタリング完了（k={optimal_k}）")
print("\nクラスター別企業数:")
print(df_cluster["クラスター"].value_counts().sort_index())

# %%
# 各クラスターの企業リスト
print("\n=== クラスター別企業リスト ===")
for cluster_id in sorted(df_cluster["クラスター"].unique()):
    companies = df_cluster[df_cluster["クラスター"] == cluster_id]["企業名"].tolist()
    print(f"\nクラスター{cluster_id}（{len(companies)}社）:")
    for company in companies:
        print(f"  - {company}")

# %%
# 元のデータフレームにクラスター情報をマージ
print(df.shape)
df = df.merge(df_cluster[["企業名", "クラスター"]], on="企業名", how="left")
print(df.shape)
# %%[markdown]
# ### PCAによる2次元可視化
# - 主成分分析で2次元に落として、クラスターを可視化

# %%
# PCAで2次元に圧縮
pca = PCA(n_components=2, random_state=42)
X_pca = pca.fit_transform(X_scaled)

# 寄与率
print(f"\n第1主成分の寄与率: {pca.explained_variance_ratio_[0]:.4f}")
print(f"第2主成分の寄与率: {pca.explained_variance_ratio_[1]:.4f}")
print(f"累積寄与率: {pca.explained_variance_ratio_.sum():.4f}")

# %%
# PCA散布図の作成
fig, ax = plt.subplots(figsize=(12, 8))

# クラスターごとに色分け
for cluster_id in sorted(df_cluster["クラスター"].unique()):
    mask = df_cluster["クラスター"] == cluster_id
    ax.scatter(
        X_pca[mask, 0],
        X_pca[mask, 1],
        label=f"クラスター{cluster_id}",
        alpha=0.7,
        s=100,
    )

# 企業名ラベル
for i, company in enumerate(df_cluster["企業名"]):
    ax.annotate(
        company,
        (X_pca[i, 0], X_pca[i, 1]),
        xytext=(5, 5),
        textcoords="offset points",
        fontsize=8,
        alpha=0.7,
    )

# ラベルとタイトル
ax.set_xlabel(
    f"第1主成分（寄与率: {pca.explained_variance_ratio_[0]:.2%}）", fontsize=12
)
ax.set_ylabel(
    f"第2主成分（寄与率: {pca.explained_variance_ratio_[1]:.2%}）", fontsize=12
)
ax.set_title("PCAによるクラスター可視化", fontsize=14, fontweight="bold")
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(
    project_root / "output" / "fig_04_pca_clusters.png", dpi=150, bbox_inches="tight"
)
plt.show()

print("PCA散布図を保存しました: output/fig_04_pca_clusters.png")

# %%[markdown]
# ### クラスター別プロファイルの作成
# - 各クラスターの特徴量の平均値を計算

# %%
# 科学的記数法を無効化
pd.set_option("display.float_format", "{:.2f}".format)

# クラスター別の平均値
cluster_profile = df.groupby("クラスター")[
    [
        "売上高",
        "総資産",
        "営業利益率",
        "当期純利益率",
        "ROA",
        "ROE",
        "自己資本比率",
        "総資産回転率",
    ]
].mean()

print("\n=== クラスター別プロファイル（平均値）===")
print(cluster_profile.round(2))

# プロファイルをXLSXで保存
io.save_table(
    cluster_profile.reset_index(),
    str(project_root / "output" / "04_cluster_profile.xlsx"),
)

# %%[markdown]
# ### クラスター別プロファイルの可視化（レーダーチャート）
# - 標準化されたデータでクラスターの特徴を可視化

# %%
# レーダーチャート用のデータ準備
profile_cols = [
    "営業利益率",
    "当期純利益率",
    "ROA",
    "ROE",
    "自己資本比率",
    "総資産回転率",
]

# クラスター別の平均値（元のdfを使用、欠損値を除外）
cluster_profile_scaled = (
    df.dropna(subset=["クラスター"]).groupby("クラスター")[profile_cols].mean()
)

# 標準化（各列を0-1にスケール）
for col in profile_cols:
    col_min = cluster_profile_scaled[col].min()
    col_max = cluster_profile_scaled[col].max()
    if col_max > col_min:
        cluster_profile_scaled[col] = (cluster_profile_scaled[col] - col_min) / (
            col_max - col_min
        )
    else:
        cluster_profile_scaled[col] = 0.5

# %%
# レーダーチャートの作成
fig = plt.figure(figsize=(10, 10))
ax = fig.add_subplot(111, projection="polar")

# 角度の設定
angles = np.linspace(0, 2 * np.pi, len(profile_cols), endpoint=False).tolist()
angles += angles[:1]  # 最初の点に戻る

# クラスターごとにプロット
for cluster_id in sorted(cluster_profile_scaled.index):
    values = cluster_profile_scaled.loc[cluster_id].tolist()
    values += values[:1]  # 最初の点に戻る

    ax.plot(angles, values, "o-", linewidth=2, label=f"クラスター{cluster_id}")
    ax.fill(angles, values, alpha=0.15)

# ラベルの設定
ax.set_xticks(angles[:-1])
ax.set_xticklabels(profile_cols, fontsize=10)
ax.set_ylim(0, 1)
ax.set_title(
    "クラスター別プロファイル（正規化済み）", fontsize=14, fontweight="bold", pad=20
)
ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
ax.grid(True)

plt.tight_layout()
plt.savefig(
    project_root / "output" / "fig_04_cluster_profile.png", dpi=150, bbox_inches="tight"
)
plt.show()

print("レーダーチャートを保存しました: output/fig_04_cluster_profile.png")

# %%[markdown]
# ### クラスター割り当てテーブルの保存
# - 企業とクラスターの対応表を保存

# %%
# クラスター割り当てテーブル
cluster_assignments = df[
    [
        "証券コード",
        "企業名",
        "クラスター",
        "売上高",
        "総資産",
        "営業利益率",
        "当期純利益率",
        "ROA",
        "ROE",
        "自己資本比率",
    ]
].copy()

# クラスター番号でソート
cluster_assignments = cluster_assignments.sort_values(
    ["クラスター", "売上高"], ascending=[True, False]
)

io.save_table(
    cluster_assignments, str(project_root / "output" / "04_cluster_assignments.xlsx")
)

# %%[markdown]
# ### クラスターの解釈
# - 各クラスターの特徴を記述

# %%
# クラスター別の特徴を抽出
cluster_interpretations = []

for cluster_id in sorted(df_cluster["クラスター"].unique()):
    cluster_data = df[df["クラスター"] == cluster_id]
    n_companies = len(cluster_data)

    # 平均値
    avg_sales = cluster_data["売上高"].mean()
    avg_roa = cluster_data["ROA"].mean()
    avg_roe = cluster_data["ROE"].mean()
    avg_opm = cluster_data["営業利益率"].mean()
    avg_equity_ratio = cluster_data["自己資本比率"].mean()

    interpretation = {
        "クラスター": f"クラスター{cluster_id}",
        "企業数": n_companies,
        "特徴": f"売上高平均{avg_sales:.1f}百万円、ROA{avg_roa:.1f}%、ROE{avg_roe:.1f}%、営業利益率{avg_opm:.1f}%、自己資本比率{avg_equity_ratio:.1f}%",
    }

    cluster_interpretations.append(interpretation)

print("\n=== クラスター解釈 ===")
for interp in cluster_interpretations:
    print(f"\n{interp['クラスター']}（{interp['企業数']}社）:")
    print(f"  {interp['特徴']}")

# %%[markdown]
# ---
# ### 結果の確認と考察の記載
#
# **ここで一旦、以下の結果を確認してください：**
# - 生成されたグラフ（デンドログラム、クラスター別散布図、特徴量の箱ひげ図）
# - 各クラスターの特徴と属する企業

# %%
# TODO: グラフとクラスター解釈を確認した上で、観察された事実を記載してください
#
# 記載方法の注意：
# - 各文字列の後に必ずカンマ(,)を付ける
# - セクション間に空行を入れたい場合は空文字列 "" を追加
# - Markdown形式で記載可能（**太字**、箇条書きなど）
#
# 記載例：
# observations = [
#     "### ① クラスター0の特徴",
#     "観察事項１",
#     "観察事項２",
#     "### ② クラスター間の比較",
#     "観察事項３",
# ]
observations = [
    "#### 利益率（営業利益率・当期純利益率）",
    "クラスター3は営業利益率・当期純利益率ともに突出して高く、高付加価値型の収益構造を示す。",
    "クラスター2は中程度だが安定した利益率を維持しており、堅実な収益モデルと解釈できる。",
    "クラスター1は規模の割に利益率が中位にとどまる。",
    "クラスター4は利益率が低く、投資負担や構造改革段階の影響が示唆される。",
    "#### 資産効率（ROA・総資産回転率）",
    "ROAはクラスター3が最も高く、次いでクラスター2が続く。一方、総資産回転率はクラスター2が最も高い。",
    "高ROAは必ずしも総資産回転率の高さによるものではなく、クラスター3では利益率が相対的に高く、クラスター2では総資産回転率が相対的に高いという形で、ROAの構成要素に差がみられる。",
    "#### 株主資本効率（ROE）",
    "クラスター3は高ROAと高自己資本比率の組合せにより高ROEを実現している。",
    "クラスター1は自己資本比率の低さによるレバレッジ効果を通じて高ROEを達成している。",
    "クラスター4はROAが低く、ROEも低水準にとどまる。  ",
    "#### 財務体質（自己資本比率）",
    "クラスター3およびクラスター2は自己資本比率が高く、財務的に極めて安定している。",
    "クラスター1は自己資本比率が低く、レバレッジを活用した資本構成である。",
    "クラスター4は中程度の自己資本比率に位置する。",
    "#### PCAによるクラスター構造",
    "第1主成分は収益性・効率性を総合的に表す軸、第2主成分は規模や投資フェーズの違いを反映する軸と解釈できる。",
    "クラスターはPCA空間上でも明確に分離しており、財務構造の差異は統計的にも一貫している。",
    "#### 仮説に対する評価",
    "H4-1（製薬業界内に複数の財務・ビジネスモデル型クラスターが存在する）は支持される。",
    "高収益型、高回転安定型、レバレッジ型、規模大・低効率型といった異なる類型が確認された。",
    "H4-2（クラスター間で、利益率・資産回転率・自己資本比率の組合せが系統的に異なる）は強く支持される。",
    "同一のROE水準であっても、その達成メカニズムはクラスターごとに明確に異なっている。",
]

# %%[markdown]
# ### サマリーレポートの作成
# - 分析結果をMarkdown形式でまとめる

# %%
# メタデータ
metadata = report.create_analysis_metadata(
    data_path=data_path.name,  # ファイル名のみを使用
    n_companies=len(df_cluster),
    target_year="2024年度",
    used_columns=required_cols + clustering_cols,
    created_metrics=["クラスター番号"],
    missing_strategy="欠損値を含む行を除外",
)

# %%
# データ品質レポート
quality_report = report.create_data_quality_report(df, required_cols)

# クラスター別企業リスト
cluster_summary = []
for cluster_id in sorted(df_cluster["クラスター"].unique()):
    companies = df_cluster[df_cluster["クラスター"] == cluster_id]["企業名"].tolist()
    cluster_summary.append(f"**クラスター{cluster_id}（{len(companies)}社）**")
    for company in companies:
        cluster_summary.append(f"{company}")
    cluster_summary.append("")

# クラスター解釈
interpretation_summary = []
for interp in cluster_interpretations:
    interpretation_summary.append(f"**{interp['クラスター']}**")
    interpretation_summary.append(f"{interp['特徴']}")
    interpretation_summary.append("")

# Markdownレポート作成
sections = [
    {"heading": "分析メタデータ", "content": metadata},
    {"heading": "データ品質", "content": quality_report},
    {
        "heading": "クラスタリング結果",
        "content": [
            f"最適なクラスター数: k={optimal_k}",
            f"シルエットスコア: {silhouette_scores[optimal_k]:.4f}",
            f"第1主成分の寄与率: {pca.explained_variance_ratio_[0]:.2%}",
            f"第2主成分の寄与率: {pca.explained_variance_ratio_[1]:.2%}",
            f"累積寄与率: {pca.explained_variance_ratio_.sum():.2%}",
        ],
    },
    {"heading": "クラスター別企業リスト", "content": cluster_summary},
    {"heading": "クラスター解釈", "content": interpretation_summary},
    {
        "heading": "クラスター別プロファイル",
        "content": report.format_table_for_markdown(
            cluster_profile, large_number_cols=["売上高", "総資産"]
        ),
    },
    {"heading": "観察事項", "content": observations},
]

report.create_markdown_summary(
    output_path=str(project_root / "output" / "04_summary.md"),
    title="分析④：業界構造の類型化（クラスタリング）",
    sections=sections,
)

# 表示オプションをデフォルトに戻す
pd.reset_option("display.float_format")

# %%[markdown]
# ### 分析完了
#
# 以下のファイルが生成されました：
# - `output/04_cluster_assignments.xlsx`: クラスター割り当てテーブル
# - `output/04_cluster_profile.xlsx`: クラスター別プロファイル
# - `output/fig_04_pca_clusters.png`: PCA散布図
# - `output/fig_04_cluster_profile.png`: レーダーチャート
# - `output/04_summary.md`: サマリーレポート

# %%
print("\n=== 分析④完了 ===")
print("生成されたファイル:")
print("output/04_cluster_assignments.xlsx")
print("output/04_cluster_profile.xlsx")
print("output/fig_04_pca_clusters.png")
print("output/fig_04_cluster_profile.png")
print("output/04_summary.md")
# %%

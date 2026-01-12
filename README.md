# 製薬業界（13社・2024年度）財務データ分析プロジェクト

製薬業界13社の2024年度決算データを用いた財務分析プロジェクトです。企業のポジショニング、収益性の質、資本効率、業界構造の類型化を行います。

## プロジェクト概要

### 目的
1. **業界内でのポジショニング**：規模 × 収益性/効率で企業の立ち位置を可視化
2. **収益性の質**：本業収益力（営業利益）と最終利益（純利益）の差を分析
3. **資本効率と財務体質**：ROA/ROE/自己資本比率/レバレッジ構造の整理
4. **業界構造の類型化**：クラスタリングによる企業タイプの分類

### 対象データ
- **データファイル**: `input/財務データ_製薬業界.xlsx`
- **対象企業**: 製薬業界13社
- **対象年度**: 2024年度決算
- **主な指標**: 売上高、営業利益、当期純利益、総資産、自己資本、ROA、ROE

---

## プロジェクト構成

```
pharmaceutical_industry_analysis/
├── input/
│   └── 財務データ_製薬業界.xlsx  # 入力データ（ユーザーが準備）
├── output/                        # 分析結果の出力先
│
├── src/                           # 分析スクリプト
│   ├── 01_positioning.py          # ポジショニング分析
│   ├── 02_profitability_decompose.py  # 収益性分解
│   ├── 03_capital_efficiency.py   # 資本効率分析
│   └── 04_clustering.py           # クラスタリング分析
├── modules/                       # 共通モジュール
│   ├── __init__.py
│   ├── io.py                      # データ入出力
│   ├── preprocess.py              # 前処理
│   ├── metrics.py                 # 指標計算
│   ├── viz.py                     # 可視化
│   └── report.py                  # レポート生成
├── main.py                        # エントリーポイント（全分析の一括実行）
├── pyproject.toml                 # 依存パッケージ管理
├── README.md                      # 本ファイル
└── プロンプト用_分析計画.md        # 詳細な分析計画書（AIコーディング用）
```

---

## セットアップ

### 1. 環境構築

このプロジェクトはPoetryで環境を管理します。

```bash
# Poetryで依存パッケージをインストール
poetry install

# Poetry環境を有効化
poetry shell
```

### 2. Jupyter環境の設定（オプション）

VS CodeやJupyterLabでインタラクティブに実行する場合は、カーネルを設定してください。

```bash
# JupyterLabの起動
poetry run jupyter lab
```

### 3. データの配置

`input/財務データ_製薬業界.xlsx` にExcel形式の財務データを配置してください。

必須列：
- 証券コード
- 企業名
- 売上高
- 営業利益
- 当期純利益
- 総資産
- 自己資本
- ROA
- ROE

---

## 分析の実行

### コマンドラインでの実行

`main.py` をエントリーポイントとして、全分析を順番に実行します。

```bash
# 全分析を順番に実行
python main.py

# エラー時に停止するモード
python main.py --stop-on-error

# 特定の分析のみ実行
python main.py --script 1  # ポジショニング分析のみ
python main.py --script 2  # 収益性分解のみ
python main.py --script 3  # 資本効率分析のみ
python main.py --script 4  # クラスタリング分析のみ
```

### Interactive Window（VS Code）での実行

VS CodeのInteractive Windowを使用して、段階的に分析を実行できます。

#### 方法1：main.pyでの一括実行

1. VS Codeで `main.py` を開く
2. 各セルを順番に `Shift + Enter` で実行して関数を定義
3. ファイル最後の実行用セルのコメントを外して実行：
   ```python
   run_all_analyses()  # 全分析を実行
   # または
   run_single_analysis(1)  # 特定の分析のみ
   ```

#### 方法2：個別スクリプトでの実行

各分析を個別に詳細確認しながら実行する場合：

1. VS Codeで `src/0X_*.py` を開く
2. セル単位で `Shift + Enter` で実行
3. 各セルの結果を確認しながら進める
4. 最後のセルで観察事項を記載して出力

### 観察事項・考察の記載について

各分析スクリプトは、**分析者自身が結果を確認した上で観察事項や考察をまとめる**設計になっています。

**記載項目**:
- 観察された特徴（外れ値、上位/下位企業、象限別の傾向など）
- データから読み取れる考察
- 仮説の検証結果
- 深掘りすべき追加分析項目の提案

**記載方法**:

各スクリプトの最後のセルに、`observations`変数を編集して観察事項を記載します。

```python
# 観察事項を記載（分析者が編集する部分）
observations = [
    "### 主要な発見",
    "企業Aは総資産が最大規模だが、ROAは業界平均（5.2%）を下回る3.8%に留まる。",
    "企業B、Cは中堅規模（総資産1-2兆円）ながらROA 7%超を実現している。",

    "### 考察",
    "規模拡大と資産効率が必ずしも比例せず、中堅企業の高効率性は研究開発の効率性を示唆。",
    "**深掘り点**: 固定資産・研究開発資産の構成比がROA差に与える影響を検証する必要がある。"
]
```

この`observations`変数は、スクリプト内で自動的に`output/XX_summary.md`に出力されます。

**重要**: 観察事項は、グラフや数値を確認した分析者が**主体的に記載**することで、分析の価値が高まります。

---

## 分析内容の詳細

### 分析①：ポジショニング分析（規模 × 効率）

**目的**: 企業の規模と収益性/効率の関係を可視化

**実施内容**:
- 総資産 vs ROA の散布図
- 売上高 vs ROE の散布図
- バブル図（売上高 × ROA × 総資産）
- 外れ値企業の特定

**出力**:
- `output/01_positioning_table.xlsx`
- `output/fig_01_roa_vs_assets.png`
- `output/fig_01_roe_vs_sales.png`
- `output/fig_01_bubble_positioning.png`
- `output/01_summary.md` ※分析者が記載した観察事項・考察を含む

### 分析②：収益性分解（マージン分析）

**目的**: 営業利益と純利益の差から利益の「質」を把握

**実施内容**:
- 営業利益率・当期純利益率の計算
- 利益率差分の分析
- 企業別の棒グラフ
- 45度線を用いた散布図

**出力**:
- `output/02_margin_table.xlsx`
- `output/fig_02_margins_by_company.png`
- `output/fig_02_opm_vs_npm.png`
- `output/02_summary.md` ※分析者が記載した観察事項・考察を含む

### 分析③：資本効率・財務体質

**目的**: ROA/ROE/自己資本比率から資本構造を理解

**実施内容**:
- ROA vs ROE の関係
- 自己資本比率 vs ROE
- レバレッジ効果の可視化
- ROA分解（資産回転率 × 利益率）
- 象限別の企業分類

**出力**:
- `output/03_capital_table.xlsx`
- `output/fig_03_roa_vs_roe.png`
- `output/fig_03_equity_ratio_vs_roe.png`
- `output/fig_03_equity_vs_roe_roa_gap.png`
- `output/fig_03_turnover_vs_margin.png`
- `output/03_summary.md` ※分析者が記載した観察事項・考察を含む

### 分析④：クラスタリング（業界構造の類型化）

**目的**: 企業を複数のタイプに分類し、業界構造を理解

**実施内容**:
- k-meansクラスタリング（k=2〜4）
- シルエットスコアによる最適k選択
- PCAによる2次元可視化
- クラスター別プロファイル（レーダーチャート）

**出力**:
- `output/04_cluster_assignments.xlsx`
- `output/04_cluster_profile.xlsx`
- `output/fig_04_pca_clusters.png`
- `output/fig_04_cluster_profile.png`
- `output/04_summary.md` ※分析者が記載した観察事項・考察を含む

---

## 🔧 共通モジュールの説明

#### `modules/io.py`
- Excel/CSV形式のデータ読み込み・保存
- 列名の正規化
- 出力ディレクトリの自動作成

#### `modules/preprocess.py`
- 数値列の型変換（カンマ除去など）
- 必須列の存在確認
- 欠損値の処理

#### `modules/metrics.py`
- 財務指標の計算
  - 営業利益率、当期純利益率
  - 自己資本比率、総資産回転率
  - ROE-ROA差分
- 外れ値検出（IQR法、Z-score法）
- 上位・下位企業の抽出

#### `modules/viz.py`
- 散布図・バブル図の作成
- 棒グラフの作成
- 企業名ラベルの自動付与
- 日本語フォント対応（japanize-matplotlib）

#### `modules/report.py`
- 基本統計量の集計
- Markdown形式のレポート生成
- データ品質レポート
- 分析メタデータの記録

---

## コーディング規約

- **言語**: Python 3.12+
- **スタイルガイド**: PEP8準拠
- **型ヒント**: 関数には必ず型ヒントを記述
- **Docstring**: NumPyスタイル（日本語）
- **コメント**: 日本語で記述
- **セル形式**: `# %%[markdown]` と `# %%` を使用

---

## 依存パッケージ

主な依存パッケージ：
- `pandas`: データ操作
- `numpy`: 数値計算
- `matplotlib`: 可視化
- `seaborn`: 統計的可視化
- `scikit-learn`: 機械学習（クラスタリング、PCA）
- `japanize-matplotlib`: 日本語フォント対応
- `openpyxl`: Excel読み込み

完全なリストは [pyproject.toml](pyproject.toml) を参照してください。

---

## セキュリティ・プライバシー

- 認証情報は環境変数から読み込むこと
- 個人情報を含むデータはリポジトリに含めないこと
- サンプルデータ使用時はその旨をREADMEに明記すること

---

## 参考資料

詳細な分析計画・仮説・実施手順については、以下を参照してください：
- [プロンプト用_分析計画.md](プロンプト用_分析計画.md)

---

**注意**: 分析の実行前に、必ず `input/財務データ_製薬業界.xlsx` を配置してください。

# %%[markdown]
# # 製薬業界分析プロジェクト - メインエントリーポイント
# 
# このスクリプトは、srcディレクトリ内の分析スクリプトを順番に実行するにゃー。
# 
# ## 実行方法
# 
# ### 1. コマンドラインでの実行
# ```bash
# python main.py                    # 全スクリプトを実行
# python main.py --stop-on-error    # エラー時に停止
# python main.py --script 1         # スクリプト1のみ実行
# ```
# 
# ### 2. インタラクティブウィンドウでの実行
# 1. まず、このファイルを開く
# 2. 「Run All Cells」または各セルを順番に実行して、全関数を定義
# 3. ファイルの**最後にある実行用セル**のコメントを外して実行
#    - `run_all_analyses()` で全分析を実行
#    - `run_single_analysis(1)` で個別分析を実行
# 
# ## 実行順序
# 1. 01_positioning.py - ポジショニング分析
# 2. 02_profitability_decompose.py - 収益性分解
# 3. 03_capital_efficiency.py - 資本効率分析
# 4. 04_clustering.py - クラスタリング分析

# %%
# インポートとプロジェクト設定
import sys
from pathlib import Path
from typing import List, Optional
import time
import traceback

# プロジェクトルートの設定
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

# %%
# 個別スクリプト実行関数


def run_script(script_path: Path, script_name: str) -> bool:
    """
    指定されたPythonスクリプトを実行する
    
    Parameters
    ----------
    script_path : Path
        実行するスクリプトのパス
    script_name : str
        スクリプトの表示名（ログ出力用）
    
    Returns
    -------
    bool
        実行が成功したかどうか
    """
    print(f"\n{'='*60}")
    print(f"{script_name} の実行を開始するにゃー")
    print(f"{'='*60}")
    
    start_time = time.time()
    
    try:
        # スクリプトの内容を読み込んで実行
        with open(script_path, 'r', encoding='utf-8') as f:
            script_content = f.read()
        
        # セルマーカー（# %%）を削除してから実行
        # これによりインタラクティブ環境でもコマンドラインでも動作する
        script_content_cleaned = script_content.replace('# %%[markdown]', '')
        script_content_cleaned = script_content_cleaned.replace('# %%', '')
        
        # グローバル名前空間を使って実行（各スクリプトが独立して動作）
        exec_globals = {
            '__file__': str(script_path),
            '__name__': '__main__',
        }
        
        exec(script_content_cleaned, exec_globals)
        
        elapsed_time = time.time() - start_time
        print(f"\n{script_name} の実行が完了したにゃー（所要時間: {elapsed_time:.2f}秒）")
        return True
        
    except Exception as e:
        elapsed_time = time.time() - start_time
        print(f"\n {script_name} の実行中にエラーが発生したにゃー（所要時間: {elapsed_time:.2f}秒）")
        print(f"エラー内容: {str(e)}")
        print("\n詳細なトレースバック:")
# %%
# 全分析実行関数

        traceback.print_exc()
        return False


def run_all_analyses(stop_on_error: bool = False) -> None:
    """
    全ての分析スクリプトを順番に実行する
    
    Parameters
    ----------
    stop_on_error : bool, default=False
        エラーが発生した場合に実行を停止するかどうか
        False: エラーが発生しても次のスクリプトを実行
        True: エラーが発生したら即座に停止
    """
    # 実行するスクリプトのリスト
    scripts = [
        ("01_positioning.py", "分析①：ポジショニング分析"),
        ("02_profitability_decompose.py", "分析②：収益性分解"),
        ("03_capital_efficiency.py", "分析③：資本効率・財務体質"),
        ("04_clustering.py", "分析④：クラスタリング"),
    ]
    
    print("\n" + "="*60)
    print("製薬業界分析プロジェクトを開始するにゃー")
    print("="*60)
    print(f"実行対象: {len(scripts)}個のスクリプト")
    print(f"エラー時の動作: {'停止する' if stop_on_error else '継続する'}")
    print("="*60)
    
    src_dir = PROJECT_ROOT / "src"
    results = []
    
    total_start_time = time.time()
    
    for i, (script_file, script_name) in enumerate(scripts, 1):
        script_path = src_dir / script_file
        
        if not script_path.exists():
            print(f"\n{script_file} が見つからないにゃー: {script_path}")
            results.append((script_name, False))
            if stop_on_error:
                break
            continue
        
        success = run_script(script_path, f"[{i}/{len(scripts)}] {script_name}")
        results.append((script_name, success))
        
        if not success and stop_on_error:
            print("\nエラーが発生したため、実行を中断するにゃー")
            break
    
    # 実行結果のサマリー
    total_elapsed_time = time.time() - total_start_time
    print("\n" + "="*60)
    print("実行結果サマリー")
    print("="*60)
    
    success_count = sum(1 for _, success in results if success)
    for script_name, success in results:
        status = "成功" if success else "失敗"
        print(f"{status}: {script_name}")
    
    print(f"\n成功: {success_count}/{len(results)}")
# %%
# 個別分析実行関数

    print(f"総所要時間: {total_elapsed_time:.2f}秒")
    print("="*60)


def run_single_analysis(script_number: int) -> None:
    """
    指定された番号の分析スクリプトのみを実行する
    
    Parameters
    ----------
    script_number : int
        実行するスクリプトの番号（1-4）
    """
    scripts = {
        1: ("01_positioning.py", "分析①：ポジショニング分析"),
        2: ("02_profitability_decompose.py", "分析②：収益性分解"),
        3: ("03_capital_efficiency.py", "分析③：資本効率・財務体質"),
        4: ("04_clustering.py", "分析④：クラスタリング"),
    }
    
    if script_number not in scripts:
        print(f"無効なスクリプト番号: {script_number}")
        print(f"有効な番号は 1-{len(scripts)} にゃー")
        return
    
    script_file, script_name = scripts[script_number]
    script_path = PROJECT_ROOT / "src" / script_file
    
    if not script_path.exists():
        print(f"{script_file} が見つからないにゃー: {script_path}")
        return
    
    run_script(script_path, script_name)

# %%
# メイン実行関数とコマンドライン実行エントリーポイント


def main() -> None:
    """
    メイン関数
    
    コマンドライン引数に応じて実行モードを切り替える:
    - 引数なし: 全スクリプトを順番に実行
    - --script N: 指定された番号（1-4）のスクリプトのみ実行
    - --stop-on-error: エラー時に実行を停止
    """
    import argparse
    
    parser = argparse.ArgumentParser(
        description="製薬業界分析プロジェクトのメインスクリプト",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python main.py                    # 全スクリプトを実行
  python main.py --stop-on-error    # エラー時に停止
  python main.py --script 1         # スクリプト1のみ実行
  python main.py --script 3         # スクリプト3のみ実行
        """
    )
    
    parser.add_argument(
        '--script',
        type=int,
        choices=[1, 2, 3, 4],
        help='実行する分析スクリプトの番号（1-4）'
    )
    
    parser.add_argument(
        '--stop-on-error',
        action='store_true',
        help='エラー発生時に実行を停止する'
    )
    
    args = parser.parse_args()
    
    if args.script:
        # 単一スクリプトの実行
        run_single_analysis(args.script)
    else:
        # 全スクリプトの実行
        run_all_analyses(stop_on_error=args.stop_on_error)

# エントリーポイント（コマンドライン実行用）
if __name__ == "__main__":
    main()

# %%[markdown]
# ---
# ## インタラクティブ実行用セル
# 
# 上記のすべてのセルを実行して関数を定義した後、
# 以下のセルのコメントを外して実行するにゃー

# %%
# 全分析を実行（エラーが発生しても継続）
# run_all_analyses()

# %%
# 全分析を実行（エラー時に停止）
# run_all_analyses(stop_on_error=True)

# %%
# 特定の分析のみ実行
# run_single_analysis(1)  # ポジショニング分析
# run_single_analysis(2)  # 収益性分解
# run_single_analysis(3)  # 資本効率分析
# run_single_analysis(4)  # クラスタリング分析

"""Excelファイルの構造を詳細に分析するスクリプト"""

import sys
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers import ExcelReader


def analyze_file_structure(file_path: str):
    """ファイルの構造を詳細に分析"""
    print(f"\n{'='*80}")
    print(f"📊 ファイル: {Path(file_path).name}")
    print(f"{'='*80}")
    
    reader = ExcelReader()
    result = reader.read_file(file_path)
    
    for sheet in result['sheets']:
        print(f"\n【シート: {sheet['name']}】")
        print(f"行数: {len(sheet['data'])}")
        
        # 最初の10行を詳細に表示
        print("\n最初の10行:")
        for i, row in enumerate(sheet['data'][:10], 1):
            print(f"\n  行{i}:")
            for j, cell in enumerate(row, 1):
                if cell and str(cell).strip() and str(cell) != 'None':
                    print(f"    列{j}: '{cell}'")


def main():
    """メイン関数"""
    print("\n" + "="*80)
    print("Excel構造分析")
    print("="*80)
    
    files = [
        "test_files/04_システム概要.xlsx",
        "test_files/05_画面遷移図.xlsx",
        "test_files/06_画面一覧.xlsx"
    ]
    
    for file_path in files:
        try:
            analyze_file_structure(file_path)
        except Exception as e:
            print(f"\n❌ エラー: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()

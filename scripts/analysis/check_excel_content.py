"""
Excelファイルの内容を確認するスクリプト
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers.excel_reader import ExcelReader


def check_excel_files():
    """Excelファイルの内容を確認"""
    reader = ExcelReader()
    
    files = [
        "test_files/04_システム概要.xlsx",
        "test_files/06_画面一覧.xlsx"
    ]
    
    for file_path in files:
        if not Path(file_path).exists():
            print(f"⚠️  ファイルが見つかりません: {file_path}")
            continue
            
        print(f"\n{'=' * 60}")
        print(f"📊 {Path(file_path).name}")
        print(f"{'=' * 60}")
        
        result = reader.read_file(file_path)
        
        if result:
            sheets = result.get("sheets", [])
            print(f"シート数: {len(sheets)}\n")
            
            for sheet in sheets:
                print(f"シート名: {sheet.get('name', '不明')}")
                data = sheet.get('data', [])
                print(f"行数: {len(data)}")
                print(f"\nデータ（最初の10行）:")
                
                for i, row in enumerate(data[:10], 1):
                    print(f"  {i}: {row}")
                
                print()


if __name__ == "__main__":
    check_excel_files()

"""
生成されたspecのExcelファイルの内容を確認するスクリプト
"""
import sys
from pathlib import Path

# プロジェクトのsrcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers.excel_reader import ExcelReader


def verify_excel_file(file_path: str):
    """
    Excelファイルの内容を確認する
    
    Args:
        file_path: Excelファイルのパス
    """
    print(f"\n{'=' * 60}")
    print(f"📊 ファイル: {Path(file_path).name}")
    print(f"{'=' * 60}")
    
    reader = ExcelReader()
    result = reader.read_file(file_path)
    
    if not result:
        print("❌ ファイルの読み込みに失敗しました")
        return
    
    sheets = result.get("sheets", [])
    print(f"\n✅ シート数: {len(sheets)}")
    
    for i, sheet in enumerate(sheets, 1):
        sheet_name = sheet.get("name", "不明")
        data = sheet.get("data", [])
        row_count = sheet.get("row_count", 0)
        column_count = sheet.get("column_count", 0)
        
        print(f"\n  シート {i}: {sheet_name}")
        print(f"    行数: {row_count}, 列数: {column_count}")
        
        # 最初の5行を表示
        print(f"    最初の5行:")
        for j, row in enumerate(data[:5], 1):
            # 空のセルを除外して表示
            non_empty_cells = [str(cell) for cell in row if cell]
            if non_empty_cells:
                print(f"      {j}: {' | '.join(non_empty_cells[:3])}")


def main():
    """
    メイン処理
    """
    print("\n🔍 生成されたspecのExcelファイルを確認します\n")
    
    test_files_dir = Path("test_files")
    spec_files = [
        "spec_requirements.xlsx",
        "spec_design.xlsx",
        "spec_tasks.xlsx"
    ]
    
    for spec_file in spec_files:
        file_path = test_files_dir / spec_file
        if file_path.exists():
            verify_excel_file(str(file_path))
        else:
            print(f"\n⚠️  ファイルが見つかりません: {file_path}")
    
    print(f"\n{'=' * 60}")
    print("✨ 確認完了")
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    main()

"""
すべてのドキュメントをreadersディレクトリに出力する統合スクリプト
"""
import sys
import shutil
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers.excel_reader import ExcelReader


def copy_spec_files(output_dir: Path):
    """
    requirements.mdとdesign.mdをコピー
    
    Args:
        output_dir: 出力ディレクトリ
    """
    spec_dir = Path(".kiro/specs/document-format-mcp-server")
    files_to_copy = ["requirements.md", "design.md"]
    
    print("\n📝 Specファイルのコピー")
    print("-" * 60)
    
    for filename in files_to_copy:
        source_file = spec_dir / filename
        dest_file = output_dir / filename
        
        if not source_file.exists():
            print(f"⚠️  ファイルが見つかりません: {source_file}")
            continue
        
        try:
            shutil.copy2(source_file, dest_file)
            file_size = dest_file.stat().st_size
            print(f"✅ {filename} ({file_size:,} bytes)")
        except Exception as e:
            print(f"❌ {filename} のコピーに失敗: {e}")


def excel_to_markdown(excel_file: str, sheet_name: str) -> str:
    """
    Excelファイルを読み込んでMarkdown文字列を生成
    
    Args:
        excel_file: 入力Excelファイルのパス
        sheet_name: 処理するシート名
        
    Returns:
        Markdown文字列
    """
    reader = ExcelReader()
    result = reader.read_file(excel_file)
    
    if not result:
        return ""
    
    sheets = result.get("sheets", [])
    md_lines = []
    
    for sheet in sheets:
        if sheet.get("name") != sheet_name:
            continue
            
        data = sheet.get("data", [])
        
        # シート名を見出しとして追加
        md_lines.append(f"# {sheet_name}\n\n")
        
        # システム概要の場合
        if sheet_name == "システム概要":
            for row in data:
                if len(row) >= 2 and row[1]:
                    md_lines.append(f"{row[1]}\n\n")
                elif len(row) >= 1 and row[0] and row[0] != sheet_name:
                    md_lines.append(f"## {row[0]}\n\n")
        
        # 画面一覧の場合
        elif sheet_name == "画面一覧":
            # ヘッダー行を探す
            header_row = None
            data_start_index = 0
            
            for i, row in enumerate(data):
                if any(cell in str(row) for cell in ["#", "機能", "画面ID", "タイトル"]):
                    header_row = row
                    data_start_index = i + 1
                    break
            
            if header_row:
                # ヘッダー行
                header_cells = [str(cell) if cell else "" for cell in header_row]
                md_lines.append("| " + " | ".join(header_cells) + " |\n")
                
                # 区切り行
                md_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |\n")
                
                # データ行
                for row in data[data_start_index:]:
                    if any(cell for cell in row):
                        cells = [str(cell) if cell else "" for cell in row]
                        while len(cells) < len(header_cells):
                            cells.append("")
                        md_lines.append("| " + " | ".join(cells[:len(header_cells)]) + " |\n")
    
    return "".join(md_lines)


def export_excel_files(output_dir: Path):
    """
    Excelファイルを読み込んでMarkdownファイルを生成
    
    Args:
        output_dir: 出力ディレクトリ
    """
    print("\n📊 ExcelファイルからMarkdownを生成")
    print("-" * 60)
    
    # 処理するファイルのリスト
    files_to_process = [
        {
            "excel": "test_files/04_システム概要.xlsx",
            "sheet": "システム概要",
            "output": "system_overview.md"
        },
        {
            "excel": "test_files/06_画面一覧.xlsx",
            "sheet": "画面一覧",
            "output": "screen_list.md"
        }
    ]
    
    for file_info in files_to_process:
        excel_file = file_info["excel"]
        sheet_name = file_info["sheet"]
        output_filename = file_info["output"]
        
        if not Path(excel_file).exists():
            print(f"⚠️  ファイルが見つかりません: {excel_file}")
            continue
        
        # Markdownを生成
        markdown_content = excel_to_markdown(excel_file, sheet_name)
        
        if markdown_content:
            # ファイルに書き込み
            output_file = output_dir / output_filename
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            file_size = output_file.stat().st_size
            print(f"✅ {output_filename} ({file_size:,} bytes)")
        else:
            print(f"❌ {output_filename} の生成に失敗")


def main():
    """
    メイン処理
    """
    print("\n" + "=" * 60)
    print("📚 ドキュメントをreadersディレクトリに出力")
    print("=" * 60)
    
    # 出力ディレクトリ
    output_dir = Path("src/document_format_mcp_server/readers")
    
    if not output_dir.exists():
        print(f"❌ 出力ディレクトリが見つかりません: {output_dir}")
        return
    
    print(f"\n出力先: {output_dir}")
    
    # Specファイルをコピー
    copy_spec_files(output_dir)
    
    # Excelファイルを変換
    export_excel_files(output_dir)
    
    print("\n" + "=" * 60)
    print("✨ すべての処理が完了しました")
    print("=" * 60)
    
    # 出力されたファイルの一覧を表示
    print("\n📁 出力されたファイル:")
    for file in sorted(output_dir.glob("*.md")):
        if file.name not in ["__init__.py"]:
            size = file.stat().st_size
            print(f"  - {file.name} ({size:,} bytes)")


if __name__ == "__main__":
    main()

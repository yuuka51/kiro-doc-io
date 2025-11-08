"""
Excelファイルから適切なMarkdownファイルを生成するスクリプト
"""
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers.excel_reader import ExcelReader


def format_timestamp(unix_timestamp):
    """
    Unixタイムスタンプを人間が読みやすい形式に変換
    
    Args:
        unix_timestamp: Unixタイムスタンプ（float or str）
        
    Returns:
        フォーマットされた日時文字列
    """
    try:
        timestamp = float(unix_timestamp)
        dt = datetime.fromtimestamp(timestamp)
        return dt.strftime("%Y年%m月%d日 %H:%M:%S")
    except:
        return str(unix_timestamp)


def excel_to_markdown(excel_file: str, output_file: str):
    """
    Excelファイルを読み込んでMarkdownファイルを生成
    
    Args:
        excel_file: 入力Excelファイルのパス
        output_file: 出力Markdownファイルのパス
    """
    reader = ExcelReader()
    result = reader.read_file(excel_file)
    
    if not result:
        print(f"❌ ファイルの読み込みに失敗: {excel_file}")
        return False
    
    sheets = result.get("sheets", [])
    
    # Markdownコンテンツを生成
    md_lines = []
    
    for sheet in sheets:
        sheet_name = sheet.get("name", "不明")
        data = sheet.get("data", [])
        
        # シート名を見出しとして追加
        md_lines.append(f"# {sheet_name}\n")
        
        # データを処理
        if len(data) > 0:
            # 最初の行がヘッダーかどうかを判定
            first_row = data[0]
            
            # システム概要の場合
            if sheet_name == "システム概要":
                for row in data:
                    if len(row) >= 2 and row[1]:
                        # 2列目にデータがある場合
                        md_lines.append(f"{row[1]}\n")
                    elif len(row) >= 1 and row[0]:
                        # 1列目にデータがある場合（見出しなど）
                        if row[0] != sheet_name:  # シート名と同じ場合はスキップ
                            md_lines.append(f"## {row[0]}\n")
            
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
                    # テーブルとして出力
                    md_lines.append("\n")
                    
                    # ヘッダー行
                    header_cells = [str(cell) if cell else "" for cell in header_row]
                    md_lines.append("| " + " | ".join(header_cells) + " |\n")
                    
                    # 区切り行
                    md_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |\n")
                    
                    # データ行
                    for row in data[data_start_index:]:
                        if any(cell for cell in row):  # 空行をスキップ
                            cells = [str(cell) if cell else "" for cell in row]
                            # 列数を合わせる
                            while len(cells) < len(header_cells):
                                cells.append("")
                            md_lines.append("| " + " | ".join(cells[:len(header_cells)]) + " |\n")
            
            # その他のシート
            else:
                for row in data:
                    if any(cell for cell in row):  # 空行をスキップ
                        # 最初の非空セルを見つける
                        non_empty_cells = [str(cell) for cell in row if cell]
                        if non_empty_cells:
                            md_lines.append(f"{' '.join(non_empty_cells)}\n")
        
        md_lines.append("\n")
    
    # ファイルに書き込み
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(md_lines)
    
    return True


def main():
    """
    メイン処理
    """
    print("\n" + "=" * 60)
    print("📊 ExcelファイルからMarkdownファイルを生成")
    print("=" * 60 + "\n")
    
    # 処理するファイルのリスト
    files_to_process = [
        {
            "excel": "test_files/04_システム概要.xlsx",
            "markdown": "output/system_overview.md"
        },
        {
            "excel": "test_files/06_画面一覧.xlsx",
            "markdown": "output/screen_list.md"
        }
    ]
    
    for file_info in files_to_process:
        excel_file = file_info["excel"]
        markdown_file = file_info["markdown"]
        
        if not Path(excel_file).exists():
            print(f"⚠️  ファイルが見つかりません: {excel_file}")
            continue
        
        print(f"📄 処理中: {Path(excel_file).name}")
        
        if excel_to_markdown(excel_file, markdown_file):
            print(f"✅ 出力成功: {markdown_file}")
            
            # ファイルサイズを表示
            size = Path(markdown_file).stat().st_size
            print(f"   サイズ: {size:,} bytes\n")
        else:
            print(f"❌ 出力失敗\n")
    
    print("=" * 60)
    print("✨ 処理完了")
    print("=" * 60)


if __name__ == "__main__":
    main()

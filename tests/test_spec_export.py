"""
Specファイル（requirements.md、design.md、tasks.md）をExcelおよびGoogleスプレッドシート形式で出力するテストスクリプト
"""
import os
import sys
from pathlib import Path

# プロジェクトのsrcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.writers.excel_writer import ExcelWriter
from document_format_mcp_server.writers.google_writer import GoogleWorkspaceWriter


def parse_markdown_to_structured_data(md_file_path: str) -> dict:
    """
    Markdownファイルを読み込み、構造化データに変換する
    
    Args:
        md_file_path: Markdownファイルのパス
        
    Returns:
        構造化されたデータ（Excel/Googleスプレッドシート用）
    """
    with open(md_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    sheets_data = []
    current_sheet = None
    current_rows = []
    
    for line in lines:
        # 見出しレベル1（# ）をシート名として使用
        if line.startswith('# '):
            if current_sheet and current_rows:
                sheets_data.append({
                    "name": current_sheet[:31],  # Excelのシート名は31文字まで
                    "data": current_rows
                })
            current_sheet = line[2:].strip()
            current_rows = [[current_sheet]]  # シート名を最初の行に
            
        # 見出しレベル2（## ）をセクション見出しとして追加
        elif line.startswith('## '):
            current_rows.append([line[3:].strip()])
            
        # 見出しレベル3（### ）をサブセクション見出しとして追加
        elif line.startswith('### '):
            current_rows.append(["", line[4:].strip()])
            
        # 見出しレベル4（#### ）をサブサブセクション見出しとして追加
        elif line.startswith('#### '):
            current_rows.append(["", "", line[5:].strip()])
            
        # 箇条書き（- ）を追加
        elif line.strip().startswith('- '):
            current_rows.append(["", "", line.strip()[2:]])
            
        # 通常のテキスト行を追加（空行は除く）
        elif line.strip():
            current_rows.append(["", "", "", line.strip()])
    
    # 最後のシートを追加
    if current_sheet and current_rows:
        sheets_data.append({
            "name": current_sheet[:31],
            "data": current_rows
        })
    
    return {"sheets": sheets_data}


def test_excel_export():
    """
    Specファイルをxlsx形式で出力するテスト
    """
    print("=" * 60)
    print("Excel形式でのエクスポートテスト")
    print("=" * 60)
    
    spec_dir = Path(".kiro/specs/document-format-mcp-server")
    output_dir = Path("test_files")
    output_dir.mkdir(exist_ok=True)
    
    writer = ExcelWriter()
    
    # 各specファイルをExcelに変換
    spec_files = ["requirements.md", "design.md", "tasks.md"]
    
    for spec_file in spec_files:
        spec_path = spec_dir / spec_file
        if not spec_path.exists():
            print(f"⚠️  ファイルが見つかりません: {spec_path}")
            continue
        
        print(f"\n📄 処理中: {spec_file}")
        
        # Markdownを構造化データに変換
        data = parse_markdown_to_structured_data(str(spec_path))
        
        # Excelファイルとして出力
        output_path = output_dir / f"spec_{spec_file.replace('.md', '.xlsx')}"
        result_path = writer.create_workbook(data, str(output_path))
        
        if result_path:
            print(f"✅ 出力成功: {result_path}")
            print(f"   シート数: {len(data['sheets'])}")
        else:
            print(f"❌ 出力失敗")


def test_google_sheets_export():
    """
    SpecファイルをGoogleスプレッドシート形式で出力するテスト
    """
    print("\n" + "=" * 60)
    print("Googleスプレッドシート形式でのエクスポートテスト")
    print("=" * 60)
    
    # Google認証情報の確認
    config_path = Path(".config/google-credentials.json")
    if not config_path.exists():
        print(f"⚠️  Google認証情報が見つかりません: {config_path}")
        print("   Googleスプレッドシートへのエクスポートをスキップします")
        print("   認証情報の設定方法は GOOGLE_API_SETUP.md を参照してください")
        return
    
    spec_dir = Path(".kiro/specs/document-format-mcp-server")
    
    try:
        writer = GoogleWorkspaceWriter(str(config_path))
    except Exception as e:
        print(f"❌ GoogleWorkspaceWriterの初期化に失敗: {e}")
        return
    
    # 各specファイルをGoogleスプレッドシートに変換
    spec_files = ["requirements.md", "design.md", "tasks.md"]
    
    for spec_file in spec_files:
        spec_path = spec_dir / spec_file
        if not spec_path.exists():
            print(f"⚠️  ファイルが見つかりません: {spec_path}")
            continue
        
        print(f"\n📄 処理中: {spec_file}")
        
        # Markdownを構造化データに変換
        data = parse_markdown_to_structured_data(str(spec_path))
        
        # Googleスプレッドシートとして出力
        title = f"Spec - {spec_file.replace('.md', '')}"
        try:
            url = writer.create_spreadsheet(data, title)
            if url:
                print(f"✅ 出力成功: {url}")
                print(f"   シート数: {len(data['sheets'])}")
            else:
                print(f"❌ 出力失敗")
        except Exception as e:
            print(f"❌ エラー: {e}")


def main():
    """
    メイン処理
    """
    print("\n🚀 Specファイルのエクスポートテストを開始します\n")
    
    # Excel形式でのエクスポートテスト
    test_excel_export()
    
    # Googleスプレッドシート形式でのエクスポートテスト
    test_google_sheets_export()
    
    print("\n" + "=" * 60)
    print("✨ テスト完了")
    print("=" * 60)
    print("\n出力ファイル:")
    print("  - Excel: test_files/spec_*.xlsx")
    print("  - Googleスプレッドシート: 上記のURLを参照")


if __name__ == "__main__":
    main()

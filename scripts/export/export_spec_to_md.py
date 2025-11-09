"""
.kiro/specs/配下のmdファイルをreadersディレクトリに出力するスクリプト
"""
import shutil
from pathlib import Path
from datetime import datetime


def export_spec_files():
    """
    requirements.mdとdesign.mdをreadersディレクトリにコピーする
    """
    # 入力ディレクトリと出力ディレクトリ
    spec_dir = Path(".kiro/specs/document-format-mcp-server")
    output_dir = Path("src/document_format_mcp_server/readers")
    
    # 出力ディレクトリが存在することを確認
    if not output_dir.exists():
        print(f"❌ 出力ディレクトリが見つかりません: {output_dir}")
        return
    
    # コピーするファイル
    files_to_copy = ["requirements.md", "design.md"]
    
    print("=" * 60)
    print("📝 Specファイルのエクスポート")
    print("=" * 60)
    print(f"\n入力元: {spec_dir}")
    print(f"出力先: {output_dir}\n")
    
    for filename in files_to_copy:
        source_file = spec_dir / filename
        dest_file = output_dir / filename
        
        if not source_file.exists():
            print(f"⚠️  ファイルが見つかりません: {source_file}")
            continue
        
        try:
            # ファイルをコピー
            shutil.copy2(source_file, dest_file)
            
            # ファイルサイズを取得
            file_size = dest_file.stat().st_size
            
            print(f"✅ {filename}")
            print(f"   サイズ: {file_size:,} bytes")
            print(f"   パス: {dest_file}")
            print()
            
        except Exception as e:
            print(f"❌ {filename} のコピーに失敗: {e}")
            print()
    
    print("=" * 60)
    print("✨ エクスポート完了")
    print("=" * 60)


if __name__ == "__main__":
    export_spec_files()

"""PowerPointWriterのテストスクリプト"""

import os
import sys

# プロジェクトのルートディレクトリをパスに追加
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from document_format_mcp_server.writers.powerpoint_writer import PowerPointWriter
from document_format_mcp_server.utils.logging_config import setup_logging


def test_powerpoint_writer():
    """PowerPointWriterの基本機能をテストする"""
    
    # ロギングのセットアップ
    setup_logging()
    
    print("=" * 60)
    print("PowerPointWriter テスト")
    print("=" * 60)
    
    writer = PowerPointWriter()
    
    # テストデータ1: タイトルスライド
    test_data_1 = {
        "slides": [
            {
                "layout": "title",
                "title": "Document Format MCP Server",
                "content": "PowerPointファイル生成機能のテスト"
            }
        ]
    }
    
    output_path_1 = "test_files/output_title_slide.pptx"
    
    try:
        print("\n[テスト1] タイトルスライドの生成")
        result = writer.create_presentation(test_data_1, output_path_1)
        print(f"✓ 成功: {result}")
        print(f"  ファイルサイズ: {os.path.getsize(result)} bytes")
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # テストデータ2: コンテンツスライド
    test_data_2 = {
        "slides": [
            {
                "layout": "title",
                "title": "プレゼンテーションタイトル",
                "content": "サブタイトル"
            },
            {
                "layout": "content",
                "title": "概要",
                "content": "このプレゼンテーションは、Document Format MCP Serverの\nPowerPoint生成機能をテストするために作成されました。"
            },
            {
                "layout": "content",
                "title": "詳細説明",
                "content": "PowerPointWriterクラスは、構造化データから\nPowerPointファイルを生成します。"
            }
        ]
    }
    
    output_path_2 = "test_files/output_content_slides.pptx"
    
    try:
        print("\n[テスト2] コンテンツスライドの生成")
        result = writer.create_presentation(test_data_2, output_path_2)
        print(f"✓ 成功: {result}")
        print(f"  ファイルサイズ: {os.path.getsize(result)} bytes")
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # テストデータ3: 箇条書きスライド（リスト形式）
    test_data_3 = {
        "slides": [
            {
                "layout": "title",
                "title": "機能一覧",
                "content": "Document Format MCP Server"
            },
            {
                "layout": "bullet",
                "title": "主な機能",
                "content": [
                    "PowerPointファイルの読み取り",
                    "Wordファイルの読み取り",
                    "Excelファイルの読み取り",
                    "Google Workspaceファイルの読み取り"
                ]
            },
            {
                "layout": "bullet",
                "title": "ライター機能",
                "content": [
                    "PowerPointファイルの生成",
                    "Wordファイルの生成",
                    "Excelファイルの生成",
                    "Google Workspaceファイルの生成"
                ]
            }
        ]
    }
    
    output_path_3 = "test_files/output_bullet_slides.pptx"
    
    try:
        print("\n[テスト3] 箇条書きスライドの生成（リスト形式）")
        result = writer.create_presentation(test_data_3, output_path_3)
        print(f"✓ 成功: {result}")
        print(f"  ファイルサイズ: {os.path.getsize(result)} bytes")
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # テストデータ4: 箇条書きスライド（文字列形式）
    test_data_4 = {
        "slides": [
            {
                "layout": "bullet",
                "title": "実装状況",
                "content": """PowerPointリーダー: 完了
Wordリーダー: 完了
Excelリーダー: 完了
Google Workspaceリーダー: 完了
PowerPointライター: 実装中"""
            }
        ]
    }
    
    output_path_4 = "test_files/output_bullet_string.pptx"
    
    try:
        print("\n[テスト4] 箇条書きスライドの生成（文字列形式）")
        result = writer.create_presentation(test_data_4, output_path_4)
        print(f"✓ 成功: {result}")
        print(f"  ファイルサイズ: {os.path.getsize(result)} bytes")
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # テストデータ5: 複合スライド
    test_data_5 = {
        "title": "総合テスト",
        "slides": [
            {
                "layout": "title",
                "title": "Document Format MCP Server",
                "content": "総合機能テスト"
            },
            {
                "layout": "content",
                "title": "プロジェクト概要",
                "content": "Kiro AIアシスタントに対して、Microsoft Office形式および\nGoogle Workspace形式のファイルを読み取り・生成する機能を提供します。"
            },
            {
                "layout": "bullet",
                "title": "対応形式",
                "content": [
                    "Microsoft PowerPoint (.pptx)",
                    "Microsoft Word (.docx)",
                    "Microsoft Excel (.xlsx)",
                    "Google スプレッドシート",
                    "Google ドキュメント",
                    "Google スライド"
                ]
            },
            {
                "layout": "content",
                "title": "技術スタック",
                "content": "Python 3.10以上\npython-pptx, python-docx, openpyxl\nGoogle API Client"
            },
            {
                "layout": "bullet",
                "title": "実装済み機能",
                "content": [
                    "PowerPoint/Word/Excelリーダー",
                    "Google Workspaceリーダー",
                    "MCPツール定義と統合",
                    "PowerPointライター（本機能）"
                ]
            }
        ]
    }
    
    output_path_5 = "test_files/output_comprehensive.pptx"
    
    try:
        print("\n[テスト5] 複合スライドの生成")
        result = writer.create_presentation(test_data_5, output_path_5)
        print(f"✓ 成功: {result}")
        print(f"  ファイルサイズ: {os.path.getsize(result)} bytes")
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # エラーハンドリングのテスト
    print("\n[テスト6] エラーハンドリング")
    
    # 不正なデータ形式
    try:
        print("  - 不正なデータ形式（slidesキーなし）")
        writer.create_presentation({"title": "test"}, "test_files/error.pptx")
        print("    ✗ エラーが検出されませんでした")
    except Exception as e:
        print(f"    ✓ 正しくエラーを検出: {type(e).__name__}")
    
    # 不正なレイアウト
    try:
        print("  - 不正なレイアウト")
        writer.create_presentation(
            {"slides": [{"layout": "invalid", "title": "test"}]},
            "test_files/error.pptx"
        )
        print("    ✗ エラーが検出されませんでした")
    except Exception as e:
        print(f"    ✓ 正しくエラーを検出: {type(e).__name__}")
    
    print("\n" + "=" * 60)
    print("テスト完了")
    print("=" * 60)
    print("\n生成されたファイル:")
    for path in [output_path_1, output_path_2, output_path_3, output_path_4, output_path_5]:
        if os.path.exists(path):
            print(f"  - {path}")


if __name__ == "__main__":
    test_powerpoint_writer()

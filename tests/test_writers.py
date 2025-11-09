"""ライター機能のテストスクリプト。"""

import os
import sys
from pathlib import Path

# プロジェクトのルートディレクトリをパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from document_format_mcp_server.writers.word_writer import WordWriter
from document_format_mcp_server.writers.excel_writer import ExcelWriter
from document_format_mcp_server.writers.powerpoint_writer import PowerPointWriter
from document_format_mcp_server.utils.logging_config import setup_logging, get_logger


# ロギングの設定
setup_logging()
logger = get_logger(__name__)


def test_word_writer():
    """WordWriterのテスト。"""
    logger.info("=" * 60)
    logger.info("WordWriterのテスト開始")
    logger.info("=" * 60)

    writer = WordWriter()

    # テストデータ
    data = {
        "title": "サンプルドキュメント",
        "sections": [
            {
                "heading": "はじめに",
                "level": 1,
                "paragraphs": [
                    "これはDocument Format MCP Serverのテスト用ドキュメントです。",
                    "WordWriterクラスの動作を確認するために作成されました。"
                ]
            },
            {
                "heading": "主な機能",
                "level": 1,
                "paragraphs": [
                    "このシステムは以下の機能を提供します："
                ],
                "bullets": [
                    "Wordファイルの読み取り",
                    "Wordファイルの生成",
                    "見出しと段落の管理",
                    "表と箇条書きのサポート"
                ]
            },
            {
                "heading": "データ表",
                "level": 2,
                "paragraphs": [
                    "以下は機能の実装状況を示す表です。"
                ],
                "tables": [
                    {
                        "data": [
                            ["機能", "状態", "備考"],
                            ["読み取り", "完了", "PowerPoint, Word, Excel対応"],
                            ["書き込み", "実装中", "Word, Excel実装完了"],
                            ["Google Workspace", "完了", "読み取りのみ"]
                        ]
                    }
                ]
            },
            {
                "heading": "まとめ",
                "level": 1,
                "paragraphs": [
                    "WordWriterクラスは正常に動作しています。",
                    "見出し、段落、箇条書き、表のすべての機能が実装されています。"
                ]
            }
        ]
    }

    # ファイルを生成
    output_path = "test_files/output_word.docx"
    try:
        result_path = writer.create_document(data, output_path)
        logger.info(f"✓ Wordファイルの生成に成功: {result_path}")
        logger.info(f"  セクション数: {len(data['sections'])}")
    except Exception as e:
        logger.error(f"✗ Wordファイルの生成に失敗: {e}")
        return False

    return True


def test_excel_writer():
    """ExcelWriterのテスト。"""
    logger.info("=" * 60)
    logger.info("ExcelWriterのテスト開始")
    logger.info("=" * 60)

    writer = ExcelWriter()

    # テストデータ
    data = {
        "sheets": [
            {
                "name": "機能一覧",
                "data": [
                    ["ID", "機能名", "カテゴリ", "状態"],
                    [1, "PowerPoint読み取り", "読み取り", "完了"],
                    [2, "Word読み取り", "読み取り", "完了"],
                    [3, "Excel読み取り", "読み取り", "完了"],
                    [4, "PowerPoint書き込み", "書き込み", "完了"],
                    [5, "Word書き込み", "書き込み", "完了"],
                    [6, "Excel書き込み", "書き込み", "完了"]
                ],
                "formatting": {
                    "header_row": True,
                    "auto_width": True
                }
            },
            {
                "name": "統計",
                "data": [
                    ["カテゴリ", "完了数"],
                    ["読み取り", 6],
                    ["書き込み", 3],
                    ["合計", 9]
                ],
                "formatting": {
                    "header_row": True,
                    "auto_width": True
                }
            },
            {
                "name": "進捗",
                "data": [
                    ["タスク", "進捗率"],
                    ["リーダー実装", "100%"],
                    ["ライター実装", "75%"],
                    ["MCPツール統合", "100%"],
                    ["テスト", "80%"]
                ],
                "formatting": {
                    "header_row": True,
                    "auto_width": True
                }
            }
        ]
    }

    # ファイルを生成
    output_path = "test_files/output_excel.xlsx"
    try:
        result_path = writer.create_workbook(data, output_path)
        logger.info(f"✓ Excelファイルの生成に成功: {result_path}")
        logger.info(f"  シート数: {len(data['sheets'])}")
        for sheet in data['sheets']:
            logger.info(f"    - {sheet['name']}: {len(sheet['data'])}行")
    except Exception as e:
        logger.error(f"✗ Excelファイルの生成に失敗: {e}")
        return False

    return True


def test_powerpoint_writer():
    """PowerPointWriterのテスト（既存機能の確認）。"""
    logger.info("=" * 60)
    logger.info("PowerPointWriterのテスト開始")
    logger.info("=" * 60)

    writer = PowerPointWriter()

    # テストデータ
    data = {
        "title": "ライター機能テスト",
        "slides": [
            {
                "layout": "title",
                "title": "Document Format MCP Server",
                "content": "ライター機能のテスト"
            },
            {
                "layout": "bullet",
                "title": "実装済みライター",
                "content": [
                    "PowerPointWriter - プレゼンテーション生成",
                    "WordWriter - ドキュメント生成",
                    "ExcelWriter - ワークブック生成"
                ]
            },
            {
                "layout": "content",
                "title": "次のステップ",
                "content": "Google Workspaceライター機能の実装を進めます。"
            }
        ]
    }

    # ファイルを生成
    output_path = "test_files/output_powerpoint_test.pptx"
    try:
        result_path = writer.create_presentation(data, output_path)
        logger.info(f"✓ PowerPointファイルの生成に成功: {result_path}")
        logger.info(f"  スライド数: {len(data['slides'])}")
    except Exception as e:
        logger.error(f"✗ PowerPointファイルの生成に失敗: {e}")
        return False

    return True


def main():
    """メイン関数。"""
    logger.info("ライター機能のテストを開始します")
    logger.info("")

    # test_filesディレクトリが存在しない場合は作成
    os.makedirs("test_files", exist_ok=True)

    results = []

    # 各ライターをテスト
    results.append(("PowerPointWriter", test_powerpoint_writer()))
    results.append(("WordWriter", test_word_writer()))
    results.append(("ExcelWriter", test_excel_writer()))

    # 結果のサマリー
    logger.info("")
    logger.info("=" * 60)
    logger.info("テスト結果サマリー")
    logger.info("=" * 60)

    for name, success in results:
        status = "✓ 成功" if success else "✗ 失敗"
        logger.info(f"{name}: {status}")

    # 全体の結果
    all_success = all(success for _, success in results)
    logger.info("")
    if all_success:
        logger.info("すべてのテストが成功しました！")
        logger.info("生成されたファイルは test_files/ ディレクトリにあります。")
    else:
        logger.error("一部のテストが失敗しました。")

    return 0 if all_success else 1


if __name__ == "__main__":
    sys.exit(main())

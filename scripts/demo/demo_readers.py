"""リーダー機能のデモンストレーション"""

import sys
import json
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers import (
    PowerPointReader,
    WordReader,
    ExcelReader,
)


def demo_powerpoint():
    """PowerPointリーダーのデモ"""
    print("\n" + "="*70)
    print("PowerPointファイルの読み込みデモ")
    print("="*70)
    
    reader = PowerPointReader()
    result = reader.read_file("test_files/sample.pptx")
    
    print(f"\n📊 抽出された情報:")
    print(f"  スライド数: {len(result['slides'])}")
    
    for slide in result['slides']:
        print(f"\n  【スライド {slide['slide_number']}】")
        print(f"    タイトル: {slide['title']}")
        print(f"    コンテンツ: {slide['content'][:80]}..." if len(slide['content']) > 80 else f"    コンテンツ: {slide['content']}")
        print(f"    ノート: {slide['notes'] if slide['notes'] else '(なし)'}")
        
        if slide['tables']:
            print(f"    表の数: {len(slide['tables'])}")
            for i, table in enumerate(slide['tables'], 1):
                print(f"      表{i}: {table['rows']}行 x {table['columns']}列")
                print(f"        データサンプル: {table['data'][0][:3]}")
    
    print("\n💡 Kiroに提供される情報:")
    print("  - 各スライドのタイトルと本文")
    print("  - スライドノート（発表者用メモ）")
    print("  - 表データ（構造化された形式）")
    print("  - スライドの順序と階層")


def demo_word():
    """Wordリーダーのデモ"""
    print("\n" + "="*70)
    print("Wordファイルの読み込みデモ")
    print("="*70)
    
    reader = WordReader()
    result = reader.read_file("test_files/sample.docx")
    
    print(f"\n📄 抽出された情報:")
    print(f"  段落数: {len(result['paragraphs'])}")
    print(f"  表の数: {len(result['tables'])}")
    
    print("\n  【段落の内容】")
    for i, para in enumerate(result['paragraphs'][:8], 1):
        level = para.get('level')
        if level is not None and level > 0:
            print(f"    {i}. [見出し{level}] {para['text']}")
        else:
            text = para['text'][:60] + "..." if len(para['text']) > 60 else para['text']
            print(f"    {i}. [段落] {text}")
    
    if result['tables']:
        print(f"\n  【表の内容】")
        for i, table in enumerate(result['tables'], 1):
            print(f"    表{i}: {table['rows']}行 x {table['columns']}列")
            print(f"      ヘッダー: {table['data'][0]}")
            if len(table['data']) > 1:
                print(f"      データ例: {table['data'][1]}")
    
    print("\n💡 Kiroに提供される情報:")
    print("  - ドキュメントの階層構造（見出しレベル）")
    print("  - 段落ごとのテキスト内容")
    print("  - 箇条書きリスト")
    print("  - 表データ（構造化された形式）")


def demo_excel():
    """Excelリーダーのデモ"""
    print("\n" + "="*70)
    print("Excelファイルの読み込みデモ")
    print("="*70)
    
    reader = ExcelReader()
    result = reader.read_file("test_files/sample.xlsx")
    
    print(f"\n📈 抽出された情報:")
    print(f"  シート数: {len(result['sheets'])}")
    
    for sheet in result['sheets']:
        print(f"\n  【シート: {sheet['name']}】")
        row_count = len(sheet['data'])
        column_count = max(len(row) for row in sheet['data']) if sheet['data'] else 0
        print(f"    行数: {row_count}, 列数: {column_count}")
        
        # 最初の数行を表示
        print(f"    データ:")
        for i, row in enumerate(sheet['data'][:5], 1):
            row_str = " | ".join(str(cell) if cell is not None else "" for cell in row[:5])
            print(f"      {i}. {row_str}")
        
        # 数式がある場合
        if sheet['formulas']:
            print(f"    数式:")
            for cell, formula in list(sheet['formulas'].items())[:3]:
                print(f"      {cell}: {formula}")
    
    print("\n💡 Kiroに提供される情報:")
    print("  - 各シートの名前とデータ")
    print("  - セルの値（数値、テキスト、日付など）")
    print("  - 数式の内容")
    print("  - データの行数・列数")


def show_json_structure():
    """JSON構造のサンプルを表示"""
    print("\n" + "="*70)
    print("データ構造のサンプル（JSON形式）")
    print("="*70)
    
    # PowerPointのサンプル
    ppt_sample = {
        "slides": [
            {
                "slide_number": 1,
                "title": "サンプルプレゼンテーション",
                "content": "Document Format MCP Server テスト用",
                "notes": "",
                "tables": []
            },
            {
                "slide_number": 2,
                "title": "機能紹介",
                "content": "主な機能:\n  PowerPointファイルの読み取り\n  ...",
                "notes": "",
                "tables": []
            }
        ]
    }
    
    print("\n【PowerPointのデータ構造】")
    print(json.dumps(ppt_sample, ensure_ascii=False, indent=2))
    
    # Wordのサンプル
    word_sample = {
        "paragraphs": [
            {
                "text": "サンプルドキュメント",
                "type": "heading",
                "level": 0,
                "style": "Title"
            },
            {
                "text": "これはDocument Format MCP Serverのテスト用...",
                "type": "paragraph",
                "style": "Normal"
            }
        ],
        "tables": [
            {
                "rows": 4,
                "columns": 3,
                "data": [
                    ["機能", "ステータス", "備考"],
                    ["ローカルファイル読み取り", "完了", "PowerPoint, Word, Excel"]
                ]
            }
        ]
    }
    
    print("\n【Wordのデータ構造】")
    print(json.dumps(word_sample, ensure_ascii=False, indent=2))
    
    # Excelのサンプル
    excel_sample = {
        "sheets": [
            {
                "name": "データ",
                "data": [
                    ["ID", "名前", "カテゴリ", "値"],
                    [1, "PowerPoint", "読み取り", "完了"]
                ],
                "row_count": 7,
                "column_count": 4,
                "formulas": {}
            },
            {
                "name": "計算",
                "data": [
                    ["項目", "値", "計算"],
                    ["数値1", 100, None],
                    ["数値2", 50, None],
                    ["合計", None, "=B2+B3"]
                ],
                "row_count": 5,
                "column_count": 3,
                "formulas": {
                    "C4": "=B2+B3",
                    "C5": "=(B2+B3)/2"
                }
            }
        ]
    }
    
    print("\n【Excelのデータ構造】")
    print(json.dumps(excel_sample, ensure_ascii=False, indent=2))


def main():
    """メイン関数"""
    print("\n" + "="*70)
    print("Document Format MCP Server - リーダー機能デモンストレーション")
    print("="*70)
    
    try:
        # 各ファイル形式のデモ
        demo_powerpoint()
        demo_word()
        demo_excel()
        
        # データ構造のサンプル
        show_json_structure()
        
        print("\n" + "="*70)
        print("まとめ")
        print("="*70)
        print("\n✅ 実装済みの機能:")
        print("  1. PowerPoint (.pptx) の読み取り")
        print("     - スライドのタイトル、本文、ノート")
        print("     - 表データの抽出")
        print("")
        print("  2. Word (.docx) の読み取り")
        print("     - 見出しと段落の階層構造")
        print("     - 箇条書きリスト")
        print("     - 表データの抽出")
        print("")
        print("  3. Excel (.xlsx) の読み取り")
        print("     - 複数シートのデータ")
        print("     - セルの値と数式")
        print("     - データの構造化")
        print("")
        print("💡 Kiroへの活用例:")
        print("  - 設計書を読み込んで、その内容に基づいたコード生成")
        print("  - データファイルを分析して、レポート作成")
        print("  - プレゼン資料の内容を要約")
        print("  - ドキュメントの構造を理解して、類似文書の生成")
        print("")
        print("🚧 今後の実装予定:")
        print("  - ファイル書き込み機能（PowerPoint、Word、Excel）")
        print("  - Google Workspace対応（スプレッドシート、ドキュメント、スライド）")
        print("  - MCPツールとしての統合")
        print("="*70 + "\n")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

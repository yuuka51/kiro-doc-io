"""Excelファイルから仕様書を生成する改善版スクリプト"""

import sys
from pathlib import Path
from datetime import datetime

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers import ExcelReader


def read_system_overview(file_path: str):
    """システム概要を読み込む"""
    reader = ExcelReader()
    result = reader.read_file(file_path)
    
    overview_text = []
    for sheet in result['sheets']:
        for row in sheet['data'][1:]:  # 1行目はタイトルなのでスキップ
            if len(row) > 1 and row[1] and str(row[1]).strip():
                text = str(row[1]).strip()
                if text != 'None':
                    overview_text.append(text)
    
    return "\n".join(overview_text)


def read_screen_list(file_path: str):
    """画面一覧を読み込む"""
    reader = ExcelReader()
    result = reader.read_file(file_path)
    
    screens = []
    for sheet in result['sheets']:
        # 行2がヘッダー（インデックス1）
        if len(sheet['data']) < 3:
            continue
        
        header_row = sheet['data'][1]  # 行2
        
        # データ行を処理（行3以降）
        for row in sheet['data'][2:]:
            if not row or all(not cell or str(cell).strip() == '' or str(cell) == 'None' for cell in row):
                continue
            
            screen = {}
            for i, cell in enumerate(row):
                if i < len(header_row) and header_row[i] and str(header_row[i]).strip():
                    header = str(header_row[i]).strip()
                    value = str(cell).strip() if cell and str(cell) != 'None' else ''
                    if value:
                        screen[header] = value
            
            if screen:
                screens.append(screen)
    
    return screens


def create_specification(system_overview, screens, output_path="specification.md"):
    """仕様書を作成"""
    lines = []
    
    # タイトル
    lines.append("# オンラインショッピングサイト 仕様書")
    lines.append("")
    lines.append(f"**生成日時**: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
    lines.append("")
    
    # システム概要
    lines.append("## 1. システム概要")
    lines.append("")
    lines.append(system_overview)
    lines.append("")
    
    # 画面一覧サマリー
    lines.append("## 2. 画面一覧")
    lines.append("")
    lines.append(f"本システムは全{len(screens)}画面で構成されています。")
    lines.append("")
    
    # 機能別の画面数を集計
    user_screens = [s for s in screens if s.get('機能') == 'ユーザ用']
    admin_screens = [s for s in screens if s.get('機能') == '管理用']
    common_screens = [s for s in screens if s.get('機能') == '共通']
    
    lines.append("### 機能別画面数")
    lines.append("")
    lines.append(f"- **ユーザ用機能**: {len(user_screens)}画面")
    lines.append(f"- **管理用機能**: {len(admin_screens)}画面")
    lines.append(f"- **共通機能**: {len(common_screens)}画面")
    lines.append("")
    
    # 画面一覧表
    lines.append("### 画面一覧表")
    lines.append("")
    lines.append("| # | 機能 | 画面分類 | 画面ID | タイトル |")
    lines.append("|---|------|----------|--------|----------|")
    
    for screen in screens:
        no = screen.get('#', '')
        func = screen.get('機能', '')
        category = screen.get('', '')  # 3列目（画面分類）
        screen_id = screen.get('画面ID', '')
        title = screen.get('タイトル', '')
        
        lines.append(f"| {no} | {func} | {category} | {screen_id} | {title} |")
    
    lines.append("")
    
    # 画面詳細
    lines.append("## 3. 画面詳細")
    lines.append("")
    
    # ユーザ用機能
    if user_screens:
        lines.append("### 3.1 ユーザ用機能")
        lines.append("")
        
        current_category = None
        for screen in user_screens:
            category = screen.get('', '')  # 3列目
            screen_id = screen.get('画面ID', '')
            title = screen.get('タイトル', '')
            
            # カテゴリが変わったら見出しを追加
            if category and category != current_category:
                lines.append(f"#### {category}")
                lines.append("")
                current_category = category
            
            lines.append(f"**{screen_id}: {title}**")
            lines.append("")
            lines.append(f"- 画面ID: `{screen_id}`")
            lines.append(f"- 画面名: {title}")
            if category:
                lines.append(f"- 機能分類: {category}")
            lines.append("")
    
    # 管理用機能
    if admin_screens:
        lines.append("### 3.2 管理用機能")
        lines.append("")
        
        current_category = None
        for screen in admin_screens:
            category = screen.get('', '')
            screen_id = screen.get('画面ID', '')
            title = screen.get('タイトル', '')
            
            if category and category != current_category:
                lines.append(f"#### {category}")
                lines.append("")
                current_category = category
            
            lines.append(f"**{screen_id}: {title}**")
            lines.append("")
            lines.append(f"- 画面ID: `{screen_id}`")
            lines.append(f"- 画面名: {title}")
            if category:
                lines.append(f"- 機能分類: {category}")
            lines.append("")
    
    # 共通機能
    if common_screens:
        lines.append("### 3.3 共通機能")
        lines.append("")
        
        for screen in common_screens:
            screen_id = screen.get('画面ID', '')
            title = screen.get('タイトル', '')
            
            lines.append(f"**{screen_id}: {title}**")
            lines.append("")
            lines.append(f"- 画面ID: `{screen_id}`")
            lines.append(f"- 画面名: {title}")
            lines.append("")
    
    # ファイルに書き込み
    content = "\n".join(lines)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return content


def main():
    """メイン関数"""
    print("\n" + "="*80)
    print("📝 Excel仕様書から Markdown仕様書を生成")
    print("="*80)
    
    try:
        # システム概要を読み込む
        print("\n📊 システム概要を読み込み中...")
        system_overview = read_system_overview("test_files/04_システム概要.xlsx")
        print(f"✅ 読み込み完了")
        print(f"   概要: {system_overview[:100]}...")
        
        # 画面一覧を読み込む
        print("\n📱 画面一覧を読み込み中...")
        screens = read_screen_list("test_files/06_画面一覧.xlsx")
        print(f"✅ 読み込み完了")
        print(f"   画面数: {len(screens)}")
        
        # 機能別に集計
        user_count = len([s for s in screens if s.get('機能') == 'ユーザ用'])
        admin_count = len([s for s in screens if s.get('機能') == '管理用'])
        common_count = len([s for s in screens if s.get('機能') == '共通'])
        
        print(f"   - ユーザ用: {user_count}画面")
        print(f"   - 管理用: {admin_count}画面")
        print(f"   - 共通: {common_count}画面")
        
        # 仕様書を生成
        print("\n📝 仕様書を生成中...")
        spec_content = create_specification(system_overview, screens)
        print(f"✅ 生成完了: specification.md")
        
        # プレビュー
        print("\n" + "="*80)
        print("📄 生成された仕様書のプレビュー（最初の50行）")
        print("="*80 + "\n")
        
        lines = spec_content.split('\n')
        for line in lines[:50]:
            print(line)
        
        if len(lines) > 50:
            print(f"\n... (残り{len(lines) - 50}行)")
        
        print("\n" + "="*80)
        print("✅ 完了")
        print("="*80)
        
        print("\n💡 このデモで実現したこと:")
        print("  ✓ Excelファイルからシステム概要を抽出")
        print("  ✓ 画面一覧データを構造化して抽出")
        print("  ✓ 機能別に画面を分類")
        print("  ✓ Markdown形式の仕様書を自動生成")
        print("")
        print("🎯 Kiroへの応用シナリオ:")
        print("  1. 既存のExcel設計書を読み込んで内容を理解")
        print("  2. 設計書の内容に基づいて画面のHTMLコードを生成")
        print("  3. 画面遷移ロジックのコードを生成")
        print("  4. テストケースを自動生成")
        print("  5. API仕様書を生成")
        print("  6. データベーススキーマを生成")
        print("")
        print("📂 生成されたファイル:")
        print("  - specification.md (Markdown形式の仕様書)")
        print("="*80 + "\n")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

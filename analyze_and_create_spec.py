"""Excelファイルを読み込んで仕様書を作成するスクリプト"""

import sys
import json
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers import ExcelReader


def read_excel_file(file_path: str):
    """Excelファイルを読み込む"""
    print(f"\n{'='*70}")
    print(f"📊 ファイル読み込み: {file_path}")
    print(f"{'='*70}")
    
    reader = ExcelReader()
    result = reader.read_file(file_path)
    
    print(f"\n✅ 読み込み成功!")
    print(f"シート数: {len(result['sheets'])}")
    
    for sheet in result['sheets']:
        row_count = len(sheet['data'])
        column_count = max(len(row) for row in sheet['data']) if sheet['data'] else 0
        print(f"  - {sheet['name']}: {row_count}行 x {column_count}列")
    
    return result


def display_sheet_content(sheet_data, max_rows=10):
    """シートの内容を表示"""
    print(f"\n【シート: {sheet_data['name']}】")
    
    if not sheet_data['data']:
        print("  (空のシート)")
        return
    
    row_count = len(sheet_data['data'])
    column_count = max(len(row) for row in sheet_data['data']) if sheet_data['data'] else 0
    print(f"サイズ: {row_count}行 x {column_count}列")
    
    print("\nデータ:")
    for i, row in enumerate(sheet_data['data'][:max_rows], 1):
        # 空の行はスキップ
        if all(cell == '' or cell == 'None' for cell in row):
            continue
        
        row_str = " | ".join(str(cell)[:30] for cell in row[:10])
        print(f"  {i:3d}. {row_str}")
    
    if row_count > max_rows:
        print(f"  ... (残り{row_count - max_rows}行)")


def analyze_system_overview(data):
    """システム概要ファイルを分析"""
    print(f"\n{'='*70}")
    print("📋 システム概要の分析")
    print(f"{'='*70}")
    
    for sheet in data['sheets']:
        display_sheet_content(sheet, max_rows=20)


def analyze_screen_transition(data):
    """画面遷移図ファイルを分析"""
    print(f"\n{'='*70}")
    print("🔄 画面遷移図の分析")
    print(f"{'='*70}")
    
    for sheet in data['sheets']:
        display_sheet_content(sheet, max_rows=20)


def analyze_screen_list(data):
    """画面一覧ファイルを分析"""
    print(f"\n{'='*70}")
    print("📱 画面一覧の分析")
    print(f"{'='*70}")
    
    for sheet in data['sheets']:
        display_sheet_content(sheet, max_rows=30)


def extract_system_info(system_data):
    """システム概要から情報を抽出"""
    info = {
        "system_name": "",
        "description": "",
        "features": [],
        "technologies": []
    }
    
    for sheet in system_data['sheets']:
        for row in sheet['data']:
            # 空行をスキップ
            if not row or all(cell == '' or cell == 'None' for cell in row):
                continue
            
            # システム名を探す
            if len(row) > 0 and 'システム名' in str(row[0]):
                if len(row) > 1:
                    info['system_name'] = str(row[1])
            
            # 概要を探す
            if len(row) > 0 and ('概要' in str(row[0]) or '説明' in str(row[0])):
                if len(row) > 1:
                    info['description'] = str(row[1])
    
    return info


def extract_screens(screen_list_data):
    """画面一覧から画面情報を抽出"""
    screens = []
    
    for sheet in screen_list_data['sheets']:
        # ヘッダー行を探す
        header_row = None
        data_start_idx = 0
        
        for i, row in enumerate(sheet['data']):
            if any('画面' in str(cell) or 'ID' in str(cell) or '名称' in str(cell) for cell in row):
                header_row = row
                data_start_idx = i + 1
                break
        
        if header_row:
            # データ行を処理
            for row in sheet['data'][data_start_idx:]:
                if row and any(cell != '' and cell != 'None' for cell in row):
                    screen = {}
                    for j, cell in enumerate(row):
                        if j < len(header_row):
                            key = str(header_row[j]).strip()
                            if key:
                                screen[key] = str(cell).strip()
                    
                    if screen:
                        screens.append(screen)
    
    return screens


def create_specification_document(system_info, screens, output_path="generated_spec.md"):
    """仕様書を作成"""
    print(f"\n{'='*70}")
    print("📝 仕様書の作成")
    print(f"{'='*70}")
    
    content = []
    
    # タイトル
    system_name = system_info.get('system_name', 'システム')
    content.append(f"# {system_name} 仕様書")
    content.append("")
    content.append(f"生成日時: {Path(__file__).stat().st_mtime}")
    content.append("")
    
    # システム概要
    content.append("## システム概要")
    content.append("")
    if system_info.get('description'):
        content.append(system_info['description'])
    else:
        content.append("(システムの概要説明)")
    content.append("")
    
    # 画面一覧
    if screens:
        content.append("## 画面一覧")
        content.append("")
        content.append(f"全{len(screens)}画面")
        content.append("")
        
        # 表形式で出力
        if screens:
            # ヘッダーを取得
            headers = list(screens[0].keys())
            
            # マークダウン表のヘッダー
            content.append("| " + " | ".join(headers) + " |")
            content.append("| " + " | ".join(["---"] * len(headers)) + " |")
            
            # データ行
            for screen in screens:
                row_data = [screen.get(h, "") for h in headers]
                content.append("| " + " | ".join(row_data) + " |")
            
            content.append("")
    
    # 画面詳細
    if screens:
        content.append("## 画面詳細")
        content.append("")
        
        for i, screen in enumerate(screens, 1):
            screen_id = screen.get('画面ID', screen.get('ID', f'画面{i}'))
            screen_name = screen.get('画面名', screen.get('名称', ''))
            
            content.append(f"### {screen_id}: {screen_name}")
            content.append("")
            
            # 画面情報を表示
            for key, value in screen.items():
                if key not in ['画面ID', 'ID', '画面名', '名称'] and value:
                    content.append(f"- **{key}**: {value}")
            
            content.append("")
    
    # ファイルに書き込み
    spec_content = "\n".join(content)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print(f"\n✅ 仕様書を作成しました: {output_path}")
    print(f"   - システム名: {system_name}")
    print(f"   - 画面数: {len(screens)}")
    
    return spec_content


def main():
    """メイン関数"""
    print("\n" + "="*70)
    print("Excel仕様書読み込み & 仕様書生成デモ")
    print("="*70)
    
    try:
        # ファイルを読み込む
        system_overview = read_excel_file("test_files/04_システム概要.xlsx")
        screen_transition = read_excel_file("test_files/05_画面遷移図.xlsx")
        screen_list = read_excel_file("test_files/06_画面一覧.xlsx")
        
        # 内容を分析
        analyze_system_overview(system_overview)
        analyze_screen_transition(screen_transition)
        analyze_screen_list(screen_list)
        
        # 情報を抽出
        print(f"\n{'='*70}")
        print("🔍 情報の抽出")
        print(f"{'='*70}")
        
        system_info = extract_system_info(system_overview)
        print(f"\nシステム情報:")
        print(f"  システム名: {system_info.get('system_name', '(未設定)')}")
        print(f"  概要: {system_info.get('description', '(未設定)')[:100]}...")
        
        screens = extract_screens(screen_list)
        print(f"\n画面情報:")
        print(f"  抽出された画面数: {len(screens)}")
        if screens:
            print(f"  サンプル: {list(screens[0].keys())}")
        
        # 仕様書を作成
        spec_content = create_specification_document(system_info, screens)
        
        # 作成した仕様書の一部を表示
        print(f"\n{'='*70}")
        print("📄 生成された仕様書のプレビュー")
        print(f"{'='*70}")
        lines = spec_content.split('\n')
        for line in lines[:30]:
            print(line)
        
        if len(lines) > 30:
            print(f"\n... (残り{len(lines) - 30}行)")
        
        print(f"\n{'='*70}")
        print("✅ 完了")
        print(f"{'='*70}")
        print("\n💡 このデモで示したこと:")
        print("  1. 複数のExcelファイルからデータを読み込み")
        print("  2. シート内のデータを構造化して抽出")
        print("  3. 抽出したデータから仕様書（Markdown）を自動生成")
        print("")
        print("🎯 Kiroへの応用:")
        print("  - 既存の設計書を読み込んで理解")
        print("  - 設計書の内容に基づいたコード生成")
        print("  - 仕様書の自動生成・更新")
        print("  - データの整形・変換")
        print("="*70 + "\n")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

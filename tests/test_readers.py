"""リーダー機能のテストスクリプト"""

import sys
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from document_format_mcp_server.readers import (
    PowerPointReader,
    WordReader,
    ExcelReader,
    GoogleWorkspaceReader,
)
from document_format_mcp_server.utils.logging_config import setup_logging, get_logger


# ロガーの設定
setup_logging()
logger = get_logger(__name__)


def test_powerpoint_reader(file_path: str):
    """PowerPointリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"PowerPointファイルを読み込み中: {file_path}")
    logger.info("="*60)
    
    try:
        reader = PowerPointReader()
        result = reader.read_file(file_path)
        
        logger.info("読み込み成功!")
        logger.info(f"スライド数: {len(result['slides'])}")
        
        for slide in result['slides'][:3]:  # 最初の3スライドのみ表示
            logger.info(f"--- スライド {slide['slide_number']} ---")
            logger.info(f"タイトル: {slide['title']}")
            content_preview = slide['content'][:100] + "..." if len(slide['content']) > 100 else slide['content']
            logger.info(f"コンテンツ: {content_preview}")
            if slide['tables']:
                logger.info(f"表の数: {len(slide['tables'])}")
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def test_word_reader(file_path: str):
    """Wordリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"Wordファイルを読み込み中: {file_path}")
    logger.info("="*60)
    
    try:
        reader = WordReader()
        result = reader.read_file(file_path)
        
        logger.info("読み込み成功!")
        logger.info(f"段落数: {len(result['paragraphs'])}")
        
        for para in result['paragraphs'][:5]:  # 最初の5段落のみ表示
            para_type = "heading" if para.get('level', 0) > 0 else "paragraph"
            logger.info(f"--- {para_type} ---")
            if para.get('level', 0) > 0:
                logger.info(f"レベル: {para['level']}")
            text_preview = para['text'][:100] + "..." if len(para['text']) > 100 else para['text']
            logger.info(f"テキスト: {text_preview}")
        
        if result['tables']:
            logger.info(f"表の数: {len(result['tables'])}")
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def test_excel_reader(file_path: str):
    """Excelリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"Excelファイルを読み込み中: {file_path}")
    logger.info("="*60)
    
    try:
        reader = ExcelReader()
        result = reader.read_file(file_path)
        
        logger.info("読み込み成功!")
        logger.info(f"シート数: {len(result['sheets'])}")
        
        for sheet in result['sheets'][:3]:  # 最初の3シートのみ表示
            logger.info(f"--- シート: {sheet['name']} ---")
            row_count = len(sheet['data'])
            col_count = max(len(row) for row in sheet['data']) if sheet['data'] else 0
            logger.info(f"行数: {row_count}, 列数: {col_count}")
            
            # 最初の数行を表示
            for i, row in enumerate(sheet['data'][:3]):
                logger.info(f"行{i+1}: {row[:5]}")  # 最初の5列のみ
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def test_google_spreadsheet(file_id_or_url: str, credentials_path: str):
    """Googleスプレッドシートリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"Googleスプレッドシートを読み込み中: {file_id_or_url}")
    logger.info("="*60)
    
    try:
        reader = GoogleWorkspaceReader(credentials_path)
        result = reader.read_spreadsheet(file_id_or_url)
        
        logger.info("読み込み成功!")
        logger.info(f"タイトル: {result['title']}")
        logger.info(f"シート数: {len(result['sheets'])}")
        
        for sheet in result['sheets'][:2]:  # 最初の2シートのみ表示
            logger.info(f"--- シート: {sheet['name']} ---")
            logger.info(f"行数: {sheet['row_count']}, 列数: {sheet['column_count']}")
            
            # 最初の数行を表示
            for i, row in enumerate(sheet['data'][:3]):
                logger.info(f"行{i+1}: {row[:5]}")  # 最初の5列のみ
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def test_google_document(file_id_or_url: str, credentials_path: str):
    """Googleドキュメントリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"Googleドキュメントを読み込み中: {file_id_or_url}")
    logger.info("="*60)
    
    try:
        reader = GoogleWorkspaceReader(credentials_path)
        result = reader.read_document(file_id_or_url)
        
        logger.info("読み込み成功!")
        logger.info(f"タイトル: {result['title']}")
        logger.info(f"コンテンツ要素数: {len(result['content'])}")
        
        for item in result['content'][:5]:  # 最初の5要素のみ表示
            logger.info(f"--- {item['type']} ---")
            if item['type'] in ['paragraph', 'heading']:
                text_preview = item['text'][:100] + "..." if len(item['text']) > 100 else item['text']
                logger.info(f"テキスト: {text_preview}")
            elif item['type'] == 'table':
                logger.info(f"表: {item['rows']}行 x {item['columns']}列")
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def test_google_slides(file_id_or_url: str, credentials_path: str):
    """Googleスライドリーダーのテスト"""
    logger.info("="*60)
    logger.info(f"Googleスライドを読み込み中: {file_id_or_url}")
    logger.info("="*60)
    
    try:
        reader = GoogleWorkspaceReader(credentials_path)
        result = reader.read_slides(file_id_or_url)
        
        logger.info("読み込み成功!")
        logger.info(f"タイトル: {result['title']}")
        logger.info(f"スライド数: {len(result['slides'])}")
        
        for slide in result['slides'][:3]:  # 最初の3スライドのみ表示
            logger.info(f"--- スライド {slide['slide_number']} ---")
            logger.info(f"要素数: {len(slide['elements'])}")
            
            for element in slide['elements'][:3]:  # 最初の3要素のみ
                if element['type'] == 'text':
                    text_preview = element['content'][:100] + "..." if len(element['content']) > 100 else element['content']
                    logger.info(f"テキスト: {text_preview}")
                elif element['type'] == 'table':
                    logger.info(f"表: {element['content']['rows']}行 x {element['content']['columns']}列")
        
        return True
    except Exception as e:
        logger.error(f"エラー: {e}", exc_info=True)
        return False


def main():
    """メイン関数"""
    logger.info("="*60)
    logger.info("Document Format MCP Server - リーダー機能テスト")
    logger.info("="*60)
    
    logger.info("使用方法:")
    logger.info("1. ローカルファイルのテスト:")
    logger.info("   python test_readers.py")
    logger.info("2. Googleファイルのテスト:")
    logger.info("   python test_readers.py --google")
    logger.info("注意: テストファイルのパスを編集してください")
    
    import sys
    
    if "--google" in sys.argv:
        # Google Workspaceのテスト
        credentials_path = "path/to/google-credentials.json"  # ここを編集
        
        logger.info("Google認証情報のパスを設定してください:")
        logger.info(f"   credentials_path = '{credentials_path}'")
        
        # テストするファイルのURLまたはID
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
        document_url = "https://docs.google.com/document/d/YOUR_DOCUMENT_ID/edit"
        slides_url = "https://docs.google.com/presentation/d/YOUR_SLIDES_ID/edit"
        
        logger.info("テストするGoogleファイルのURLを設定してください")
        logger.debug(f"スプレッドシート: {spreadsheet_url}")
        logger.debug(f"ドキュメント: {document_url}")
        logger.debug(f"スライド: {slides_url}")
        
        # test_google_spreadsheet(spreadsheet_url, credentials_path)
        # test_google_document(document_url, credentials_path)
        # test_google_slides(slides_url, credentials_path)
        
    else:
        # ローカルファイルのテスト
        logger.info("テストファイルのパスを設定してください:")
        
        # テストファイルのパス（ここを編集）
        pptx_file = "test_files/sample.pptx"
        docx_file = "test_files/sample.docx"
        xlsx_file = "test_files/sample.xlsx"
        
        logger.info(f"   PowerPoint: {pptx_file}")
        logger.info(f"   Word: {docx_file}")
        logger.info(f"   Excel: {xlsx_file}")
        
        # ファイルが存在する場合のみテスト
        if Path(pptx_file).exists():
            test_powerpoint_reader(pptx_file)
        else:
            logger.warning(f"PowerPointファイルが見つかりません: {pptx_file}")
        
        if Path(docx_file).exists():
            test_word_reader(docx_file)
        else:
            logger.warning(f"Wordファイルが見つかりません: {docx_file}")
        
        if Path(xlsx_file).exists():
            test_excel_reader(xlsx_file)
        else:
            logger.warning(f"Excelファイルが見つかりません: {xlsx_file}")
    
    logger.info("="*60)
    logger.info("テスト完了")
    logger.info("="*60)


if __name__ == "__main__":
    main()

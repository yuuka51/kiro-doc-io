"""Google Workspaceリーダーのユニットテスト

要件4.1: GoogleスプレッドシートのURLまたはファイルIDからデータを取得する
要件4.5: Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、
        明確なエラーメッセージをユーザーに返す
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import os

from document_format_mcp_server.readers.google_reader import GoogleWorkspaceReader
from document_format_mcp_server.utils.models import ReadResult, DocumentContent
from document_format_mcp_server.utils.errors import (
    AuthenticationError,
    ConfigurationError,
    APIError,
)


# テスト用のモックデータ
MOCK_SPREADSHEET_DATA = {
    'properties': {'title': 'テストスプレッドシート'},
    'sheets': [
        {'properties': {'title': 'シート1'}},
        {'properties': {'title': 'シート2'}},
    ]
}

MOCK_SHEET_VALUES = {
    'values': [
        ['ヘッダー1', 'ヘッダー2', 'ヘッダー3'],
        ['データ1', 'データ2', 'データ3'],
        ['データ4', 'データ5', 'データ6'],
    ]
}


MOCK_DOCUMENT_DATA = {
    'title': 'テストドキュメント',
    'body': {
        'content': [
            {
                'paragraph': {
                    'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                    'elements': [
                        {'textRun': {'content': '見出し1\n'}}
                    ]
                }
            },
            {
                'paragraph': {
                    'paragraphStyle': {'namedStyleType': 'NORMAL_TEXT'},
                    'elements': [
                        {'textRun': {'content': '本文テキスト\n'}}
                    ]
                }
            },
        ]
    }
}

MOCK_SLIDES_DATA = {
    'title': 'テストプレゼンテーション',
    'slides': [
        {
            'pageElements': [
                {
                    'shape': {
                        'text': {
                            'textElements': [
                                {'textRun': {'content': 'スライド1タイトル\n'}}
                            ]
                        }
                    }
                }
            ]
        },
        {
            'pageElements': [
                {
                    'shape': {
                        'text': {
                            'textElements': [
                                {'textRun': {'content': 'スライド2タイトル\n'}}
                            ]
                        }
                    }
                }
            ]
        },
    ]
}


class MockCredentials:
    """モック認証情報クラス"""
    valid = True
    expired = False
    refresh_token = None
    
    def to_json(self):
        return '{"mock": "credentials"}'


class MockHttpError(Exception):
    """モックHTTPエラークラス"""
    def __init__(self, status):
        self.resp = Mock()
        self.resp.status = status
        super().__init__(f"HTTP Error {status}")


@pytest.fixture
def mock_credentials_file(tmp_path):
    """一時的な認証情報ファイルを作成するフィクスチャ"""
    creds_file = tmp_path / "credentials.json"
    creds_file.write_text('{"installed": {"client_id": "test", "client_secret": "test"}}')
    return str(creds_file)


@pytest.fixture
def mock_token_file(tmp_path):
    """一時的なトークンファイルを作成するフィクスチャ"""
    token_file = tmp_path / "token.json"
    token_file.write_text('{"token": "test", "refresh_token": "test"}')
    return str(token_file)


@pytest.fixture
def google_reader(mock_credentials_file, tmp_path):
    """モック認証を使用したGoogleWorkspaceReaderインスタンスを返すフィクスチャ"""
    # トークンファイルも作成
    token_file = tmp_path / "token.json"
    token_file.write_text('{"token": "test", "refresh_token": "test"}')
    
    with patch('document_format_mcp_server.readers.google_reader.Credentials') as mock_creds_class:
        mock_creds = MockCredentials()
        mock_creds_class.from_authorized_user_file.return_value = mock_creds
        
        reader = GoogleWorkspaceReader(
            credentials_path=mock_credentials_file,
            api_timeout_seconds=30,
            max_retries=3
        )
        return reader


# ===== 認証関連のテスト =====

def test_google_reader_raises_config_error_for_missing_credentials():
    """認証情報ファイルが存在しない場合にConfigurationErrorが発生することを検証（要件4.5）"""
    with pytest.raises(ConfigurationError) as exc_info:
        GoogleWorkspaceReader(
            credentials_path="/nonexistent/path/credentials.json"
        )
    
    assert "見つかりません" in str(exc_info.value.message)


def test_google_reader_initialization_with_valid_credentials(mock_credentials_file, tmp_path):
    """有効な認証情報でGoogleWorkspaceReaderが初期化できることを検証"""
    # トークンファイルを作成
    token_file = tmp_path / "token.json"
    token_file.write_text('{"token": "test", "refresh_token": "test"}')
    
    with patch('document_format_mcp_server.readers.google_reader.Credentials') as mock_creds_class:
        mock_creds = MockCredentials()
        mock_creds_class.from_authorized_user_file.return_value = mock_creds
        
        reader = GoogleWorkspaceReader(
            credentials_path=mock_credentials_file,
            api_timeout_seconds=60,
            max_retries=3
        )
        
        assert reader.api_timeout == 60
        assert reader.max_retries == 3
        assert reader.credentials is not None


# ===== ファイルID抽出のテスト =====

def test_extract_file_id_from_url(google_reader):
    """URLからファイルIDを正しく抽出できることを検証（要件4.1）"""
    # スプレッドシートURL
    url1 = "https://docs.google.com/spreadsheets/d/1abc123xyz/edit"
    assert google_reader._extract_file_id(url1) == "1abc123xyz"
    
    # ドキュメントURL
    url2 = "https://docs.google.com/document/d/2def456abc/edit"
    assert google_reader._extract_file_id(url2) == "2def456abc"
    
    # スライドURL
    url3 = "https://docs.google.com/presentation/d/3ghi789def/edit"
    assert google_reader._extract_file_id(url3) == "3ghi789def"


def test_extract_file_id_from_direct_id(google_reader):
    """直接ファイルIDを渡した場合にそのまま返されることを検証（要件4.1）"""
    file_id = "1abc123xyz"
    assert google_reader._extract_file_id(file_id) == file_id


def test_extract_file_id_from_url_with_query_params(google_reader):
    """クエリパラメータ付きURLからファイルIDを抽出できることを検証"""
    url = "https://docs.google.com/spreadsheets/d/1abc123xyz/edit?usp=sharing"
    assert google_reader._extract_file_id(url) == "1abc123xyz"


# ===== スプレッドシート読み取りのテスト =====

def test_read_spreadsheet_returns_read_result(google_reader):
    """read_spreadsheetがReadResultを返すことを検証（要件4.1）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        # モックサービスを設定
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        # スプレッドシートメタデータのモック
        mock_service.spreadsheets().get().execute.return_value = MOCK_SPREADSHEET_DATA
        
        # シートデータのモック
        mock_service.spreadsheets().values().get().execute.return_value = MOCK_SHEET_VALUES
        
        result = google_reader.read_spreadsheet("test_file_id")
        
        assert isinstance(result, ReadResult)
        assert hasattr(result, 'success')
        assert hasattr(result, 'content')
        assert hasattr(result, 'error')
        assert hasattr(result, 'file_path')


def test_read_spreadsheet_success_structure(google_reader):
    """スプレッドシート読み取り成功時のレスポンス構造を検証（要件4.1）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.spreadsheets().get().execute.return_value = MOCK_SPREADSHEET_DATA
        mock_service.spreadsheets().values().get().execute.return_value = MOCK_SHEET_VALUES
        
        result = google_reader.read_spreadsheet("test_file_id")
        
        # 成功フラグを検証
        assert result.success is True
        assert result.error is None
        
        # DocumentContentを検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "google_sheets"
        assert isinstance(result.content.metadata, dict)
        assert isinstance(result.content.content, dict)
        
        # メタデータの検証
        assert "title" in result.content.metadata
        assert "sheet_count" in result.content.metadata
        assert "file_id" in result.content.metadata
        
        # コンテンツの検証
        assert "sheets" in result.content.content
        assert isinstance(result.content.content["sheets"], list)


def test_read_spreadsheet_extracts_all_sheets(google_reader):
    """スプレッドシートの全シートが抽出されることを検証（要件4.1）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.spreadsheets().get().execute.return_value = MOCK_SPREADSHEET_DATA
        mock_service.spreadsheets().values().get().execute.return_value = MOCK_SHEET_VALUES
        
        result = google_reader.read_spreadsheet("test_file_id")
        
        assert result.success is True
        sheets = result.content.content["sheets"]
        
        # 2つのシートが抽出されることを検証
        assert len(sheets) == 2
        assert sheets[0]["name"] == "シート1"
        assert sheets[1]["name"] == "シート2"


def test_read_spreadsheet_not_found_error(google_reader):
    """存在しないスプレッドシートの場合のエラーハンドリングを検証（要件4.5）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        from googleapiclient.errors import HttpError
        
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        # 404エラーをシミュレート
        mock_resp = Mock()
        mock_resp.status = 404
        mock_service.spreadsheets().get().execute.side_effect = HttpError(
            resp=mock_resp, content=b'Not Found'
        )
        
        result = google_reader.read_spreadsheet("nonexistent_file_id")
        
        assert result.success is False
        assert result.error is not None
        assert "見つかりません" in result.error


def test_read_spreadsheet_permission_error(google_reader):
    """アクセス権限がない場合のエラーハンドリングを検証（要件4.5）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        from googleapiclient.errors import HttpError
        
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        # 403エラーをシミュレート
        mock_resp = Mock()
        mock_resp.status = 403
        mock_service.spreadsheets().get().execute.side_effect = HttpError(
            resp=mock_resp, content=b'Forbidden'
        )
        
        result = google_reader.read_spreadsheet("forbidden_file_id")
        
        assert result.success is False
        assert result.error is not None
        assert "権限" in result.error


# ===== ドキュメント読み取りのテスト =====

def test_read_document_returns_read_result(google_reader):
    """read_documentがReadResultを返すことを検証（要件4.2）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.documents().get().execute.return_value = MOCK_DOCUMENT_DATA
        
        result = google_reader.read_document("test_file_id")
        
        assert isinstance(result, ReadResult)
        assert result.success is True


def test_read_document_success_structure(google_reader):
    """ドキュメント読み取り成功時のレスポンス構造を検証（要件4.2）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.documents().get().execute.return_value = MOCK_DOCUMENT_DATA
        
        result = google_reader.read_document("test_file_id")
        
        # 成功フラグを検証
        assert result.success is True
        assert result.error is None
        
        # DocumentContentを検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "google_docs"
        
        # メタデータの検証
        assert "title" in result.content.metadata
        assert result.content.metadata["title"] == "テストドキュメント"
        
        # コンテンツの検証
        assert "content" in result.content.content
        content_items = result.content.content["content"]
        assert len(content_items) == 2


def test_read_document_extracts_headings(google_reader):
    """ドキュメントの見出しが正しく抽出されることを検証（要件4.2）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.documents().get().execute.return_value = MOCK_DOCUMENT_DATA
        
        result = google_reader.read_document("test_file_id")
        
        assert result.success is True
        content_items = result.content.content["content"]
        
        # 見出しが正しく抽出されることを検証
        heading = content_items[0]
        assert heading["type"] == "heading"
        assert heading["text"] == "見出し1"
        assert heading["level"] == 1


def test_read_document_not_found_error(google_reader):
    """存在しないドキュメントの場合のエラーハンドリングを検証（要件4.5）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        from googleapiclient.errors import HttpError
        
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        mock_resp = Mock()
        mock_resp.status = 404
        mock_service.documents().get().execute.side_effect = HttpError(
            resp=mock_resp, content=b'Not Found'
        )
        
        result = google_reader.read_document("nonexistent_file_id")
        
        assert result.success is False
        assert "見つかりません" in result.error


# ===== スライド読み取りのテスト =====

def test_read_slides_returns_read_result(google_reader):
    """read_slidesがReadResultを返すことを検証（要件4.3）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.presentations().get().execute.return_value = MOCK_SLIDES_DATA
        
        result = google_reader.read_slides("test_file_id")
        
        assert isinstance(result, ReadResult)
        assert result.success is True


def test_read_slides_success_structure(google_reader):
    """スライド読み取り成功時のレスポンス構造を検証（要件4.3）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.presentations().get().execute.return_value = MOCK_SLIDES_DATA
        
        result = google_reader.read_slides("test_file_id")
        
        # 成功フラグを検証
        assert result.success is True
        assert result.error is None
        
        # DocumentContentを検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "google_slides"
        
        # メタデータの検証
        assert "title" in result.content.metadata
        assert "slide_count" in result.content.metadata
        assert result.content.metadata["slide_count"] == 2
        
        # コンテンツの検証
        assert "slides" in result.content.content
        slides = result.content.content["slides"]
        assert len(slides) == 2


def test_read_slides_extracts_slide_content(google_reader):
    """スライドのコンテンツが正しく抽出されることを検証（要件4.3）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.presentations().get().execute.return_value = MOCK_SLIDES_DATA
        
        result = google_reader.read_slides("test_file_id")
        
        assert result.success is True
        slides = result.content.content["slides"]
        
        # スライド番号が正しく設定されることを検証
        assert slides[0]["slide_number"] == 1
        assert slides[1]["slide_number"] == 2
        
        # 要素が抽出されることを検証
        assert len(slides[0]["elements"]) > 0


def test_read_slides_not_found_error(google_reader):
    """存在しないスライドの場合のエラーハンドリングを検証（要件4.5）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        from googleapiclient.errors import HttpError
        
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        mock_resp = Mock()
        mock_resp.status = 404
        mock_service.presentations().get().execute.side_effect = HttpError(
            resp=mock_resp, content=b'Not Found'
        )
        
        result = google_reader.read_slides("nonexistent_file_id")
        
        assert result.success is False
        assert "見つかりません" in result.error


def test_read_slides_permission_error(google_reader):
    """アクセス権限がない場合のエラーハンドリングを検証（要件4.5）"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        from googleapiclient.errors import HttpError
        
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        mock_resp = Mock()
        mock_resp.status = 403
        mock_service.presentations().get().execute.side_effect = HttpError(
            resp=mock_resp, content=b'Forbidden'
        )
        
        result = google_reader.read_slides("forbidden_file_id")
        
        assert result.success is False
        assert "権限" in result.error


# ===== リトライロジックのテスト =====

def test_execute_with_retry_success_on_first_attempt(google_reader):
    """最初の試行で成功した場合にリトライしないことを検証"""
    mock_func = Mock(return_value="success")
    
    result = google_reader._execute_with_retry(mock_func)
    
    assert result == "success"
    assert mock_func.call_count == 1


def test_execute_with_retry_retries_on_server_error(google_reader):
    """サーバーエラー時にリトライすることを検証"""
    from googleapiclient.errors import HttpError
    
    mock_resp = Mock()
    mock_resp.status = 500
    
    # 最初の2回は失敗、3回目で成功
    mock_func = Mock(side_effect=[
        HttpError(resp=mock_resp, content=b'Server Error'),
        HttpError(resp=mock_resp, content=b'Server Error'),
        "success"
    ])
    
    with patch('time.sleep'):  # sleepをスキップ
        result = google_reader._execute_with_retry(mock_func)
    
    assert result == "success"
    assert mock_func.call_count == 3


def test_execute_with_retry_raises_after_max_retries(google_reader):
    """最大リトライ回数後にAPIErrorが発生することを検証"""
    from googleapiclient.errors import HttpError
    
    mock_resp = Mock()
    mock_resp.status = 500
    
    mock_func = Mock(side_effect=HttpError(resp=mock_resp, content=b'Server Error'))
    
    with patch('time.sleep'):  # sleepをスキップ
        with pytest.raises(APIError) as exc_info:
            google_reader._execute_with_retry(mock_func)
    
    assert "3回失敗" in str(exc_info.value.message)
    assert mock_func.call_count == google_reader.max_retries


def test_execute_with_retry_no_retry_on_404(google_reader):
    """404エラー時にリトライしないことを検証"""
    from googleapiclient.errors import HttpError
    
    mock_resp = Mock()
    mock_resp.status = 404
    
    mock_func = Mock(side_effect=HttpError(resp=mock_resp, content=b'Not Found'))
    
    with pytest.raises(HttpError):
        google_reader._execute_with_retry(mock_func)
    
    # 404はリトライ不可能なので1回のみ呼び出される
    assert mock_func.call_count == 1


def test_execute_with_retry_no_retry_on_403(google_reader):
    """403エラー時にリトライしないことを検証"""
    from googleapiclient.errors import HttpError
    
    mock_resp = Mock()
    mock_resp.status = 403
    
    mock_func = Mock(side_effect=HttpError(resp=mock_resp, content=b'Forbidden'))
    
    with pytest.raises(HttpError):
        google_reader._execute_with_retry(mock_func)
    
    # 403はリトライ不可能なので1回のみ呼び出される
    assert mock_func.call_count == 1


# ===== ファイルパスの検証 =====

def test_file_path_in_result(google_reader):
    """ReadResultにファイルIDが含まれることを検証"""
    with patch('document_format_mcp_server.readers.google_reader.build') as mock_build:
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_service.spreadsheets().get().execute.return_value = MOCK_SPREADSHEET_DATA
        mock_service.spreadsheets().values().get().execute.return_value = MOCK_SHEET_VALUES
        
        file_id = "test_file_id_123"
        result = google_reader.read_spreadsheet(file_id)
        
        assert result.file_path == file_id

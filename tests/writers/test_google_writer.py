"""Google Workspaceライターのユニットテスト

要件8.1: Google Sheets APIを使用して新しいスプレッドシートを作成する
要件8.2: Google Docs APIを使用して新しいドキュメントを作成する
要件8.3: Google Slides APIを使用して新しいプレゼンテーションを作成する
要件8.4: 作成したファイルのURLをユーザーに返す
要件8.5: Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、
        明確なエラーメッセージをユーザーに返す
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import os

from document_format_mcp_server.writers.google_writer import GoogleWorkspaceWriter
from document_format_mcp_server.utils.models import WriteResult
from document_format_mcp_server.utils.errors import (
    AuthenticationError,
    ConfigurationError,
    APIError,
)


# テスト用のモックデータ
MOCK_SPREADSHEET_RESPONSE = {
    'spreadsheetId': 'test_spreadsheet_id_123',
    'spreadsheetUrl': 'https://docs.google.com/spreadsheets/d/test_spreadsheet_id_123/edit'
}

MOCK_DOCUMENT_RESPONSE = {
    'documentId': 'test_document_id_456'
}

MOCK_PRESENTATION_RESPONSE = {
    'presentationId': 'test_presentation_id_789',
    'slides': [
        {
            'pageElements': [
                {
                    'objectId': 'title_placeholder',
                    'shape': {
                        'placeholder': {'type': 'CENTERED_TITLE'}
                    }
                },
                {
                    'objectId': 'subtitle_placeholder',
                    'shape': {
                        'placeholder': {'type': 'SUBTITLE'}
                    }
                }
            ]
        }
    ]
}


class MockCredentials:
    """モック認証情報クラス"""
    valid = True
    expired = False
    refresh_token = None
    
    def to_json(self):
        return '{"mock": "credentials"}'


@pytest.fixture
def mock_credentials_file(tmp_path):
    """一時的な認証情報ファイルを作成するフィクスチャ"""
    creds_file = tmp_path / "credentials.json"
    creds_file.write_text('{"installed": {"client_id": "test", "client_secret": "test"}}')
    return str(creds_file)


@pytest.fixture
def mock_token_file(tmp_path):
    """一時的なトークンファイルを作成するフィクスチャ"""
    token_file = tmp_path / "token_writer.json"
    token_file.write_text('{"token": "test", "refresh_token": "test"}')
    return str(token_file)


@pytest.fixture
def google_writer(mock_credentials_file, tmp_path):
    """モック認証を使用したGoogleWorkspaceWriterインスタンスを返すフィクスチャ"""
    # トークンファイルも作成
    token_file = tmp_path / "token_writer.json"
    token_file.write_text('{"token": "test", "refresh_token": "test"}')
    
    with patch('document_format_mcp_server.writers.google_writer.Credentials') as mock_creds_class:
        mock_creds = MockCredentials()
        mock_creds_class.from_authorized_user_file.return_value = mock_creds
        
        writer = GoogleWorkspaceWriter(
            credentials_path=mock_credentials_file,
            api_timeout_seconds=30,
            max_retries=3
        )
        return writer


@pytest.fixture
def basic_spreadsheet_data():
    """基本的なスプレッドシートデータを返すフィクスチャ"""
    return {
        "sheets": [
            {
                "name": "シート1",
                "data": [
                    ["ヘッダー1", "ヘッダー2", "ヘッダー3"],
                    ["データ1", "データ2", "データ3"],
                    ["データ4", "データ5", "データ6"]
                ]
            }
        ]
    }


@pytest.fixture
def multi_sheet_data():
    """複数シートのデータを返すフィクスチャ"""
    return {
        "sheets": [
            {
                "name": "売上データ",
                "data": [
                    ["月", "売上", "利益"],
                    ["1月", 100000, 20000],
                    ["2月", 120000, 25000]
                ]
            },
            {
                "name": "経費データ",
                "data": [
                    ["項目", "金額"],
                    ["人件費", 50000],
                    ["設備費", 30000]
                ]
            }
        ]
    }


@pytest.fixture
def basic_document_data():
    """基本的なドキュメントデータを返すフィクスチャ"""
    return {
        "sections": [
            {
                "heading": "セクション1",
                "level": 1,
                "paragraphs": ["これはテスト段落です。"]
            }
        ]
    }


@pytest.fixture
def basic_slides_data():
    """基本的なスライドデータを返すフィクスチャ"""
    return {
        "slides": [
            {
                "layout": "title",
                "title": "プレゼンテーションタイトル",
                "content": "サブタイトル"
            },
            {
                "layout": "content",
                "title": "スライド2",
                "content": "コンテンツ内容"
            }
        ]
    }


# ===== 認証関連のテスト =====

class TestGoogleWriterAuthentication:
    """認証関連のテスト"""

    def test_raises_config_error_for_missing_credentials(self):
        """認証情報ファイルが存在しない場合にConfigurationErrorが発生することを検証（要件8.5）"""
        with pytest.raises(ConfigurationError) as exc_info:
            GoogleWorkspaceWriter(
                credentials_path="/nonexistent/path/credentials.json"
            )
        
        assert "見つかりません" in str(exc_info.value.message)

    def test_initialization_with_valid_credentials(self, mock_credentials_file, tmp_path):
        """有効な認証情報でGoogleWorkspaceWriterが初期化できることを検証"""
        # トークンファイルを作成
        token_file = tmp_path / "token_writer.json"
        token_file.write_text('{"token": "test", "refresh_token": "test"}')
        
        with patch('document_format_mcp_server.writers.google_writer.Credentials') as mock_creds_class:
            mock_creds = MockCredentials()
            mock_creds_class.from_authorized_user_file.return_value = mock_creds
            
            writer = GoogleWorkspaceWriter(
                credentials_path=mock_credentials_file,
                api_timeout_seconds=60,
                max_retries=3
            )
            
            assert writer.api_timeout == 60
            assert writer.max_retries == 3
            assert writer.credentials is not None


# ===== スプレッドシート作成のテスト =====

class TestGoogleWriterSpreadsheet:
    """スプレッドシート作成のテスト"""

    def test_create_spreadsheet_returns_write_result(
        self, google_writer, basic_spreadsheet_data
    ):
        """create_spreadsheetがWriteResultを返すことを検証（要件8.1）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            # スプレッドシート作成のモック
            mock_service.spreadsheets().create().execute.return_value = MOCK_SPREADSHEET_RESPONSE
            mock_service.spreadsheets().values().update().execute.return_value = {}
            
            result = google_writer.create_spreadsheet(
                basic_spreadsheet_data, "テストスプレッドシート"
            )
            
            assert isinstance(result, WriteResult)
            assert hasattr(result, 'success')
            assert hasattr(result, 'output_path')
            assert hasattr(result, 'url')
            assert hasattr(result, 'error')

    def test_create_spreadsheet_success(
        self, google_writer, basic_spreadsheet_data
    ):
        """スプレッドシートが正常に作成されることを検証（要件8.1）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.spreadsheets().create().execute.return_value = MOCK_SPREADSHEET_RESPONSE
            mock_service.spreadsheets().values().update().execute.return_value = {}
            
            result = google_writer.create_spreadsheet(
                basic_spreadsheet_data, "テストスプレッドシート"
            )
            
            assert result.success is True
            assert result.error is None

    def test_create_spreadsheet_returns_url(
        self, google_writer, basic_spreadsheet_data
    ):
        """作成されたスプレッドシートのURLが返されることを検証（要件8.4）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.spreadsheets().create().execute.return_value = MOCK_SPREADSHEET_RESPONSE
            mock_service.spreadsheets().values().update().execute.return_value = {}
            
            result = google_writer.create_spreadsheet(
                basic_spreadsheet_data, "テストスプレッドシート"
            )
            
            assert result.url is not None
            assert "docs.google.com/spreadsheets" in result.url
            assert result.output_path is None  # ローカルファイルではないのでNone

    def test_create_spreadsheet_with_multiple_sheets(
        self, google_writer, multi_sheet_data
    ):
        """複数シートのスプレッドシートが作成されることを検証（要件8.1）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.spreadsheets().create().execute.return_value = MOCK_SPREADSHEET_RESPONSE
            mock_service.spreadsheets().values().update().execute.return_value = {}
            mock_service.spreadsheets().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_spreadsheet(
                multi_sheet_data, "複数シートテスト"
            )
            
            assert result.success is True
            assert result.url is not None

    def test_create_spreadsheet_permission_error(self, google_writer, basic_spreadsheet_data):
        """アクセス権限がない場合のエラーハンドリングを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            from googleapiclient.errors import HttpError
            
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            # 403エラーをシミュレート
            mock_resp = Mock()
            mock_resp.status = 403
            mock_service.spreadsheets().create().execute.side_effect = HttpError(
                resp=mock_resp, content=b'Forbidden'
            )
            
            result = google_writer.create_spreadsheet(
                basic_spreadsheet_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert "権限" in result.error

    def test_create_spreadsheet_with_empty_sheets(self, google_writer):
        """空のシートリストでもスプレッドシートが作成されることを検証"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.spreadsheets().create().execute.return_value = MOCK_SPREADSHEET_RESPONSE
            
            result = google_writer.create_spreadsheet(
                {"sheets": []}, "空のスプレッドシート"
            )
            
            assert result.success is True


# ===== ドキュメント作成のテスト =====

class TestGoogleWriterDocument:
    """ドキュメント作成のテスト"""

    def test_create_document_returns_write_result(
        self, google_writer, basic_document_data
    ):
        """create_documentがWriteResultを返すことを検証（要件8.2）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.return_value = MOCK_DOCUMENT_RESPONSE
            mock_service.documents().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_document(
                basic_document_data, "テストドキュメント"
            )
            
            assert isinstance(result, WriteResult)

    def test_create_document_success(
        self, google_writer, basic_document_data
    ):
        """ドキュメントが正常に作成されることを検証（要件8.2）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.return_value = MOCK_DOCUMENT_RESPONSE
            mock_service.documents().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_document(
                basic_document_data, "テストドキュメント"
            )
            
            assert result.success is True
            assert result.error is None

    def test_create_document_returns_url(
        self, google_writer, basic_document_data
    ):
        """作成されたドキュメントのURLが返されることを検証（要件8.4）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.return_value = MOCK_DOCUMENT_RESPONSE
            mock_service.documents().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_document(
                basic_document_data, "テストドキュメント"
            )
            
            assert result.url is not None
            assert "docs.google.com/document" in result.url
            assert MOCK_DOCUMENT_RESPONSE['documentId'] in result.url

    def test_create_document_with_multiple_sections(self, google_writer):
        """複数セクションのドキュメントが作成されることを検証（要件8.2）"""
        data = {
            "sections": [
                {
                    "heading": "概要",
                    "level": 1,
                    "paragraphs": ["概要の内容です。"]
                },
                {
                    "heading": "詳細",
                    "level": 1,
                    "paragraphs": ["詳細の内容です。", "追加の段落です。"]
                },
                {
                    "heading": "サブセクション",
                    "level": 2,
                    "paragraphs": ["サブセクションの内容です。"]
                }
            ]
        }
        
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.return_value = MOCK_DOCUMENT_RESPONSE
            mock_service.documents().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_document(data, "複数セクションテスト")
            
            assert result.success is True

    def test_create_document_permission_error(self, google_writer, basic_document_data):
        """アクセス権限がない場合のエラーハンドリングを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            from googleapiclient.errors import HttpError
            
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_resp = Mock()
            mock_resp.status = 403
            mock_service.documents().create().execute.side_effect = HttpError(
                resp=mock_resp, content=b'Forbidden'
            )
            
            result = google_writer.create_document(
                basic_document_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert "権限" in result.error

    def test_create_document_with_empty_sections(self, google_writer):
        """空のセクションリストでもドキュメントが作成されることを検証"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.return_value = MOCK_DOCUMENT_RESPONSE
            
            result = google_writer.create_document(
                {"sections": []}, "空のドキュメント"
            )
            
            assert result.success is True


# ===== スライド作成のテスト =====

class TestGoogleWriterSlides:
    """スライド作成のテスト"""

    def test_create_slides_returns_write_result(
        self, google_writer, basic_slides_data
    ):
        """create_slidesがWriteResultを返すことを検証（要件8.3）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().get().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_slides(
                basic_slides_data, "テストプレゼンテーション"
            )
            
            assert isinstance(result, WriteResult)

    def test_create_slides_success(
        self, google_writer, basic_slides_data
    ):
        """スライドが正常に作成されることを検証（要件8.3）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().get().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_slides(
                basic_slides_data, "テストプレゼンテーション"
            )
            
            assert result.success is True
            assert result.error is None

    def test_create_slides_returns_url(
        self, google_writer, basic_slides_data
    ):
        """作成されたスライドのURLが返されることを検証（要件8.4）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().get().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_slides(
                basic_slides_data, "テストプレゼンテーション"
            )
            
            assert result.url is not None
            assert "docs.google.com/presentation" in result.url
            assert MOCK_PRESENTATION_RESPONSE['presentationId'] in result.url

    def test_create_slides_with_multiple_slides(self, google_writer):
        """複数スライドのプレゼンテーションが作成されることを検証（要件8.3）"""
        data = {
            "slides": [
                {"layout": "title", "title": "タイトル", "content": "サブタイトル"},
                {"layout": "content", "title": "スライド2", "content": "内容2"},
                {"layout": "bullet", "title": "スライド3", "content": ["項目1", "項目2"]}
            ]
        }
        
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().get().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().batchUpdate().execute.return_value = {}
            
            result = google_writer.create_slides(data, "複数スライドテスト")
            
            assert result.success is True

    def test_create_slides_permission_error(self, google_writer, basic_slides_data):
        """アクセス権限がない場合のエラーハンドリングを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            from googleapiclient.errors import HttpError
            
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_resp = Mock()
            mock_resp.status = 403
            mock_service.presentations().create().execute.side_effect = HttpError(
                resp=mock_resp, content=b'Forbidden'
            )
            
            result = google_writer.create_slides(
                basic_slides_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert "権限" in result.error

    def test_create_slides_with_empty_slides(self, google_writer):
        """空のスライドリストでもプレゼンテーションが作成されることを検証"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.return_value = MOCK_PRESENTATION_RESPONSE
            mock_service.presentations().get().execute.return_value = MOCK_PRESENTATION_RESPONSE
            
            result = google_writer.create_slides(
                {"slides": []}, "空のプレゼンテーション"
            )
            
            assert result.success is True


# ===== リトライロジックのテスト =====

class TestGoogleWriterRetryLogic:
    """リトライロジックのテスト"""

    def test_execute_with_retry_success_on_first_attempt(self, google_writer):
        """最初の試行で成功した場合にリトライしないことを検証"""
        mock_func = Mock(return_value="success")
        
        result = google_writer._execute_with_retry(mock_func)
        
        assert result == "success"
        assert mock_func.call_count == 1

    def test_execute_with_retry_retries_on_server_error(self, google_writer):
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
            result = google_writer._execute_with_retry(mock_func)
        
        assert result == "success"
        assert mock_func.call_count == 3

    def test_execute_with_retry_raises_after_max_retries(self, google_writer):
        """最大リトライ回数後にAPIErrorが発生することを検証"""
        from googleapiclient.errors import HttpError
        
        mock_resp = Mock()
        mock_resp.status = 500
        
        mock_func = Mock(side_effect=HttpError(resp=mock_resp, content=b'Server Error'))
        
        with patch('time.sleep'):  # sleepをスキップ
            with pytest.raises(APIError) as exc_info:
                google_writer._execute_with_retry(mock_func)
        
        assert "3回失敗" in str(exc_info.value.message)
        assert mock_func.call_count == google_writer.max_retries

    def test_execute_with_retry_retries_on_rate_limit(self, google_writer):
        """レート制限エラー（429）時にリトライすることを検証"""
        from googleapiclient.errors import HttpError
        
        mock_resp = Mock()
        mock_resp.status = 429
        
        # 最初の1回は失敗、2回目で成功
        mock_func = Mock(side_effect=[
            HttpError(resp=mock_resp, content=b'Rate Limit'),
            "success"
        ])
        
        with patch('time.sleep'):
            result = google_writer._execute_with_retry(mock_func)
        
        assert result == "success"
        assert mock_func.call_count == 2

    def test_execute_with_retry_no_retry_on_403(self, google_writer):
        """403エラー時にリトライしないことを検証"""
        from googleapiclient.errors import HttpError
        
        mock_resp = Mock()
        mock_resp.status = 403
        
        mock_func = Mock(side_effect=HttpError(resp=mock_resp, content=b'Forbidden'))
        
        with pytest.raises(HttpError):
            google_writer._execute_with_retry(mock_func)
        
        # 403はリトライ不可能なので1回のみ呼び出される
        assert mock_func.call_count == 1


# ===== エラーハンドリングのテスト =====

class TestGoogleWriterErrorHandling:
    """エラーハンドリングのテスト"""

    def test_spreadsheet_api_error_returns_write_result_with_error(
        self, google_writer, basic_spreadsheet_data
    ):
        """API エラー時にエラー付きWriteResultが返されることを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.spreadsheets().create().execute.side_effect = Exception(
                "予期しないエラー"
            )
            
            result = google_writer.create_spreadsheet(
                basic_spreadsheet_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert result.url is None

    def test_document_api_error_returns_write_result_with_error(
        self, google_writer, basic_document_data
    ):
        """API エラー時にエラー付きWriteResultが返されることを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.documents().create().execute.side_effect = Exception(
                "予期しないエラー"
            )
            
            result = google_writer.create_document(
                basic_document_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert result.url is None

    def test_slides_api_error_returns_write_result_with_error(
        self, google_writer, basic_slides_data
    ):
        """API エラー時にエラー付きWriteResultが返されることを検証（要件8.5）"""
        with patch('document_format_mcp_server.writers.google_writer.build') as mock_build:
            mock_service = MagicMock()
            mock_build.return_value = mock_service
            
            mock_service.presentations().create().execute.side_effect = Exception(
                "予期しないエラー"
            )
            
            result = google_writer.create_slides(
                basic_slides_data, "テスト"
            )
            
            assert result.success is False
            assert result.error is not None
            assert result.url is None

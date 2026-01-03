"""
MCPサーバとツールの統合テスト

このテストモジュールは、MCPサーバとツールのエンドツーエンドテストを実装します。
実際のファイルを使用してツールハンドラーの動作を検証します。

要件: 9.1, 9.5
"""

import asyncio
import os
import tempfile
from pathlib import Path

import pytest

from document_format_mcp_server.server import DocumentMCPServer
from document_format_mcp_server.tools.tool_handlers import ToolHandlers
from document_format_mcp_server.tools.tool_definitions import ALL_TOOL_SCHEMAS, TOOL_DEFINITIONS
from document_format_mcp_server.utils.config import Config


# テスト用のサンプルファイルパス
SAMPLE_PPTX = "test_files/samples/sample.pptx"
SAMPLE_DOCX = "test_files/samples/sample.docx"
SAMPLE_XLSX = "test_files/samples/sample.xlsx"


def run_async(coro):
    """非同期関数を同期的に実行するヘルパー関数"""
    return asyncio.run(coro)


@pytest.fixture
def config():
    """テスト用の設定を返すフィクスチャ"""
    return Config()


@pytest.fixture
def tool_handlers(config):
    """ToolHandlersインスタンスを返すフィクスチャ"""
    return ToolHandlers(config)


@pytest.fixture
def mcp_server(config):
    """DocumentMCPServerインスタンスを返すフィクスチャ"""
    return DocumentMCPServer(config)


# =============================================================================
# MCPサーバ初期化テスト（要件9.1、9.4）
# =============================================================================

class TestMCPServerInitialization:
    """MCPサーバの初期化に関するテスト"""
    
    def test_server_initialization_without_error(self, config):
        """MCPサーバがTypeErrorなしで正常に初期化されることを検証（要件9.4、11.3）"""
        # サーバを初期化
        server = DocumentMCPServer(config)
        
        # サーバが正常に初期化されたことを検証
        assert server is not None
        assert server.server is not None
        assert server.tool_handlers is not None
    
    def test_server_has_tool_handlers(self, mcp_server):
        """MCPサーバがToolHandlersを持つことを検証"""
        assert mcp_server.tool_handlers is not None
        assert isinstance(mcp_server.tool_handlers, ToolHandlers)
    
    def test_server_has_all_readers(self, mcp_server):
        """MCPサーバが全てのリーダーを持つことを検証"""
        handlers = mcp_server.tool_handlers
        
        # ローカルファイルリーダーの存在を検証
        assert handlers.powerpoint_reader is not None
        assert handlers.word_reader is not None
        assert handlers.excel_reader is not None
    
    def test_server_has_all_writers(self, mcp_server):
        """MCPサーバが全てのライターを持つことを検証"""
        handlers = mcp_server.tool_handlers
        
        # ローカルファイルライターの存在を検証
        assert handlers.powerpoint_writer is not None
        assert handlers.word_writer is not None
        assert handlers.excel_writer is not None


# =============================================================================
# ツール定義テスト（要件9.2）
# =============================================================================

class TestToolDefinitions:
    """ツール定義に関するテスト"""
    
    def test_all_tool_schemas_count(self):
        """12個のツールスキーマが定義されていることを検証（要件9.2）"""
        assert len(ALL_TOOL_SCHEMAS) == 12
    
    def test_tool_definitions_export(self):
        """TOOL_DEFINITIONSがエクスポートされていることを検証"""
        assert TOOL_DEFINITIONS is not None
        assert len(TOOL_DEFINITIONS) == 12
    
    def test_read_tools_defined(self):
        """読み取りツールが定義されていることを検証"""
        tool_names = [schema["name"] for schema in ALL_TOOL_SCHEMAS]
        
        expected_read_tools = [
            "read_powerpoint",
            "read_word",
            "read_excel",
            "read_google_spreadsheet",
            "read_google_document",
            "read_google_slides",
        ]
        
        for tool_name in expected_read_tools:
            assert tool_name in tool_names, f"ツール {tool_name} が定義されていません"
    
    def test_write_tools_defined(self):
        """書き込みツールが定義されていることを検証"""
        tool_names = [schema["name"] for schema in ALL_TOOL_SCHEMAS]
        
        expected_write_tools = [
            "write_powerpoint",
            "write_word",
            "write_excel",
            "write_google_spreadsheet",
            "write_google_document",
            "write_google_slides",
        ]
        
        for tool_name in expected_write_tools:
            assert tool_name in tool_names, f"ツール {tool_name} が定義されていません"
    
    def test_tool_schema_structure(self):
        """各ツールスキーマが必要な構造を持つことを検証"""
        for schema in ALL_TOOL_SCHEMAS:
            assert "name" in schema, "スキーマにnameがありません"
            assert "description" in schema, "スキーマにdescriptionがありません"
            assert "inputSchema" in schema, "スキーマにinputSchemaがありません"
            
            # inputSchemaの構造を検証
            input_schema = schema["inputSchema"]
            assert "type" in input_schema
            assert input_schema["type"] == "object"
            assert "properties" in input_schema
            assert "required" in input_schema


# =============================================================================
# 読み取りツールハンドラー統合テスト（要件9.5）
# =============================================================================

class TestReadToolHandlersIntegration:
    """読み取りツールハンドラーの統合テスト"""
    
    def test_handle_read_powerpoint_success(self, tool_handlers):
        """PowerPoint読み取りハンドラーが成功レスポンスを返すことを検証"""
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        # ツールを呼び出し
        result = run_async(tool_handlers.handle_read_powerpoint({"file_path": SAMPLE_PPTX}))
        
        # 成功レスポンスを検証
        assert result["success"] is True
        assert "format_type" in result
        assert result["format_type"] == "pptx"
        assert "content" in result
        assert "slides" in result["content"]
    
    def test_handle_read_powerpoint_file_not_found(self, tool_handlers):
        """存在しないファイルに対してエラーレスポンスを返すことを検証"""
        result = run_async(tool_handlers.handle_read_powerpoint({"file_path": "nonexistent.pptx"}))
        
        # エラーレスポンスを検証
        assert result["success"] is False
        assert "error" in result
        assert "type" in result["error"]
        assert "message" in result["error"]
    
    def test_handle_read_powerpoint_missing_param(self, tool_handlers):
        """必須パラメータが不足している場合にエラーを返すことを検証"""
        result = run_async(tool_handlers.handle_read_powerpoint({}))
        
        # エラーレスポンスを検証
        assert result["success"] is False
        assert "error" in result
    
    def test_handle_read_word_success(self, tool_handlers):
        """Word読み取りハンドラーが成功レスポンスを返すことを検証"""
        if not Path(SAMPLE_DOCX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_DOCX}")
        
        # ツールを呼び出し
        result = run_async(tool_handlers.handle_read_word({"file_path": SAMPLE_DOCX}))
        
        # 成功レスポンスを検証
        assert result["success"] is True
        assert "format_type" in result
        assert result["format_type"] == "docx"
        assert "content" in result
        assert "paragraphs" in result["content"]
    
    def test_handle_read_word_file_not_found(self, tool_handlers):
        """存在しないファイルに対してエラーレスポンスを返すことを検証"""
        result = run_async(tool_handlers.handle_read_word({"file_path": "nonexistent.docx"}))
        
        # エラーレスポンスを検証
        assert result["success"] is False
        assert "error" in result
    
    def test_handle_read_excel_success(self, tool_handlers):
        """Excel読み取りハンドラーが成功レスポンスを返すことを検証"""
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        # ツールを呼び出し
        result = run_async(tool_handlers.handle_read_excel({"file_path": SAMPLE_XLSX}))
        
        # 成功レスポンスを検証
        assert result["success"] is True
        assert "format_type" in result
        assert result["format_type"] == "xlsx"
        assert "content" in result
        assert "sheets" in result["content"]
    
    def test_handle_read_excel_file_not_found(self, tool_handlers):
        """存在しないファイルに対してエラーレスポンスを返すことを検証"""
        result = run_async(tool_handlers.handle_read_excel({"file_path": "nonexistent.xlsx"}))
        
        # エラーレスポンスを検証
        assert result["success"] is False
        assert "error" in result


# =============================================================================
# 書き込みツールハンドラー統合テスト（要件9.5）
# =============================================================================

class TestWriteToolHandlersIntegration:
    """書き込みツールハンドラーの統合テスト"""
    
    def test_handle_write_powerpoint_success(self, tool_handlers):
        """PowerPoint書き込みハンドラーが成功レスポンスを返すことを検証"""
        # 一時ファイルパスを作成
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        
        try:
            # テストデータ
            data = {
                "title": "テストプレゼンテーション",
                "slides": [
                    {
                        "layout": "title",
                        "title": "タイトルスライド",
                        "content": "サブタイトル"
                    },
                    {
                        "layout": "content",
                        "title": "コンテンツスライド",
                        "content": "本文テキスト"
                    }
                ]
            }
            
            # ツールを呼び出し
            result = run_async(tool_handlers.handle_write_powerpoint({
                "data": data,
                "output_path": output_path
            }))
            
            # 成功レスポンスを検証
            assert result["success"] is True
            assert "output_path" in result
            
            # ファイルが作成されたことを検証
            assert Path(output_path).exists()
        finally:
            # 一時ファイルを削除
            if Path(output_path).exists():
                os.unlink(output_path)
    
    def test_handle_write_powerpoint_missing_data(self, tool_handlers):
        """必須パラメータが不足している場合にエラーを返すことを検証"""
        result = run_async(tool_handlers.handle_write_powerpoint({
            "output_path": "test.pptx"
        }))
        
        # エラーレスポンスを検証
        assert result["success"] is False
        assert "error" in result
    
    def test_handle_write_word_success(self, tool_handlers):
        """Word書き込みハンドラーが成功レスポンスを返すことを検証"""
        # 一時ファイルパスを作成
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name
        
        try:
            # テストデータ
            data = {
                "title": "テストドキュメント",
                "sections": [
                    {
                        "heading": "セクション1",
                        "level": 1,
                        "paragraphs": ["段落1", "段落2"]
                    }
                ]
            }
            
            # ツールを呼び出し
            result = run_async(tool_handlers.handle_write_word({
                "data": data,
                "output_path": output_path
            }))
            
            # 成功レスポンスを検証
            assert result["success"] is True
            assert "output_path" in result
            
            # ファイルが作成されたことを検証
            assert Path(output_path).exists()
        finally:
            # 一時ファイルを削除
            if Path(output_path).exists():
                os.unlink(output_path)
    
    def test_handle_write_excel_success(self, tool_handlers):
        """Excel書き込みハンドラーが成功レスポンスを返すことを検証"""
        # 一時ファイルパスを作成
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name
        
        try:
            # テストデータ
            data = {
                "sheets": [
                    {
                        "name": "Sheet1",
                        "data": [
                            ["A1", "B1", "C1"],
                            ["A2", "B2", "C2"]
                        ]
                    }
                ]
            }
            
            # ツールを呼び出し
            result = run_async(tool_handlers.handle_write_excel({
                "data": data,
                "output_path": output_path
            }))
            
            # 成功レスポンスを検証
            assert result["success"] is True
            assert "output_path" in result
            
            # ファイルが作成されたことを検証
            assert Path(output_path).exists()
        finally:
            # 一時ファイルを削除
            if Path(output_path).exists():
                os.unlink(output_path)


# =============================================================================
# エラーレスポンス形式テスト（要件9.5、11.6）
# =============================================================================

class TestErrorResponseFormat:
    """エラーレスポンス形式のテスト"""
    
    def test_error_response_is_dict(self, tool_handlers):
        """エラーレスポンスがJSON辞書形式であることを検証（要件11.6）"""
        result = run_async(tool_handlers.handle_read_powerpoint({"file_path": "nonexistent.pptx"}))
        
        # レスポンスが辞書であることを検証（文字列ではない）
        assert isinstance(result, dict)
        assert not isinstance(result, str)
    
    def test_error_response_structure(self, tool_handlers):
        """エラーレスポンスが正しい構造を持つことを検証"""
        result = run_async(tool_handlers.handle_read_powerpoint({"file_path": "nonexistent.pptx"}))
        
        # 構造を検証
        assert "success" in result
        assert result["success"] is False
        assert "error" in result
        assert "type" in result["error"]
        assert "message" in result["error"]
        assert "details" in result["error"]
    
    def test_success_response_is_dict(self, tool_handlers):
        """成功レスポンスがJSON辞書形式であることを検証"""
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        result = run_async(tool_handlers.handle_read_powerpoint({"file_path": SAMPLE_PPTX}))
        
        # レスポンスが辞書であることを検証（文字列ではない）
        assert isinstance(result, dict)
        assert not isinstance(result, str)


# =============================================================================
# 読み取り→書き込みラウンドトリップテスト
# =============================================================================

class TestReadWriteRoundTrip:
    """読み取りと書き込みのラウンドトリップテスト"""
    
    def test_powerpoint_read_write_roundtrip(self, tool_handlers):
        """PowerPointの読み取り→書き込み→読み取りが正常に動作することを検証"""
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        # 元のファイルを読み取り
        read_result = run_async(tool_handlers.handle_read_powerpoint({"file_path": SAMPLE_PPTX}))
        assert read_result["success"] is True
        
        # 一時ファイルに書き込み
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output_path = f.name
        
        try:
            # 読み取ったデータを使って新しいファイルを作成
            original_slides = read_result["content"]["slides"]
            write_data = {
                "title": "ラウンドトリップテスト",
                "slides": [
                    {
                        "layout": "content",
                        "title": slide.get("title", ""),
                        "content": slide.get("content", "")
                    }
                    for slide in original_slides[:3]  # 最初の3スライドのみ
                ]
            }
            
            write_result = run_async(tool_handlers.handle_write_powerpoint({
                "data": write_data,
                "output_path": output_path
            }))
            assert write_result["success"] is True
            
            # 書き込んだファイルを読み取り
            read_result2 = run_async(tool_handlers.handle_read_powerpoint({"file_path": output_path}))
            assert read_result2["success"] is True
            assert "slides" in read_result2["content"]
        finally:
            if Path(output_path).exists():
                os.unlink(output_path)
    
    def test_word_read_write_roundtrip(self, tool_handlers):
        """Wordの読み取り→書き込み→読み取りが正常に動作することを検証"""
        if not Path(SAMPLE_DOCX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_DOCX}")
        
        # 元のファイルを読み取り
        read_result = run_async(tool_handlers.handle_read_word({"file_path": SAMPLE_DOCX}))
        assert read_result["success"] is True
        
        # 一時ファイルに書き込み
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name
        
        try:
            # 読み取ったデータを使って新しいファイルを作成
            original_paragraphs = read_result["content"]["paragraphs"]
            write_data = {
                "title": "ラウンドトリップテスト",
                "sections": [
                    {
                        "heading": "セクション1",
                        "level": 1,
                        "paragraphs": [
                            p.get("text", "") 
                            for p in original_paragraphs[:5]
                            if p.get("text", "").strip()
                        ]
                    }
                ]
            }
            
            write_result = run_async(tool_handlers.handle_write_word({
                "data": write_data,
                "output_path": output_path
            }))
            assert write_result["success"] is True
            
            # 書き込んだファイルを読み取り
            read_result2 = run_async(tool_handlers.handle_read_word({"file_path": output_path}))
            assert read_result2["success"] is True
            assert "paragraphs" in read_result2["content"]
        finally:
            if Path(output_path).exists():
                os.unlink(output_path)
    
    def test_excel_read_write_roundtrip(self, tool_handlers):
        """Excelの読み取り→書き込み→読み取りが正常に動作することを検証"""
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        # 元のファイルを読み取り
        read_result = run_async(tool_handlers.handle_read_excel({"file_path": SAMPLE_XLSX}))
        assert read_result["success"] is True
        
        # 一時ファイルに書き込み
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name
        
        try:
            # 読み取ったデータを使って新しいファイルを作成
            original_sheets = read_result["content"]["sheets"]
            write_data = {
                "sheets": [
                    {
                        "name": sheet.get("name", f"Sheet{i+1}"),
                        "data": sheet.get("data", [])[:10]  # 最初の10行のみ
                    }
                    for i, sheet in enumerate(original_sheets[:3])  # 最初の3シートのみ
                ]
            }
            
            write_result = run_async(tool_handlers.handle_write_excel({
                "data": write_data,
                "output_path": output_path
            }))
            assert write_result["success"] is True
            
            # 書き込んだファイルを読み取り
            read_result2 = run_async(tool_handlers.handle_read_excel({"file_path": output_path}))
            assert read_result2["success"] is True
            assert "sheets" in read_result2["content"]
        finally:
            if Path(output_path).exists():
                os.unlink(output_path)


# =============================================================================
# Google Workspaceツールテスト（Google認証が無効な場合のエラーハンドリング）
# =============================================================================

class TestGoogleWorkspaceToolsWithoutAuth:
    """Google Workspace認証が無効な場合のツールテスト"""
    
    @pytest.fixture
    def config_without_google(self):
        """Google Workspaceが無効な設定を返すフィクスチャ"""
        config = Config()
        config._config["enable_google_workspace"] = False
        return config
    
    @pytest.fixture
    def tool_handlers_without_google(self, config_without_google):
        """Google Workspaceが無効なToolHandlersを返すフィクスチャ"""
        return ToolHandlers(config_without_google)
    
    def test_google_spreadsheet_read_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_read_google_spreadsheet({
            "file_id": "test_file_id"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]
    
    def test_google_document_read_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_read_google_document({
            "file_id": "test_file_id"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]
    
    def test_google_slides_read_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_read_google_slides({
            "file_id": "test_file_id"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]
    
    def test_google_spreadsheet_write_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_write_google_spreadsheet({
            "data": {"sheets": []},
            "title": "テスト"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]
    
    def test_google_document_write_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_write_google_document({
            "data": {"sections": []},
            "title": "テスト"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]
    
    def test_google_slides_write_without_auth(self, tool_handlers_without_google):
        """Google認証が無効な場合にエラーを返すことを検証"""
        result = run_async(tool_handlers_without_google.handle_write_google_slides({
            "data": {"slides": []},
            "title": "テスト"
        }))
        
        assert result["success"] is False
        assert "error" in result
        assert "Google Workspace" in result["error"]["message"]

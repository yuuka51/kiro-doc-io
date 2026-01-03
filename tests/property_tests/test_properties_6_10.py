"""
プロパティベーステスト: プロパティ6-10（ドキュメント生成とMCP）

**Feature: document-format-mcp-server**

このモジュールは、ドキュメント生成機能とMCPツールに関するプロパティベーステストを実装します。
hypothesisライブラリを使用し、最小100回の反復で実行します。
"""

import os
import tempfile
from pathlib import Path

import pytest
from hypothesis import given, settings, assume, HealthCheck
from hypothesis import strategies as st

from document_format_mcp_server.writers.powerpoint_writer import PowerPointWriter
from document_format_mcp_server.writers.word_writer import WordWriter
from document_format_mcp_server.writers.excel_writer import ExcelWriter
from document_format_mcp_server.utils.models import WriteResult


# ============================================================================
# プロパティ6: ドキュメント生成の完全性
# *任意の*有効な構造化データに対して、Document_Writerは対応する形式の
# ドキュメントファイルを生成し、必要な要素（タイトル、コンテンツ、書式）を含むべきである
# **検証対象: 要件 5.1, 5.2, 5.4, 6.1, 6.2, 6.3, 7.1, 7.2, 7.3**
# ============================================================================

# スライドデータ生成用のストラテジー
slide_layout_strategy = st.sampled_from(["title", "content", "bullet"])
slide_title_strategy = st.text(min_size=0, max_size=100).filter(lambda x: "\x00" not in x)
slide_content_strategy = st.one_of(
    st.text(min_size=0, max_size=200).filter(lambda x: "\x00" not in x),
    st.lists(st.text(min_size=1, max_size=50).filter(lambda x: "\x00" not in x), min_size=0, max_size=5)
)

slide_strategy = st.fixed_dictionaries({
    "layout": slide_layout_strategy,
    "title": slide_title_strategy,
    "content": slide_content_strategy
})

powerpoint_data_strategy = st.fixed_dictionaries({
    "title": st.text(min_size=0, max_size=100).filter(lambda x: "\x00" not in x),
    "slides": st.lists(slide_strategy, min_size=1, max_size=5)
})


class TestProperty6DocumentGenerationCompleteness:
    """
    **Property 6: ドキュメント生成の完全性**
    
    *任意の*有効な構造化データに対して、ライターは対応する形式の
    ドキュメントファイルを生成するべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(data=powerpoint_data_strategy)
    def test_powerpoint_writer_generates_valid_file(self, data: dict):
        """
        **Feature: document-format-mcp-server, Property 6: ドキュメント生成の完全性**
        
        *任意の*有効なスライドデータに対して、PowerPointライターは
        有効なファイルを生成するべきである。
        **検証対象: 要件 5.1, 5.2, 5.4**
        """
        writer = PowerPointWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, "test_output.pptx")
            result = writer.create_presentation(data, output_path)
            
            # WriteResultの構造を検証
            assert isinstance(result, WriteResult)
            assert result.success is True
            assert result.output_path == output_path
            assert result.error is None
            
            # ファイルが生成されたことを検証
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 0
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        title=st.text(min_size=0, max_size=100).filter(lambda x: "\x00" not in x),
        heading=st.text(min_size=0, max_size=100).filter(lambda x: "\x00" not in x),
        level=st.integers(min_value=1, max_value=3),
        paragraphs=st.lists(st.text(min_size=0, max_size=200).filter(lambda x: "\x00" not in x), min_size=0, max_size=3)
    )
    def test_word_writer_generates_valid_file(self, title: str, heading: str, level: int, paragraphs: list):
        """
        **Feature: document-format-mcp-server, Property 6: ドキュメント生成の完全性**
        
        *任意の*有効なセクションデータに対して、Wordライターは
        有効なファイルを生成するべきである。
        **検証対象: 要件 6.1, 6.2, 6.3**
        """
        data = {
            "title": title,
            "sections": [
                {
                    "heading": heading,
                    "level": level,
                    "paragraphs": paragraphs
                }
            ]
        }
        
        writer = WordWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, "test_output.docx")
            result = writer.create_document(data, output_path)
            
            # WriteResultの構造を検証
            assert isinstance(result, WriteResult)
            assert result.success is True
            assert result.output_path == output_path
            assert result.error is None
            
            # ファイルが生成されたことを検証
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 0
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        sheet_name=st.text(min_size=1, max_size=31).filter(lambda x: x.strip() and "\x00" not in x and "/" not in x and "\\" not in x and "[" not in x and "]" not in x and ":" not in x and "*" not in x and "?" not in x),
        rows=st.integers(min_value=1, max_value=10),
        cols=st.integers(min_value=1, max_value=5)
    )
    def test_excel_writer_generates_valid_file(self, sheet_name: str, rows: int, cols: int):
        """
        **Feature: document-format-mcp-server, Property 6: ドキュメント生成の完全性**
        
        *任意の*有効なシートデータに対して、Excelライターは
        有効なファイルを生成するべきである。
        **検証対象: 要件 7.1, 7.2, 7.3**
        """
        # テストデータを生成
        data_rows = []
        for i in range(rows):
            row = [f"Cell_{i}_{j}" for j in range(cols)]
            data_rows.append(row)
        
        data = {
            "sheets": [
                {
                    "name": sheet_name,
                    "data": data_rows
                }
            ]
        }
        
        writer = ExcelWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, "test_output.xlsx")
            result = writer.create_workbook(data, output_path)
            
            # WriteResultの構造を検証
            assert isinstance(result, WriteResult)
            assert result.success is True
            assert result.output_path == output_path
            assert result.error is None
            
            # ファイルが生成されたことを検証
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 0


# ============================================================================
# プロパティ7: ファイル保存の確実性
# *任意の*ドキュメント生成操作において、Document_Writerは生成したファイルを
# ユーザーがアクセス可能な場所に保存し、そのパスまたはURLを返すべきである
# **検証対象: 要件 5.3, 6.4, 7.4, 8.4**
# ============================================================================

class TestProperty7FileSaveReliability:
    """
    **Property 7: ファイル保存の確実性**
    
    *任意の*ドキュメント生成操作において、ライターは生成したファイルを
    指定された場所に保存し、そのパスを返すべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        subdir=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum()),
        filename=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum())
    )
    def test_powerpoint_writer_saves_to_specified_path(self, subdir: str, filename: str):
        """
        **Feature: document-format-mcp-server, Property 7: ファイル保存の確実性**
        
        *任意の*有効な出力パスに対して、PowerPointライターは
        指定された場所にファイルを保存するべきである。
        **検証対象: 要件 5.3**
        """
        data = {
            "slides": [
                {"layout": "title", "title": "Test", "content": "Content"}
            ]
        }
        
        writer = PowerPointWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, subdir, f"{filename}.pptx")
            result = writer.create_presentation(data, output_path)
            
            # 成功を検証
            assert result.success is True
            assert result.output_path == output_path
            
            # ファイルが指定されたパスに存在することを検証
            assert os.path.exists(output_path)
            assert os.path.isfile(output_path)
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        subdir=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum()),
        filename=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum())
    )
    def test_word_writer_saves_to_specified_path(self, subdir: str, filename: str):
        """
        **Feature: document-format-mcp-server, Property 7: ファイル保存の確実性**
        
        *任意の*有効な出力パスに対して、Wordライターは
        指定された場所にファイルを保存するべきである。
        **検証対象: 要件 6.4**
        """
        data = {
            "sections": [
                {"heading": "Test", "level": 1, "paragraphs": ["Content"]}
            ]
        }
        
        writer = WordWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, subdir, f"{filename}.docx")
            result = writer.create_document(data, output_path)
            
            # 成功を検証
            assert result.success is True
            assert result.output_path == output_path
            
            # ファイルが指定されたパスに存在することを検証
            assert os.path.exists(output_path)
            assert os.path.isfile(output_path)
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        subdir=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum()),
        filename=st.text(min_size=1, max_size=20).filter(lambda x: x.strip() and x.isalnum())
    )
    def test_excel_writer_saves_to_specified_path(self, subdir: str, filename: str):
        """
        **Feature: document-format-mcp-server, Property 7: ファイル保存の確実性**
        
        *任意の*有効な出力パスに対して、Excelライターは
        指定された場所にファイルを保存するべきである。
        **検証対象: 要件 7.4**
        """
        data = {
            "sheets": [
                {"name": "Sheet1", "data": [["A", "B"], ["1", "2"]]}
            ]
        }
        
        writer = ExcelWriter()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, subdir, f"{filename}.xlsx")
            result = writer.create_workbook(data, output_path)
            
            # 成功を検証
            assert result.success is True
            assert result.output_path == output_path
            
            # ファイルが指定されたパスに存在することを検証
            assert os.path.exists(output_path)
            assert os.path.isfile(output_path)


# ============================================================================
# プロパティ8: Google Workspace作成の一貫性
# *任意の*有効な構造化データに対して、Document_Writerは対応するGoogle APIを
# 使用して新しいファイルを作成し、アクセス可能なURLを返すべきである
# **検証対象: 要件 8.1, 8.2, 8.3, 8.4**
# ============================================================================

class TestProperty8GoogleWorkspaceCreation:
    """
    **Property 8: Google Workspace作成の一貫性**
    
    *任意の*無効な認証情報に対して、Google Workspaceライターは
    一貫したエラーを返すべきである。
    
    注: 実際のGoogle API呼び出しはモックなしでは実行できないため、
    このテストは認証エラーのハンドリングを検証します。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(credentials_path=st.text(min_size=1, max_size=50).filter(lambda x: x.strip()))
    def test_google_writer_handles_invalid_credentials(self, credentials_path: str):
        """
        **Feature: document-format-mcp-server, Property 8: Google Workspace作成の一貫性**
        
        *任意の*無効な認証情報パスに対して、Google Workspaceライターは
        一貫したエラーを返すべきである。
        **検証対象: 要件 8.1, 8.2, 8.3, 8.4**
        """
        try:
            from document_format_mcp_server.writers.google_writer import GoogleWorkspaceWriter
            
            # 無効な認証情報パスでの初期化を試行
            # 存在しないパスの場合、初期化時にエラーが発生するはず
            if not os.path.exists(credentials_path):
                with pytest.raises(Exception):
                    writer = GoogleWorkspaceWriter(credentials_path=credentials_path)
            else:
                # パスが存在する場合（偶然）、テストをスキップ
                pytest.skip("認証情報パスが存在します")
        
        except ImportError:
            pytest.skip("Google Workspaceライターがインポートできません")


# ============================================================================
# プロパティ9: MCPツール公開の完全性
# *任意の*MCP_Server起動時において、全ての必要なツール（12個の読み取り/書き込み機能）が
# Kiroに公開されるべきである
# **検証対象: 要件 9.2**
# ============================================================================

class TestProperty9MCPToolExposure:
    """
    **Property 9: MCPツール公開の完全性**
    
    MCPサーバは全ての必要なツールを公開するべきである。
    """
    
    def test_all_tools_are_defined(self):
        """
        **Feature: document-format-mcp-server, Property 9: MCPツール公開の完全性**
        
        全ての必要なツール（12個）が定義されているべきである。
        **検証対象: 要件 9.2**
        """
        from document_format_mcp_server.tools.tool_definitions import TOOL_DEFINITIONS
        
        # 必要なツール名のリスト
        required_tools = [
            "read_powerpoint",
            "read_word",
            "read_excel",
            "read_google_spreadsheet",
            "read_google_document",
            "read_google_slides",
            "write_powerpoint",
            "write_word",
            "write_excel",
            "write_google_spreadsheet",
            "write_google_document",
            "write_google_slides",
        ]
        
        # 全てのツールが定義されていることを検証
        defined_tool_names = [tool["name"] for tool in TOOL_DEFINITIONS]
        
        for tool_name in required_tools:
            assert tool_name in defined_tool_names, f"ツール '{tool_name}' が定義されていません"
        
        # ツール数が12個であることを検証
        assert len(TOOL_DEFINITIONS) == 12, f"ツール数が12個ではありません: {len(TOOL_DEFINITIONS)}"


# ============================================================================
# プロパティ10: 通信プロトコルの準拠
# *任意の*ツール呼び出しに対して、MCP_Serverは標準入出力を介して
# 明確な成功または失敗の応答を返すべきである
# **検証対象: 要件 9.3, 9.5**
# ============================================================================

class TestProperty10CommunicationProtocol:
    """
    **Property 10: 通信プロトコルの準拠**
    
    *任意の*ツール呼び出しに対して、ハンドラーは明確な成功または失敗の応答を返すべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(file_path=st.text(min_size=1, max_size=100).filter(lambda x: x.strip()))
    def test_tool_handlers_return_consistent_response_format(self, file_path: str):
        """
        **Feature: document-format-mcp-server, Property 10: 通信プロトコルの準拠**
        
        *任意の*ファイルパスに対して、ツールハンドラーは一貫したレスポンス形式を返すべきである。
        **検証対象: 要件 9.3, 9.5**
        """
        import asyncio
        from document_format_mcp_server.tools.tool_handlers import ToolHandlers
        from document_format_mcp_server.utils.config import Config
        
        # デフォルト設定でハンドラーを初期化
        config = Config()
        handlers = ToolHandlers(config)
        
        # 非同期関数を実行
        async def run_test():
            result = await handlers.handle_read_powerpoint({"file_path": file_path})
            return result
        
        result = asyncio.run(run_test())
        
        # レスポンス形式を検証
        assert isinstance(result, dict)
        assert "success" in result
        assert isinstance(result["success"], bool)
        
        if result["success"]:
            # 成功時のレスポンス形式
            assert "format_type" in result or "data" in result
        else:
            # 失敗時のレスポンス形式
            assert "error" in result
            assert isinstance(result["error"], dict)
            assert "type" in result["error"]
            assert "message" in result["error"]

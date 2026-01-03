"""
プロパティベーステスト: プロパティ11-15（性能と信頼性）

**Feature: document-format-mcp-server**

このモジュールは、性能と信頼性に関するプロパティベーステストを実装します。
hypothesisライブラリを使用し、最小100回の反復で実行します。
"""

import os
import time
from pathlib import Path

import pytest
from hypothesis import given, settings, assume, HealthCheck
from hypothesis import strategies as st

from document_format_mcp_server.readers.powerpoint_reader import PowerPointReader
from document_format_mcp_server.readers.word_reader import WordReader
from document_format_mcp_server.readers.excel_reader import ExcelReader
from document_format_mcp_server.utils.models import ReadResult, WriteResult, DocumentContent
from document_format_mcp_server.utils.config import Config


# サンプルファイルのパス
SAMPLE_PPTX = "test_files/samples/sample.pptx"
SAMPLE_DOCX = "test_files/samples/sample.docx"
SAMPLE_XLSX = "test_files/samples/sample.xlsx"


# ============================================================================
# プロパティ11: 初期化性能の保証
# *任意の*MCP_Server起動において、初期化は30秒以内に完了し、
# TypeErrorなしで正常に動作するべきである
# **検証対象: 要件 9.4, 11.3**
# ============================================================================

class TestProperty11InitializationPerformance:
    """
    **Property 11: 初期化性能の保証**
    
    *任意の*設定値に対して、リーダー/ライターの初期化は高速に完了し、
    TypeErrorなしで正常に動作するべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        max_slides=st.integers(min_value=1, max_value=1000),
        max_file_size_mb=st.integers(min_value=1, max_value=500)
    )
    def test_powerpoint_reader_initializes_without_error(self, max_slides: int, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 11: 初期化性能の保証**
        
        *任意の*設定値に対して、PowerPointリーダーはTypeErrorなしで初期化されるべきである。
        **検証対象: 要件 9.4, 11.3**
        """
        start_time = time.time()
        
        # 初期化
        reader = PowerPointReader(max_slides=max_slides, max_file_size_mb=max_file_size_mb)
        
        elapsed_time = time.time() - start_time
        
        # 初期化が高速であることを検証（1秒以内）
        assert elapsed_time < 1.0, f"初期化に時間がかかりすぎています: {elapsed_time:.2f}秒"
        
        # 設定値が正しく設定されていることを検証
        assert reader.max_slides == max_slides
        assert reader.max_file_size_mb == max_file_size_mb
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_file_size_mb=st.integers(min_value=1, max_value=500))
    def test_word_reader_initializes_without_error(self, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 11: 初期化性能の保証**
        
        *任意の*設定値に対して、WordリーダーはTypeErrorなしで初期化されるべきである。
        **検証対象: 要件 9.4, 11.3**
        """
        start_time = time.time()
        
        # 初期化
        reader = WordReader(max_file_size_mb=max_file_size_mb)
        
        elapsed_time = time.time() - start_time
        
        # 初期化が高速であることを検証（1秒以内）
        assert elapsed_time < 1.0, f"初期化に時間がかかりすぎています: {elapsed_time:.2f}秒"
        
        # 設定値が正しく設定されていることを検証
        assert reader.max_file_size_mb == max_file_size_mb
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        max_sheets=st.integers(min_value=1, max_value=200),
        max_file_size_mb=st.integers(min_value=1, max_value=500)
    )
    def test_excel_reader_initializes_without_error(self, max_sheets: int, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 11: 初期化性能の保証**
        
        *任意の*設定値に対して、ExcelリーダーはTypeErrorなしで初期化されるべきである。
        **検証対象: 要件 9.4, 11.3**
        """
        start_time = time.time()
        
        # 初期化
        reader = ExcelReader(max_sheets=max_sheets, max_file_size_mb=max_file_size_mb)
        
        elapsed_time = time.time() - start_time
        
        # 初期化が高速であることを検証（1秒以内）
        assert elapsed_time < 1.0, f"初期化に時間がかかりすぎています: {elapsed_time:.2f}秒"
        
        # 設定値が正しく設定されていることを検証
        assert reader.max_sheets == max_sheets
        assert reader.max_file_size_mb == max_file_size_mb


# ============================================================================
# プロパティ12: 開発環境スクリプトの動作保証
# *任意の*開発環境において、サンプルファイル生成スクリプトは必要なテストファイル
# （PowerPoint、Word、Excel）を生成し、テストスクリプトは各リーダーの動作を検証するべきである
# **検証対象: 要件 10.2, 10.4**
# ============================================================================

class TestProperty12DevelopmentEnvironment:
    """
    **Property 12: 開発環境スクリプトの動作保証**
    
    サンプルファイルが存在し、リーダーで読み取れることを検証する。
    """
    
    def test_sample_files_exist_and_readable(self):
        """
        **Feature: document-format-mcp-server, Property 12: 開発環境スクリプトの動作保証**
        
        サンプルファイルが存在し、各リーダーで読み取れるべきである。
        **検証対象: 要件 10.2, 10.4**
        """
        # サンプルファイルの存在を確認
        sample_files = [SAMPLE_PPTX, SAMPLE_DOCX, SAMPLE_XLSX]
        
        for sample_file in sample_files:
            if not Path(sample_file).exists():
                pytest.skip(f"サンプルファイルが見つかりません: {sample_file}")
        
        # PowerPointリーダーでの読み取り
        pptx_reader = PowerPointReader()
        pptx_result = pptx_reader.read_file(SAMPLE_PPTX)
        assert pptx_result.success is True
        assert pptx_result.content is not None
        
        # Wordリーダーでの読み取り
        docx_reader = WordReader()
        docx_result = docx_reader.read_file(SAMPLE_DOCX)
        assert docx_result.success is True
        assert docx_result.content is not None
        
        # Excelリーダーでの読み取り
        xlsx_reader = ExcelReader()
        xlsx_result = xlsx_reader.read_file(SAMPLE_XLSX)
        assert xlsx_result.success is True
        assert xlsx_result.content is not None


# ============================================================================
# プロパティ13: 設定値の適用
# *任意の*リーダー初期化において、設定値（max_sheets、max_file_size_mb、max_slides）が
# 正しく適用され、制限の検証が実行されるべきである
# **検証対象: 要件 11.1, 11.2, 11.8**
# ============================================================================

class TestProperty13ConfigurationApplication:
    """
    **Property 13: 設定値の適用**
    
    *任意の*設定値に対して、リーダーは制限を正しく適用するべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_sheets=st.integers(min_value=1, max_value=100))
    def test_excel_reader_applies_sheet_limit(self, max_sheets: int):
        """
        **Feature: document-format-mcp-server, Property 13: 設定値の適用**
        
        *任意の*max_sheets設定値に対して、Excelリーダーは制限を正しく適用するべきである。
        **検証対象: 要件 11.1, 11.8**
        """
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        reader = ExcelReader(max_sheets=max_sheets)
        result = reader.read_file(SAMPLE_XLSX)
        
        if result.success:
            # 処理されたシート数が制限以下であることを検証
            processed_sheets = len(result.content.content["sheets"])
            assert processed_sheets <= max_sheets
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_slides=st.integers(min_value=1, max_value=500))
    def test_powerpoint_reader_applies_slide_limit(self, max_slides: int):
        """
        **Feature: document-format-mcp-server, Property 13: 設定値の適用**
        
        *任意の*max_slides設定値に対して、PowerPointリーダーは制限を正しく適用するべきである。
        **検証対象: 要件 11.2, 11.8**
        """
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        reader = PowerPointReader(max_slides=max_slides)
        result = reader.read_file(SAMPLE_PPTX)
        
        if result.success:
            # 処理されたスライド数が制限以下であることを検証
            processed_slides = len(result.content.content["slides"])
            assert processed_slides <= max_slides
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(
        max_file_size_mb=st.integers(min_value=1, max_value=500),
        max_sheets=st.integers(min_value=1, max_value=200),
        max_slides=st.integers(min_value=1, max_value=1000)
    )
    def test_config_values_are_applied_to_tool_handlers(
        self, max_file_size_mb: int, max_sheets: int, max_slides: int
    ):
        """
        **Feature: document-format-mcp-server, Property 13: 設定値の適用**
        
        *任意の*設定値に対して、ToolHandlersは設定を正しく適用するべきである。
        **検証対象: 要件 11.1, 11.2, 11.8**
        """
        from document_format_mcp_server.tools.tool_handlers import ToolHandlers
        
        # カスタム設定を作成（内部辞書を直接変更）
        config = Config()
        config._config["max_file_size_mb"] = max_file_size_mb
        config._config["max_sheets"] = max_sheets
        config._config["max_slides"] = max_slides
        
        # ToolHandlersを初期化
        handlers = ToolHandlers(config)
        
        # 設定値が正しく適用されていることを検証
        assert handlers.powerpoint_reader.max_slides == max_slides
        assert handlers.powerpoint_reader.max_file_size_mb == max_file_size_mb
        assert handlers.word_reader.max_file_size_mb == max_file_size_mb
        assert handlers.excel_reader.max_sheets == max_sheets
        assert handlers.excel_reader.max_file_size_mb == max_file_size_mb


# ============================================================================
# プロパティ14: データモデルの一貫性
# *任意の*読み取り/書き込み操作において、共通データモデル（DocumentContent、ReadResult、WriteResult）が
# 使用され、一貫した形式でデータが処理されるべきである
# **検証対象: 要件 11.5**
# ============================================================================

class TestProperty14DataModelConsistency:
    """
    **Property 14: データモデルの一貫性**
    
    *任意の*読み取り操作において、共通データモデルが使用されるべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_slides=st.integers(min_value=1, max_value=500))
    def test_powerpoint_reader_returns_read_result(self, max_slides: int):
        """
        **Feature: document-format-mcp-server, Property 14: データモデルの一貫性**
        
        *任意の*設定値に対して、PowerPointリーダーはReadResultを返すべきである。
        **検証対象: 要件 11.5**
        """
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        reader = PowerPointReader(max_slides=max_slides)
        result = reader.read_file(SAMPLE_PPTX)
        
        # ReadResultデータクラスであることを検証
        assert isinstance(result, ReadResult)
        assert hasattr(result, 'success')
        assert hasattr(result, 'content')
        assert hasattr(result, 'error')
        assert hasattr(result, 'file_path')
        
        if result.success:
            # DocumentContentデータクラスであることを検証
            assert isinstance(result.content, DocumentContent)
            assert hasattr(result.content, 'format_type')
            assert hasattr(result.content, 'metadata')
            assert hasattr(result.content, 'content')
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_file_size_mb=st.integers(min_value=1, max_value=500))
    def test_word_reader_returns_read_result(self, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 14: データモデルの一貫性**
        
        *任意の*設定値に対して、WordリーダーはReadResultを返すべきである。
        **検証対象: 要件 11.5**
        """
        if not Path(SAMPLE_DOCX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_DOCX}")
        
        reader = WordReader(max_file_size_mb=max_file_size_mb)
        result = reader.read_file(SAMPLE_DOCX)
        
        # ReadResultデータクラスであることを検証
        assert isinstance(result, ReadResult)
        
        if result.success:
            # DocumentContentデータクラスであることを検証
            assert isinstance(result.content, DocumentContent)
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_sheets=st.integers(min_value=1, max_value=200))
    def test_excel_reader_returns_read_result(self, max_sheets: int):
        """
        **Feature: document-format-mcp-server, Property 14: データモデルの一貫性**
        
        *任意の*設定値に対して、ExcelリーダーはReadResultを返すべきである。
        **検証対象: 要件 11.5**
        """
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        reader = ExcelReader(max_sheets=max_sheets)
        result = reader.read_file(SAMPLE_XLSX)
        
        # ReadResultデータクラスであることを検証
        assert isinstance(result, ReadResult)
        
        if result.success:
            # DocumentContentデータクラスであることを検証
            assert isinstance(result.content, DocumentContent)


# ============================================================================
# プロパティ15: API呼び出しの信頼性
# *任意の*Google API呼び出しにおいて、最大3回のリトライと60秒のタイムアウトが実装され、
# 一時的な障害に対する耐性を提供するべきである
# **検証対象: 要件 11.7**
# ============================================================================

class TestProperty15APIReliability:
    """
    **Property 15: API呼び出しの信頼性**
    
    Google Workspaceリーダー/ライターはリトライとタイムアウトを実装するべきである。
    """
    
    def test_google_reader_has_retry_configuration(self):
        """
        **Feature: document-format-mcp-server, Property 15: API呼び出しの信頼性**
        
        Google Workspaceリーダーはリトライとタイムアウトの設定を持つべきである。
        **検証対象: 要件 11.7**
        """
        try:
            from document_format_mcp_server.readers.google_reader import GoogleWorkspaceReader
            
            # クラスの属性を確認（インスタンス化せずに）
            # __init__メソッドのシグネチャを確認
            import inspect
            sig = inspect.signature(GoogleWorkspaceReader.__init__)
            params = sig.parameters
            
            # api_timeout_secondsとmax_retriesパラメータが存在することを検証
            assert 'api_timeout_seconds' in params, "api_timeout_secondsパラメータが存在しません"
            assert 'max_retries' in params, "max_retriesパラメータが存在しません"
            
            # デフォルト値を検証
            assert params['api_timeout_seconds'].default == 60, "api_timeout_secondsのデフォルト値が60ではありません"
            assert params['max_retries'].default == 3, "max_retriesのデフォルト値が3ではありません"
        
        except ImportError:
            pytest.skip("Google Workspaceリーダーがインポートできません")
    
    def test_google_writer_has_retry_configuration(self):
        """
        **Feature: document-format-mcp-server, Property 15: API呼び出しの信頼性**
        
        Google Workspaceライターはリトライとタイムアウトの設定を持つべきである。
        **検証対象: 要件 11.7**
        """
        try:
            from document_format_mcp_server.writers.google_writer import GoogleWorkspaceWriter
            
            # クラスの属性を確認（インスタンス化せずに）
            import inspect
            sig = inspect.signature(GoogleWorkspaceWriter.__init__)
            params = sig.parameters
            
            # api_timeout_secondsとmax_retriesパラメータが存在することを検証
            assert 'api_timeout_seconds' in params, "api_timeout_secondsパラメータが存在しません"
            assert 'max_retries' in params, "max_retriesパラメータが存在しません"
            
            # デフォルト値を検証
            assert params['api_timeout_seconds'].default == 60, "api_timeout_secondsのデフォルト値が60ではありません"
            assert params['max_retries'].default == 3, "max_retriesのデフォルト値が3ではありません"
        
        except ImportError:
            pytest.skip("Google Workspaceライターがインポートできません")

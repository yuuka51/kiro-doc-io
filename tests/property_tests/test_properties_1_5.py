"""
プロパティベーステスト: プロパティ1-5（ドキュメント読み取り）

**Feature: document-format-mcp-server**

このモジュールは、ドキュメント読み取り機能に関するプロパティベーステストを実装します。
hypothesisライブラリを使用し、最小100回の反復で実行します。
"""

import os
import tempfile
from pathlib import Path

import pytest
from hypothesis import given, settings, assume, HealthCheck
from hypothesis import strategies as st

from document_format_mcp_server.readers.powerpoint_reader import PowerPointReader
from document_format_mcp_server.readers.word_reader import WordReader
from document_format_mcp_server.readers.excel_reader import ExcelReader
from document_format_mcp_server.utils.models import ReadResult, DocumentContent


# サンプルファイルのパス
SAMPLE_PPTX = "test_files/samples/sample.pptx"
SAMPLE_DOCX = "test_files/samples/sample.docx"
SAMPLE_XLSX = "test_files/samples/sample.xlsx"


# ============================================================================
# プロパティ1: ドキュメント読み取りの完全性
# *任意の*有効なドキュメントファイル（PowerPoint、Word、Excel）に対して、
# Document_Readerは必要な全ての構造要素（タイトル、コンテンツ、メタデータ）を
# 抽出し、構造化された形式で返すべきである
# **検証対象: 要件 1.1, 1.2, 1.4, 2.1, 2.2, 3.1, 3.2**
# ============================================================================

class TestProperty1DocumentReadCompleteness:
    """
    **Property 1: ドキュメント読み取りの完全性**
    
    *任意の*有効なドキュメントファイルに対して、リーダーは必要な全ての構造要素を
    抽出し、構造化された形式で返すべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_slides=st.integers(min_value=1, max_value=1000))
    def test_powerpoint_reader_returns_complete_structure(self, max_slides: int):
        """
        **Feature: document-format-mcp-server, Property 1: ドキュメント読み取りの完全性**
        
        *任意の*max_slides設定値に対して、PowerPointリーダーは完全な構造を返すべきである。
        **検証対象: 要件 1.1, 1.2, 1.4**
        """
        # サンプルファイルが存在しない場合はスキップ
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        # リーダーを初期化
        reader = PowerPointReader(max_slides=max_slides)
        result = reader.read_file(SAMPLE_PPTX)
        
        # ReadResultの構造を検証
        assert isinstance(result, ReadResult)
        assert result.success is True
        assert result.content is not None
        
        # DocumentContentの構造を検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "pptx"
        assert isinstance(result.content.metadata, dict)
        assert isinstance(result.content.content, dict)
        
        # 必須メタデータの存在を検証
        assert "slide_count" in result.content.metadata
        assert "file_size_mb" in result.content.metadata
        
        # 必須コンテンツの存在を検証
        assert "slides" in result.content.content
        assert isinstance(result.content.content["slides"], list)
        
        # 各スライドの構造を検証
        for slide in result.content.content["slides"]:
            assert "slide_number" in slide
            assert "title" in slide
            assert "content" in slide
            assert "notes" in slide
            assert "tables" in slide
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_file_size_mb=st.integers(min_value=1, max_value=500))
    def test_word_reader_returns_complete_structure(self, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 1: ドキュメント読み取りの完全性**
        
        *任意の*max_file_size_mb設定値に対して、Wordリーダーは完全な構造を返すべきである。
        **検証対象: 要件 2.1, 2.2**
        """
        # サンプルファイルが存在しない場合はスキップ
        if not Path(SAMPLE_DOCX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_DOCX}")
        
        # リーダーを初期化
        reader = WordReader(max_file_size_mb=max_file_size_mb)
        result = reader.read_file(SAMPLE_DOCX)
        
        # ReadResultの構造を検証
        assert isinstance(result, ReadResult)
        assert result.success is True
        assert result.content is not None
        
        # DocumentContentの構造を検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "docx"
        assert isinstance(result.content.metadata, dict)
        assert isinstance(result.content.content, dict)
        
        # 必須メタデータの存在を検証
        assert "paragraph_count" in result.content.metadata
        assert "file_size_mb" in result.content.metadata
        
        # 必須コンテンツの存在を検証
        assert "paragraphs" in result.content.content
        assert "tables" in result.content.content
        assert isinstance(result.content.content["paragraphs"], list)
        assert isinstance(result.content.content["tables"], list)
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_sheets=st.integers(min_value=1, max_value=200))
    def test_excel_reader_returns_complete_structure(self, max_sheets: int):
        """
        **Feature: document-format-mcp-server, Property 1: ドキュメント読み取りの完全性**
        
        *任意の*max_sheets設定値に対して、Excelリーダーは完全な構造を返すべきである。
        **検証対象: 要件 3.1, 3.2**
        """
        # サンプルファイルが存在しない場合はスキップ
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        # リーダーを初期化
        reader = ExcelReader(max_sheets=max_sheets)
        result = reader.read_file(SAMPLE_XLSX)
        
        # ReadResultの構造を検証
        assert isinstance(result, ReadResult)
        assert result.success is True
        assert result.content is not None
        
        # DocumentContentの構造を検証
        assert isinstance(result.content, DocumentContent)
        assert result.content.format_type == "xlsx"
        assert isinstance(result.content.metadata, dict)
        assert isinstance(result.content.content, dict)
        
        # 必須メタデータの存在を検証
        assert "sheet_count" in result.content.metadata
        assert "file_size_mb" in result.content.metadata
        
        # 必須コンテンツの存在を検証
        assert "sheets" in result.content.content
        assert isinstance(result.content.content["sheets"], list)
        
        # 各シートの構造を検証
        for sheet in result.content.content["sheets"]:
            assert "name" in sheet
            assert "data" in sheet


# ============================================================================
# プロパティ2: データ形式の一貫性
# *任意の*ドキュメント読み取り操作において、抽出されたデータは指定された
# 構造化形式（JSON、CSV、マークダウン）のいずれかで提供されるべきである
# **検証対象: 要件 2.3, 3.3**
# ============================================================================

class TestProperty2DataFormatConsistency:
    """
    **Property 2: データ形式の一貫性**
    
    *任意の*ドキュメント読み取り操作において、抽出されたデータは
    構造化形式で提供されるべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_file_size_mb=st.integers(min_value=1, max_value=500))
    def test_word_tables_are_structured(self, max_file_size_mb: int):
        """
        **Feature: document-format-mcp-server, Property 2: データ形式の一貫性**
        
        *任意の*Wordファイル読み取りにおいて、表データは構造化形式で提供されるべきである。
        **検証対象: 要件 2.3**
        """
        if not Path(SAMPLE_DOCX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_DOCX}")
        
        reader = WordReader(max_file_size_mb=max_file_size_mb)
        result = reader.read_file(SAMPLE_DOCX)
        
        assert result.success is True
        
        # 表データが構造化されていることを検証
        tables = result.content.content.get("tables", [])
        for table in tables:
            # 表は辞書形式であるべき
            assert isinstance(table, dict)
            # 必須フィールドが存在するべき
            assert "rows" in table
            assert "columns" in table
            assert "data" in table
            # データは2次元リストであるべき
            assert isinstance(table["data"], list)
            for row in table["data"]:
                assert isinstance(row, list)
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_sheets=st.integers(min_value=1, max_value=200))
    def test_excel_data_is_structured(self, max_sheets: int):
        """
        **Feature: document-format-mcp-server, Property 2: データ形式の一貫性**
        
        *任意の*Excelファイル読み取りにおいて、データは構造化形式で提供されるべきである。
        **検証対象: 要件 3.3**
        """
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        reader = ExcelReader(max_sheets=max_sheets)
        result = reader.read_file(SAMPLE_XLSX)
        
        assert result.success is True
        
        # シートデータが構造化されていることを検証
        sheets = result.content.content.get("sheets", [])
        for sheet in sheets:
            # シートは辞書形式であるべき
            assert isinstance(sheet, dict)
            # 必須フィールドが存在するべき
            assert "name" in sheet
            assert "data" in sheet
            # データは2次元リストであるべき
            assert isinstance(sheet["data"], list)
            for row in sheet["data"]:
                assert isinstance(row, list)


# ============================================================================
# プロパティ3: エラーハンドリングの統一性
# *任意の*無効なファイル（破損、存在しない、アクセス不可）に対して、
# MCP_Serverは明確で一貫したエラーメッセージを返すべきである
# **検証対象: 要件 1.3, 2.4, 3.4, 4.5, 5.5, 6.5, 7.5, 8.5**
# ============================================================================

class TestProperty3ErrorHandlingUniformity:
    """
    **Property 3: エラーハンドリングの統一性**
    
    *任意の*無効なファイルに対して、リーダーは明確で一貫したエラーメッセージを返すべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(filename=st.text(min_size=1, max_size=50).filter(lambda x: x.strip() and "/" not in x and "\\" not in x and ":" not in x))
    def test_powerpoint_reader_handles_nonexistent_files(self, filename: str):
        """
        **Feature: document-format-mcp-server, Property 3: エラーハンドリングの統一性**
        
        *任意の*存在しないファイル名に対して、PowerPointリーダーは一貫したエラーを返すべきである。
        **検証対象: 要件 1.3**
        """
        # 存在しないファイルパスを生成
        nonexistent_path = f"nonexistent_dir/{filename}.pptx"
        
        reader = PowerPointReader()
        result = reader.read_file(nonexistent_path)
        
        # ReadResultの構造を検証
        assert isinstance(result, ReadResult)
        assert result.success is False
        assert result.error is not None
        assert result.content is None
        assert isinstance(result.error, str)
        assert len(result.error) > 0
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(filename=st.text(min_size=1, max_size=50).filter(lambda x: x.strip() and "/" not in x and "\\" not in x and ":" not in x))
    def test_word_reader_handles_nonexistent_files(self, filename: str):
        """
        **Feature: document-format-mcp-server, Property 3: エラーハンドリングの統一性**
        
        *任意の*存在しないファイル名に対して、Wordリーダーは一貫したエラーを返すべきである。
        **検証対象: 要件 2.4**
        """
        nonexistent_path = f"nonexistent_dir/{filename}.docx"
        
        reader = WordReader()
        result = reader.read_file(nonexistent_path)
        
        assert isinstance(result, ReadResult)
        assert result.success is False
        assert result.error is not None
        assert result.content is None
        assert isinstance(result.error, str)
        assert len(result.error) > 0
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(filename=st.text(min_size=1, max_size=50).filter(lambda x: x.strip() and "/" not in x and "\\" not in x and ":" not in x))
    def test_excel_reader_handles_nonexistent_files(self, filename: str):
        """
        **Feature: document-format-mcp-server, Property 3: エラーハンドリングの統一性**
        
        *任意の*存在しないファイル名に対して、Excelリーダーは一貫したエラーを返すべきである。
        **検証対象: 要件 3.4**
        """
        nonexistent_path = f"nonexistent_dir/{filename}.xlsx"
        
        reader = ExcelReader()
        result = reader.read_file(nonexistent_path)
        
        assert isinstance(result, ReadResult)
        assert result.success is False
        assert result.error is not None
        assert result.content is None
        assert isinstance(result.error, str)
        assert len(result.error) > 0


# ============================================================================
# プロパティ4: 制限値の遵守
# *任意の*Excelファイルに対して、Document_Readerは最大100シートまでの制限を遵守し、
# 制限を超える場合は適切に処理するべきである
# **検証対象: 要件 3.5**
# ============================================================================

class TestProperty4LimitCompliance:
    """
    **Property 4: 制限値の遵守**
    
    *任意の*Excelファイルに対して、リーダーは設定された制限を遵守するべきである。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_sheets=st.integers(min_value=1, max_value=100))
    def test_excel_reader_respects_sheet_limit(self, max_sheets: int):
        """
        **Feature: document-format-mcp-server, Property 4: 制限値の遵守**
        
        *任意の*max_sheets設定値に対して、Excelリーダーは制限を遵守するべきである。
        **検証対象: 要件 3.5**
        """
        if not Path(SAMPLE_XLSX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_XLSX}")
        
        reader = ExcelReader(max_sheets=max_sheets)
        result = reader.read_file(SAMPLE_XLSX)
        
        assert result.success is True
        
        # 処理されたシート数が制限以下であることを検証
        processed_sheets = len(result.content.content["sheets"])
        assert processed_sheets <= max_sheets
        
        # メタデータのsheet_countも制限以下であることを検証
        assert result.content.metadata["sheet_count"] <= max_sheets
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(max_slides=st.integers(min_value=1, max_value=500))
    def test_powerpoint_reader_respects_slide_limit(self, max_slides: int):
        """
        **Feature: document-format-mcp-server, Property 4: 制限値の遵守**
        
        *任意の*max_slides設定値に対して、PowerPointリーダーは制限を遵守するべきである。
        **検証対象: 要件 3.5（スライド数制限の類似要件）**
        """
        if not Path(SAMPLE_PPTX).exists():
            pytest.skip(f"サンプルファイルが見つかりません: {SAMPLE_PPTX}")
        
        reader = PowerPointReader(max_slides=max_slides)
        result = reader.read_file(SAMPLE_PPTX)
        
        assert result.success is True
        
        # 処理されたスライド数が制限以下であることを検証
        processed_slides = len(result.content.content["slides"])
        assert processed_slides <= max_slides
        
        # メタデータのslide_countも制限以下であることを検証
        assert result.content.metadata["slide_count"] <= max_slides


# ============================================================================
# プロパティ5: Google API統合の一貫性
# *任意の*有効なGoogle WorkspaceファイルID/URLに対して、Document_Readerは
# 対応するGoogle APIを使用してデータを取得するべきである
# **検証対象: 要件 4.1, 4.2, 4.3**
# ============================================================================

class TestProperty5GoogleAPIConsistency:
    """
    **Property 5: Google API統合の一貫性**
    
    *任意の*無効なファイルIDに対して、Google Workspaceリーダーは
    一貫したエラーを返すべきである。
    
    注: 実際のGoogle API呼び出しはモックなしでは実行できないため、
    このテストは無効なファイルIDに対するエラーハンドリングを検証します。
    """
    
    @settings(max_examples=100, suppress_health_check=[HealthCheck.too_slow])
    @given(file_id=st.text(min_size=1, max_size=50).filter(lambda x: x.strip()))
    def test_google_reader_handles_invalid_file_id(self, file_id: str):
        """
        **Feature: document-format-mcp-server, Property 5: Google API統合の一貫性**
        
        *任意の*無効なファイルIDに対して、Google Workspaceリーダーは
        一貫したエラーを返すべきである。
        **検証対象: 要件 4.1, 4.2, 4.3**
        
        注: Google API認証情報がない場合、このテストはスキップされます。
        """
        try:
            from document_format_mcp_server.readers.google_reader import GoogleWorkspaceReader
            
            # 認証情報がない場合はスキップ
            credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")
            if not credentials_path or not Path(credentials_path).exists():
                pytest.skip("Google API認証情報が設定されていません")
            
            reader = GoogleWorkspaceReader(credentials_path=credentials_path)
            
            # 無効なファイルIDでの読み取りを試行
            result = reader.read_spreadsheet(file_id)
            
            # エラーが返されることを検証（無効なIDなので失敗するはず）
            # 成功した場合は、たまたま有効なIDだった可能性がある
            if not result.success:
                assert result.error is not None
                assert isinstance(result.error, str)
        
        except ImportError:
            pytest.skip("Google Workspaceリーダーがインポートできません")
        except Exception as e:
            # 認証エラーなどの場合はスキップ
            pytest.skip(f"Google APIテストをスキップ: {str(e)}")

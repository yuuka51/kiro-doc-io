"""Excelリーダーのユニットテスト

要件3.1: Excelファイルからすべてのシートのデータを抽出する
要件3.2: 各シートの名前、セルデータ、数式を識別して抽出する
"""

import pytest
from pathlib import Path

from document_format_mcp_server.readers.excel_reader import ExcelReader
from document_format_mcp_server.utils.models import ReadResult, DocumentContent


@pytest.fixture
def sample_xlsx_path():
    """サンプルExcelファイルのパスを返すフィクスチャ"""
    return "test_files/samples/sample.xlsx"


@pytest.fixture
def excel_reader():
    """ExcelReaderインスタンスを返すフィクスチャ"""
    return ExcelReader()


def test_excel_reader_returns_read_result(excel_reader, sample_xlsx_path):
    """ExcelリーダーがReadResultを返すことを検証"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    # ReadResultデータクラスであることを検証
    assert isinstance(result, ReadResult)
    assert hasattr(result, 'success')
    assert hasattr(result, 'content')
    assert hasattr(result, 'error')
    assert hasattr(result, 'file_path')


def test_excel_reader_success_structure(excel_reader, sample_xlsx_path):
    """Excelリーダーの成功時のレスポンス構造を検証"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    # 成功フラグを検証
    assert result.success is True
    assert result.error is None
    
    # DocumentContentを検証
    assert isinstance(result.content, DocumentContent)
    assert result.content.format_type == "xlsx"
    assert isinstance(result.content.metadata, dict)
    assert isinstance(result.content.content, dict)
    
    # メタデータの検証
    assert "sheet_count" in result.content.metadata
    assert "file_size_mb" in result.content.metadata
    
    # コンテンツの検証
    assert "sheets" in result.content.content
    assert isinstance(result.content.content["sheets"], list)


def test_excel_reader_file_not_found(excel_reader):
    """存在しないファイルを読み取った場合のエラーハンドリングを検証"""
    result = excel_reader.read_file("nonexistent_file.xlsx")
    
    # 失敗フラグを検証
    assert result.success is False
    assert result.error is not None
    assert result.content is None


def test_excel_reader_content_structure(excel_reader, sample_xlsx_path):
    """Excelリーダーのコンテンツ構造を検証"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    if result.success and result.content.content["sheets"]:
        # 最初のシートの構造を検証
        first_sheet = result.content.content["sheets"][0]
        assert "name" in first_sheet
        assert "data" in first_sheet
        assert "formulas" in first_sheet
        
        # データ型を検証
        assert isinstance(first_sheet["name"], str)
        assert isinstance(first_sheet["data"], list)
        assert isinstance(first_sheet["formulas"], dict)


def test_excel_reader_extracts_all_sheets(excel_reader, sample_xlsx_path):
    """Excelリーダーがすべてのシートを抽出することを検証（要件3.1）"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.success is True
    sheets = result.content.content["sheets"]
    
    # シートが1つ以上あることを検証
    assert len(sheets) >= 1
    
    # メタデータのsheet_countと一致することを検証
    assert result.content.metadata["sheet_count"] == len(sheets)


def test_excel_reader_extracts_sheet_names(excel_reader, sample_xlsx_path):
    """Excelリーダーが各シートの名前を抽出することを検証（要件3.2）"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.success is True
    sheets = result.content.content["sheets"]
    
    # 各シートに名前があることを検証
    for sheet in sheets:
        assert "name" in sheet
        assert isinstance(sheet["name"], str)
        assert len(sheet["name"]) > 0


def test_excel_reader_extracts_cell_data(excel_reader, sample_xlsx_path):
    """Excelリーダーがセルデータを抽出することを検証（要件3.2）"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.success is True
    sheets = result.content.content["sheets"]
    
    # 各シートにデータがあることを検証
    for sheet in sheets:
        assert "data" in sheet
        assert isinstance(sheet["data"], list)
        
        # データがある場合、各行がリストであることを検証
        if sheet["data"]:
            for row in sheet["data"]:
                assert isinstance(row, list)


def test_excel_reader_extracts_formulas(excel_reader, sample_xlsx_path):
    """Excelリーダーが数式を抽出することを検証（要件3.2）"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.success is True
    sheets = result.content.content["sheets"]
    
    # 各シートにformulasキーがあることを検証
    for sheet in sheets:
        assert "formulas" in sheet
        assert isinstance(sheet["formulas"], dict)


def test_excel_reader_file_path_in_result(excel_reader, sample_xlsx_path):
    """ReadResultにファイルパスが含まれることを検証"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.file_path == sample_xlsx_path


def test_excel_reader_with_custom_max_sheets():
    """カスタムmax_sheetsパラメータでExcelReaderを初期化できることを検証"""
    reader = ExcelReader(max_sheets=50)
    assert reader.max_sheets == 50


def test_excel_reader_with_custom_max_file_size():
    """カスタムmax_file_size_mbパラメータでExcelReaderを初期化できることを検証"""
    reader = ExcelReader(max_file_size_mb=50)
    assert reader.max_file_size_mb == 50


def test_excel_reader_metadata_contains_total_sheets(excel_reader, sample_xlsx_path):
    """Excelリーダーのメタデータに総シート数が含まれることを検証"""
    if not Path(sample_xlsx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_xlsx_path}")
    
    result = excel_reader.read_file(sample_xlsx_path)
    
    assert result.success is True
    assert "total_sheets" in result.content.metadata
    assert isinstance(result.content.metadata["total_sheets"], int)

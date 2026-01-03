"""Wordリーダーのユニットテスト

要件2.1: Wordファイルから本文テキスト、見出し構造、表、リストを抽出する
"""

import pytest
from pathlib import Path

from document_format_mcp_server.readers.word_reader import WordReader
from document_format_mcp_server.utils.models import ReadResult, DocumentContent


@pytest.fixture
def sample_docx_path():
    """サンプルWordファイルのパスを返すフィクスチャ"""
    return "test_files/samples/sample.docx"


@pytest.fixture
def word_reader():
    """WordReaderインスタンスを返すフィクスチャ"""
    return WordReader()


def test_word_reader_returns_read_result(word_reader, sample_docx_path):
    """WordリーダーがReadResultを返すことを検証"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    # ReadResultデータクラスであることを検証
    assert isinstance(result, ReadResult)
    assert hasattr(result, 'success')
    assert hasattr(result, 'content')
    assert hasattr(result, 'error')
    assert hasattr(result, 'file_path')


def test_word_reader_success_structure(word_reader, sample_docx_path):
    """Wordリーダーの成功時のレスポンス構造を検証"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    # 成功フラグを検証
    assert result.success is True
    assert result.error is None
    
    # DocumentContentを検証
    assert isinstance(result.content, DocumentContent)
    assert result.content.format_type == "docx"
    assert isinstance(result.content.metadata, dict)
    assert isinstance(result.content.content, dict)
    
    # メタデータの検証
    assert "paragraph_count" in result.content.metadata
    assert "file_size_mb" in result.content.metadata
    
    # コンテンツの検証
    assert "paragraphs" in result.content.content
    assert isinstance(result.content.content["paragraphs"], list)


def test_word_reader_file_not_found(word_reader):
    """存在しないファイルを読み取った場合のエラーハンドリングを検証"""
    result = word_reader.read_file("nonexistent_file.docx")
    
    # 失敗フラグを検証
    assert result.success is False
    assert result.error is not None
    assert result.content is None


def test_word_reader_content_structure(word_reader, sample_docx_path):
    """Wordリーダーのコンテンツ構造を検証"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    if result.success and result.content.content["paragraphs"]:
        # 最初の段落の構造を検証
        first_para = result.content.content["paragraphs"][0]
        assert "text" in first_para
        assert "style" in first_para
        assert "level" in first_para
        
        # データ型を検証
        assert isinstance(first_para["text"], str)
        assert isinstance(first_para["style"], str)
        assert isinstance(first_para["level"], int)


def test_word_reader_extracts_heading_structure(word_reader, sample_docx_path):
    """Wordリーダーが見出し構造を正しく抽出することを検証（要件2.1, 2.2）"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    assert result.success is True
    paragraphs = result.content.content["paragraphs"]
    
    # 見出しレベルが正しく設定されていることを検証
    for para in paragraphs:
        level = para["level"]
        style = para["style"]
        
        # levelは0以上の整数であること
        assert isinstance(level, int)
        assert level >= 0
        
        # Headingスタイルの場合、levelが1以上であること
        if style.startswith("Heading"):
            assert level >= 1


def test_word_reader_extracts_tables(word_reader, sample_docx_path):
    """Wordリーダーが表データを正しく抽出することを検証（要件2.1）"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    assert result.success is True
    
    # tablesキーが存在することを検証
    assert "tables" in result.content.content
    tables = result.content.content["tables"]
    assert isinstance(tables, list)
    
    # 表がある場合、構造を検証
    if tables:
        for table in tables:
            assert "rows" in table
            assert "columns" in table
            assert "data" in table
            assert isinstance(table["rows"], int)
            assert isinstance(table["columns"], int)
            assert isinstance(table["data"], list)


def test_word_reader_metadata_contains_table_count(word_reader, sample_docx_path):
    """Wordリーダーのメタデータに表数が含まれることを検証"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    assert result.success is True
    assert "table_count" in result.content.metadata
    assert isinstance(result.content.metadata["table_count"], int)


def test_word_reader_file_path_in_result(word_reader, sample_docx_path):
    """ReadResultにファイルパスが含まれることを検証"""
    if not Path(sample_docx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_docx_path}")
    
    result = word_reader.read_file(sample_docx_path)
    
    assert result.file_path == sample_docx_path


def test_word_reader_with_custom_max_file_size():
    """カスタムmax_file_size_mbパラメータでWordReaderを初期化できることを検証"""
    reader = WordReader(max_file_size_mb=50)
    assert reader.max_file_size_mb == 50

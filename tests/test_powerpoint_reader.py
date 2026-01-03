"""PowerPointリーダーのユニットテスト"""

import pytest
import os
import tempfile
from pathlib import Path
from pptx import Presentation

from document_format_mcp_server.readers.powerpoint_reader import PowerPointReader
from document_format_mcp_server.utils.models import ReadResult, DocumentContent


@pytest.fixture
def sample_pptx_path():
    """サンプルPowerPointファイルのパスを返すフィクスチャ"""
    return "test_files/samples/sample.pptx"


@pytest.fixture
def powerpoint_reader():
    """PowerPointReaderインスタンスを返すフィクスチャ"""
    return PowerPointReader()


@pytest.fixture
def powerpoint_reader_with_limits():
    """制限付きPowerPointReaderインスタンスを返すフィクスチャ"""
    return PowerPointReader(max_slides=2, max_file_size_mb=1)


def test_powerpoint_reader_returns_read_result(powerpoint_reader, sample_pptx_path):
    """PowerPointリーダーがReadResultを返すことを検証"""
    # ファイルが存在しない場合はスキップ
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    # ファイルを読み取り
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    # ReadResultデータクラスであることを検証
    assert isinstance(result, ReadResult)
    assert hasattr(result, 'success')
    assert hasattr(result, 'content')
    assert hasattr(result, 'error')
    assert hasattr(result, 'file_path')


def test_powerpoint_reader_success_structure(powerpoint_reader, sample_pptx_path):
    """PowerPointリーダーの成功時のレスポンス構造を検証"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    # 成功フラグを検証
    assert result.success is True
    assert result.error is None
    
    # DocumentContentを検証
    assert isinstance(result.content, DocumentContent)
    assert result.content.format_type == "pptx"
    assert isinstance(result.content.metadata, dict)
    assert isinstance(result.content.content, dict)
    
    # メタデータの検証
    assert "slide_count" in result.content.metadata
    assert "file_size_mb" in result.content.metadata
    
    # コンテンツの検証
    assert "slides" in result.content.content
    assert isinstance(result.content.content["slides"], list)


def test_powerpoint_reader_file_not_found(powerpoint_reader):
    """存在しないファイルを読み取った場合のエラーハンドリングを検証（要件1.3）"""
    result = powerpoint_reader.read_file("nonexistent_file.pptx")
    
    # 失敗フラグを検証
    assert result.success is False
    assert result.error is not None
    assert result.content is None
    assert "見つかりません" in result.error or "not found" in result.error.lower()


def test_powerpoint_reader_content_structure(powerpoint_reader, sample_pptx_path):
    """PowerPointリーダーのコンテンツ構造を検証（要件1.1、1.4）"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    if result.success and result.content.content["slides"]:
        # 最初のスライドの構造を検証
        first_slide = result.content.content["slides"][0]
        assert "slide_number" in first_slide
        assert "title" in first_slide
        assert "content" in first_slide
        assert "notes" in first_slide
        assert "tables" in first_slide
        
        # データ型を検証
        assert isinstance(first_slide["slide_number"], int)
        assert isinstance(first_slide["title"], str)
        assert isinstance(first_slide["content"], str)
        assert isinstance(first_slide["notes"], str)
        assert isinstance(first_slide["tables"], list)


def test_powerpoint_reader_corrupted_file(powerpoint_reader):
    """破損したファイルを読み取った場合のエラーハンドリングを検証（要件1.3）"""
    # 一時的な破損ファイルを作成
    with tempfile.NamedTemporaryFile(mode='w', suffix='.pptx', delete=False) as f:
        f.write("これは破損したPowerPointファイルです")
        temp_path = f.name
    
    try:
        result = powerpoint_reader.read_file(temp_path)
        
        # 失敗フラグを検証
        assert result.success is False
        assert result.error is not None
        assert result.content is None
        assert "破損" in result.error or "読み取り不可能" in result.error or "corrupted" in result.error.lower()
    finally:
        # 一時ファイルを削除
        if os.path.exists(temp_path):
            os.unlink(temp_path)


def test_powerpoint_reader_slide_limit(powerpoint_reader_with_limits, sample_pptx_path):
    """スライド数制限を超えた場合の処理を検証"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    # サンプルファイルのスライド数を確認
    prs = Presentation(sample_pptx_path)
    total_slides = len(prs.slides)
    
    # スライド数が制限を超える場合のみテスト
    if total_slides > 2:
        result = powerpoint_reader_with_limits.read_file(sample_pptx_path)
        
        # 成功するが、警告が含まれることを検証
        assert result.success is True
        assert result.content is not None
        
        # 処理されたスライド数が制限以下であることを検証
        assert len(result.content.content["slides"]) <= 2
        
        # 警告メッセージが含まれることを検証
        if "warning" in result.content.content:
            assert "スライド" in result.content.content["warning"]


def test_powerpoint_reader_extracts_title_and_content(powerpoint_reader, sample_pptx_path):
    """タイトルと本文が正しく抽出されることを検証（要件1.1、1.4）"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    if result.success and result.content.content["slides"]:
        # 少なくとも1つのスライドにタイトルまたはコンテンツがあることを検証
        has_content = False
        for slide in result.content.content["slides"]:
            if slide["title"] or slide["content"]:
                has_content = True
                break
        
        assert has_content, "少なくとも1つのスライドにタイトルまたはコンテンツが必要です"


def test_powerpoint_reader_extracts_tables(powerpoint_reader, sample_pptx_path):
    """表データが正しく抽出されることを検証（要件1.4）"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    if result.success:
        # 表を含むスライドを探す
        for slide in result.content.content["slides"]:
            if slide["tables"]:
                # 表の構造を検証
                table = slide["tables"][0]
                assert "rows" in table
                assert "columns" in table
                assert "data" in table
                assert isinstance(table["data"], list)
                assert len(table["data"]) == table["rows"]
                break


def test_powerpoint_reader_metadata(powerpoint_reader, sample_pptx_path):
    """メタデータが正しく設定されることを検証"""
    if not Path(sample_pptx_path).exists():
        pytest.skip(f"サンプルファイルが見つかりません: {sample_pptx_path}")
    
    result = powerpoint_reader.read_file(sample_pptx_path)
    
    if result.success:
        metadata = result.content.metadata
        
        # メタデータの存在を検証
        assert "slide_count" in metadata
        assert "total_slides" in metadata
        assert "file_size_mb" in metadata
        
        # メタデータの値を検証
        assert metadata["slide_count"] > 0
        assert metadata["total_slides"] > 0
        assert metadata["file_size_mb"] > 0
        assert metadata["slide_count"] <= metadata["total_slides"]

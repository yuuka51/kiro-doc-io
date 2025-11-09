"""共通データモデル定義。"""

from dataclasses import dataclass
from typing import Any, Optional


@dataclass
class DocumentContent:
    """ドキュメントコンテンツの基本クラス。"""
    
    format_type: str  # "pptx", "docx", "xlsx", "google_sheets", etc.
    metadata: dict[str, Any]
    content: dict[str, Any]


@dataclass
class ReadResult:
    """ドキュメント読み取り操作の結果。"""
    
    success: bool
    content: Optional[DocumentContent]
    error: Optional[str]
    file_path: str


@dataclass
class WriteResult:
    """ドキュメント書き込み操作の結果。"""
    
    success: bool
    output_path: Optional[str]
    url: Optional[str]  # Google Workspaceファイル用
    error: Optional[str]

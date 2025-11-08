"""ログ設定モジュール - Document Format MCP Server用のログ設定を提供します。"""

import logging
import os
import sys
from typing import Optional


# ログレベルのマッピング
LOG_LEVELS = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}


def setup_logging(
    name: str = "document_format_mcp_server",
    level: Optional[str] = None,
    format_string: Optional[str] = None,
) -> logging.Logger:
    """
    ログ設定を初期化し、ロガーを返します。
    
    Args:
        name: ロガー名
        level: ログレベル（DEBUG、INFO、WARNING、ERROR、CRITICAL）
               Noneの場合は環境変数MCP_LOG_LEVELから読み込み、デフォルトはINFO
        format_string: ログフォーマット文字列
                      Noneの場合はデフォルトフォーマットを使用
    
    Returns:
        設定済みのロガーインスタンス
    """
    # ログレベルの決定（優先順位: 引数 > 環境変数 > デフォルト）
    if level is None:
        level = os.environ.get("MCP_LOG_LEVEL", "INFO").upper()
    
    # ログレベルの検証
    if level not in LOG_LEVELS:
        level = "INFO"
    
    log_level = LOG_LEVELS[level]
    
    # デフォルトのログフォーマット
    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    # ロガーの取得
    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    
    # 既存のハンドラーをクリア（重複を避けるため）
    if logger.handlers:
        logger.handlers.clear()
    
    # コンソールハンドラーの作成
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setLevel(log_level)
    
    # フォーマッターの作成と設定
    formatter = logging.Formatter(
        format_string,
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    console_handler.setFormatter(formatter)
    
    # ハンドラーをロガーに追加
    logger.addHandler(console_handler)
    
    # 親ロガーへの伝播を防ぐ（重複ログを避けるため）
    logger.propagate = False
    
    return logger


def get_logger(name: str) -> logging.Logger:
    """
    指定された名前のロガーを取得します。
    
    Args:
        name: ロガー名（通常は__name__を使用）
    
    Returns:
        ロガーインスタンス
    """
    # 親ロガーが設定されていない場合は設定
    root_logger = logging.getLogger("document_format_mcp_server")
    if not root_logger.handlers:
        setup_logging()
    
    return logging.getLogger(name)


def set_log_level(level: str) -> None:
    """
    すべてのロガーのログレベルを変更します。
    
    Args:
        level: 新しいログレベル（DEBUG、INFO、WARNING、ERROR、CRITICAL）
    """
    level = level.upper()
    if level not in LOG_LEVELS:
        raise ValueError(f"Invalid log level: {level}. Must be one of {list(LOG_LEVELS.keys())}")
    
    log_level = LOG_LEVELS[level]
    
    # ルートロガーのレベルを変更
    root_logger = logging.getLogger("document_format_mcp_server")
    root_logger.setLevel(log_level)
    
    # すべてのハンドラーのレベルも変更
    for handler in root_logger.handlers:
        handler.setLevel(log_level)

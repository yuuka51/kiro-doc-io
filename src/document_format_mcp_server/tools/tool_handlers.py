"""Tool handlers for MCP document format tools."""

import json
from pathlib import Path
from typing import Any

from ..readers.powerpoint_reader import PowerPointReader
from ..readers.word_reader import WordReader
from ..readers.excel_reader import ExcelReader
from ..readers.google_reader import GoogleWorkspaceReader
from ..utils.config import Config
from ..utils.errors import (
    DocumentMCPError,
    ValidationError,
    FileNotFoundError as DocFileNotFoundError,
)


class ToolHandlers:
    """Handlers for all MCP tools."""
    
    def __init__(self, config: Config):
        """
        Initialize tool handlers.
        
        Args:
            config: Configuration object
        """
        self.config = config
        
        # リーダーの初期化
        self.powerpoint_reader = PowerPointReader(
            max_slides=config.max_slides,
            max_file_size_mb=config.max_file_size_mb
        )
        self.word_reader = WordReader(
            max_file_size_mb=config.max_file_size_mb
        )
        self.excel_reader = ExcelReader(
            max_sheets=config.max_sheets,
            max_file_size_mb=config.max_file_size_mb
        )
        
        # Google Workspaceリーダーの初期化（有効な場合のみ）
        self.google_reader = None
        if config.enable_google_workspace:
            try:
                self.google_reader = GoogleWorkspaceReader(
                    credentials_path=config.google_credentials_path
                )
            except Exception:
                # Google認証情報が利用できない場合はスキップ
                pass
        
        # ライターの初期化
        from ..writers.powerpoint_writer import PowerPointWriter
        from ..writers.word_writer import WordWriter
        from ..writers.excel_writer import ExcelWriter
        from ..writers.google_writer import GoogleWorkspaceWriter
        
        self.powerpoint_writer = PowerPointWriter()
        self.word_writer = WordWriter()
        self.excel_writer = ExcelWriter()
        
        # Google Workspaceライターの初期化（有効な場合のみ）
        self.google_writer = None
        if config.enable_google_workspace:
            try:
                self.google_writer = GoogleWorkspaceWriter(
                    credentials_path=config.google_credentials_path
                )
            except Exception:
                # Google認証情報が利用できない場合はスキップ
                pass
    
    # 読み取りツールハンドラー
    
    async def handle_read_powerpoint(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        PowerPointファイルを読み取る。
        
        Args:
            arguments: ツール引数（file_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            file_path = self._validate_file_path(arguments, "file_path")
            
            # ファイル読み取り
            result = self.powerpoint_reader.read_file(file_path)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_read_word(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        Wordファイルを読み取る。
        
        Args:
            arguments: ツール引数（file_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            file_path = self._validate_file_path(arguments, "file_path")
            
            # ファイル読み取り
            result = self.word_reader.read_file(file_path)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_read_excel(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        Excelファイルを読み取る。
        
        Args:
            arguments: ツール引数（file_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            file_path = self._validate_file_path(arguments, "file_path")
            
            # ファイル読み取り
            result = self.excel_reader.read_file(file_path)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_read_google_spreadsheet(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleスプレッドシートを読み取る。
        
        Args:
            arguments: ツール引数（file_idを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_reader:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            file_id = self._validate_string_param(arguments, "file_id")
            
            # ファイル読み取り
            result = self.google_reader.read_spreadsheet(file_id)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_read_google_document(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleドキュメントを読み取る。
        
        Args:
            arguments: ツール引数（file_idを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_reader:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            file_id = self._validate_string_param(arguments, "file_id")
            
            # ファイル読み取り
            result = self.google_reader.read_document(file_id)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_read_google_slides(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleスライドを読み取る。
        
        Args:
            arguments: ツール引数（file_idを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_reader:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            file_id = self._validate_string_param(arguments, "file_id")
            
            # ファイル読み取り
            result = self.google_reader.read_slides(file_id)
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    # 書き込みツールハンドラー（将来の実装用）
    
    async def handle_write_powerpoint(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        PowerPointファイルを生成する。
        
        Args:
            arguments: ツール引数（data、output_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "output_path" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: output_path",
                    {"missing_parameter": "output_path"}
                )
            
            data = arguments["data"]
            output_path = arguments["output_path"]
            
            # ファイル生成
            result_path = self.powerpoint_writer.create_presentation(data, output_path)
            
            result = {
                "success": True,
                "output_path": result_path,
                "message": f"PowerPointファイルを生成しました: {result_path}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_write_word(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        Wordファイルを生成する。
        
        Args:
            arguments: ツール引数（data、output_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "output_path" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: output_path",
                    {"missing_parameter": "output_path"}
                )
            
            data = arguments["data"]
            output_path = arguments["output_path"]
            
            # ファイル生成
            result_path = self.word_writer.create_document(data, output_path)
            
            result = {
                "success": True,
                "output_path": result_path,
                "message": f"Wordファイルを生成しました: {result_path}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_write_excel(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        Excelファイルを生成する。
        
        Args:
            arguments: ツール引数（data、output_pathを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "output_path" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: output_path",
                    {"missing_parameter": "output_path"}
                )
            
            data = arguments["data"]
            output_path = arguments["output_path"]
            
            # ファイル生成
            result_path = self.excel_writer.create_workbook(data, output_path)
            
            result = {
                "success": True,
                "output_path": result_path,
                "message": f"Excelファイルを生成しました: {result_path}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_write_google_spreadsheet(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleスプレッドシートを生成する。
        
        Args:
            arguments: ツール引数（data、titleを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_writer:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "title" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: title",
                    {"missing_parameter": "title"}
                )
            
            data = arguments["data"]
            title = arguments["title"]
            
            # ファイル生成
            result_url = self.google_writer.create_spreadsheet(data, title)
            
            result = {
                "success": True,
                "url": result_url,
                "message": f"Googleスプレッドシートを生成しました: {result_url}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_write_google_document(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleドキュメントを生成する。
        
        Args:
            arguments: ツール引数（data、titleを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_writer:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "title" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: title",
                    {"missing_parameter": "title"}
                )
            
            data = arguments["data"]
            title = arguments["title"]
            
            # ファイル生成
            result_url = self.google_writer.create_document(data, title)
            
            result = {
                "success": True,
                "url": result_url,
                "message": f"Googleドキュメントを生成しました: {result_url}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    async def handle_write_google_slides(
        self, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        """
        Googleスライドを生成する。
        
        Args:
            arguments: ツール引数（data、titleを含む）
            
        Returns:
            成功/失敗のレスポンス
        """
        try:
            # Google Workspaceが有効か確認
            if not self.google_writer:
                raise ValidationError(
                    "Google Workspaceが有効になっていません",
                    {"enable_google_workspace": self.config.enable_google_workspace}
                )
            
            # パラメータ検証
            if "data" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: data",
                    {"missing_parameter": "data"}
                )
            
            if "title" not in arguments:
                raise ValidationError(
                    "必須パラメータが不足しています: title",
                    {"missing_parameter": "title"}
                )
            
            data = arguments["data"]
            title = arguments["title"]
            
            # ファイル生成
            result_url = self.google_writer.create_slides(data, title)
            
            result = {
                "success": True,
                "url": result_url,
                "message": f"Googleスライドを生成しました: {result_url}"
            }
            
            return self._success_response(result)
        
        except DocumentMCPError as e:
            return self._error_response(e)
        except Exception as e:
            return self._error_response(
                DocumentMCPError(f"予期しないエラー: {str(e)}")
            )
    
    # ヘルパーメソッド
    
    def _validate_file_path(self, arguments: dict[str, Any], key: str) -> str:
        """
        ファイルパスパラメータを検証する。
        
        Args:
            arguments: ツール引数
            key: パラメータキー
            
        Returns:
            検証済みファイルパス
            
        Raises:
            ValidationError: パラメータが無効な場合
            FileNotFoundError: ファイルが存在しない場合
        """
        if key not in arguments:
            raise ValidationError(
                f"必須パラメータが不足しています: {key}",
                {"missing_parameter": key}
            )
        
        file_path = arguments[key]
        
        if not isinstance(file_path, str):
            raise ValidationError(
                f"パラメータは文字列である必要があります: {key}",
                {"parameter": key, "type": type(file_path).__name__}
            )
        
        if not file_path.strip():
            raise ValidationError(
                f"パラメータが空です: {key}",
                {"parameter": key}
            )
        
        # ファイルの存在確認
        path = Path(file_path).expanduser()
        if not path.exists():
            raise DocFileNotFoundError(
                f"ファイルが見つかりません: {file_path}",
                {"file_path": str(path)}
            )
        
        if not path.is_file():
            raise ValidationError(
                f"指定されたパスはファイルではありません: {file_path}",
                {"file_path": str(path)}
            )
        
        return str(path)
    
    def _validate_string_param(self, arguments: dict[str, Any], key: str) -> str:
        """
        文字列パラメータを検証する。
        
        Args:
            arguments: ツール引数
            key: パラメータキー
            
        Returns:
            検証済み文字列
            
        Raises:
            ValidationError: パラメータが無効な場合
        """
        if key not in arguments:
            raise ValidationError(
                f"必須パラメータが不足しています: {key}",
                {"missing_parameter": key}
            )
        
        value = arguments[key]
        
        if not isinstance(value, str):
            raise ValidationError(
                f"パラメータは文字列である必要があります: {key}",
                {"parameter": key, "type": type(value).__name__}
            )
        
        if not value.strip():
            raise ValidationError(
                f"パラメータが空です: {key}",
                {"parameter": key}
            )
        
        return value
    
    def _success_response(self, content: Any) -> dict[str, Any]:
        """
        成功レスポンスを生成する。
        
        Args:
            content: レスポンスコンテンツ
            
        Returns:
            成功レスポンス
        """
        return {
            "content": [
                {
                    "type": "text",
                    "text": json.dumps(content, ensure_ascii=False, indent=2)
                }
            ]
        }
    
    def _error_response(self, error: DocumentMCPError) -> dict[str, Any]:
        """
        エラーレスポンスを生成する。
        
        Args:
            error: エラーオブジェクト
            
        Returns:
            エラーレスポンス
        """
        error_data = {
            "success": False,
            "error": {
                "type": type(error).__name__,
                "message": error.message,
                "details": error.details
            }
        }
        
        return {
            "content": [
                {
                    "type": "text",
                    "text": json.dumps(error_data, ensure_ascii=False, indent=2)
                }
            ],
            "isError": True
        }

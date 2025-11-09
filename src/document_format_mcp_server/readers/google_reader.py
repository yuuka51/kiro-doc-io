"""Google Workspace file reader."""

import os
import re
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from ..utils.errors import (
    AuthenticationError,
    APIError,
    ConfigurationError,
)
from ..utils.logging_config import get_logger
from ..utils.models import ReadResult, DocumentContent


# ロガーの取得
logger = get_logger(__name__)


# Google APIのスコープ
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/presentations.readonly',
]


class GoogleWorkspaceReader:
    """Google Workspaceファイルを読み取るクラス。"""

    def __init__(
        self,
        credentials_path: str,
        api_timeout_seconds: int = 60,
        max_retries: int = 3
    ):
        """
        GoogleWorkspaceReaderを初期化する。

        Args:
            credentials_path: Google API認証情報ファイルのパス
            api_timeout_seconds: API呼び出しのタイムアウト（秒）
            max_retries: 最大リトライ回数

        Raises:
            ConfigurationError: 認証情報ファイルが見つからない場合
            AuthenticationError: 認証に失敗した場合
        """
        self.credentials_path = credentials_path
        self.api_timeout = api_timeout_seconds
        self.max_retries = max_retries
        self.credentials = None
        self._authenticate()

    def _authenticate(self) -> None:
        """
        Google APIの認証を行う。

        Raises:
            ConfigurationError: 認証情報ファイルが見つからない場合
            AuthenticationError: 認証に失敗した場合
        """
        logger.info("Google API認証を開始")
        
        # 認証情報ファイルの存在確認
        if not os.path.exists(self.credentials_path):
            logger.error(f"認証情報ファイルが見つかりません: {self.credentials_path}")
            raise ConfigurationError(
                f"Google API認証情報ファイルが見つかりません: {self.credentials_path}",
                details={"credentials_path": self.credentials_path}
            )

        try:
            creds = None
            token_path = os.path.join(
                os.path.dirname(self.credentials_path),
                'token.json'
            )

            # トークンファイルが存在する場合は読み込む
            if os.path.exists(token_path):
                logger.debug("既存の認証トークンを読み込み")
                creds = Credentials.from_authorized_user_file(token_path, SCOPES)

            # 有効な認証情報がない場合は新規認証
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    try:
                        logger.info("認証トークンをリフレッシュ")
                        creds.refresh(Request())
                    except Exception as e:
                        logger.error("認証トークンのリフレッシュに失敗", exc_info=True)
                        raise AuthenticationError(
                            "Google API認証トークンのリフレッシュに失敗しました",
                            details={"error": str(e)}
                        )
                else:
                    try:
                        logger.info("新規認証フローを開始")
                        flow = InstalledAppFlow.from_client_secrets_file(
                            self.credentials_path, SCOPES
                        )
                        creds = flow.run_local_server(port=0)
                    except Exception as e:
                        logger.error("Google API認証に失敗", exc_info=True)
                        raise AuthenticationError(
                            "Google API認証に失敗しました",
                            details={"error": str(e)}
                        )

                # トークンを保存
                try:
                    with open(token_path, 'w') as token:
                        token.write(creds.to_json())
                    logger.debug("認証トークンを保存")
                except Exception:
                    # トークン保存失敗は警告のみ（認証自体は成功）
                    logger.warning("認証トークンの保存に失敗しましたが、認証は成功しました")
                    pass

            self.credentials = creds
            logger.info("Google API認証が完了")

        except (ConfigurationError, AuthenticationError):
            raise
        except Exception as e:
            logger.error("Google API認証処理中にエラーが発生", exc_info=True)
            raise AuthenticationError(
                "Google API認証処理中にエラーが発生しました",
                details={"error": str(e)}
            )

    def _execute_with_retry(self, request_func, *args, **kwargs):
        """
        Google API呼び出しをリトライロジック付きで実行する。

        Args:
            request_func: 実行するAPI呼び出し関数
            *args: 関数の位置引数
            **kwargs: 関数のキーワード引数

        Returns:
            API呼び出しの結果

        Raises:
            APIError: リトライ後も失敗した場合
        """
        import socket
        import concurrent.futures
        from googleapiclient.errors import HttpError
        
        last_exception = None
        
        for attempt in range(self.max_retries):
            try:
                # タイムアウト付きでAPI呼び出しを実行
                with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                    future = executor.submit(request_func, *args, **kwargs)
                    try:
                        result = future.result(timeout=self.api_timeout)
                        logger.debug(f"API呼び出しが成功（試行回数: {attempt + 1}）")
                        return result
                    except concurrent.futures.TimeoutError:
                        logger.warning(
                            f"API呼び出しがタイムアウトしました（{self.api_timeout}秒）。"
                            f"リトライします（試行回数: {attempt + 1}/{self.max_retries}）"
                        )
                        last_exception = TimeoutError(f"API呼び出しが{self.api_timeout}秒でタイムアウトしました")
                        if attempt < self.max_retries - 1:
                            wait_time = (2 ** attempt)
                            time.sleep(wait_time)
                        continue
                
            except HttpError as e:
                last_exception = e
                
                # リトライ可能なエラーかチェック
                if e.resp.status in [429, 500, 502, 503, 504]:
                    # レート制限またはサーバーエラー
                    wait_time = (2 ** attempt)  # 指数バックオフ
                    logger.warning(
                        f"API呼び出しが失敗（ステータス: {e.resp.status}）。"
                        f"{wait_time}秒後にリトライします（試行回数: {attempt + 1}/{self.max_retries}）"
                    )
                    time.sleep(wait_time)
                else:
                    # リトライ不可能なエラー（404, 403など）
                    logger.error(f"リトライ不可能なエラー（ステータス: {e.resp.status}）")
                    raise
                    
            except (socket.error, socket.timeout, TimeoutError) as e:
                last_exception = e
                
                # ネットワークエラー
                wait_time = (2 ** attempt)  # 指数バックオフ
                logger.warning(
                    f"ネットワークエラーが発生: {str(e)}。"
                    f"{wait_time}秒後にリトライします（試行回数: {attempt + 1}/{self.max_retries}）"
                )
                time.sleep(wait_time)
                
            except Exception as e:
                # その他の予期しないエラー
                logger.error(f"予期しないエラーが発生: {str(e)}", exc_info=True)
                raise
        
        # 最大リトライ回数に達した
        logger.error(f"最大リトライ回数（{self.max_retries}）に達しました")
        raise APIError(
            f"API呼び出しが{self.max_retries}回失敗しました",
            details={"last_error": str(last_exception)}
        )

    def _extract_file_id(self, file_id_or_url: str) -> str:
        """
        URLまたはファイルIDからファイルIDを抽出する。

        Args:
            file_id_or_url: GoogleファイルのURLまたはファイルID

        Returns:
            抽出されたファイルID

        Examples:
            >>> reader._extract_file_id("1abc123")
            "1abc123"
            >>> reader._extract_file_id("https://docs.google.com/spreadsheets/d/1abc123/edit")
            "1abc123"
        """
        # URLの場合はファイルIDを抽出
        patterns = [
            r'/d/([a-zA-Z0-9-_]+)',  # /d/{file_id}
            r'id=([a-zA-Z0-9-_]+)',  # id={file_id}
        ]

        for pattern in patterns:
            match = re.search(pattern, file_id_or_url)
            if match:
                return match.group(1)

        # パターンにマッチしない場合はそのまま返す（ファイルIDとして扱う）
        return file_id_or_url

    def read_spreadsheet(self, file_id_or_url: str) -> ReadResult:
        """
        Googleスプレッドシートを読み取る。

        Args:
            file_id_or_url: スプレッドシートのURLまたはファイルID

        Returns:
            ReadResult: 読み取り結果を含むデータクラス

        Raises:
            FileNotFoundError: ファイルが見つからない場合
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        file_id = self._extract_file_id(file_id_or_url)
        logger.info(f"Googleスプレッドシートの読み込みを開始: {file_id}")
        start_time = time.time()

        try:
            service = build('sheets', 'v4', credentials=self.credentials)

            # スプレッドシートのメタデータを取得
            spreadsheet = self._execute_with_retry(
                lambda: service.spreadsheets().get(
                    spreadsheetId=file_id
                ).execute()
            )

            title = spreadsheet.get('properties', {}).get('title', '')
            sheets_data = []

            # 各シートのデータを取得
            for sheet in spreadsheet.get('sheets', []):
                sheet_properties = sheet.get('properties', {})
                sheet_title = sheet_properties.get('title', '')

                # シートのデータを取得
                result = self._execute_with_retry(
                    lambda st=sheet_title: service.spreadsheets().values().get(
                        spreadsheetId=file_id,
                        range=st
                    ).execute()
                )

                values = result.get('values', [])

                sheet_data = {
                    "name": sheet_title,
                    "data": values,
                    "row_count": len(values),
                    "column_count": max(len(row) for row in values) if values else 0
                }
                sheets_data.append(sheet_data)

            content_dict = {
                "title": title,
                "sheets": sheets_data
            }
            
            # メタデータを作成
            metadata = {
                "title": title,
                "sheet_count": len(sheets_data),
                "file_id": file_id
            }
            
            # DocumentContentを作成
            document_content = DocumentContent(
                format_type="google_sheets",
                metadata=metadata,
                content=content_dict
            )
            
            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleスプレッドシートの読み込みが完了: {file_id} "
                f"(タイトル: {title}, シート数: {len(sheets_data)}, 処理時間: {elapsed_time:.2f}秒)"
            )
            logger.debug(f"抽出されたデータの概要: シート数={len(sheets_data)}")
            
            # ReadResultを返す
            return ReadResult(
                success=True,
                content=document_content,
                error=None,
                file_path=file_id
            )

        except HttpError as e:
            if e.resp.status == 404:
                logger.error(f"Googleスプレッドシートが見つかりません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"指定されたGoogleスプレッドシートが見つかりません: {file_id}",
                    file_path=file_id
                )
            elif e.resp.status == 403:
                logger.error(f"Googleスプレッドシートへのアクセス権限がありません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleスプレッドシートへのアクセス権限がありません: {file_id}",
                    file_path=file_id
                )
            else:
                logger.error(
                    f"Googleスプレッドシートの読み取り中にエラーが発生: {file_id}",
                    exc_info=True
                )
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleスプレッドシートの読み取り中にエラーが発生しました: {str(e)}",
                    file_path=file_id
                )
        except Exception as e:
            logger.error(
                f"Googleスプレッドシートの読み取り中に予期しないエラーが発生: {file_id}",
                exc_info=True
            )
            return ReadResult(
                success=False,
                content=None,
                error=f"Googleスプレッドシートの読み取り中に予期しないエラーが発生しました: {str(e)}",
                file_path=file_id
            )

    def read_document(self, file_id_or_url: str) -> ReadResult:
        """
        Googleドキュメントを読み取る。

        Args:
            file_id_or_url: ドキュメントのURLまたはファイルID

        Returns:
            ReadResult: 読み取り結果を含むデータクラス

        Raises:
            FileNotFoundError: ファイルが見つからない場合
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        file_id = self._extract_file_id(file_id_or_url)
        logger.info(f"Googleドキュメントの読み込みを開始: {file_id}")
        start_time = time.time()

        try:
            service = build('docs', 'v1', credentials=self.credentials)

            # ドキュメントを取得
            document = self._execute_with_retry(
                lambda: service.documents().get(documentId=file_id).execute()
            )

            title = document.get('title', '')
            content_data = []

            # ドキュメントの内容を解析
            for element in document.get('body', {}).get('content', []):
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    paragraph_style = paragraph.get('paragraphStyle', {})
                    named_style = paragraph_style.get('namedStyleType', 'NORMAL_TEXT')

                    # テキストを抽出
                    text_parts = []
                    for text_element in paragraph.get('elements', []):
                        if 'textRun' in text_element:
                            text_parts.append(text_element['textRun'].get('content', ''))

                    text = ''.join(text_parts).strip()

                    if text:
                        content_item = {
                            "type": "heading" if "HEADING" in named_style else "paragraph",
                            "text": text,
                            "style": named_style
                        }

                        # 見出しレベルを抽出
                        if "HEADING" in named_style:
                            level_match = re.search(r'HEADING_(\d+)', named_style)
                            if level_match:
                                content_item["level"] = int(level_match.group(1))

                        content_data.append(content_item)

                elif 'table' in element:
                    table = element['table']
                    table_data = []

                    for row in table.get('tableRows', []):
                        row_data = []
                        for cell in row.get('tableCells', []):
                            cell_text = []
                            for cell_element in cell.get('content', []):
                                if 'paragraph' in cell_element:
                                    for text_element in cell_element['paragraph'].get('elements', []):
                                        if 'textRun' in text_element:
                                            cell_text.append(text_element['textRun'].get('content', ''))
                            row_data.append(''.join(cell_text).strip())
                        table_data.append(row_data)

                    content_data.append({
                        "type": "table",
                        "data": table_data,
                        "rows": len(table_data),
                        "columns": max(len(row) for row in table_data) if table_data else 0
                    })

            content_dict = {
                "title": title,
                "content": content_data
            }
            
            # メタデータを作成
            metadata = {
                "title": title,
                "content_count": len(content_data),
                "file_id": file_id
            }
            
            # DocumentContentを作成
            document_content = DocumentContent(
                format_type="google_docs",
                metadata=metadata,
                content=content_dict
            )
            
            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleドキュメントの読み込みが完了: {file_id} "
                f"(タイトル: {title}, コンテンツ数: {len(content_data)}, 処理時間: {elapsed_time:.2f}秒)"
            )
            logger.debug(f"抽出されたデータの概要: コンテンツ数={len(content_data)}")
            
            # ReadResultを返す
            return ReadResult(
                success=True,
                content=document_content,
                error=None,
                file_path=file_id
            )

        except HttpError as e:
            if e.resp.status == 404:
                logger.error(f"Googleドキュメントが見つかりません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"指定されたGoogleドキュメントが見つかりません: {file_id}",
                    file_path=file_id
                )
            elif e.resp.status == 403:
                logger.error(f"Googleドキュメントへのアクセス権限がありません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleドキュメントへのアクセス権限がありません: {file_id}",
                    file_path=file_id
                )
            else:
                logger.error(
                    f"Googleドキュメントの読み取り中にエラーが発生: {file_id}",
                    exc_info=True
                )
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleドキュメントの読み取り中にエラーが発生しました: {str(e)}",
                    file_path=file_id
                )
        except Exception as e:
            logger.error(
                f"Googleドキュメントの読み取り中に予期しないエラーが発生: {file_id}",
                exc_info=True
            )
            return ReadResult(
                success=False,
                content=None,
                error=f"Googleドキュメントの読み取り中に予期しないエラーが発生しました: {str(e)}",
                file_path=file_id
            )

    def read_slides(self, file_id_or_url: str) -> ReadResult:
        """
        Googleスライドを読み取る。

        Args:
            file_id_or_url: スライドのURLまたはファイルID

        Returns:
            ReadResult: 読み取り結果を含むデータクラス

        Raises:
            FileNotFoundError: ファイルが見つからない場合
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        file_id = self._extract_file_id(file_id_or_url)
        logger.info(f"Googleスライドの読み込みを開始: {file_id}")
        start_time = time.time()

        try:
            service = build('slides', 'v1', credentials=self.credentials)

            # プレゼンテーションを取得
            presentation = self._execute_with_retry(
                lambda: service.presentations().get(
                    presentationId=file_id
                ).execute()
            )

            title = presentation.get('title', '')
            slides_data = []

            # 各スライドを処理
            for idx, slide in enumerate(presentation.get('slides', []), start=1):
                slide_elements = []

                # スライド内の要素を処理
                for page_element in slide.get('pageElements', []):
                    if 'shape' in page_element:
                        shape = page_element['shape']

                        # テキストを抽出
                        if 'text' in shape:
                            text_parts = []
                            for text_element in shape['text'].get('textElements', []):
                                if 'textRun' in text_element:
                                    text_parts.append(text_element['textRun'].get('content', ''))

                            text = ''.join(text_parts).strip()
                            if text:
                                slide_elements.append({
                                    "type": "text",
                                    "content": text
                                })

                    elif 'table' in page_element:
                        table = page_element['table']
                        table_data = []

                        for row in table.get('tableRows', []):
                            row_data = []
                            for cell in row.get('tableCells', []):
                                cell_text = []
                                if 'text' in cell:
                                    for text_element in cell['text'].get('textElements', []):
                                        if 'textRun' in text_element:
                                            cell_text.append(text_element['textRun'].get('content', ''))
                                row_data.append(''.join(cell_text).strip())
                            table_data.append(row_data)

                        slide_elements.append({
                            "type": "table",
                            "content": {
                                "data": table_data,
                                "rows": len(table_data),
                                "columns": max(len(row) for row in table_data) if table_data else 0
                            }
                        })

                    elif 'image' in page_element:
                        image = page_element['image']
                        slide_elements.append({
                            "type": "image",
                            "content": {
                                "description": image.get('contentUrl', ''),
                                "title": image.get('title', '')
                            }
                        })

                slides_data.append({
                    "slide_number": idx,
                    "elements": slide_elements
                })

            content_dict = {
                "title": title,
                "slides": slides_data
            }
            
            # メタデータを作成
            metadata = {
                "title": title,
                "slide_count": len(slides_data),
                "file_id": file_id
            }
            
            # DocumentContentを作成
            document_content = DocumentContent(
                format_type="google_slides",
                metadata=metadata,
                content=content_dict
            )
            
            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleスライドの読み込みが完了: {file_id} "
                f"(タイトル: {title}, スライド数: {len(slides_data)}, 処理時間: {elapsed_time:.2f}秒)"
            )
            logger.debug(f"抽出されたデータの概要: スライド数={len(slides_data)}")
            
            # ReadResultを返す
            return ReadResult(
                success=True,
                content=document_content,
                error=None,
                file_path=file_id
            )

        except HttpError as e:
            if e.resp.status == 404:
                logger.error(f"Googleスライドが見つかりません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"指定されたGoogleスライドが見つかりません: {file_id}",
                    file_path=file_id
                )
            elif e.resp.status == 403:
                logger.error(f"Googleスライドへのアクセス権限がありません: {file_id}")
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleスライドへのアクセス権限がありません: {file_id}",
                    file_path=file_id
                )
            else:
                logger.error(
                    f"Googleスライドの読み取り中にエラーが発生: {file_id}",
                    exc_info=True
                )
                return ReadResult(
                    success=False,
                    content=None,
                    error=f"Googleスライドの読み取り中にエラーが発生しました: {str(e)}",
                    file_path=file_id
                )
        except Exception as e:
            logger.error(
                f"Googleスライドの読み取り中に予期しないエラーが発生: {file_id}",
                exc_info=True
            )
            return ReadResult(
                success=False,
                content=None,
                error=f"Googleスライドの読み取り中に予期しないエラーが発生しました: {str(e)}",
                file_path=file_id
            )

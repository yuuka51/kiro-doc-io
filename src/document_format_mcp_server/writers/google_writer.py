"""Google Workspace file writer."""

import os
import time
from typing import Any

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
from ..utils.models import WriteResult


# ロガーの取得
logger = get_logger(__name__)


# Google APIのスコープ（書き込み権限を含む）
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file',
]


class GoogleWorkspaceWriter:
    """Google Workspaceファイルを作成するクラス。"""

    def __init__(
        self,
        credentials_path: str,
        api_timeout_seconds: int = 60,
        max_retries: int = 3
    ):
        """
        GoogleWorkspaceWriterを初期化する。

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
        logger.info("Google API認証を開始（書き込み権限）")
        
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
                'token_writer.json'
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
            logger.info("Google API認証が完了（書き込み権限）")

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
        from googleapiclient.errors import HttpError
        
        last_exception = None
        
        for attempt in range(self.max_retries):
            try:
                # タイムアウト付きでAPI呼び出しを実行
                result = request_func(*args, **kwargs)
                
                logger.debug(f"API呼び出しが成功（試行回数: {attempt + 1}）")
                return result
                
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

    def create_spreadsheet(self, data: dict[str, Any], title: str) -> WriteResult:
        """
        Googleスプレッドシートを作成する。

        Args:
            data: スプレッドシートのデータ
                {
                    "sheets": [
                        {
                            "name": str,
                            "data": [[cell_value, ...], ...]
                        }
                    ]
                }
            title: スプレッドシートのタイトル

        Returns:
            作成されたスプレッドシートのURL

        Raises:
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        logger.info(f"Googleスプレッドシートの作成を開始: {title}")
        start_time = time.time()

        try:
            service = build('sheets', 'v4', credentials=self.credentials)

            # スプレッドシートを作成
            spreadsheet_body = {
                'properties': {
                    'title': title
                }
            }

            spreadsheet = self._execute_with_retry(
                lambda: service.spreadsheets().create(
                    body=spreadsheet_body
                ).execute()
            )

            spreadsheet_id = spreadsheet['spreadsheetId']
            spreadsheet_url = spreadsheet['spreadsheetUrl']

            logger.info(f"スプレッドシートを作成: {spreadsheet_id}")

            # シートデータを書き込む
            sheets = data.get('sheets', [])
            
            if sheets:
                # 最初のシートのデータを書き込む
                first_sheet = sheets[0]
                sheet_name = first_sheet.get('name', 'Sheet1')
                sheet_data = first_sheet.get('data', [])

                if sheet_data:
                    # デフォルトのSheet1にデータを書き込む
                    self._execute_with_retry(
                        lambda: service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range='Sheet1!A1',
                            valueInputOption='RAW',
                            body={'values': sheet_data}
                        ).execute()
                    )

                    # Sheet1の名前を変更
                    if sheet_name != 'Sheet1':
                        requests = [{
                            'updateSheetProperties': {
                                'properties': {
                                    'sheetId': 0,
                                    'title': sheet_name
                                },
                                'fields': 'title'
                            }
                        }]
                        self._execute_with_retry(
                            lambda: service.spreadsheets().batchUpdate(
                                spreadsheetId=spreadsheet_id,
                                body={'requests': requests}
                            ).execute()
                        )

                # 追加のシートを作成
                for idx, sheet in enumerate(sheets[1:], start=1):
                    sheet_name = sheet.get('name', f'Sheet{idx + 1}')
                    sheet_data = sheet.get('data', [])

                    # 新しいシートを追加
                    requests = [{
                        'addSheet': {
                            'properties': {
                                'title': sheet_name
                            }
                        }
                    }]
                    self._execute_with_retry(
                        lambda req=requests: service.spreadsheets().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={'requests': req}
                        ).execute()
                    )

                    # データを書き込む
                    if sheet_data:
                        self._execute_with_retry(
                            lambda sn=sheet_name, sd=sheet_data: service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id,
                                range=f'{sn}!A1',
                                valueInputOption='RAW',
                                body={'values': sd}
                            ).execute()
                        )

            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleスプレッドシートの作成が完了: {spreadsheet_id} "
                f"(URL: {spreadsheet_url}, シート数: {len(sheets)}, 処理時間: {elapsed_time:.2f}秒)"
            )

            # WriteResultを返す
            return WriteResult(
                success=True,
                output_path=None,
                url=spreadsheet_url,
                error=None
            )

        except HttpError as e:
            if e.resp.status == 403:
                logger.error("Googleスプレッドシートの作成権限がありません")
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error="Googleスプレッドシートの作成権限がありません"
                )
            else:
                logger.error(
                    "Googleスプレッドシートの作成中にエラーが発生",
                    exc_info=True
                )
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error=f"Googleスプレッドシートの作成中にエラーが発生しました: {str(e)}"
                )
        except Exception as e:
            logger.error(
                "Googleスプレッドシートの作成中に予期しないエラーが発生",
                exc_info=True
            )
            return WriteResult(
                success=False,
                output_path=None,
                url=None,
                error=f"Googleスプレッドシートの作成中に予期しないエラーが発生しました: {str(e)}"
            )

    def create_document(self, data: dict[str, Any], title: str) -> WriteResult:
        """
        Googleドキュメントを作成する。

        Args:
            data: ドキュメントのデータ
                {
                    "sections": [
                        {
                            "heading": str,
                            "level": int,
                            "paragraphs": [str, ...],
                            "tables": [
                                {
                                    "data": [[cell_value, ...], ...]
                                }
                            ]
                        }
                    ]
                }
            title: ドキュメントのタイトル

        Returns:
            作成されたドキュメントのURL

        Raises:
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        logger.info(f"Googleドキュメントの作成を開始: {title}")
        start_time = time.time()

        try:
            service = build('docs', 'v1', credentials=self.credentials)

            # ドキュメントを作成
            document_body = {
                'title': title
            }

            document = self._execute_with_retry(
                lambda: service.documents().create(body=document_body).execute()
            )
            document_id = document['documentId']
            document_url = f"https://docs.google.com/document/d/{document_id}/edit"

            logger.info(f"ドキュメントを作成: {document_id}")

            # コンテンツを書き込む
            sections = data.get('sections', [])
            
            if sections:
                requests = []
                current_index = 1  # ドキュメントの開始位置

                for section in sections:
                    # 見出しを追加
                    heading = section.get('heading', '')
                    level = section.get('level', 1)
                    
                    if heading:
                        requests.append({
                            'insertText': {
                                'location': {'index': current_index},
                                'text': heading + '\n'
                            }
                        })
                        
                        # 見出しスタイルを適用
                        heading_style = f'HEADING_{level}'
                        requests.append({
                            'updateParagraphStyle': {
                                'range': {
                                    'startIndex': current_index,
                                    'endIndex': current_index + len(heading) + 1
                                },
                                'paragraphStyle': {
                                    'namedStyleType': heading_style
                                },
                                'fields': 'namedStyleType'
                            }
                        })
                        
                        current_index += len(heading) + 1

                    # 段落を追加
                    paragraphs = section.get('paragraphs', [])
                    for paragraph in paragraphs:
                        if paragraph:
                            requests.append({
                                'insertText': {
                                    'location': {'index': current_index},
                                    'text': paragraph + '\n'
                                }
                            })
                            current_index += len(paragraph) + 1

                    # 表を追加
                    tables = section.get('tables', [])
                    for table in tables:
                        table_data = table.get('data', [])
                        if table_data:
                            rows = len(table_data)
                            columns = max(len(row) for row in table_data) if table_data else 0
                            
                            # 表を挿入
                            requests.append({
                                'insertTable': {
                                    'location': {'index': current_index},
                                    'rows': rows,
                                    'columns': columns
                                }
                            })
                            
                            # 表のセルにデータを入力するためのインデックスを計算
                            # （表の挿入後に別のリクエストで更新する必要がある）
                            current_index += 3  # 表の基本構造分

                # リクエストを実行
                if requests:
                    self._execute_with_retry(
                        lambda: service.documents().batchUpdate(
                            documentId=document_id,
                            body={'requests': requests}
                        ).execute()
                    )

            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleドキュメントの作成が完了: {document_id} "
                f"(URL: {document_url}, セクション数: {len(sections)}, 処理時間: {elapsed_time:.2f}秒)"
            )

            # WriteResultを返す
            return WriteResult(
                success=True,
                output_path=None,
                url=document_url,
                error=None
            )

        except HttpError as e:
            if e.resp.status == 403:
                logger.error("Googleドキュメントの作成権限がありません")
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error="Googleドキュメントの作成権限がありません"
                )
            else:
                logger.error(
                    "Googleドキュメントの作成中にエラーが発生",
                    exc_info=True
                )
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error=f"Googleドキュメントの作成中にエラーが発生しました: {str(e)}"
                )
        except Exception as e:
            logger.error(
                "Googleドキュメントの作成中に予期しないエラーが発生",
                exc_info=True
            )
            return WriteResult(
                success=False,
                output_path=None,
                url=None,
                error=f"Googleドキュメントの作成中に予期しないエラーが発生しました: {str(e)}"
            )

    def create_slides(self, data: dict[str, Any], title: str) -> WriteResult:
        """
        Googleスライドを作成する。

        Args:
            data: スライドのデータ
                {
                    "slides": [
                        {
                            "layout": "title" | "content" | "bullet",
                            "title": str,
                            "content": str | list
                        }
                    ]
                }
            title: プレゼンテーションのタイトル

        Returns:
            作成されたプレゼンテーションのURL

        Raises:
            PermissionError: アクセス権限がない場合
            APIError: API呼び出しに失敗した場合
        """
        logger.info(f"Googleスライドの作成を開始: {title}")
        start_time = time.time()

        try:
            service = build('slides', 'v1', credentials=self.credentials)

            # プレゼンテーションを作成
            presentation_body = {
                'title': title
            }

            presentation = self._execute_with_retry(
                lambda: service.presentations().create(
                    body=presentation_body
                ).execute()
            )

            presentation_id = presentation['presentationId']
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"

            logger.info(f"プレゼンテーションを作成: {presentation_id}")

            # スライドを追加
            slides_data = data.get('slides', [])
            
            if slides_data:
                requests = []
                
                # 最初のスライド（タイトルスライド）を更新
                if slides_data:
                    first_slide = slides_data[0]
                    slide_title = first_slide.get('title', '')
                    slide_content = first_slide.get('content', '')
                    
                    # デフォルトのスライドIDを取得
                    presentation_info = self._execute_with_retry(
                        lambda: service.presentations().get(
                            presentationId=presentation_id
                        ).execute()
                    )
                    
                    if presentation_info.get('slides'):
                        # タイトルとサブタイトルのプレースホルダーを探す
                        page_elements = presentation_info['slides'][0].get('pageElements', [])
                        title_placeholder_id = None
                        subtitle_placeholder_id = None
                        
                        for element in page_elements:
                            if 'shape' in element:
                                shape = element['shape']
                                placeholder_type = shape.get('placeholder', {}).get('type', '')
                                
                                if placeholder_type == 'CENTERED_TITLE' or placeholder_type == 'TITLE':
                                    title_placeholder_id = element['objectId']
                                elif placeholder_type == 'SUBTITLE':
                                    subtitle_placeholder_id = element['objectId']
                        
                        # タイトルを設定
                        if title_placeholder_id and slide_title:
                            requests.append({
                                'insertText': {
                                    'objectId': title_placeholder_id,
                                    'text': slide_title
                                }
                            })
                        
                        # サブタイトルを設定
                        if subtitle_placeholder_id and slide_content:
                            content_text = slide_content if isinstance(slide_content, str) else '\n'.join(slide_content)
                            requests.append({
                                'insertText': {
                                    'objectId': subtitle_placeholder_id,
                                    'text': content_text
                                }
                            })

                # 追加のスライドを作成
                for slide_data in slides_data[1:]:
                    layout = slide_data.get('layout', 'content')
                    slide_title = slide_data.get('title', '')
                    slide_content = slide_data.get('content', '')
                    
                    # レイアウトに応じたスライドを作成
                    if layout == 'title':
                        predefined_layout = 'TITLE'
                    elif layout == 'bullet':
                        predefined_layout = 'TITLE_AND_BODY'
                    else:
                        predefined_layout = 'TITLE_AND_BODY'
                    
                    # 新しいスライドを追加
                    slide_id = f'slide_{len(requests)}'
                    requests.append({
                        'createSlide': {
                            'objectId': slide_id,
                            'slideLayoutReference': {
                                'predefinedLayout': predefined_layout
                            }
                        }
                    })

                # リクエストを実行
                if requests:
                    self._execute_with_retry(
                        lambda: service.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={'requests': requests}
                        ).execute()
                    )

            # 処理時間を計算
            elapsed_time = time.time() - start_time
            logger.info(
                f"Googleスライドの作成が完了: {presentation_id} "
                f"(URL: {presentation_url}, スライド数: {len(slides_data)}, 処理時間: {elapsed_time:.2f}秒)"
            )

            # WriteResultを返す
            return WriteResult(
                success=True,
                output_path=None,
                url=presentation_url,
                error=None
            )

        except HttpError as e:
            if e.resp.status == 403:
                logger.error("Googleスライドの作成権限がありません")
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error="Googleスライドの作成権限がありません"
                )
            else:
                logger.error(
                    "Googleスライドの作成中にエラーが発生",
                    exc_info=True
                )
                return WriteResult(
                    success=False,
                    output_path=None,
                    url=None,
                    error=f"Googleスライドの作成中にエラーが発生しました: {str(e)}"
                )
        except Exception as e:
            logger.error(
                "Googleスライドの作成中に予期しないエラーが発生",
                exc_info=True
            )
            return WriteResult(
                success=False,
                output_path=None,
                url=None,
                error=f"Googleスライドの作成中に予期しないエラーが発生しました: {str(e)}"
            )

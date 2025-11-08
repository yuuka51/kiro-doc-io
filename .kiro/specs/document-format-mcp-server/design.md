# 設計書

## 概要

本MCPサーバは、Kiro AIアシスタントに対してMicrosoft Office形式（PowerPoint、Word、Excel）およびGoogle Workspace形式（スプレッドシート、ドキュメント、スライド）のファイルを読み取り・生成する機能を提供します。

Pythonで実装し、Model Context Protocol（MCP）仕様に準拠したサーバとして動作します。標準入出力（stdio）を介してKiroと通信し、ツール呼び出しを通じてドキュメント操作機能を公開します。

### 技術スタック

- **言語**: Python 3.10以上
- **MCPフレームワーク**: `mcp` パッケージ
- **Microsoft Office処理**: 
  - `python-pptx` (PowerPoint)
  - `python-docx` (Word)
  - `openpyxl` (Excel)
- **Google API処理**: 
  - `google-api-python-client`
  - `google-auth-oauthlib`
  - `google-auth-httplib2`

## アーキテクチャ

### システム構成図

```mermaid
graph TB
    Kiro[Kiro AI Assistant] -->|stdio/MCP| MCPServer[MCP Server]
    MCPServer --> ToolRegistry[Tool Registry]
    ToolRegistry --> ReadTools[Document Reader Tools]
    ToolRegistry --> WriteTools[Document Writer Tools]
    
    ReadTools --> MSOfficeReader[MS Office Reader]
    ReadTools --> GoogleReader[Google Workspace Reader]
    
    WriteTools --> MSOfficeWriter[MS Office Writer]
    WriteTools --> GoogleWriter[Google Workspace Writer]
    
    MSOfficeReader --> PPTXLib[python-pptx]
    MSOfficeReader --> DOCXLib[python-docx]
    MSOfficeReader --> XLSXLib[openpyxl]
    
    GoogleReader --> GoogleAPI[Google APIs]
    GoogleWriter --> GoogleAPI
    
    MSOfficeWriter --> PPTXLib
    MSOfficeWriter --> DOCXLib
    MSOfficeWriter --> XLSXLib
```

### レイヤー構造

1. **MCPサーバレイヤー**: MCP仕様に準拠した通信処理
2. **ツールレジストリレイヤー**: 利用可能なツールの登録と管理
3. **ドキュメント処理レイヤー**: 読み取り・書き込みロジックの実装
4. **ライブラリレイヤー**: 外部ライブラリとのインターフェース

## コンポーネントとインターフェース

### 1. MCPサーバコンポーネント (`server.py`)

MCPサーバのエントリーポイント。標準入出力を介してKiroと通信します。

```python
class DocumentMCPServer:
    """MCP Server for document format handling"""
    
    def __init__(self):
        self.server = Server("document-format-server")
        self._register_tools()
    
    def _register_tools(self):
        """Register all available tools"""
        pass
    
    async def run(self):
        """Start the MCP server"""
        pass
```

### 2. ドキュメントリーダーコンポーネント

#### PowerPointリーダー (`readers/powerpoint_reader.py`)

```python
class PowerPointReader:
    """Read PowerPoint (.pptx) files"""
    
    def read_file(self, file_path: str) -> dict:
        """
        Extract content from PowerPoint file
        
        Returns:
            {
                "slides": [
                    {
                        "slide_number": int,
                        "title": str,
                        "content": str,
                        "notes": str,
                        "tables": [...]
                    }
                ]
            }
        """
        pass
```

#### Wordリーダー (`readers/word_reader.py`)

```python
class WordReader:
    """Read Word (.docx) files"""
    
    def read_file(self, file_path: str) -> dict:
        """
        Extract content from Word file
        
        Returns:
            {
                "paragraphs": [
                    {
                        "text": str,
                        "style": str,  # "Heading 1", "Normal", etc.
                        "level": int
                    }
                ],
                "tables": [...]
            }
        """
        pass
```

#### Excelリーダー (`readers/excel_reader.py`)

```python
class ExcelReader:
    """Read Excel (.xlsx) files"""
    
    def read_file(self, file_path: str) -> dict:
        """
        Extract content from Excel file
        
        Returns:
            {
                "sheets": [
                    {
                        "name": str,
                        "data": [[cell_value, ...], ...],
                        "formulas": {...}
                    }
                ]
            }
        """
        pass
```

#### Google Workspaceリーダー (`readers/google_reader.py`)

```python
class GoogleWorkspaceReader:
    """Read Google Workspace files"""
    
    def __init__(self, credentials_path: str):
        self.credentials = self._load_credentials(credentials_path)
    
    def read_spreadsheet(self, file_id: str) -> dict:
        """Read Google Spreadsheet"""
        pass
    
    def read_document(self, file_id: str) -> dict:
        """Read Google Document"""
        pass
    
    def read_slides(self, file_id: str) -> dict:
        """Read Google Slides"""
        pass
```

### 3. ドキュメントライターコンポーネント

#### PowerPointライター (`writers/powerpoint_writer.py`)

```python
class PowerPointWriter:
    """Write PowerPoint (.pptx) files"""
    
    def create_presentation(self, data: dict, output_path: str) -> str:
        """
        Create PowerPoint file from structured data
        
        Args:
            data: {
                "title": str,
                "slides": [
                    {
                        "layout": "title" | "content" | "bullet",
                        "title": str,
                        "content": str | list
                    }
                ]
            }
        
        Returns:
            Path to created file
        """
        pass
```

#### Wordライター (`writers/word_writer.py`)

```python
class WordWriter:
    """Write Word (.docx) files"""
    
    def create_document(self, data: dict, output_path: str) -> str:
        """
        Create Word file from structured data
        
        Args:
            data: {
                "title": str,
                "sections": [
                    {
                        "heading": str,
                        "level": int,
                        "paragraphs": [str, ...],
                        "tables": [...]
                    }
                ]
            }
        
        Returns:
            Path to created file
        """
        pass
```

#### Excelライター (`writers/excel_writer.py`)

```python
class ExcelWriter:
    """Write Excel (.xlsx) files"""
    
    def create_workbook(self, data: dict, output_path: str) -> str:
        """
        Create Excel file from structured data
        
        Args:
            data: {
                "sheets": [
                    {
                        "name": str,
                        "data": [[cell_value, ...], ...],
                        "formatting": {...}
                    }
                ]
            }
        
        Returns:
            Path to created file
        """
        pass
```

#### Google Workspaceライター (`writers/google_writer.py`)

```python
class GoogleWorkspaceWriter:
    """Write Google Workspace files"""
    
    def __init__(self, credentials_path: str):
        self.credentials = self._load_credentials(credentials_path)
    
    def create_spreadsheet(self, data: dict, title: str) -> str:
        """Create Google Spreadsheet and return URL"""
        pass
    
    def create_document(self, data: dict, title: str) -> str:
        """Create Google Document and return URL"""
        pass
    
    def create_slides(self, data: dict, title: str) -> str:
        """Create Google Slides and return URL"""
        pass
```

### 4. ツール定義 (`tools/`)

MCPツールとして公開される機能：

- `read_powerpoint`: PowerPointファイルを読み取る
- `read_word`: Wordファイルを読み取る
- `read_excel`: Excelファイルを読み取る
- `read_google_spreadsheet`: Googleスプレッドシートを読み取る
- `read_google_document`: Googleドキュメントを読み取る
- `read_google_slides`: Googleスライドを読み取る
- `write_powerpoint`: PowerPointファイルを生成する
- `write_word`: Wordファイルを生成する
- `write_excel`: Excelファイルを生成する
- `write_google_spreadsheet`: Googleスプレッドシートを生成する
- `write_google_document`: Googleドキュメントを生成する
- `write_google_slides`: Googleスライドを生成する

## データモデル

### 共通データ構造

#### DocumentContent

```python
@dataclass
class DocumentContent:
    """Base class for document content"""
    format_type: str  # "pptx", "docx", "xlsx", "google_sheets", etc.
    metadata: dict
    content: dict
```

#### ReadResult

```python
@dataclass
class ReadResult:
    """Result of document read operation"""
    success: bool
    content: Optional[DocumentContent]
    error: Optional[str]
    file_path: str
```

#### WriteResult

```python
@dataclass
class WriteResult:
    """Result of document write operation"""
    success: bool
    output_path: Optional[str]
    url: Optional[str]  # For Google Workspace files
    error: Optional[str]
```

## エラーハンドリング

### エラータイプ

1. **FileNotFoundError**: ファイルが存在しない
2. **CorruptedFileError**: ファイルが破損している
3. **AuthenticationError**: Google API認証エラー
4. **PermissionError**: ファイルアクセス権限エラー
5. **APIError**: Google APIエラー
6. **ValidationError**: 入力データの検証エラー

### エラーレスポンス形式

```python
{
    "success": false,
    "error": {
        "type": "FileNotFoundError",
        "message": "指定されたファイルが見つかりません: /path/to/file.pptx",
        "details": {...}
    }
}
```

### エラーハンドリング戦略

- すべての例外をキャッチし、適切なエラーメッセージを返す
- ファイル操作前にファイルの存在と読み取り権限を確認
- Google API呼び出しはリトライロジックを実装（最大3回）
- タイムアウト設定: ファイル読み取り30秒、API呼び出し60秒

## テスト戦略

### ユニットテスト

各コンポーネントの個別機能をテスト：

- `tests/readers/test_powerpoint_reader.py`
- `tests/readers/test_word_reader.py`
- `tests/readers/test_excel_reader.py`
- `tests/readers/test_google_reader.py`
- `tests/writers/test_powerpoint_writer.py`
- `tests/writers/test_word_writer.py`
- `tests/writers/test_excel_writer.py`
- `tests/writers/test_google_writer.py`

### 統合テスト

- MCPサーバとツールの統合テスト
- 実際のファイルを使用したエンドツーエンドテスト
- Google APIのモックを使用したテスト

### テストデータ

- `tests/fixtures/`: サンプルファイル（.pptx、.docx、.xlsx）
- Google APIはモックまたはテスト用アカウントを使用

## セキュリティ考慮事項

### 認証情報の管理

- Google API認証情報は環境変数または設定ファイルから読み込む
- 認証情報ファイルのパスは設定可能
- 認証情報をログに出力しない

### ファイルアクセス制限

- 読み取り可能なファイルパスを制限（サンドボックス化）
- 書き込み先ディレクトリを制限
- パストラバーサル攻撃を防ぐ

### データ検証

- ファイルサイズ制限: 最大100MB
- シート数制限: 最大100シート（Excel）
- スライド数制限: 最大500スライド（PowerPoint）

## 設定管理

### 設定ファイル (`config.json`)

```json
{
  "google_credentials_path": "~/.config/kiro-mcp/google-credentials.json",
  "output_directory": "~/Documents/kiro-output",
  "max_file_size_mb": 100,
  "max_sheets": 100,
  "max_slides": 500,
  "api_timeout_seconds": 60,
  "enable_google_workspace": true
}
```

### 環境変数

- `GOOGLE_APPLICATION_CREDENTIALS`: Google API認証情報ファイルパス
- `MCP_OUTPUT_DIR`: 出力ファイルディレクトリ
- `MCP_LOG_LEVEL`: ログレベル（DEBUG、INFO、WARNING、ERROR）

## デプロイメント

### パッケージ構造

```
document-format-mcp-server/
├── src/
│   ├── __init__.py
│   ├── server.py
│   ├── readers/
│   │   ├── __init__.py
│   │   ├── powerpoint_reader.py
│   │   ├── word_reader.py
│   │   ├── excel_reader.py
│   │   └── google_reader.py
│   ├── writers/
│   │   ├── __init__.py
│   │   ├── powerpoint_writer.py
│   │   ├── word_writer.py
│   │   ├── excel_writer.py
│   │   └── google_writer.py
│   ├── tools/
│   │   ├── __init__.py
│   │   └── tool_definitions.py
│   └── utils/
│       ├── __init__.py
│       ├── config.py
│       └── errors.py
├── tests/
├── pyproject.toml
├── README.md
└── config.json.example
```

### インストール方法

```bash
# uvxを使用してインストール
uvx document-format-mcp-server

# または、pipを使用
pip install document-format-mcp-server
```

### Kiro設定 (`.kiro/settings/mcp.json`)

```json
{
  "mcpServers": {
    "document-format": {
      "command": "uvx",
      "args": ["document-format-mcp-server"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "~/.config/kiro-mcp/google-credentials.json",
        "MCP_OUTPUT_DIR": "~/Documents/kiro-output"
      },
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

## パフォーマンス考慮事項

### 最適化戦略

- 大きなファイルは段階的に読み込む（ストリーミング）
- キャッシュ機構: 同じファイルの再読み込みを避ける
- 並列処理: 複数シートの処理を並列化

### リソース制限

- メモリ使用量: 最大500MB
- 処理時間: 1ファイルあたり最大5分
- 同時処理数: 最大3ファイル

## 開発・テスト環境

### ローカル開発環境のセットアップ

#### 環境構築ツール

プロジェクトは以下のツールをサポート：

- **uv** (推奨): 高速なPythonパッケージマネージャー
- **pip**: 標準的なPythonパッケージマネージャー
- **venv**: Python標準の仮想環境

#### セットアップ手順

```bash
# uvを使用する場合（推奨）
uv venv
.venv\Scripts\activate  # Windows
uv pip install -r requirements.txt

# pipを使用する場合
python -m venv .venv
.venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

### テストツール

#### 1. サンプルファイル生成スクリプト (`create_sample_files.py`)

テスト用のサンプルファイルを自動生成するスクリプト。

**機能:**
- PowerPointファイル（3スライド、表を含む）の生成
- Wordファイル（見出し、段落、箇条書き、表を含む）の生成
- Excelファイル（3シート、データと数式を含む）の生成

**使用方法:**
```bash
python create_sample_files.py
```

**出力:**
- `test_files/sample.pptx`
- `test_files/sample.docx`
- `test_files/sample.xlsx`

#### 2. リーダー機能テストスクリプト (`test_readers.py`)

実装済みのリーダー機能を検証するスクリプト。

**機能:**
- ローカルファイル（PowerPoint、Word、Excel）の読み取りテスト
- Google Workspaceファイル（スプレッドシート、ドキュメント、スライド）の読み取りテスト
- 読み込んだ内容の表示と検証

**使用方法:**
```bash
# ローカルファイルのテスト
python test_readers.py

# Google Workspaceファイルのテスト
python test_readers.py --google
```

**テスト項目:**
- ファイルの正常読み込み
- データ構造の検証
- エラーハンドリングの確認
- 抽出されたコンテンツの表示

### テストデータ構造

#### PowerPointテストデータ

```python
{
    "slides": [
        {
            "slide_number": 1,
            "title": "サンプルプレゼンテーション",
            "content": "Document Format MCP Server テスト用",
            "notes": "",
            "tables": []
        },
        {
            "slide_number": 2,
            "title": "機能紹介",
            "content": "主な機能:\n  PowerPointファイルの読み取り\n  ...",
            "notes": "",
            "tables": []
        },
        {
            "slide_number": 3,
            "title": "データ表",
            "content": "",
            "notes": "",
            "tables": [
                {
                    "rows": 4,
                    "columns": 3,
                    "data": [
                        ["項目", "値", "備考"],
                        ["読み取り", "対応", "完了"],
                        ...
                    ]
                }
            ]
        }
    ]
}
```

#### Wordテストデータ

```python
{
    "paragraphs": [
        {
            "text": "サンプルドキュメント",
            "type": "heading",
            "level": 0,
            "style": "Title"
        },
        {
            "text": "これはDocument Format MCP Serverのテスト用...",
            "type": "paragraph",
            "style": "Normal"
        },
        ...
    ],
    "tables": [
        {
            "rows": 4,
            "columns": 3,
            "data": [...]
        }
    ]
}
```

#### Excelテストデータ

```python
{
    "sheets": [
        {
            "name": "データ",
            "data": [
                ["ID", "名前", "カテゴリ", "値"],
                [1, "PowerPoint", "読み取り", "完了"],
                ...
            ],
            "row_count": 7,
            "column_count": 4,
            "formulas": {}
        },
        {
            "name": "統計",
            "data": [...],
            "row_count": 4,
            "column_count": 2,
            "formulas": {}
        },
        {
            "name": "計算",
            "data": [...],
            "row_count": 5,
            "column_count": 3,
            "formulas": {
                "C4": "=B2+B3",
                "C5": "=(B2+B3)/2"
            }
        }
    ]
}
```

### Google Workspace認証設定

#### 認証情報の取得

1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクトを作成
2. 以下のAPIを有効化:
   - Google Sheets API
   - Google Docs API
   - Google Slides API
3. OAuth 2.0クライアントID（デスクトップアプリ）を作成
4. 認証情報JSONファイルをダウンロード

#### 認証情報の配置

```bash
# 推奨ディレクトリ構造
.config/
  └── google-credentials.json  # OAuth 2.0クライアント認証情報
  └── token.json              # 自動生成される認証トークン
```

#### 認証フロー

1. 初回実行時、ブラウザが開いてGoogle認証が求められる
2. Googleアカウントでログインし、アクセス許可を承認
3. 認証トークンが`token.json`に保存される
4. 以降の実行では保存されたトークンを使用

### ドキュメント

#### クイックスタートガイド (`QUICKSTART.md`)

5分で動作確認できる簡潔なガイド：
- 環境セットアップ（2分）
- サンプルファイル生成（1分）
- テスト実行（2分）
- Google Workspaceテスト（追加10分）

#### セットアップガイド (`SETUP.md`)

詳細なセットアップ手順：
- uvまたはpipでの環境構築
- テストファイルの準備方法
- Google認証情報の取得と設定
- トラブルシューティング
- 次のステップへの案内

### 開発ワークフロー

#### 1. 環境セットアップ

```bash
uv venv
.venv\Scripts\activate
uv pip install -r requirements.txt
```

#### 2. サンプルファイル生成

```bash
python create_sample_files.py
```

#### 3. リーダー機能のテスト

```bash
python test_readers.py
```

#### 4. 実装の検証

- 各リーダークラスが正しくファイルを読み込めることを確認
- データ構造が設計通りであることを確認
- エラーハンドリングが適切に動作することを確認

#### 5. 次の実装へ

- ライター機能の実装（タスク5-7）
- MCPツール統合（タスク8）
- パッケージング（タスク9-10）

### CI/CD統合（将来）

- GitHub Actionsでの自動テスト
- Pytestによるユニットテスト
- カバレッジレポート
- 自動デプロイメント

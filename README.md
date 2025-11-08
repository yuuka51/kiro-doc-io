# Document Format MCP Server

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Kiro AIアシスタント向けのMCPサーバーで、Microsoft Office形式（PowerPoint、Word、Excel）およびGoogle Workspace形式（スプレッドシート、ドキュメント、スライド）のファイルを読み取り・生成する機能を提供します。

## 特徴

- 🚀 **簡単セットアップ**: 5分で動作確認可能
- 📄 **多様なフォーマット対応**: PowerPoint、Word、Excel、Google Workspace
- 🔧 **柔軟な設定**: 環境変数や設定ファイルでカスタマイズ可能
- 🔒 **セキュア**: OAuth 2.0による安全なGoogle API認証
- 📝 **詳細なログ**: デバッグ情報の出力をサポート
- 🎯 **MCP準拠**: Model Context Protocol標準に完全準拠

## 機能

### 読み取り機能（実装済み）
- ✅ PowerPoint (.pptx) ファイルの読み取り
- ✅ Word (.docx) ファイルの読み取り
- ✅ Excel (.xlsx) ファイルの読み取り
- ✅ Google Workspace ファイルの読み取り（スプレッドシート、ドキュメント、スライド）

### 書き込み機能（実装済み）
- ✅ PowerPoint (.pptx) ファイルの生成
- ✅ Word (.docx) ファイルの生成
- ✅ Excel (.xlsx) ファイルの生成
- ✅ Google Workspace ファイルの生成（スプレッドシート、ドキュメント、スライド）

### MCPツール統合（実装済み）
- ✅ 12個のMCPツールを定義・実装
- ✅ MCPサーバーへの統合完了
- ✅ ログ出力とエラーハンドリング

## Quick Start

### 5-Minute Test

```bash
# 1. Setup environment (2 min)
uv venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux
uv pip install -r requirements.txt

# 2. Generate sample files (1 min)
python create_sample_files.py

# 3. Run tests (2 min)
python test_readers.py
```

詳細は [QUICKSTART.md](QUICKSTART.md) を参照してください。

## インストール方法

### 前提条件

- Python 3.10以上
- pip または uv パッケージマネージャー

### uvxを使用する方法（推奨）

uvxは高速なPythonパッケージマネージャーです。

```bash
# uvのインストール（初回のみ）
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# MCPサーバーの実行
uvx document-format-mcp-server
```

### pipを使用する方法

```bash
# パッケージのインストール
pip install document-format-mcp-server

# MCPサーバーの実行
python -m document_format_mcp_server.server
```

### ローカル開発環境でのインストール

```bash
# リポジトリのクローン
git clone https://github.com/your-repo/document-format-mcp-server.git
cd document-format-mcp-server

# 仮想環境の作成と有効化
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# 依存関係のインストール
pip install -r requirements.txt

# MCPサーバーの実行
python -m src.document_format_mcp_server.server
```

### 依存関係

このMCPサーバーは以下のライブラリを使用します：

- `mcp` - Model Context Protocol実装
- `python-pptx` - PowerPointファイル処理
- `python-docx` - Wordファイル処理
- `openpyxl` - Excelファイル処理
- `google-api-python-client` - Google API クライアント
- `google-auth-oauthlib` - Google OAuth認証
- `google-auth-httplib2` - Google API HTTP通信

詳細は `requirements.txt` を参照してください。

## 設定方法

### Kiro MCP設定

`.kiro/settings/mcp.json` に以下の設定を追加してください：

```json
{
  "mcpServers": {
    "document-format": {
      "command": "uvx",
      "args": ["document-format-mcp-server"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "~/.config/kiro-mcp/google-credentials.json",
        "MCP_OUTPUT_DIR": "~/Documents/kiro-output",
        "MCP_LOG_LEVEL": "INFO"
      },
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

#### 設定項目の説明

- **command**: MCPサーバーの起動コマンド
  - `uvx`: uvパッケージマネージャーを使用（推奨）
  - または `python -m document_format_mcp_server.server` でローカル実行
  
- **args**: コマンドライン引数
  - `["document-format-mcp-server"]`: パッケージ名
  
- **env**: 環境変数
  - `GOOGLE_APPLICATION_CREDENTIALS`: Google API認証情報ファイルのパス
  - `MCP_OUTPUT_DIR`: 生成ファイルの出力先ディレクトリ
  - `MCP_LOG_LEVEL`: ログレベル（DEBUG、INFO、WARNING、ERROR）
  
- **disabled**: サーバーの有効/無効
  - `false`: 有効（デフォルト）
  - `true`: 無効
  
- **autoApprove**: 自動承認するツール名のリスト
  - 空配列の場合、すべてのツール呼び出しで確認が必要

### 設定ファイル（オプション）

プロジェクトルートに `config.json` ファイルを作成できます（オプション）：

```json
{
  "google_credentials_path": "~/.config/kiro-mcp/google-credentials.json",
  "output_directory": "~/Documents/kiro-output",
  "max_file_size_mb": 100,
  "max_sheets": 100,
  "max_slides": 500,
  "api_timeout_seconds": 60,
  "enable_google_workspace": true,
  "log_level": "INFO",
  "file_read_timeout_seconds": 30
}
```

詳細は `config.json.example` を参照してください。

### 環境変数

以下の環境変数で設定を上書きできます：

- `GOOGLE_APPLICATION_CREDENTIALS`: Google API認証情報ファイルのパス
- `MCP_OUTPUT_DIR`: 生成ファイルの出力先ディレクトリ
- `MCP_LOG_LEVEL`: ログレベル（DEBUG、INFO、WARNING、ERROR）
- `MCP_MAX_FILE_SIZE_MB`: 最大ファイルサイズ（MB）
- `MCP_API_TIMEOUT`: APIタイムアウト時間（秒）
- `MCP_ENABLE_GOOGLE`: Google Workspace機能の有効/無効（true/false）

## Google API認証情報の取得方法

Google Workspace機能（スプレッドシート、ドキュメント、スライド）を使用するには、Google Cloud Consoleで認証情報を取得する必要があります。

### ステップ1: Google Cloud プロジェクトの作成

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成、または既存のプロジェクトを選択
3. プロジェクト名を入力（例: "Kiro Document MCP"）

### ステップ2: APIの有効化

1. 左側のメニューから「APIとサービス」→「ライブラリ」を選択
2. 以下のAPIを検索して有効化：
   - **Google Sheets API** - スプレッドシートの読み取り・書き込み
   - **Google Docs API** - ドキュメントの読み取り・書き込み
   - **Google Slides API** - スライドの読み取り・書き込み
   - **Google Drive API** - ファイルの作成・管理（推奨）

### ステップ3: OAuth 2.0 認証情報の作成

1. 「APIとサービス」→「認証情報」を選択
2. 「認証情報を作成」→「OAuth クライアント ID」をクリック
3. 同意画面の設定（初回のみ）：
   - ユーザータイプ: 「外部」を選択（個人利用の場合）
   - アプリ名: "Kiro Document MCP Server"
   - ユーザーサポートメール: 自分のメールアドレス
   - デベロッパーの連絡先情報: 自分のメールアドレス
   - 「保存して次へ」をクリック
4. スコープの追加（オプション）：
   - 「スコープを追加または削除」をクリック
   - 以下のスコープを追加（推奨）：
     - `https://www.googleapis.com/auth/spreadsheets`
     - `https://www.googleapis.com/auth/documents`
     - `https://www.googleapis.com/auth/presentations`
     - `https://www.googleapis.com/auth/drive.file`
5. テストユーザーの追加：
   - 自分のGoogleアカウントのメールアドレスを追加
6. OAuth クライアント IDの作成：
   - アプリケーションの種類: 「デスクトップアプリ」を選択
   - 名前: "Kiro MCP Client"
   - 「作成」をクリック

### ステップ4: 認証情報ファイルのダウンロード

1. 作成したOAuth 2.0 クライアントIDの右側にあるダウンロードアイコンをクリック
2. JSONファイルがダウンロードされます（例: `client_secret_xxxxx.json`）
3. ファイル名を `google-credentials.json` に変更
4. 以下のディレクトリに配置：
   ```
   ~/.config/kiro-mcp/google-credentials.json
   ```
   または
   ```
   .config/google-credentials.json  # プロジェクトルート
   ```

### ステップ5: 初回認証

1. MCPサーバーを起動すると、ブラウザが自動的に開きます
2. Googleアカウントでログイン
3. アクセス許可を確認して「許可」をクリック
4. 認証トークンが自動的に保存されます（`token.json`）
5. 以降の実行では、保存されたトークンが使用されます

### トラブルシューティング

#### 「このアプリは確認されていません」と表示される場合

1. 「詳細」をクリック
2. 「（アプリ名）に移動（安全ではないページ）」をクリック
3. これは自分で作成したアプリなので安全です

#### 認証トークンの再生成

認証エラーが発生した場合、以下のファイルを削除して再認証してください：
```bash
rm token.json
```

#### 認証情報ファイルのパス設定

環境変数で認証情報ファイルのパスを指定できます：
```bash
export GOOGLE_APPLICATION_CREDENTIALS="~/.config/kiro-mcp/google-credentials.json"
```

または、`.kiro/settings/mcp.json` の `env` セクションで設定：
```json
{
  "env": {
    "GOOGLE_APPLICATION_CREDENTIALS": "~/.config/kiro-mcp/google-credentials.json"
  }
}
```

## 利用可能なツール

このMCPサーバーは、Kiro AIアシスタントに対して以下の12個のツールを提供します。

### 読み取りツール

#### `read_powerpoint`
PowerPoint (.pptx) ファイルを読み取り、スライドの内容を抽出します。

**パラメータ:**
- `file_path` (string, 必須): 読み取るPowerPointファイルのパス

**戻り値:**
- スライドごとのタイトル、本文、ノート、表データ

#### `read_word`
Word (.docx) ファイルを読み取り、ドキュメントの内容を抽出します。

**パラメータ:**
- `file_path` (string, 必須): 読み取るWordファイルのパス

**戻り値:**
- 段落、見出し、表、箇条書きリスト

#### `read_excel`
Excel (.xlsx) ファイルを読み取り、シートのデータを抽出します。

**パラメータ:**
- `file_path` (string, 必須): 読み取るExcelファイルのパス

**戻り値:**
- シートごとの名前、セルデータ、数式

#### `read_google_spreadsheet`
Googleスプレッドシートを読み取ります。

**パラメータ:**
- `file_id` (string, 必須): スプレッドシートのIDまたはURL

**戻り値:**
- シートごとのデータ

#### `read_google_document`
Googleドキュメントを読み取ります。

**パラメータ:**
- `file_id` (string, 必須): ドキュメントのIDまたはURL

**戻り値:**
- ドキュメントの内容

#### `read_google_slides`
Googleスライドを読み取ります。

**パラメータ:**
- `file_id` (string, 必須): スライドのIDまたはURL

**戻り値:**
- スライドごとの内容

### 書き込みツール

#### `write_powerpoint`
PowerPoint (.pptx) ファイルを生成します。

**パラメータ:**
- `data` (object, 必須): プレゼンテーションデータ
- `output_path` (string, 必須): 出力ファイルパス

**戻り値:**
- 生成されたファイルのパス

#### `write_word`
Word (.docx) ファイルを生成します。

**パラメータ:**
- `data` (object, 必須): ドキュメントデータ
- `output_path` (string, 必須): 出力ファイルパス

**戻り値:**
- 生成されたファイルのパス

#### `write_excel`
Excel (.xlsx) ファイルを生成します。

**パラメータ:**
- `data` (object, 必須): ワークブックデータ
- `output_path` (string, 必須): 出力ファイルパス

**戻り値:**
- 生成されたファイルのパス

#### `write_google_spreadsheet`
Googleスプレッドシートを生成します。

**パラメータ:**
- `data` (object, 必須): スプレッドシートデータ
- `title` (string, 必須): スプレッドシートのタイトル

**戻り値:**
- 生成されたスプレッドシートのURL

#### `write_google_document`
Googleドキュメントを生成します。

**パラメータ:**
- `data` (object, 必須): ドキュメントデータ
- `title` (string, 必須): ドキュメントのタイトル

**戻り値:**
- 生成されたドキュメントのURL

#### `write_google_slides`
Googleスライドを生成します。

**パラメータ:**
- `data` (object, 必須): プレゼンテーションデータ
- `title` (string, 必須): プレゼンテーションのタイトル

**戻り値:**
- 生成されたスライドのURL

## Development

### Local Development Setup

詳細なセットアップ手順は [SETUP.md](SETUP.md) を参照してください。

#### Using uv (Recommended)

```bash
# Create virtual environment
uv venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies
uv pip install -r requirements.txt
```

#### Using pip

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

### Testing Reader Functions

#### Generate Sample Files

```bash
python create_sample_files.py
```

This creates:
- `test_files/sample.pptx` - PowerPoint file with 3 slides
- `test_files/sample.docx` - Word file with headings, paragraphs, and tables
- `test_files/sample.xlsx` - Excel file with 3 sheets

#### Run Tests

```bash
# Test local files (PowerPoint, Word, Excel)
python test_readers.py

# Test Google Workspace files
python test_readers.py --google
```

### プロジェクト構造

```
document-format-mcp-server/
├── src/
│   └── document_format_mcp_server/
│       ├── server.py              # MCPサーバーエントリーポイント
│       ├── readers/               # ドキュメント読み取り機能
│       │   ├── __init__.py
│       │   ├── powerpoint_reader.py  # PowerPoint読み取り
│       │   ├── word_reader.py        # Word読み取り
│       │   ├── excel_reader.py       # Excel読み取り
│       │   └── google_reader.py      # Google Workspace読み取り
│       ├── writers/               # ドキュメント書き込み機能
│       │   ├── __init__.py
│       │   ├── powerpoint_writer.py  # PowerPoint生成
│       │   ├── word_writer.py        # Word生成
│       │   ├── excel_writer.py       # Excel生成
│       │   └── google_writer.py      # Google Workspace生成
│       ├── tools/                 # MCPツール定義
│       │   ├── __init__.py
│       │   ├── tool_definitions.py   # ツールスキーマ定義
│       │   └── tool_handlers.py      # ツールハンドラー実装
│       └── utils/                 # ユーティリティ
│           ├── __init__.py
│           ├── config.py             # 設定管理
│           ├── errors.py             # エラー定義
│           └── logging_config.py     # ログ設定
├── tests/                         # ユニットテスト
│   ├── readers/                   # リーダーテスト
│   └── writers/                   # ライターテスト
├── test_files/                    # テスト用サンプルファイル
├── .config/                       # 設定ファイル（ローカル）
│   └── google-credentials.json    # Google API認証情報
├── test_readers.py                # リーダー機能テストスクリプト
├── test_writers.py                # ライター機能テストスクリプト
├── create_sample_files.py         # サンプルファイル生成スクリプト
├── config.json.example            # 設定ファイルのサンプル
├── QUICKSTART.md                  # 5分クイックスタートガイド
├── SETUP.md                       # 詳細セットアップガイド
├── README.md                      # このファイル
├── requirements.txt               # Python依存関係
└── pyproject.toml                 # プロジェクト設定
```

### Running Unit Tests

```bash
pytest
```

### Code Formatting

```bash
black src/ tests/
```

## 使用例

### Kiroでの使用例

MCPサーバーを設定した後、Kiroで以下のように使用できます：

```
# PowerPointファイルを読み取る
「test_files/sample.pptx を読み取って、内容を要約してください」

# Wordファイルを読み取る
「設計書.docx を読み取って、要件を抽出してください」

# Excelファイルを読み取る
「データ.xlsx を読み取って、統計情報を教えてください」

# PowerPointファイルを生成する
「プロジェクト概要のプレゼンテーションを作成して、output.pptx として保存してください」

# Googleスプレッドシートを読み取る
「このGoogleスプレッドシート（URL）のデータを分析してください」
```

### データ形式の例

#### PowerPoint生成データ

```json
{
  "title": "プロジェクト概要",
  "slides": [
    {
      "layout": "title",
      "title": "プロジェクト概要",
      "content": "2024年度 新規プロジェクト"
    },
    {
      "layout": "content",
      "title": "目的",
      "content": "システムの効率化と自動化"
    },
    {
      "layout": "bullet",
      "title": "主な機能",
      "content": [
        "ドキュメント読み取り",
        "ドキュメント生成",
        "API統合"
      ]
    }
  ]
}
```

#### Word生成データ

```json
{
  "title": "設計書",
  "sections": [
    {
      "heading": "概要",
      "level": 1,
      "paragraphs": [
        "本システムは...",
        "主な機能は..."
      ]
    },
    {
      "heading": "要件",
      "level": 1,
      "paragraphs": ["要件1", "要件2"]
    }
  ]
}
```

#### Excel生成データ

```json
{
  "sheets": [
    {
      "name": "データ",
      "data": [
        ["ID", "名前", "値"],
        [1, "項目A", 100],
        [2, "項目B", 200]
      ]
    }
  ]
}
```

## ドキュメント

### ユーザー向けドキュメント
- [QUICKSTART.md](QUICKSTART.md) - 5分で動作確認できるクイックスタートガイド
- [SETUP.md](SETUP.md) - 詳細なセットアップとトラブルシューティング
- [GOOGLE_API_SETUP.md](GOOGLE_API_SETUP.md) - Google API認証情報の取得方法（詳細ガイド）
- [config.json.example](config.json.example) - 設定ファイルのサンプル
- [.kiro/settings/mcp.json.example](.kiro/settings/mcp.json.example) - Kiro MCP設定のサンプル

### 開発者向けドキュメント
- [Design Document](.kiro/specs/document-format-mcp-server/design.md) - アーキテクチャと設計
- [Requirements](.kiro/specs/document-format-mcp-server/requirements.md) - 要件定義
- [Tasks](.kiro/specs/document-format-mcp-server/tasks.md) - 実装タスクリスト

## よくある質問（FAQ）

### Q: Google Workspace機能を使わない場合は？

A: `config.json` で `enable_google_workspace: false` に設定するか、環境変数 `MCP_ENABLE_GOOGLE=false` を設定してください。

### Q: 出力ファイルの保存先を変更するには？

A: 環境変数 `MCP_OUTPUT_DIR` または `config.json` の `output_directory` を変更してください。

### Q: ログレベルを変更するには？

A: 環境変数 `MCP_LOG_LEVEL` を設定してください（DEBUG、INFO、WARNING、ERROR）。

### Q: 大きなファイルを処理できますか？

A: デフォルトでは100MBまでのファイルを処理できます。`config.json` の `max_file_size_mb` で変更可能です。

### Q: エラーが発生した場合は？

A: ログを確認してください。`MCP_LOG_LEVEL=DEBUG` に設定すると詳細なログが出力されます。

## トラブルシューティング

### 「モジュールが見つかりません」エラー

```bash
# 依存関係を再インストール
pip install -r requirements.txt
```

### Google API認証エラー

```bash
# 認証トークンを削除して再認証
rm token.json
```

### ファイル読み取りエラー

- ファイルパスが正しいか確認
- ファイルが破損していないか確認
- ファイルサイズが制限内か確認

## 貢献

プルリクエストを歓迎します。大きな変更の場合は、まずissueを開いて変更内容を議論してください。

## ライセンス

MIT License

Copyright (c) 2024 Document Format MCP Server

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

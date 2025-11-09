# セットアップとテスト手順

このドキュメントでは、ローカル環境でDocument Format MCP Serverをセットアップし、テストする方法を説明します。

## 前提条件

- Python 3.10以上
- uv（推奨）またはpip

## セットアップ方法

### 1. uvを使用する場合（推奨）

```bash
# uvのインストール（まだの場合）
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# プロジェクトディレクトリに移動
cd document-format-mcp-server

# 仮想環境を作成して依存関係をインストール
uv venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# 依存関係をインストール
uv pip install -r requirements.txt
```

### 2. pipを使用する場合

```bash
# プロジェクトディレクトリに移動
cd document-format-mcp-server

# 仮想環境を作成
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# 依存関係をインストール
pip install -r requirements.txt
```

## テストファイルの準備

### ローカルファイルのテスト

1. サンプルファイルを生成:
```bash
python scripts/setup/create_sample_files.py
```

2. 以下のファイルが`test_files/samples/`に生成されます:
   - `sample.pptx` - PowerPointファイル
   - `sample.docx` - Wordファイル
   - `sample.xlsx` - Excelファイル

### Google Workspaceファイルのテスト

1. Google Cloud Consoleで認証情報を取得:
   - [Google Cloud Console](https://console.cloud.google.com/)にアクセス
   - プロジェクトを作成
   - Google Sheets API、Google Docs API、Google Slides APIを有効化
   - OAuth 2.0クライアントIDを作成（デスクトップアプリケーション）
   - 認証情報JSONファイルをダウンロード

2. 認証情報ファイルを配置:
```bash
mkdir -p .config
# ダウンロードしたファイルを .config/google-credentials.json に配置
```

## テストの実行

### ローカルファイルのテスト

1. `tests/test_readers.py`を編集してファイルパスを設定:
```python
pptx_file = "test_files/samples/sample.pptx"
docx_file = "test_files/samples/sample.docx"
xlsx_file = "test_files/samples/sample.xlsx"
```

2. テストを実行:
```bash
python tests/test_readers.py
```

### Google Workspaceファイルのテスト

1. `tests/test_readers.py`を編集して認証情報とファイルURLを設定:
```python
credentials_path = ".config/google-credentials.json"
spreadsheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
document_url = "https://docs.google.com/document/d/YOUR_DOCUMENT_ID/edit"
slides_url = "https://docs.google.com/presentation/d/YOUR_SLIDES_ID/edit"
```

2. テストを実行:
```bash
python tests/test_readers.py --google
```

3. 初回実行時、ブラウザが開いてGoogle認証が求められます
   - Googleアカウントでログイン
   - アクセス許可を承認
   - 認証トークンが`.config/token.json`に保存されます

## トラブルシューティング

### ImportError: No module named 'xxx'

依存関係が正しくインストールされていません:
```bash
uv pip install -r requirements.txt
# または
pip install -r requirements.txt
```

### Google認証エラー

1. 認証情報ファイルのパスが正しいか確認
2. Google Cloud ConsoleでAPIが有効化されているか確認
3. OAuth 2.0クライアントIDの種類が「デスクトップアプリケーション」になっているか確認
4. `.config/token.json`を削除して再認証を試す

### ファイルが見つからないエラー

1. ファイルパスが正しいか確認
2. ファイルが実際に存在するか確認
3. 相対パスではなく絶対パスを試す

## 次のステップ

ローカルテストが成功したら、次のタスクに進めます:

1. **タスク5-7**: ファイル書き込み機能の実装
2. **タスク8**: MCPツール定義とサーバへの統合
3. **タスク9**: 設定ファイルとドキュメントの整備
4. **タスク10**: パッケージングとデプロイメント

MCPサーバとして使用するには、タスク8の完了が必須です。

## 参考情報

- [uv公式ドキュメント](https://docs.astral.sh/uv/)
- [Google Workspace API](https://developers.google.com/workspace)
- [MCP (Model Context Protocol)](https://modelcontextprotocol.io/)

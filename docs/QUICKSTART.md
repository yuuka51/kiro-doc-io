# クイックスタートガイド

このガイドでは、最速でDocument Format MCP Serverのリーダー機能をテストする方法を説明します。

## 5分でテストする

### ステップ1: 環境セットアップ（2分）

```bash
# uvを使用（推奨）
uv venv
.venv\Scripts\activate
uv pip install -r requirements.txt

# または pip を使用
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### ステップ2: サンプルファイル生成（1分）

```bash
python scripts/setup/create_sample_files.py
```

これで`test_files/samples/`ディレクトリに以下のファイルが生成されます:
- `sample.pptx` - PowerPointファイル
- `sample.docx` - Wordファイル
- `sample.xlsx` - Excelファイル

### ステップ3: テスト実行（2分）

```bash
python tests/test_readers.py
```

成功すると、各ファイルの内容が表示されます:

```
============================================================
PowerPointファイルを読み込み中: test_files/samples/sample.pptx
============================================================

✅ 読み込み成功!
スライド数: 3

--- スライド 1 ---
タイトル: サンプルプレゼンテーション
コンテンツ: Document Format MCP Server テスト用
...
```

## Google Workspaceファイルのテスト（追加10分）

### ステップ1: Google認証情報の取得

1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新しいプロジェクトを作成
3. 以下のAPIを有効化:
   - Google Sheets API
   - Google Docs API
   - Google Slides API
4. 「認証情報」→「認証情報を作成」→「OAuth クライアント ID」
5. アプリケーションの種類: **デスクトップアプリ**
6. JSONファイルをダウンロード

### ステップ2: 認証情報を配置

```bash
mkdir .config
# ダウンロードしたファイルを .config/google-credentials.json にコピー
```

### ステップ3: テストファイルのURLを設定

`tests/test_readers.py`を編集:

```python
# 85行目あたり
credentials_path = ".config/google-credentials.json"

# テストするGoogleファイルのURL（自分のファイルに変更）
spreadsheet_url = "https://docs.google.com/spreadsheets/d/YOUR_ID/edit"
document_url = "https://docs.google.com/document/d/YOUR_ID/edit"
slides_url = "https://docs.google.com/presentation/d/YOUR_ID/edit"

# 95-97行目のコメントを解除
test_google_spreadsheet(spreadsheet_url, credentials_path)
test_google_document(document_url, credentials_path)
test_google_slides(slides_url, credentials_path)
```

### ステップ4: テスト実行

```bash
python tests/test_readers.py --google
```

初回実行時、ブラウザが開いてGoogle認証が求められます。

## トラブルシューティング

### モジュールが見つからない

```bash
# 依存関係を再インストール
uv pip install -r requirements.txt
```

### ファイルが見つからない

```bash
# サンプルファイルを再生成
python scripts/setup/create_sample_files.py
```

### Google認証エラー

1. `.config/token.json`を削除して再認証
2. Google Cloud ConsoleでAPIが有効化されているか確認
3. OAuth 2.0クライアントIDが「デスクトップアプリ」になっているか確認

## 次のステップ

テストが成功したら:

1. **タスク8を実装**: MCPツール定義とサーバ統合
   - これにより、KiroやClaude Desktopから使用可能になります

2. **タスク5-7を実装**: ファイル書き込み機能
   - PowerPoint、Word、Excel、Google Workspaceファイルの生成

3. **本番環境へのデプロイ**: タスク9-10
   - 設定ファイルの整備
   - パッケージング

## 参考

詳細なセットアップ手順は`docs/SETUP.md`を参照してください。

# Google API認証情報の取得方法

このガイドでは、Document Format MCP ServerでGoogle Workspace機能（スプレッドシート、ドキュメント、スライド）を使用するために必要なGoogle API認証情報の取得方法を詳しく説明します。

## 概要

Google Workspace機能を使用するには、以下の手順が必要です：

1. Google Cloud プロジェクトの作成
2. 必要なAPIの有効化
3. OAuth 2.0 認証情報の作成
4. 認証情報ファイルのダウンロードと配置
5. 初回認証の実行

所要時間: 約10〜15分

## 前提条件

- Googleアカウント（無料アカウントで可）
- インターネット接続

## 詳細手順

### ステップ1: Google Cloud Consoleへのアクセス

1. ブラウザで [Google Cloud Console](https://console.cloud.google.com/) を開く
2. Googleアカウントでログイン
3. 初めての場合、利用規約に同意

### ステップ2: 新しいプロジェクトの作成

1. 画面上部のプロジェクト選択ドロップダウンをクリック
2. 「新しいプロジェクト」をクリック
3. プロジェクト情報を入力：
   - **プロジェクト名**: `Kiro Document MCP` （任意の名前）
   - **組織**: なし（個人利用の場合）
   - **場所**: なし
4. 「作成」をクリック
5. プロジェクトが作成されるまで数秒待つ
6. 作成されたプロジェクトを選択

### ステップ3: APIの有効化

#### 3.1 Google Sheets APIの有効化

1. 左側のメニューから「APIとサービス」→「ライブラリ」を選択
2. 検索ボックスに「Google Sheets API」と入力
3. 「Google Sheets API」をクリック
4. 「有効にする」をクリック
5. 有効化が完了するまで待つ

#### 3.2 Google Docs APIの有効化

1. 「ライブラリ」に戻る（左上の「←」ボタン）
2. 検索ボックスに「Google Docs API」と入力
3. 「Google Docs API」をクリック
4. 「有効にする」をクリック

#### 3.3 Google Slides APIの有効化

1. 「ライブラリ」に戻る
2. 検索ボックスに「Google Slides API」と入力
3. 「Google Slides API」をクリック
4. 「有効にする」をクリック

#### 3.4 Google Drive API の有効化（推奨）

1. 「ライブラリ」に戻る
2. 検索ボックスに「Google Drive API」と入力
3. 「Google Drive API」をクリック
4. 「有効にする」をクリック

### ステップ4: OAuth同意画面の設定

1. 左側のメニューから「APIとサービス」→「OAuth同意画面」を選択
2. ユーザータイプを選択：
   - **外部**: 個人利用の場合（推奨）
   - **内部**: Google Workspaceアカウントの場合のみ
3. 「作成」をクリック
4. アプリ情報を入力：
   - **アプリ名**: `Kiro Document MCP Server`
   - **ユーザーサポートメール**: 自分のメールアドレスを選択
   - **アプリのロゴ**: （オプション）スキップ可
   - **アプリドメイン**: （オプション）スキップ可
   - **承認済みドメイン**: （オプション）スキップ可
   - **デベロッパーの連絡先情報**: 自分のメールアドレスを入力
5. 「保存して次へ」をクリック

### ステップ5: スコープの追加（オプション）

1. 「スコープを追加または削除」をクリック
2. 以下のスコープを検索して追加（推奨）：
   - `https://www.googleapis.com/auth/spreadsheets` - スプレッドシートの読み取り・書き込み
   - `https://www.googleapis.com/auth/documents` - ドキュメントの読み取り・書き込み
   - `https://www.googleapis.com/auth/presentations` - スライドの読み取り・書き込み
   - `https://www.googleapis.com/auth/drive.file` - Driveファイルの作成・管理
3. 「更新」をクリック
4. 「保存して次へ」をクリック

### ステップ6: テストユーザーの追加

1. 「テストユーザー」セクションで「ADD USERS」をクリック
2. 自分のGoogleアカウントのメールアドレスを入力
3. 「追加」をクリック
4. 「保存して次へ」をクリック
5. 概要を確認して「ダッシュボードに戻る」をクリック

### ステップ7: OAuth 2.0 クライアントIDの作成

1. 左側のメニューから「APIとサービス」→「認証情報」を選択
2. 画面上部の「認証情報を作成」をクリック
3. 「OAuth クライアント ID」を選択
4. アプリケーションの種類を選択：
   - **デスクトップアプリ**を選択
5. 名前を入力：
   - **名前**: `Kiro MCP Client`
6. 「作成」をクリック
7. 「OAuth クライアントを作成しました」ダイアログが表示される
8. 「JSONをダウンロード」をクリック
9. ダイアログを閉じる

### ステップ8: 認証情報ファイルの配置

#### 8.1 ファイル名の変更

1. ダウンロードしたJSONファイル（例: `client_secret_xxxxx.apps.googleusercontent.com.json`）を見つける
2. ファイル名を `google-credentials.json` に変更

#### 8.2 ファイルの配置

**オプション1: ユーザー設定ディレクトリ（推奨）**

```bash
# Windowsの場合
mkdir %USERPROFILE%\.config\kiro-mcp
move Downloads\google-credentials.json %USERPROFILE%\.config\kiro-mcp\

# macOS/Linuxの場合
mkdir -p ~/.config/kiro-mcp
mv ~/Downloads/google-credentials.json ~/.config/kiro-mcp/
```

**オプション2: プロジェクトディレクトリ**

```bash
# プロジェクトルートに配置
mkdir .config
move Downloads\google-credentials.json .config\
```

### ステップ9: 環境変数の設定

#### 9.1 Kiro MCP設定ファイル

`.kiro/settings/mcp.json` に以下を追加：

```json
{
  "mcpServers": {
    "document-format": {
      "command": "uvx",
      "args": ["document-format-mcp-server"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "~/.config/kiro-mcp/google-credentials.json"
      },
      "disabled": false
    }
  }
}
```

#### 9.2 システム環境変数（オプション）

**Windowsの場合:**
```powershell
setx GOOGLE_APPLICATION_CREDENTIALS "%USERPROFILE%\.config\kiro-mcp\google-credentials.json"
```

**macOS/Linuxの場合:**
```bash
export GOOGLE_APPLICATION_CREDENTIALS="~/.config/kiro-mcp/google-credentials.json"
# .bashrc または .zshrc に追加して永続化
echo 'export GOOGLE_APPLICATION_CREDENTIALS="~/.config/kiro-mcp/google-credentials.json"' >> ~/.bashrc
```

### ステップ10: 初回認証の実行

1. MCPサーバーを起動するか、Google Workspace機能を初めて使用する
2. ブラウザが自動的に開く
3. Googleアカウントでログイン（既にログイン済みの場合はスキップ）
4. 「このアプリは確認されていません」と表示される場合：
   - 「詳細」をクリック
   - 「（アプリ名）に移動（安全ではないページ）」をクリック
   - これは自分で作成したアプリなので安全です
5. アクセス許可の確認画面が表示される：
   - 「Googleスプレッドシートのすべてのスプレッドシートの参照、編集、作成、削除」
   - 「Googleドキュメントのすべてのドキュメントの参照、編集、作成、削除」
   - 「Googleスライドのすべてのプレゼンテーションの参照、編集、作成、削除」
   - 「Googleドライブで作成したファイルの参照、編集、作成、削除」
6. 「許可」をクリック
7. 「認証が完了しました」と表示される
8. ブラウザを閉じる
9. 認証トークンが自動的に `token.json` として保存される

### ステップ11: 動作確認

#### 11.1 テストスクリプトで確認

```bash
# Google Workspace機能のテスト
python test_readers.py --google
```

#### 11.2 Kiroで確認

```
「Googleスプレッドシート（URL）を読み取ってください」
```

## トラブルシューティング

### 「このアプリは確認されていません」エラー

**原因**: 自分で作成したアプリは、Googleの確認プロセスを経ていません。

**解決方法**:
1. 「詳細」をクリック
2. 「（アプリ名）に移動（安全ではないページ）」をクリック
3. これは自分で作成したアプリなので安全です

### 「アクセスがブロックされました」エラー

**原因**: OAuth同意画面でテストユーザーが追加されていない。

**解決方法**:
1. Google Cloud Consoleの「OAuth同意画面」に移動
2. 「テストユーザー」セクションで自分のメールアドレスを追加
3. 再度認証を試行

### 「認証情報ファイルが見つかりません」エラー

**原因**: `google-credentials.json` のパスが正しくない。

**解決方法**:
1. ファイルが正しい場所に配置されているか確認
2. 環境変数 `GOOGLE_APPLICATION_CREDENTIALS` が正しく設定されているか確認
3. パスに `~` を使用している場合、絶対パスに変更してみる

### 「認証トークンが無効です」エラー

**原因**: 認証トークンの有効期限が切れた、または破損している。

**解決方法**:
```bash
# 認証トークンを削除して再認証
rm token.json
# または
del token.json  # Windows
```

### 「APIが有効化されていません」エラー

**原因**: 必要なAPIが有効化されていない。

**解決方法**:
1. Google Cloud Consoleの「APIとサービス」→「ライブラリ」に移動
2. 以下のAPIが有効化されているか確認：
   - Google Sheets API
   - Google Docs API
   - Google Slides API
   - Google Drive API
3. 有効化されていない場合、各APIを有効化

### 「クォータを超過しました」エラー

**原因**: APIの使用量制限を超えた。

**解決方法**:
1. Google Cloud Consoleの「APIとサービス」→「クォータ」で使用状況を確認
2. 無料枠の場合、1日あたりの制限があります
3. 翌日まで待つか、課金を有効化してクォータを増やす

## セキュリティのベストプラクティス

### 認証情報ファイルの保護

1. **認証情報ファイルを共有しない**
   - `google-credentials.json` は秘密情報です
   - Gitリポジトリにコミットしない
   - `.gitignore` に追加する

2. **適切なファイルパーミッション**
   ```bash
   # macOS/Linuxの場合
   chmod 600 ~/.config/kiro-mcp/google-credentials.json
   ```

3. **認証トークンの管理**
   - `token.json` も秘密情報です
   - 定期的に再生成することを推奨

### スコープの最小化

必要最小限のスコープのみを要求してください：
- 読み取りのみの場合: `.readonly` スコープを使用
- 特定のAPIのみ使用する場合: 不要なAPIを無効化

## 参考リンク

- [Google Cloud Console](https://console.cloud.google.com/)
- [Google Sheets API ドキュメント](https://developers.google.com/sheets/api)
- [Google Docs API ドキュメント](https://developers.google.com/docs/api)
- [Google Slides API ドキュメント](https://developers.google.com/slides/api)
- [OAuth 2.0 認証ガイド](https://developers.google.com/identity/protocols/oauth2)

## サポート

問題が解決しない場合は、以下を確認してください：

1. エラーメッセージの全文
2. 使用しているPythonのバージョン
3. 実行したコマンド
4. ログファイルの内容（`MCP_LOG_LEVEL=DEBUG` で実行）

GitHubのIssueで質問することもできます。

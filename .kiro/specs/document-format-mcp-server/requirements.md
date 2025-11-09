# 要件定義書

## はじめに

この機能は、Kiro AIアシスタントが様々なドキュメント形式のファイルを読み取り、生成できるようにするMCP（Model Context Protocol）サーバを実装します。対象となるファイル形式は、Microsoft Office形式（PowerPoint、Word、Excel）およびGoogle Workspace形式（スプレッドシート、ドキュメント、スライド）です。これにより、ユーザーは既存の設計書をKiroに読み込ませたり、Kiroが生成した内容をこれらの形式で出力したりできるようになります。

## 用語集

- **MCP Server**: Model Context Protocolに準拠したサーバで、AIモデルに追加機能を提供するもの
- **Kiro**: 本AIアシスタントシステム
- **Document Reader**: ドキュメントファイルを読み取り、テキストや構造化データを抽出する機能
- **Document Writer**: 構造化データからドキュメントファイルを生成する機能
- **Microsoft Office形式**: .pptx、.docx、.xlsx形式のファイル
- **Google Workspace形式**: Googleスプレッドシート、Googleドキュメント、Googleスライドのファイル

## 要件

### 要件1

**ユーザーストーリー:** ユーザーとして、既存のPowerPointファイル（.pptx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、設計書の内容に基づいた開発作業を依頼できる。

#### 受入基準

1. WHEN ユーザーがPowerPointファイル（.pptx）の読み取りを要求したとき、THE Document Reader SHALL ファイルからテキスト内容、スライド構造、および画像の説明を抽出する
2. THE Document Reader SHALL 抽出したコンテンツを構造化されたテキスト形式でKiroに提供する
3. IF PowerPointファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す
4. THE Document Reader SHALL 各スライドのタイトル、本文、ノート、および表データを識別して抽出する

### 要件2

**ユーザーストーリー:** ユーザーとして、既存のWordファイル（.docx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、ドキュメントの内容を参照しながら作業を進められる。

#### 受入基準

1. WHEN ユーザーがWordファイル（.docx）の読み取りを要求したとき、THE Document Reader SHALL ファイルから本文テキスト、見出し構造、表、およびリストを抽出する
2. THE Document Reader SHALL ドキュメントの階層構造（見出しレベル）を保持して抽出する
3. THE Document Reader SHALL 表データをマークダウン形式またはCSV形式で抽出する
4. IF Wordファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す

### 要件3

**ユーザーストーリー:** ユーザーとして、既存のExcelファイル（.xlsx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、データ分析や処理のタスクを依頼できる。

#### 受入基準

1. WHEN ユーザーがExcelファイル（.xlsx）の読み取りを要求したとき、THE Document Reader SHALL すべてのシートからデータを抽出する
2. THE Document Reader SHALL 各シートの名前、セルデータ、および数式を識別して抽出する
3. THE Document Reader SHALL 抽出したデータを構造化形式（JSON、CSV、またはマークダウン表）で提供する
4. IF Excelファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す
5. THE Document Reader SHALL 最大100シートまでのExcelファイルを処理する

### 要件4

**ユーザーストーリー:** ユーザーとして、Googleスプレッドシート、Googleドキュメント、GoogleスライドのファイルをKiroに読み込ませたい。そうすることで、クラウド上の設計書を直接参照できる。

#### 受入基準

1. WHEN ユーザーがGoogleスプレッドシートのURLまたはファイルIDを提供したとき、THE Document Reader SHALL Google Sheets APIを使用してデータを取得する
2. WHEN ユーザーがGoogleドキュメントのURLまたはファイルIDを提供したとき、THE Document Reader SHALL Google Docs APIを使用してコンテンツを取得する
3. WHEN ユーザーがGoogleスライドのURLまたはファイルIDを提供したとき、THE Document Reader SHALL Google Slides APIを使用してコンテンツを取得する
4. THE MCP Server SHALL Google APIへの認証を安全に処理する
5. IF Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、THEN THE MCP Server SHALL 明確なエラーメッセージをユーザーに返す

### 要件5

**ユーザーストーリー:** ユーザーとして、Kiroが生成した内容をPowerPointファイル（.pptx）として出力したい。そうすることで、プレゼンテーション資料を自動生成できる。

#### 受入基準

1. WHEN ユーザーがPowerPoint形式での出力を要求したとき、THE Document Writer SHALL 提供された構造化データからPowerPointファイルを生成する
2. THE Document Writer SHALL タイトルスライド、コンテンツスライド、および箇条書きリストを含むスライドを作成する
3. THE Document Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
4. THE Document Writer SHALL 基本的なフォーマット（タイトル、本文、箇条書き）を適用する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す

### 要件6

**ユーザーストーリー:** ユーザーとして、Kiroが生成した内容をWordファイル（.docx）として出力したい。そうすることで、設計書やドキュメントを自動生成できる。

#### 受入基準

1. WHEN ユーザーがWord形式での出力を要求したとき、THE Document Writer SHALL 提供された構造化データからWordファイルを生成する
2. THE Document Writer SHALL 見出し、段落、表、および箇条書きリストを含むドキュメントを作成する
3. THE Document Writer SHALL 見出しレベル（H1、H2、H3など）を適切に適用する
4. THE Document Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す

### 要件7

**ユーザーストーリー:** ユーザーとして、Kiroが生成した内容をExcelファイル（.xlsx）として出力したい。そうすることで、データや表を自動生成できる。

#### 受入基準

1. WHEN ユーザーがExcel形式での出力を要求したとき、THE Document Writer SHALL 提供された構造化データからExcelファイルを生成する
2. THE Document Writer SHALL 複数のシートを含むワークブックを作成する
3. THE Document Writer SHALL セルデータ、基本的な書式設定、および列幅の自動調整を適用する
4. THE Document Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP Server SHALL エラーメッセージをユーザーに返す

### 要件8

**ユーザーストーリー:** ユーザーとして、Kiroが生成した内容をGoogleスプレッドシート、Googleドキュメント、Googleスライドとして出力したい。そうすることで、クラウド上で直接編集可能なファイルを作成できる。

#### 受入基準

1. WHEN ユーザーがGoogleスプレッドシート形式での出力を要求したとき、THE Document Writer SHALL Google Sheets APIを使用して新しいスプレッドシートを作成する
2. WHEN ユーザーがGoogleドキュメント形式での出力を要求したとき、THE Document Writer SHALL Google Docs APIを使用して新しいドキュメントを作成する
3. WHEN ユーザーがGoogleスライド形式での出力を要求したとき、THE Document Writer SHALL Google Slides APIを使用して新しいプレゼンテーションを作成する
4. THE Document Writer SHALL 作成したファイルのURLをユーザーに返す
5. IF Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、THEN THE MCP Server SHALL 明確なエラーメッセージをユーザーに返す

### 要件9

**ユーザーストーリー:** ユーザーとして、MCPサーバをKiroに簡単に統合したい。そうすることで、追加の設定なしで機能を使用できる。

#### 受入基準

1. THE MCP Server SHALL Model Context Protocol仕様に準拠する
2. THE MCP Server SHALL 利用可能なツール（読み取りおよび書き込み機能）をKiroに公開する
3. THE MCP Server SHALL 標準入出力（stdio）を介してKiroと通信する
4. THE MCP Server SHALL 起動時に30秒以内に初期化を完了する
5. THE MCP Server SHALL 各ツール呼び出しに対して明確な成功または失敗の応答を返す

### 要件10

**ユーザーストーリー:** 開発者として、ローカル環境で実装した機能を簡単にテストしたい。そうすることで、MCPサーバとして統合する前に各コンポーネントの動作を確認できる。

#### 受入基準

1. THE Development Environment SHALL サンプルファイル生成スクリプトを提供する
2. WHEN 開発者がサンプルファイル生成スクリプトを実行したとき、THE Script SHALL PowerPoint、Word、Excelのテストファイルを生成する
3. THE Development Environment SHALL リーダー機能をテストするスクリプトを提供する
4. WHEN 開発者がテストスクリプトを実行したとき、THE Script SHALL 各リーダークラスの動作を検証し、結果を表示する
5. THE Development Environment SHALL uvおよびpipの両方でのセットアップをサポートする
6. THE Development Environment SHALL クイックスタートガイドを提供し、5分以内にテスト環境をセットアップできるようにする
7. THE Development Environment SHALL 詳細なセットアップガイドを提供し、トラブルシューティング情報を含める

### 要件11

**ユーザーストーリー:** 開発者として、実装が設計仕様に準拠していることを確認したい。そうすることで、本番環境で正しく動作することを保証できる。

#### 受入基準

1. THE ExcelReader SHALL 設定可能なmax_sheets引数を受け取る__init__メソッドを持つ
2. THE ToolHandlers SHALL 設定値を使用してExcelReaderを正しく初期化する
3. THE MCP Server SHALL 起動時にTypeErrorなしで正常に初期化される
4. THE Document Writer Tools SHALL 実際のライター実装を呼び出してファイルを生成する
5. THE MCP Server SHALL 共通データモデル（DocumentContent、ReadResult、WriteResult）を定義し使用する
6. THE MCP Server SHALL エラーレスポンスを{"success": false, "error": {...}}形式のJSONとして返す
7. THE Google API Calls SHALL 最大3回のリトライと60秒のタイムアウトを実装する
8. THE Document Readers SHALL ファイルサイズ、シート数、スライド数の制限を設定値に基づいて検証する
9. THE Test Suite SHALL 各リーダー/ライター用の具体的なテストモジュールを含む

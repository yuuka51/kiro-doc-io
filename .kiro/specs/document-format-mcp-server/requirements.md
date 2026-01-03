# Requirements Document

## Introduction

この機能は、Kiro AIアシスタントが様々なドキュメント形式のファイルを読み取り、生成できるようにするMCP（Model Context Protocol）サーバを実装します。対象となるファイル形式は、Microsoft Office形式（PowerPoint、Word、Excel）およびGoogle Workspace形式（スプレッドシート、ドキュメント、スライド）です。これにより、ユーザーは既存の設計書をKiroに読み込ませたり、Kiroが生成した内容をこれらの形式で出力したりできるようになります。

本システムは、プロパティベーステストによる正確性検証を含む包括的な実装を行い、開発者向けのテスト環境とツールも提供します。

## Glossary

- **MCP_Server**: Model Context Protocolに準拠したサーバで、AIモデルに追加機能を提供するもの
- **Kiro**: 本AIアシスタントシステム
- **Document_Reader**: ドキュメントファイルを読み取り、テキストや構造化データを抽出する機能
- **Document_Writer**: 構造化データからドキュメントファイルを生成する機能
- **Microsoft_Office_Format**: .pptx、.docx、.xlsx形式のファイル
- **Google_Workspace_Format**: Googleスプレッドシート、Googleドキュメント、Googleスライドのファイル
- **Development_Environment**: 開発者がローカル環境で機能をテストするための環境とツール群
- **Script**: 特定の処理を自動化するプログラム
- **ExcelReader**: Excelファイルを読み取る専用のリーダークラス
- **ToolHandlers**: MCPツールの呼び出しを処理するハンドラークラス
- **Google_API_Calls**: Google Workspace APIへの呼び出し処理
- **Test_Suite**: システムの動作を検証するテストの集合
- **Property_Based_Test**: 全ての有効な入力に対して成り立つべき普遍的なプロパティを検証するテスト手法

## Requirements

### Requirement 1

**User Story:** ユーザーとして、既存のPowerPointファイル（.pptx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、設計書の内容に基づいた開発作業を依頼できる。

#### Acceptance Criteria

1. WHEN ユーザーがPowerPointファイル（.pptx）の読み取りを要求したとき、THE Document_Reader SHALL ファイルからテキスト内容、スライド構造、および画像の説明を抽出する
2. THE Document_Reader SHALL 抽出したコンテンツを構造化されたテキスト形式でKiroに提供する
3. IF PowerPointファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す
4. THE Document_Reader SHALL 各スライドのタイトル、本文、ノート、および表データを識別して抽出する

### Requirement 2

**User Story:** ユーザーとして、既存のWordファイル（.docx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、ドキュメントの内容を参照しながら作業を進められる。

#### Acceptance Criteria

1. WHEN ユーザーがWordファイル（.docx）の読み取りを要求したとき、THE Document_Reader SHALL ファイルから本文テキスト、見出し構造、表、およびリストを抽出する
2. THE Document_Reader SHALL ドキュメントの階層構造（見出しレベル）を保持して抽出する
3. THE Document_Reader SHALL 表データをマークダウン形式またはCSV形式で抽出する
4. IF Wordファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す

### Requirement 3

**User Story:** ユーザーとして、既存のExcelファイル（.xlsx）をKiroに読み込ませて、その内容を理解させたい。そうすることで、データ分析や処理のタスクを依頼できる。

#### Acceptance Criteria

1. WHEN ユーザーがExcelファイル（.xlsx）の読み取りを要求したとき、THE Document_Reader SHALL すべてのシートからデータを抽出する
2. THE Document_Reader SHALL 各シートの名前、セルデータ、および数式を識別して抽出する
3. THE Document_Reader SHALL 抽出したデータを構造化形式（JSON、CSV、またはマークダウン表）で提供する
4. IF Excelファイルが破損しているまたは読み取り不可能な場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す
5. THE Document_Reader SHALL 最大100シートまでのExcelファイルを処理する

### Requirement 4

**User Story:** ユーザーとして、Googleスプレッドシート、Googleドキュメント、GoogleスライドのファイルをKiroに読み込ませたい。そうすることで、クラウド上の設計書を直接参照できる。

#### Acceptance Criteria

1. WHEN ユーザーがGoogleスプレッドシートのURLまたはファイルIDを提供したとき、THE Document_Reader SHALL Google Sheets APIを使用してデータを取得する
2. WHEN ユーザーがGoogleドキュメントのURLまたはファイルIDを提供したとき、THE Document_Reader SHALL Google Docs APIを使用してコンテンツを取得する
3. WHEN ユーザーがGoogleスライドのURLまたはファイルIDを提供したとき、THE Document_Reader SHALL Google Slides APIを使用してコンテンツを取得する
4. THE MCP_Server SHALL Google APIへの認証を安全に処理する
5. IF Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、THEN THE MCP_Server SHALL 明確なエラーメッセージをユーザーに返す

### Requirement 5

**User Story:** ユーザーとして、Kiroが生成した内容をPowerPointファイル（.pptx）として出力したい。そうすることで、プレゼンテーション資料を自動生成できる。

#### Acceptance Criteria

1. WHEN ユーザーがPowerPoint形式での出力を要求したとき、THE Document_Writer SHALL 提供された構造化データからPowerPointファイルを生成する
2. THE Document_Writer SHALL タイトルスライド、コンテンツスライド、および箇条書きリストを含むスライドを作成する
3. THE Document_Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
4. THE Document_Writer SHALL 基本的なフォーマット（タイトル、本文、箇条書き）を適用する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す

### Requirement 6

**User Story:** ユーザーとして、Kiroが生成した内容をWordファイル（.docx）として出力したい。そうすることで、設計書やドキュメントを自動生成できる。

#### Acceptance Criteria

1. WHEN ユーザーがWord形式での出力を要求したとき、THE Document_Writer SHALL 提供された構造化データからWordファイルを生成する
2. THE Document_Writer SHALL 見出し、段落、表、および箇条書きリストを含むドキュメントを作成する
3. THE Document_Writer SHALL 見出しレベル（H1、H2、H3など）を適切に適用する
4. THE Document_Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す

### Requirement 7

**User Story:** ユーザーとして、Kiroが生成した内容をExcelファイル（.xlsx）として出力したい。そうすることで、データや表を自動生成できる。

#### Acceptance Criteria

1. WHEN ユーザーがExcel形式での出力を要求したとき、THE Document_Writer SHALL 提供された構造化データからExcelファイルを生成する
2. THE Document_Writer SHALL 複数のシートを含むワークブックを作成する
3. THE Document_Writer SHALL セルデータ、基本的な書式設定、および列幅の自動調整を適用する
4. THE Document_Writer SHALL 生成したファイルをユーザーがアクセス可能な場所に保存する
5. IF ファイル生成中にエラーが発生した場合、THEN THE MCP_Server SHALL エラーメッセージをユーザーに返す

### Requirement 8

**User Story:** ユーザーとして、Kiroが生成した内容をGoogleスプレッドシート、Googleドキュメント、Googleスライドとして出力したい。そうすることで、クラウド上で直接編集可能なファイルを作成できる。

#### Acceptance Criteria

1. WHEN ユーザーがGoogleスプレッドシート形式での出力を要求したとき、THE Document_Writer SHALL Google Sheets APIを使用して新しいスプレッドシートを作成する
2. WHEN ユーザーがGoogleドキュメント形式での出力を要求したとき、THE Document_Writer SHALL Google Docs APIを使用して新しいドキュメントを作成する
3. WHEN ユーザーがGoogleスライド形式での出力を要求したとき、THE Document_Writer SHALL Google Slides APIを使用して新しいプレゼンテーションを作成する
4. THE Document_Writer SHALL 作成したファイルのURLをユーザーに返す
5. IF Google APIへのアクセスが認証エラーまたは権限エラーで失敗した場合、THEN THE MCP_Server SHALL 明確なエラーメッセージをユーザーに返す

### Requirement 9

**User Story:** ユーザーとして、MCPサーバをKiroに簡単に統合したい。そうすることで、追加の設定なしで機能を使用できる。

#### Acceptance Criteria

1. THE MCP_Server SHALL Model Context Protocol仕様に準拠する
2. THE MCP_Server SHALL 利用可能なツール（読み取りおよび書き込み機能）をKiroに公開する
3. THE MCP_Server SHALL 標準入出力（stdio）を介してKiroと通信する
4. THE MCP_Server SHALL 起動時に30秒以内に初期化を完了する
5. THE MCP_Server SHALL 各ツール呼び出しに対して明確な成功または失敗の応答を返す

### Requirement 10

**User Story:** 開発者として、ローカル環境で実装した機能を簡単にテストしたい。そうすることで、MCPサーバとして統合する前に各コンポーネントの動作を確認できる。

#### Acceptance Criteria

1. THE Development_Environment SHALL サンプルファイル生成スクリプトを提供する
2. WHEN 開発者がサンプルファイル生成スクリプトを実行したとき、THE Script SHALL PowerPoint、Word、Excelのテストファイルを生成する
3. THE Development_Environment SHALL リーダー機能をテストするスクリプトを提供する
4. WHEN 開発者がテストスクリプトを実行したとき、THE Script SHALL 各リーダークラスの動作を検証し、結果を表示する
5. THE Development_Environment SHALL uvおよびpipの両方でのセットアップをサポートする
6. THE Development_Environment SHALL クイックスタートガイドを提供し、5分以内にテスト環境をセットアップできるようにする
7. THE Development_Environment SHALL 詳細なセットアップガイドを提供し、トラブルシューティング情報を含める

### Requirement 11

**User Story:** 開発者として、実装が設計仕様に準拠していることを確認したい。そうすることで、本番環境で正しく動作することを保証できる。

#### Acceptance Criteria

1. THE ExcelReader SHALL 設定可能なmax_sheets引数を受け取る__init__メソッドを持つ
2. THE ToolHandlers SHALL 設定値を使用してExcelReaderを正しく初期化する
3. WHEN MCP_Serverが起動するとき、THE MCP_Server SHALL TypeErrorなしで正常に初期化される
4. THE Document_Writer_Tools SHALL 実際のライター実装を呼び出してファイルを生成する
5. THE MCP_Server SHALL 共通データモデル（DocumentContent、ReadResult、WriteResult）を定義し使用する
6. THE MCP_Server SHALL エラーレスポンスを{"success": false, "error": {...}}形式のJSONとして返す
7. THE Google_API_Calls SHALL 最大3回のリトライと60秒のタイムアウトを実装する
8. THE Document_Readers SHALL ファイルサイズ、シート数、スライド数の制限を設定値に基づいて検証する
9. THE Test_Suite SHALL 各リーダー/ライター用の具体的なテストモジュールを含む

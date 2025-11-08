# 実装計画

- [x] 1. プロジェクト構造とMCPサーバの基盤を構築する




  - プロジェクトディレクトリ構造を作成（src/、tests/、設定ファイル）
  - pyproject.tomlを作成し、必要な依存関係を定義（mcp、python-pptx、python-docx、openpyxl、google-api-python-client）
  - 基本的なMCPサーバクラス（server.py）を実装し、標準入出力での通信を確立
  - 設定管理モジュール（utils/config.py）を実装し、環境変数と設定ファイルの読み込み機能を追加
  - エラー定義モジュール（utils/errors.py）を作成し、カスタム例外クラスを定義
  - _要件: 9.1, 9.3, 9.4_

- [x] 2. PowerPointリーダー機能を実装する



  - [x] 2.1 PowerPointReaderクラスを実装する


    - readers/powerpoint_reader.pyを作成
    - python-pptxを使用してスライドからテキスト、タイトル、ノート、表データを抽出するロジックを実装
    - 抽出したデータを構造化されたdict形式で返すメソッドを実装
    - _要件: 1.1, 1.2, 1.4_
  - [x] 2.2 エラーハンドリングを追加する

    - ファイルが存在しない場合のFileNotFoundError処理を実装
    - 破損したファイルのCorruptedFileError処理を実装
    - _要件: 1.3_
  - [ ]* 2.3 PowerPointリーダーのユニットテストを作成する
    - tests/readers/test_powerpoint_reader.pyを作成
    - サンプル.pptxファイルを使用したテストケースを実装
    - エラーケースのテストを追加
    - _要件: 1.1, 1.3_

- [x] 3. WordリーダーとExcelリーダー機能を実装する




  - [x] 3.1 WordReaderクラスを実装する


    - readers/word_reader.pyを作成
    - python-docxを使用して段落、見出し、表、リストを抽出するロジックを実装
    - 見出しレベルの階層構造を保持する機能を実装
    - エラーハンドリング（FileNotFoundError、CorruptedFileError）を追加
    - _要件: 2.1, 2.2, 2.3, 2.4_
  - [x] 3.2 ExcelReaderクラスを実装する


    - readers/excel_reader.pyを作成
    - openpyxlを使用してシート名、セルデータ、数式を抽出するロジックを実装
    - 最大100シートまでの制限を実装
    - エラーハンドリング（FileNotFoundError、CorruptedFileError）を追加
    - _要件: 3.1, 3.2, 3.3, 3.4, 3.5_
  - [ ]* 3.3 WordとExcelリーダーのユニットテストを作成する
    - tests/readers/test_word_reader.pyとtest_excel_reader.pyを作成
    - サンプルファイルを使用したテストケースを実装
    - _要件: 2.1, 3.1_

- [x] 4. Google Workspaceリーダー機能を実装する





  - [x] 4.1 GoogleWorkspaceReaderクラスの基盤を実装する


    - readers/google_reader.pyを作成
    - Google API認証処理を実装（credentials_pathから認証情報を読み込み）
    - 認証エラーと権限エラーのハンドリングを実装
    - _要件: 4.4, 4.5_
  - [x] 4.2 Googleスプレッドシート、ドキュメント、スライドの読み取りメソッドを実装する

    - read_spreadsheet()メソッドを実装（Google Sheets API使用）
    - read_document()メソッドを実装（Google Docs API使用）
    - read_slides()メソッドを実装（Google Slides API使用）
    - URLまたはファイルIDからファイルIDを抽出する処理を実装
    - _要件: 4.1, 4.2, 4.3_
  - [ ]* 4.3 Google Workspaceリーダーのユニットテストを作成する
    - tests/readers/test_google_reader.pyを作成
    - Google APIのモックを使用したテストケースを実装
    - _要件: 4.1, 4.5_

- [ ] 5. PowerPointライター機能を実装する
  - [ ] 5.1 PowerPointWriterクラスを実装する
    - writers/powerpoint_writer.pyを作成
    - python-pptxを使用してプレゼンテーションを生成するロジックを実装
    - タイトルスライド、コンテンツスライド、箇条書きスライドのレイアウトを実装
    - 基本的なフォーマット（タイトル、本文、箇条書き）を適用
    - _要件: 5.1, 5.2, 5.4_
  - [ ] 5.2 ファイル保存とエラーハンドリングを追加する
    - 指定されたoutput_pathにファイルを保存する処理を実装
    - ファイル生成エラーのハンドリングを実装
    - _要件: 5.3, 5.5_
  - [ ]* 5.3 PowerPointライターのユニットテストを作成する
    - tests/writers/test_powerpoint_writer.pyを作成
    - 生成されたファイルの検証テストを実装
    - _要件: 5.1_

- [ ] 6. WordライターとExcelライター機能を実装する
  - [ ] 6.1 WordWriterクラスを実装する
    - writers/word_writer.pyを作成
    - python-docxを使用してドキュメントを生成するロジックを実装
    - 見出し、段落、表、箇条書きリストの作成機能を実装
    - 見出しレベル（H1、H2、H3）の適用を実装
    - ファイル保存とエラーハンドリングを追加
    - _要件: 6.1, 6.2, 6.3, 6.4, 6.5_
  - [ ] 6.2 ExcelWriterクラスを実装する
    - writers/excel_writer.pyを作成
    - openpyxlを使用してワークブックを生成するロジックを実装
    - 複数シートの作成、セルデータの書き込み、基本的な書式設定を実装
    - 列幅の自動調整機能を実装
    - ファイル保存とエラーハンドリングを追加
    - _要件: 7.1, 7.2, 7.3, 7.4, 7.5_
  - [ ]* 6.3 WordとExcelライターのユニットテストを作成する
    - tests/writers/test_word_writer.pyとtest_excel_writer.pyを作成
    - 生成されたファイルの検証テストを実装
    - _要件: 6.1, 7.1_

- [ ] 7. Google Workspaceライター機能を実装する
  - [ ] 7.1 GoogleWorkspaceWriterクラスの基盤を実装する
    - writers/google_writer.pyを作成
    - Google API認証処理を実装
    - 認証エラーと権限エラーのハンドリングを実装
    - _要件: 8.5_
  - [ ] 7.2 Googleスプレッドシート、ドキュメント、スライドの作成メソッドを実装する
    - create_spreadsheet()メソッドを実装（Google Sheets API使用）
    - create_document()メソッドを実装（Google Docs API使用）
    - create_slides()メソッドを実装（Google Slides API使用）
    - 作成したファイルのURLを返す処理を実装
    - _要件: 8.1, 8.2, 8.3, 8.4_
  - [ ]* 7.3 Google Workspaceライターのユニットテストを作成する
    - tests/writers/test_google_writer.pyを作成
    - Google APIのモックを使用したテストケースを実装
    - _要件: 8.1, 8.5_

- [x] 8. MCPツール定義を実装し、サーバに統合する



  - [x] 8.1 ツール定義を作成する


    - tools/tool_definitions.pyを作成
    - 12個のMCPツール（read_powerpoint、read_word、read_excel、read_google_spreadsheet、read_google_document、read_google_slides、write_powerpoint、write_word、write_excel、write_google_spreadsheet、write_google_document、write_google_slides）を定義
    - 各ツールのスキーマ（入力パラメータ、出力形式）を定義
    - _要件: 9.2_
  - [x] 8.2 ツールハンドラーを実装する


    - 各ツールのハンドラー関数を実装し、対応するリーダー/ライタークラスを呼び出す
    - 入力パラメータの検証を実装
    - 成功/失敗のレスポンスを統一形式で返す処理を実装
    - _要件: 9.5_
  - [x] 8.3 MCPサーバにツールを登録する


    - server.pyの_register_tools()メソッドを実装
    - すべてのツールをMCPサーバに登録
    - サーバ起動時の初期化処理を実装（30秒以内に完了）
    - _要件: 9.2, 9.4_
  - [ ]* 8.4 統合テストを作成する
    - tests/test_integration.pyを作成
    - MCPサーバとツールのエンドツーエンドテストを実装
    - 実際のファイルを使用したテストケースを追加
    - _要件: 9.1, 9.5_

- [ ] 9. 設定ファイルとドキュメントを整備する
  - config.json.exampleを作成し、設定項目のサンプルを提供
  - README.mdを作成し、インストール方法、使用方法、設定方法を記述
  - Kiro設定例（.kiro/settings/mcp.json）をREADMEに記載
  - Google API認証情報の取得方法をドキュメント化
  - _要件: 9.1, 9.2_

- [ ] 10. パッケージングとデプロイメント準備を完了する
  - pyproject.tomlを最終調整し、パッケージメタデータを完成させる
  - エントリーポイント（コマンドライン実行）を設定
  - LICENSEファイルを追加
  - .gitignoreを作成し、不要なファイルを除外
  - パッケージをビルドし、ローカルでテスト実行
  - _要件: 9.1, 9.3_

- [x] 11. 開発・テスト環境を整備する
  - [x] 11.1 サンプルファイル生成スクリプトを作成する
    - create_sample_files.pyを作成
    - PowerPoint、Word、Excelのサンプルファイルを生成する機能を実装
    - test_files/ディレクトリにファイルを出力
    - _要件: 10.1, 10.2_
  - [x] 11.2 リーダー機能テストスクリプトを作成する
    - test_readers.pyを作成
    - ローカルファイル（PowerPoint、Word、Excel）のテスト機能を実装
    - Google Workspaceファイルのテスト機能を実装
    - 読み込んだ内容を表示する機能を実装
    - _要件: 10.3, 10.4_
  - [x] 11.3 セットアップドキュメントを作成する
    - QUICKSTART.mdを作成（5分でテストできるガイド）
    - SETUP.mdを作成（詳細なセットアップとトラブルシューティング）
    - README.mdを更新（クイックスタート情報を追加）
    - uvとpipの両方のセットアップ手順を記載
    - _要件: 10.5, 10.6, 10.7_
  - [x] 11.4 設計書にテスト環境の情報を追加する
    - design.mdに「開発・テスト環境」セクションを追加
    - テストツールの説明を追加
    - テストデータ構造の例を追加
    - Google Workspace認証設定の説明を追加
    - _要件: 10.7_

- [-] 12. ログ出力とエラーハンドリングを改善する




  - [ ] 12.1 ログ設定モジュールを作成する
    - utils/logging_config.pyを作成
    - Pythonの標準loggingモジュールを使用したログ設定を実装
    - ログレベル（DEBUG、INFO、WARNING、ERROR）の設定機能を追加
    - ログフォーマット（タイムスタンプ、ログレベル、メッセージ）を定義
    - 環境変数（MCP_LOG_LEVEL）からログレベルを読み込む機能を実装

    - _要件: 9.1_
  - [x] 12.2 リーダークラスにログ出力を追加する

    - PowerPointReader、WordReader、ExcelReader、GoogleWorkspaceReaderにロガーを追加
    - ファイル読み込み開始時にINFOレベルでログ出力
    - 読み込み成功時にINFOレベルでログ出力（処理時間を含む）
    - エラー発生時にERRORレベルでログ出力（スタックトレースを含む）
    - デバッグ情報（抽出したデータの概要）をDEBUGレベルで出力
    - _要件: 1.3, 2.4, 3.4, 4.5_
  - [ ] 12.3 ライタークラスにログ出力を追加する
    - PowerPointWriter、WordWriter、ExcelWriter、GoogleWorkspaceWriterにロガーを追加
    - ファイル生成開始時にINFOレベルでログ出力
    - 生成成功時にINFOレベルでログ出力（出力パスまたはURLを含む）
    - エラー発生時にERRORレベルでログ出力（スタックトレースを含む）
    - _要件: 5.5, 6.5, 7.5, 8.5_

  - [x] 12.4 MCPサーバにログ出力を追加する

    - server.pyにロガーを追加
    - サーバ起動時にINFOレベルでログ出力
    - ツール呼び出し時にINFOレベルでログ出力（ツール名とパラメータ）
    - エラー発生時にERRORレベルでログ出力
    - 現在のprint文をロガーに置き換える
    - _要件: 9.4, 9.5_
  - [x] 12.5 テストスクリプトのprint文をログ出力に変更する



    - test_readers.pyのprint文をロガーに置き換える
    - テスト結果の表示はINFOレベルで出力
    - エラー情報はERRORレベルで出力
    - デバッグ情報はDEBUGレベルで出力
    - _要件: 10.4_

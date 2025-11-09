# スクリプトディレクトリ

このディレクトリには、プロジェクトで使用するユーティリティスクリプトが含まれています。

## ディレクトリ構成

### analysis/
データ分析や検証を行うスクリプト
- `analyze_and_create_spec.py` - 仕様書の分析と作成
- `analyze_excel_structure.py` - Excelファイルの構造分析
- `check_excel_content.py` - Excelファイルの内容確認

### export/
データのエクスポート処理を行うスクリプト
- `export_all_to_readers.py` - 全データをreader形式でエクスポート
- `export_excel_to_md.py` - ExcelからMarkdownへの変換
- `export_spec_to_md.py` - 仕様書をMarkdownへエクスポート

### demo/
デモンストレーションや検証用スクリプト
- `demo_readers.py` - readerの動作デモ
- `verify_generated_pptx.py` - 生成されたPowerPointファイルの検証
- `verify_spec_export.py` - 仕様書エクスポートの検証

### setup/
初期設定やサンプルデータ作成用スクリプト
- `create_sample_files.py` - サンプルファイルの作成
- `create_specification.py` - 仕様書の作成

## 使用方法

各スクリプトはプロジェクトルートから実行してください：

```bash
python scripts/analysis/analyze_excel_structure.py
python scripts/export/export_excel_to_md.py
python scripts/demo/demo_readers.py
python scripts/setup/create_sample_files.py
```

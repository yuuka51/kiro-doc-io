"""生成されたPowerPointファイルを検証する"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from document_format_mcp_server.readers.powerpoint_reader import PowerPointReader

reader = PowerPointReader()

print("生成されたPowerPointファイルの検証")
print("=" * 60)

file_path = "test_files/output_comprehensive.pptx"
result = reader.read_file(file_path)

print(f"\nファイル: {file_path}")
print(f"スライド数: {len(result['slides'])}")

for slide in result['slides']:
    print(f"\nスライド {slide['slide_number']}: {slide['title']}")
    if slide['content']:
        content_preview = slide['content'][:100]
        print(f"  内容: {content_preview}...")

print("\n✓ 生成されたファイルは正常に読み込めました")

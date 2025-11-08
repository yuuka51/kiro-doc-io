"""Document reader modules."""

from .powerpoint_reader import PowerPointReader
from .word_reader import WordReader
from .excel_reader import ExcelReader
from .google_reader import GoogleWorkspaceReader

__all__ = [
    'PowerPointReader',
    'WordReader',
    'ExcelReader',
    'GoogleWorkspaceReader',
]

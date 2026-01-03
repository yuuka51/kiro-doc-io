"""MCP tool definitions for document format handling."""

# ツールスキーマ定義

# 読み取りツール

READ_POWERPOINT_SCHEMA = {
    "name": "read_powerpoint",
    "description": "PowerPointファイル(.pptx)を読み取り、スライドの内容を抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "読み取るPowerPointファイルのパス"
            }
        },
        "required": ["file_path"]
    }
}

READ_WORD_SCHEMA = {
    "name": "read_word",
    "description": "Wordファイル(.docx)を読み取り、ドキュメントの内容を抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "読み取るWordファイルのパス"
            }
        },
        "required": ["file_path"]
    }
}

READ_EXCEL_SCHEMA = {
    "name": "read_excel",
    "description": "Excelファイル(.xlsx)を読み取り、シートとセルデータを抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "読み取るExcelファイルのパス"
            }
        },
        "required": ["file_path"]
    }
}

READ_GOOGLE_SPREADSHEET_SCHEMA = {
    "name": "read_google_spreadsheet",
    "description": "Googleスプレッドシートを読み取り、シートとセルデータを抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_id": {
                "type": "string",
                "description": "GoogleスプレッドシートのファイルIDまたはURL"
            }
        },
        "required": ["file_id"]
    }
}

READ_GOOGLE_DOCUMENT_SCHEMA = {
    "name": "read_google_document",
    "description": "Googleドキュメントを読み取り、ドキュメントの内容を抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_id": {
                "type": "string",
                "description": "GoogleドキュメントのファイルIDまたはURL"
            }
        },
        "required": ["file_id"]
    }
}

READ_GOOGLE_SLIDES_SCHEMA = {
    "name": "read_google_slides",
    "description": "Googleスライドを読み取り、スライドの内容を抽出します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_id": {
                "type": "string",
                "description": "GoogleスライドのファイルIDまたはURL"
            }
        },
        "required": ["file_id"]
    }
}

# 書き込みツール

WRITE_POWERPOINT_SCHEMA = {
    "name": "write_powerpoint",
    "description": "構造化データからPowerPointファイル(.pptx)を生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "プレゼンテーションデータ（title、slidesを含む）",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "プレゼンテーションのタイトル"
                    },
                    "slides": {
                        "type": "array",
                        "description": "スライドの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "layout": {
                                    "type": "string",
                                    "enum": ["title", "content", "bullet"],
                                    "description": "スライドのレイアウトタイプ"
                                },
                                "title": {
                                    "type": "string",
                                    "description": "スライドのタイトル"
                                },
                                "content": {
                                    "description": "スライドの内容（文字列または配列）"
                                }
                            }
                        }
                    }
                }
            },
            "output_path": {
                "type": "string",
                "description": "出力するPowerPointファイルのパス"
            }
        },
        "required": ["data", "output_path"]
    }
}

WRITE_WORD_SCHEMA = {
    "name": "write_word",
    "description": "構造化データからWordファイル(.docx)を生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "ドキュメントデータ（title、sectionsを含む）",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "ドキュメントのタイトル"
                    },
                    "sections": {
                        "type": "array",
                        "description": "セクションの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "heading": {
                                    "type": "string",
                                    "description": "セクションの見出し"
                                },
                                "level": {
                                    "type": "integer",
                                    "description": "見出しレベル（1-3）"
                                },
                                "paragraphs": {
                                    "type": "array",
                                    "description": "段落の配列",
                                    "items": {
                                        "type": "string"
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "output_path": {
                "type": "string",
                "description": "出力するWordファイルのパス"
            }
        },
        "required": ["data", "output_path"]
    }
}

WRITE_EXCEL_SCHEMA = {
    "name": "write_excel",
    "description": "構造化データからExcelファイル(.xlsx)を生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "ワークブックデータ（sheetsを含む）",
                "properties": {
                    "sheets": {
                        "type": "array",
                        "description": "シートの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {
                                    "type": "string",
                                    "description": "シート名"
                                },
                                "data": {
                                    "type": "array",
                                    "description": "セルデータの2次元配列",
                                    "items": {
                                        "type": "array"
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "output_path": {
                "type": "string",
                "description": "出力するExcelファイルのパス"
            }
        },
        "required": ["data", "output_path"]
    }
}

WRITE_GOOGLE_SPREADSHEET_SCHEMA = {
    "name": "write_google_spreadsheet",
    "description": "構造化データからGoogleスプレッドシートを生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "スプレッドシートデータ（sheetsを含む）",
                "properties": {
                    "sheets": {
                        "type": "array",
                        "description": "シートの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {
                                    "type": "string",
                                    "description": "シート名"
                                },
                                "data": {
                                    "type": "array",
                                    "description": "セルデータの2次元配列",
                                    "items": {
                                        "type": "array"
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "title": {
                "type": "string",
                "description": "スプレッドシートのタイトル"
            }
        },
        "required": ["data", "title"]
    }
}

WRITE_GOOGLE_DOCUMENT_SCHEMA = {
    "name": "write_google_document",
    "description": "構造化データからGoogleドキュメントを生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "ドキュメントデータ（title、sectionsを含む）",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "ドキュメントのタイトル"
                    },
                    "sections": {
                        "type": "array",
                        "description": "セクションの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "heading": {
                                    "type": "string",
                                    "description": "セクションの見出し"
                                },
                                "level": {
                                    "type": "integer",
                                    "description": "見出しレベル（1-3）"
                                },
                                "paragraphs": {
                                    "type": "array",
                                    "description": "段落の配列",
                                    "items": {
                                        "type": "string"
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "title": {
                "type": "string",
                "description": "ドキュメントのタイトル"
            }
        },
        "required": ["data", "title"]
    }
}

WRITE_GOOGLE_SLIDES_SCHEMA = {
    "name": "write_google_slides",
    "description": "構造化データからGoogleスライドを生成します",
    "inputSchema": {
        "type": "object",
        "properties": {
            "data": {
                "type": "object",
                "description": "プレゼンテーションデータ（title、slidesを含む）",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "プレゼンテーションのタイトル"
                    },
                    "slides": {
                        "type": "array",
                        "description": "スライドの配列",
                        "items": {
                            "type": "object",
                            "properties": {
                                "layout": {
                                    "type": "string",
                                    "enum": ["title", "content", "bullet"],
                                    "description": "スライドのレイアウトタイプ"
                                },
                                "title": {
                                    "type": "string",
                                    "description": "スライドのタイトル"
                                },
                                "content": {
                                    "description": "スライドの内容（文字列または配列）"
                                }
                            }
                        }
                    }
                }
            },
            "title": {
                "type": "string",
                "description": "プレゼンテーションのタイトル"
            }
        },
        "required": ["data", "title"]
    }
}

# すべてのツールスキーマをリストとして公開
ALL_TOOL_SCHEMAS = [
    READ_POWERPOINT_SCHEMA,
    READ_WORD_SCHEMA,
    READ_EXCEL_SCHEMA,
    READ_GOOGLE_SPREADSHEET_SCHEMA,
    READ_GOOGLE_DOCUMENT_SCHEMA,
    READ_GOOGLE_SLIDES_SCHEMA,
    WRITE_POWERPOINT_SCHEMA,
    WRITE_WORD_SCHEMA,
    WRITE_EXCEL_SCHEMA,
    WRITE_GOOGLE_SPREADSHEET_SCHEMA,
    WRITE_GOOGLE_DOCUMENT_SCHEMA,
    WRITE_GOOGLE_SLIDES_SCHEMA,
]

# テスト用にTOOL_DEFINITIONSとしてもエクスポート
TOOL_DEFINITIONS = ALL_TOOL_SCHEMAS

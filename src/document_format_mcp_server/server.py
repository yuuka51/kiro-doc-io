"""MCP Server for document format handling."""

import asyncio
import sys
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool

from .tools.tool_definitions import ALL_TOOL_SCHEMAS
from .tools.tool_handlers import ToolHandlers
from .utils.config import Config
from .utils.errors import DocumentMCPError
from .utils.logging_config import setup_logging, get_logger


# ロガーの設定
logger = get_logger(__name__)


class DocumentMCPServer:
    """MCP Server for document format handling."""
    
    def __init__(self, config: Config | None = None):
        """
        Initialize the MCP server.
        
        Args:
            config: Optional configuration object
        """
        logger.info("MCPサーバを初期化中")
        self.config = config or Config()
        self.server = Server("document-format-server")
        self.tool_handlers = ToolHandlers(self.config)
        self._register_tools()
        logger.info("MCPサーバの初期化が完了")
    
    def _register_tools(self) -> None:
        """Register all available tools."""
        # ツールリストハンドラーを登録
        @self.server.list_tools()
        async def list_tools() -> list[Tool]:
            """利用可能なツールのリストを返す。"""
            tools = []
            for schema in ALL_TOOL_SCHEMAS:
                tools.append(
                    Tool(
                        name=schema["name"],
                        description=schema["description"],
                        inputSchema=schema["inputSchema"]
                    )
                )
            return tools
        
        # ツール呼び出しハンドラーを登録
        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[Any]:
            """ツールを呼び出す。"""
            logger.info(f"ツール呼び出し: {name}")
            logger.debug(f"パラメータ: {arguments}")
            
            # ツール名に対応するハンドラーを取得
            handler_map = {
                "read_powerpoint": self.tool_handlers.handle_read_powerpoint,
                "read_word": self.tool_handlers.handle_read_word,
                "read_excel": self.tool_handlers.handle_read_excel,
                "read_google_spreadsheet": self.tool_handlers.handle_read_google_spreadsheet,
                "read_google_document": self.tool_handlers.handle_read_google_document,
                "read_google_slides": self.tool_handlers.handle_read_google_slides,
                "write_powerpoint": self.tool_handlers.handle_write_powerpoint,
                "write_word": self.tool_handlers.handle_write_word,
                "write_excel": self.tool_handlers.handle_write_excel,
                "write_google_spreadsheet": self.tool_handlers.handle_write_google_spreadsheet,
                "write_google_document": self.tool_handlers.handle_write_google_document,
                "write_google_slides": self.tool_handlers.handle_write_google_slides,
            }
            
            handler = handler_map.get(name)
            if not handler:
                logger.error(f"不明なツール: {name}")
                return [
                    {
                        "type": "text",
                        "text": f"Unknown tool: {name}"
                    }
                ]
            
            try:
                # ハンドラーを呼び出し
                result = await handler(arguments)
                logger.info(f"ツール呼び出しが完了: {name}")
                
                # レスポンスを返す
                if "content" in result:
                    return result["content"]
                else:
                    return [result]
            
            except Exception as e:
                logger.error(f"ツール呼び出し中にエラーが発生: {name}", exc_info=True)
                raise
    
    async def run(self) -> None:
        """Start the MCP server with stdio transport."""
        logger.info("MCPサーバを起動中")
        async with stdio_server() as (read_stream, write_stream):
            logger.info("MCPサーバが起動しました（stdio通信）")
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options()
            )


def main() -> None:
    """Main entry point for the MCP server."""
    try:
        # ログ設定を初期化
        setup_logging()
        logger.info("Document Format MCP Serverを起動")
        
        # Initialize configuration
        config = Config()
        
        # Create and run server
        server = DocumentMCPServer(config)
        asyncio.run(server.run())
    
    except DocumentMCPError as e:
        logger.error(f"MCPエラー: {e.message}", exc_info=True)
        if e.details:
            logger.error(f"詳細: {e.details}")
        sys.exit(1)
    
    except KeyboardInterrupt:
        logger.info("サーバがユーザーによって停止されました")
        sys.exit(0)
    
    except Exception as e:
        logger.error(f"予期しないエラー: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

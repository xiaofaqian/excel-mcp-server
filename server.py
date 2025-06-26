#!/usr/bin/env python3
"""
Excel MCP Server using FastMCP
An MCP server with Excel file reading capabilities and basic tools.
"""

import logging
from mcp.server.fastmcp import FastMCP

# 导入工具模块
from tools import read_excel_file

# Configure logging for comprehensive error tracking
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("excel-mcp-server")

# Create FastMCP server instance
mcp = FastMCP("excel-mcp-server")

# 注册工具
mcp.tool()(read_excel_file)

if __name__ == "__main__":
    logger.info("[Setup] Initializing Excel MCP Server...")
    mcp.run()

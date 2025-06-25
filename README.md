# Excel MCP Server

一个基于FastMCP框架的Excel文件处理MCP服务器，提供Excel文件读取和处理功能。

## 功能特性

- 📊 支持读取Excel文件（.xlsx和.xls格式）
- 🔧 基于FastMCP框架构建
- 📋 支持指定工作表读取
- 🚀 高性能数据处理
- 🛡️ 完善的错误处理机制

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 启动服务器

```bash
python server.py
```

### 可用工具

#### read_excel_file

读取指定的Excel文件并返回结构化数据。

**参数：**
- `file_path` (必需): Excel文件的完整路径
- `sheet_name` (可选): 要读取的工作表名称，默认读取第一个工作表
- `max_rows` (可选): 最大读取行数，默认1000行

**返回格式：**
```json
{
  "success": true,
  "error": null,
  "data": {
    "file_path": "文件路径",
    "current_sheet": "当前工作表名称",
    "available_sheets": ["可用工作表列表"],
    "total_rows": 100,
    "total_columns": 5,
    "columns": ["列名列表"],
    "records": [{"数据记录"}],
    "max_rows_limit": 1000,
    "truncated": false
  }
}
```

## 项目结构

```
excel-mcp-server/
├── server.py              # 主服务器文件
├── tools/                 # 工具模块目录
│   ├── __init__.py        # 模块初始化文件
│   └── read_excel_file.py # Excel文件读取工具
├── requirements.txt       # 项目依赖
└── README.md             # 项目说明文档
```

## 技术栈

- **FastMCP**: MCP服务器框架
- **Pandas**: 数据处理库
- **OpenPyXL**: Excel文件处理
- **Pydantic**: 数据验证

## 开发说明

### 添加新工具

1. 在`tools/`目录下创建新的工具文件
2. 在`tools/__init__.py`中导入新工具
3. 在`server.py`中注册新工具

### 日志配置

项目使用Python标准logging模块，日志级别设置为INFO。

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。

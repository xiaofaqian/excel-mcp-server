# Excel MCP Server

一个基于FastMCP框架的Excel文件处理MCP服务器，为Cline、Cursor等支持MCP协议的工具提供Excel文件读取、搜索、编辑等功能。

## ⚠️ 重要说明

**本项目目前只在Windows环境下进行过测试，其他操作系统的兼容性未经验证。**

## 系统要求

- **操作系统**: Windows 10/11
- **Python版本**: 3.8 或更高版本
- **网络连接**: 安装过程需要下载Python依赖包

## 安装步骤

### 1. 克隆项目并进入目录

```bash
git clone <项目地址>
cd excel-mcp-server
```

### 2. 运行自动安装脚本

双击运行 `setup_mcp_config.bat` 文件，或在命令行中执行：

```cmd
setup_mcp_config.bat
```

该脚本将自动完成以下操作：
- ✅ 检查Python环境
- 📦 安装项目依赖
- ⚙️ 生成MCP配置文件 (`mcp_config.json`)
- 🧪 验证服务器可用性

## 配置MCP客户端

安装完成后，将生成的 `mcp_config.json` 文件内容复制到您使用的MCP客户端配置中：

### Cline (VSCode扩展)
1. 打开VSCode设置
2. 搜索"Cline MCP"
3. 将配置内容添加到MCP服务器配置中

### Cursor
1. 打开Cursor设置
2. 找到MCP配置选项
3. 添加生成的配置内容

### 其他MCP客户端
参考各客户端的MCP配置文档，将 `mcp_config.json` 中的配置添加到相应位置。

## 配置文件示例

生成的配置文件格式如下：

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "C:\\Python\\python.exe",
      "args": ["C:\\path\\to\\excel-mcp-server\\server.py"],
      "env": {}
    }
  }
}
```

## 验证安装

配置完成后，重启您的MCP客户端，应该能看到以下Excel处理工具：

- `read_excel_file` - 读取Excel文件数据
- `get_excel_summary` - 获取Excel文件概览
- `search_excel_data` - 搜索Excel数据
- `insert_excel_row` - 插入Excel行数据

## 故障排除

### 常见问题

**Q: 运行bat文件时提示"未找到Python"**
A: 请确保已安装Python 3.8+并添加到系统PATH环境变量中

**Q: 依赖安装失败**
A: 检查网络连接，或尝试使用国内pip镜像源

**Q: MCP客户端无法连接服务器**
A: 确认配置文件路径正确，且Python解释器路径有效

### 手动测试服务器

如需手动测试服务器是否正常工作：

```cmd
python server.py
```

服务器启动成功会显示相关日志信息。

## 许可证

MIT License

---

💡 **提示**: 如果在使用过程中遇到问题，请检查Python版本、网络连接和文件路径是否正确。

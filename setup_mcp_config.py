#!/usr/bin/env python3
"""
Windows MCP配置生成脚本
用于检查Python环境、安装依赖并生成MCP标准配置文件
"""

import sys
import os
import subprocess
import json
from pathlib import Path
import platform

def print_separator(title=""):
    """打印分隔线"""
    print("=" * 60)
    if title:
        print(f" {title} ")
        print("=" * 60)

def check_python_environment():
    """检查Python环境"""
    print_separator("检查Python环境")
    
    # 检查Python版本
    python_version = sys.version_info
    print(f"Python版本: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 8):
        print("❌ 错误: 需要Python 3.8或更高版本")
        return False
    else:
        print("✅ Python版本检查通过")
    
    # 检查操作系统
    os_name = platform.system()
    print(f"操作系统: {os_name}")
    
    if os_name != "Windows":
        print("⚠️  警告: 此脚本专为Windows设计")
    
    # 检查pip
    try:
        import pip
        print("✅ pip可用")
    except ImportError:
        print("❌ 错误: pip不可用")
        return False
    
    # 检查当前工作目录
    current_dir = Path.cwd()
    print(f"当前工作目录: {current_dir}")
    
    return True

def install_dependencies():
    """安装项目依赖"""
    print_separator("安装项目依赖")
    
    # 检查requirements.txt是否存在
    requirements_file = Path("requirements.txt")
    if not requirements_file.exists():
        print("❌ 错误: 找不到requirements.txt文件")
        return False
    
    print("📋 读取requirements.txt...")
    with open(requirements_file, 'r', encoding='utf-8') as f:
        requirements = f.read().strip().split('\n')
    
    print("依赖列表:")
    for req in requirements:
        if req.strip():
            print(f"  - {req.strip()}")
    
    # 安装依赖
    print("\n🔧 开始安装依赖...")
    try:
        # 使用subprocess运行pip install
        cmd = [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        
        print("✅ 依赖安装成功")
        if result.stdout:
            print("安装输出:")
            print(result.stdout)
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 依赖安装失败: {e}")
        if e.stderr:
            print("错误信息:")
            print(e.stderr)
        return False
    except Exception as e:
        print(f"❌ 安装过程中发生错误: {e}")
        return False

def get_server_path():
    """获取server.py的完整系统路径"""
    print_separator("获取服务器路径")
    
    # 检查server.py是否存在
    server_file = Path("server.py")
    if not server_file.exists():
        print("❌ 错误: 找不到server.py文件")
        return None
    
    # 获取绝对路径
    absolute_path = server_file.resolve()
    print(f"server.py绝对路径: {absolute_path}")
    
    # 验证文件是否可读
    try:
        with open(absolute_path, 'r', encoding='utf-8') as f:
            content = f.read(100)  # 读取前100个字符验证
        print("✅ server.py文件验证成功")
        return str(absolute_path)
    except Exception as e:
        print(f"❌ 无法读取server.py文件: {e}")
        return None

def generate_mcp_config(server_path):
    """生成MCP标准配置"""
    print_separator("生成MCP配置")
    
    # 获取Python解释器路径
    python_executable = sys.executable
    print(f"Python解释器路径: {python_executable}")
    
    # 生成配置
    config = {
        "mcpServers": {
            "excel-mcp-server": {
                "command": python_executable,
                "args": [server_path],
                "env": {}
            }
        }
    }
    
    # 格式化JSON输出
    config_json = json.dumps(config, indent=2, ensure_ascii=False)
    
    print("✅ MCP配置生成成功")
    print("\n" + "=" * 60)
    print(" MCP配置文件内容 ")
    print("=" * 60)
    print(config_json)
    print("=" * 60)
    
    # 保存到文件
    config_file = Path("mcp_config.json")
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(config_json)
        print(f"\n💾 配置已保存到: {config_file.resolve()}")
    except Exception as e:
        print(f"\n⚠️  无法保存配置文件: {e}")
    
    return config

def test_server():
    """测试服务器是否可以正常启动"""
    print_separator("测试服务器")
    
    try:
        # 尝试导入服务器模块来验证依赖
        print("🧪 测试服务器导入...")
        
        # 检查主要依赖是否可用
        dependencies = ['fastmcp', 'pandas', 'openpyxl', 'xlrd', 'pydantic']
        for dep in dependencies:
            try:
                __import__(dep)
                print(f"✅ {dep} 导入成功")
            except ImportError as e:
                print(f"❌ {dep} 导入失败: {e}")
                return False
        
        print("✅ 所有依赖验证通过")
        print("💡 提示: 您可以运行 'python server.py' 来启动MCP服务器")
        return True
        
    except Exception as e:
        print(f"❌ 服务器测试失败: {e}")
        return False

def main():
    """主函数"""
    print("🚀 Excel MCP服务器配置生成器")
    print("适用于Windows系统")
    print()
    
    try:
        # 1. 检查Python环境
        if not check_python_environment():
            print("\n❌ Python环境检查失败，请修复后重试")
            return False
        
        # 2. 安装依赖
        if not install_dependencies():
            print("\n❌ 依赖安装失败，请检查网络连接和权限")
            return False
        
        # 3. 获取服务器路径
        server_path = get_server_path()
        if not server_path:
            print("\n❌ 无法获取服务器路径")
            return False
        
        # 4. 生成配置
        config = generate_mcp_config(server_path)
        if not config:
            print("\n❌ 配置生成失败")
            return False
        
        print("\n" + "=" * 60)
        print(" 🎉 配置生成完成！ ")
        print("=" * 60)
        print("📋 使用说明:")
        print("1. 复制上面的JSON配置到您的MCP客户端配置文件中")
        print("2. 或者使用生成的 mcp_config.json 文件")
        print("3. 重启您的MCP客户端以加载新配置")
        print("4. 运行 'python server.py' 来启动服务器")
        
        return True
        
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断操作")
        return False
    except Exception as e:
        print(f"\n❌ 发生未预期的错误: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n💡 如果遇到问题，请检查:")
        print("  - Python版本是否为3.8+")
        print("  - 网络连接是否正常")
        print("  - 是否有足够的权限安装包")
        print("  - server.py文件是否存在")
        sys.exit(1)
    else:
        print("\n✅ 所有操作完成成功！")
        sys.exit(0)

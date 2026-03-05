# 修改后的 server.py
from fastmcp import FastMCP
import sys
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

mcp = FastMCP("My MCP Server with Graph API 🚀")

# 保留您的测试方法
@mcp.tool
def greet(name: str) -> str:
    """问候指定的人员
    
    Args:
        name: 要问候的人的姓名
        
    Returns:
        问候语字符串
    """
    return f"Hello, {name}!"

@mcp.tool  
def add(a: int, b: int) -> int:
    """计算两个整数的和
    
    Args:
        a: 第一个整数
        b: 第二个整数
        
    Returns:
        两数之和
    """
    return a + b

@mcp.tool
def subtract(a: int, b: int) -> int:
    """计算两个整数的差
    
    Args:
        a: 被减数
        b: 减数
        
    Returns:
        两数之差
    """
    return a - b

# 健康检查端点
@mcp.tool
def health_check() -> str:
    """健康检查
    
    Returns:
        服务状态
    """
    return "MCP Server is running! 🚀"

# 注册 Graph API 工具
try:
    # 导入和注册 Graph API 工具
    from tools import register_graph_tools
    print("[MCP] 成功导入 Graph API 工具模块", file=sys.stderr)
    register_graph_tools(mcp)
    print("[MCP] 成功注册 Graph API 工具", file=sys.stderr)
    
    # 导入和注册 OneDrive 文档工具
    from tools.onedrive_doc_tools import register_doc_tools
    print("[MCP] 成功导入 OneDrive 文档工具模块", file=sys.stderr)
    register_doc_tools(mcp)
    print("[MCP] 成功注册 OneDrive 文档工具", file=sys.stderr)
    
    # 导入和注册 OneDrive 邮件工具
    from tools.onedrive_mail_tools import register_mail_tools
    print("[MCP] 成功导入 OneDrive 邮件工具模块", file=sys.stderr)
    register_mail_tools(mcp)
    print("[MCP] 成功注册 OneDrive 邮件工具", file=sys.stderr)
    
    # 导入和注册 OneDrive OneNote 工具
    from tools.onedrive_note_tools import register_note_tools
    print("[MCP] 成功导入 OneDrive OneNote 工具模块", file=sys.stderr)
    register_note_tools(mcp)
    print("[MCP] 成功注册 OneDrive OneNote 工具", file=sys.stderr)
    
    # 导入和注册 OneDrive Teams 工具
    from tools.onedrive_teams_tools import register_teams_tools
    print("[MCP] 成功导入 OneDrive Teams 工具模块", file=sys.stderr)
    register_teams_tools(mcp)
    print("[MCP] 成功注册 OneDrive Teams 工具", file=sys.stderr)
    
    print("[MCP] ✅ 所有工具模块导入和注册完成！", file=sys.stderr)
    print("[MCP] 已注册工具类型: Graph API, OneDrive文档, 邮件, OneNote, Teams", file=sys.stderr)
    
except ImportError as e:
    print(f"[MCP] ❌ 模块导入失败: {e}", file=sys.stderr)
    print("[MCP] 请检查以下文件是否存在:", file=sys.stderr)
    print("[MCP] - tools/__init__.py", file=sys.stderr)
    print("[MCP] - tools/onedrive_doc_tools.py", file=sys.stderr)
    print("[MCP] - tools/onedrive_mail_tools.py", file=sys.stderr)
    print("[MCP] - tools/onedrive_note_tools.py", file=sys.stderr)
    print("[MCP] - tools/onedrive_teams_tools.py", file=sys.stderr)
    print("[MCP] - services/onedrive_service.py", file=sys.stderr)
    print("[MCP] - services/mongo_service.py", file=sys.stderr)
    import traceback
    traceback.print_exc()
    
except Exception as e:
    print(f"[MCP] ❌ 工具注册失败: {e}", file=sys.stderr)
    print("[MCP] 详细错误信息:", file=sys.stderr)
    import traceback
    traceback.print_exc()

if __name__ == "__main__":
    print("[MCP] 🚀 启动 MCP 服务器...", file=sys.stderr)
    print(f"[MCP] 服务将在端口 8001 上运行", file=sys.stderr)
    # 确保绑定到 0.0.0.0 以允许容器外部访问
    mcp.run(transport="http", port=8001, host="0.0.0.0")
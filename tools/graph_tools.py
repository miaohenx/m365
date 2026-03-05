"""
Microsoft Graph API 工具模块
纯粹的工具函数，依赖服务层
"""
from typing import Any, Dict, List, Optional
from datetime import datetime
import sys

from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_graph_tools(mcp_instance):
    """注册 Graph API 工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def ping() -> Dict[str, Any]:
        """简单的连接测试工具，返回所有上下文调试信息"""
        print(f"[MCP DEBUG] ping 工具被调用", file=sys.stderr)
        
        def safe_str(obj):
            """安全地转换对象为字符串"""
            try:
                return str(obj)
            except:
                return "无法转换为字符串"
        
        def safe_dict(obj):
            """安全地转换对象为字典"""
            try:
                if hasattr(obj, 'items'):
                    return {str(k): safe_str(v) for k, v in obj.items()}
                else:
                    return safe_str(obj)
            except:
                return "无法转换为字典"
        
        debug_info = {
            "status": "pong",
            "success": True,
            "timestamp": datetime.now().isoformat(),
            "context_debug": {}
        }
        
        # 1. 检查 contextvars
        try:
            from contextvars import copy_context
            
            ctx = copy_context()
            context_vars = {}
            
            for var, value in ctx.items():
                var_info = {
                    "var_name": safe_str(getattr(var, 'name', 'unknown')),
                    "var_type": safe_str(type(var)),
                    "value_type": safe_str(type(value)),
                    "has_headers": hasattr(value, 'headers'),
                    "value_preview": None
                }
                
                # 如果是request对象，获取更多信息
                if hasattr(value, 'headers'):
                    try:
                        var_info["headers"] = safe_dict(value.headers)
                        var_info["method"] = safe_str(getattr(value, 'method', None))
                        var_info["url"] = safe_str(getattr(value, 'url', None))
                        var_info["path"] = safe_str(getattr(value, 'path', None))
                        var_info["query_params"] = safe_str(getattr(value, 'query_params', None))
                    except Exception as e:
                        var_info["headers"] = f"获取headers失败: {safe_str(e)}"
                
                # 获取对象的所有属性
                try:
                    var_info["attributes"] = [attr for attr in dir(value) if not attr.startswith('_')][:10]
                except:
                    var_info["attributes"] = "无法获取属性"
                
                context_vars[f"var_{len(context_vars)}"] = var_info
            
            debug_info["context_debug"]["contextvars"] = {
                "count": len(context_vars),
                "vars": context_vars
            }
            
        except Exception as e:
            debug_info["context_debug"]["contextvars"] = {
                "error": safe_str(e)
            }
        
        # 2. 检查全局变量
        try:
            global_vars = {}
            
            current_globals = globals()
            for key, value in current_globals.items():
                if 'request' in key.lower() or 'context' in key.lower():
                    global_vars[key] = {
                        "type": safe_str(type(value)),
                        "has_headers": hasattr(value, 'headers')
                    }
            
            debug_info["context_debug"]["globals"] = global_vars
            
        except Exception as e:
            debug_info["context_debug"]["globals"] = {
                "error": safe_str(e)
            }
        
        # 3. 检查mcp_instance属性
        try:
            mcp_attrs = {}
            for attr in dir(mcp_instance)[:20]:  # 限制数量避免太多数据
                if not attr.startswith('_'):
                    try:
                        attr_value = getattr(mcp_instance, attr)
                        mcp_attrs[attr] = {
                            "type": safe_str(type(attr_value)),
                            "has_headers": hasattr(attr_value, 'headers'),
                            "callable": callable(attr_value),
                            "has_request": hasattr(attr_value, 'request'),
                            "has_current_request": hasattr(attr_value, 'current_request')
                        }
                    except Exception as e:
                        mcp_attrs[attr] = f"访问失败: {safe_str(e)}"
            
            debug_info["context_debug"]["mcp_instance"] = mcp_attrs
            
        except Exception as e:
            debug_info["context_debug"]["mcp_instance"] = {
                "error": safe_str(e)
            }
        
        # 4. 检查线程本地存储
        try:
            import threading
            local_data = threading.local()
            thread_info = {
                "current_thread": safe_str(threading.current_thread()),
                "local_attrs": [attr for attr in dir(local_data) if not attr.startswith('_')]
            }
            debug_info["context_debug"]["threading"] = thread_info
            
        except Exception as e:
            debug_info["context_debug"]["threading"] = {
                "error": safe_str(e)
            }
        
        # 5. 检查 fastmcp 模块
        try:
            import fastmcp
            fastmcp_attrs = {}
            interesting_attrs = [attr for attr in dir(fastmcp) if not attr.startswith('_')][:15]  # 限制数量
            
            for attr in interesting_attrs:
                try:
                    attr_value = getattr(fastmcp, attr)
                    fastmcp_attrs[attr] = {
                        "type": safe_str(type(attr_value)),
                        "has_headers": hasattr(attr_value, 'headers'),
                        "callable": callable(attr_value)
                    }
                except:
                    fastmcp_attrs[attr] = "无法访问"
            
            debug_info["context_debug"]["fastmcp_module"] = fastmcp_attrs
            
        except Exception as e:
            debug_info["context_debug"]["fastmcp_module"] = {
                "error": safe_str(e)
            }
        
        # 6. 尝试获取token（使用现有方法）
        try:
            token = get_token_from_context()
            debug_info["token_info"] = {
                "found": token is not None,
                "length": len(token) if token else 0,
                "preview": token[:20] + "..." if token and len(token) > 20 else token
            }
        except Exception as e:
            debug_info["token_info"] = {
                "error": safe_str(e)
            }
        
        print(f"[MCP DEBUG] 上下文调试信息收集完成", file=sys.stderr)
        return debug_info

    @mcp_instance.tool
    async def get_user_info() -> Dict[str, Any]:
        """
        获取当前登录用户的基本信息
        
        返回: 包含success状态、用户数据和错误信息的字典
        用户信息包括: 显示名称、邮箱地址、用户ID等
        使用场景: 验证用户身份或获取用户基本资料时使用
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_user_info 被调用，token: {token[:20]}...", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 调用 Graph API 获取用户信息
            url = f"{onedrive.BASE_URL}/me"
            result = onedrive._make_request('GET', url)
            user_data = result.json()
            
            print(f"[MCP] 用户信息获取成功: {user_data.get('displayName', 'N/A')}", file=sys.stderr)
            return {
                "success": True,
                "data": user_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 用户信息获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 用户信息获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_email_by_id(
        message_id: str,
        select: str = "subject,receivedDateTime,from,body,attachments,toRecipients,ccRecipients,importance,isRead"
    ) -> Dict[str, Any]:
        """
        通过邮件ID获取特定邮件的详细信息
        
        参数说明:
        - message_id (str): 邮件的唯一标识符ID，通常从邮件列表中获取
        - select (str): 指定要返回的邮件字段，多个字段用逗号分隔
        常用字段: subject(主题), receivedDateTime(接收时间), from(发件人), 
        body(邮件正文), attachments(附件), toRecipients(收件人), 
        ccRecipients(抄送人), importance(重要性), isRead(是否已读)
        
        返回: 包含success状态、邮件数据和错误信息的字典
        使用场景: 当需要查看特定邮件的完整详情时使用
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_email_by_id 被调用，message_id: {message_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 使用 get_single_mail 方法
            select_fields = select.split(',') if select else None
            mail_data = onedrive.get_single_mail(message_id, select_fields)
            
            print(f"[MCP] 邮件获取成功，ID: {message_id}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_email_attachments(
        message_id: str
    ) -> Dict[str, Any]:
        """
        获取指定邮件的所有附件信息
        
        参数说明:
        - message_id (str): 邮件的唯一标识符ID，要获取附件的邮件ID
        
        返回: 包含success状态、附件列表数据和错误信息的字典
        附件信息包括: 附件名称、大小、类型、ID等
        使用场景: 当需要查看邮件包含哪些附件，或准备下载附件时使用
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_email_attachments 被调用，message_id: {message_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 使用 get_mail_attachments 方法
            attachments_data = onedrive.get_mail_attachments(message_id)
            
            print(f"[MCP] 附件获取成功，数量: {len(attachments_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": attachments_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 附件获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 附件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_email_full_content(
        message_id: str
    ) -> Dict[str, Any]:
        """
        获取邮件的完整内容，包括HTML格式的正文和所有详细信息
        
        参数说明:
        - message_id (str): 邮件的唯一标识符ID，要获取完整内容的邮件ID
        
        返回: 包含success状态、完整邮件数据和错误信息的字典
        包含信息: 主题、发件人、收件人、抄送人、正文(HTML和文本)、附件、重要性、已读状态等
        使用场景: 当需要获取邮件的所有详细信息，包括完整正文内容时使用
        注意: 返回的数据量较大，包含完整的HTML正文
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_email_full_content 被调用，message_id: {message_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 使用 get_single_mail 方法，获取所有字段
            select_fields = [
                'subject', 'receivedDateTime', 'sentDateTime', 
                'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
                'body', 'bodyPreview', 'attachments', 
                'importance', 'isRead', 'flag', 'categories', 'internetMessageId'
            ]
            mail_data = onedrive.get_single_mail(message_id, select_fields)
            
            print(f"[MCP] 邮件完整内容获取成功，ID: {message_id}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件完整内容获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件完整内容获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
        
    @mcp_instance.tool
    async def test_tool(message: str = "hello") -> Dict[str, Any]:
        """测试工具"""
        print(f"[MCP DEBUG] test_tool 被调用，收到消息: {message}", file=sys.stderr)
        return {
            "success": True,
            "echo": f"收到消息: {message}",
            "timestamp": datetime.now().isoformat()
        }

    @mcp_instance.tool
    async def test_mongo_connection() -> Dict[str, Any]:
        """测试MongoDB连接"""
        print(f"[MCP DEBUG] test_mongo_connection 被调用", file=sys.stderr)
        
        try:
            from services.mongo_service import MongoTokenService
            success = await MongoTokenService.test_connection()
            
            if success:
                return {
                    "success": True,
                    "message": "MongoDB连接测试成功",
                    "timestamp": datetime.now().isoformat()
                }
            else:
                return {
                    "success": False,
                    "message": "MongoDB连接测试失败",
                    "timestamp": datetime.now().isoformat()
                }
                
        except Exception as e:
            return {
                "success": False,
                "message": f"MongoDB连接测试异常: {str(e)}",
                "timestamp": datetime.now().isoformat()
            }
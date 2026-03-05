#!/usr/bin/env python3
"""
Microsoft Graph API 工具模块
提供 Microsoft Graph API 调用功能
"""

import aiohttp
from typing import Any, Dict, Optional
from datetime import datetime
import sys

class GraphAPIError(Exception):
    """Graph API 错误类"""
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        self.message = message
        super().__init__(f"GraphAPI Error {status_code}: {message}")

async def make_graph_request(
    token: str,
    endpoint: str,
    method: str = "GET",
    params: Optional[Dict[str, Any]] = None,
    data: Optional[Dict[str, Any]] = None,
    content: Optional[bytes] = None
) -> Dict[str, Any]:
    """统一的 Graph API 请求函数"""
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json' if data else 'application/octet-stream'
    }
    
    url = f'https://graph.microsoft.com/v1.0{endpoint}'
    
    async with aiohttp.ClientSession() as session:
        try:
            if method == 'GET':
                async with session.get(url, headers=headers, params=params) as response:
                    return await handle_response(response)
            elif method == 'POST':
                if content:
                    headers['Content-Type'] = 'application/octet-stream'
                    async with session.post(url, headers=headers, data=content) as response:
                        return await handle_response(response)
                else:
                    async with session.post(url, headers=headers, json=data) as response:
                        return await handle_response(response)
            elif method == 'PUT':
                if content:
                    headers['Content-Type'] = 'application/octet-stream'
                    async with session.put(url, headers=headers, data=content) as response:
                        return await handle_response(response)
                else:
                    async with session.put(url, headers=headers, json=data) as response:
                        return await handle_response(response)
            else:
                raise GraphAPIError(400, f"Unsupported HTTP method: {method}")
                
        except aiohttp.ClientError as e:
            raise GraphAPIError(500, f"Request failed: {str(e)}")

async def handle_response(response: aiohttp.ClientResponse) -> Dict[str, Any]:
    """处理 API 响应"""
    if response.status in [200, 201, 202]:
        if response.content_length == 0:
            return {}
        try:
            return await response.json()
        except:
            return {"status": "success"}
    elif response.status == 401:
        raise GraphAPIError(401, "Unauthorized - Token may be expired")
    elif response.status == 403:
        raise GraphAPIError(403, "Forbidden - Insufficient permissions")
    elif response.status == 404:
        raise GraphAPIError(404, "Resource not found")
    else:
        error_text = await response.text()
        raise GraphAPIError(response.status, error_text)

def register_graph_tools(mcp_instance):
    """注册 Graph API 工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def ping() -> Dict[str, Any]:
        """简单的连接测试工具"""
        print(f"[MCP DEBUG] ping 工具被调用", file=sys.stderr)
        return {
            "status": "pong", 
            "success": True,
            "timestamp": datetime.now().isoformat()
        }

    @mcp_instance.tool
    async def get_user_info(token: str) -> Dict[str, Any]:
        """获取用户信息"""
        print(f"[MCP DEBUG] get_user_info 被调用，token: {token[:20]}...", file=sys.stderr)
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint='/me'
            )
            print(f"[MCP] 用户信息获取成功: {result.get('displayName', 'N/A')}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 用户信息获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }

    @mcp_instance.tool
    async def get_emails(
        token: str, 
        top: int = 10, 
        select: str = "subject,receivedDateTime,from,bodyPreview", 
        orderby: str = "receivedDateTime desc"
    ) -> Dict[str, Any]:
        """获取用户邮件列表"""
        print(f"[MCP DEBUG] get_emails 被调用，参数: token={token[:20]}..., top={top}", file=sys.stderr)
        
        params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint='/me/messages',
                params=params
            )
            print(f"[MCP] 邮件获取成功，数量: {len(result.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 邮件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 邮件获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
            
    @mcp_instance.tool
    async def get_email_by_id(
        token: str, 
        message_id: str,
        select: str = "subject,receivedDateTime,from,body,attachments,toRecipients,ccRecipients,importance,isRead"
    ) -> Dict[str, Any]:
        """
        通过邮件ID获取特定邮件的详细信息
        
        参数说明:
        - token (str): Microsoft Graph API访问令牌，用于身份验证
        - message_id (str): 邮件的唯一标识符ID，通常从邮件列表中获取
        - select (str): 指定要返回的邮件字段，多个字段用逗号分隔
        常用字段: subject(主题), receivedDateTime(接收时间), from(发件人), 
        body(邮件正文), attachments(附件), toRecipients(收件人), 
        ccRecipients(抄送人), importance(重要性), isRead(是否已读)
        
        返回: 包含success状态、邮件数据和错误信息的字典
        使用场景: 当需要查看特定邮件的完整详情时使用
        """
        print(f"[MCP DEBUG] get_email_by_id 被调用，参数: token={token[:20]}..., message_id={message_id[:20]}...", file=sys.stderr)
        
        params = {
            '$select': select
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint=f'/me/messages/{message_id}',
                params=params
            )
            print(f"[MCP] 邮件获取成功，ID: {message_id}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 邮件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 邮件获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_email_attachments(
        token: str, 
        message_id: str
    ) -> Dict[str, Any]:
        """
        获取指定邮件的所有附件信息
        
        参数说明:
        - token (str): Microsoft Graph API访问令牌，用于身份验证
        - message_id (str): 邮件的唯一标识符ID，要获取附件的邮件ID
        
        返回: 包含success状态、附件列表数据和错误信息的字典
        附件信息包括: 附件名称、大小、类型、ID等
        使用场景: 当需要查看邮件包含哪些附件，或准备下载附件时使用
        """
        print(f"[MCP DEBUG] get_email_attachments 被调用，参数: token={token[:20]}..., message_id={message_id[:20]}...", file=sys.stderr)
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint=f'/me/messages/{message_id}/attachments'
            )
            print(f"[MCP] 附件获取成功，数量: {len(result.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 附件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 附件获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_email_full_content(
        token: str, 
        message_id: str
    ) -> Dict[str, Any]:
        """
        获取邮件的完整内容，包括HTML格式的正文和所有详细信息
        
        参数说明:
        - token (str): Microsoft Graph API访问令牌，用于身份验证
        - message_id (str): 邮件的唯一标识符ID，要获取完整内容的邮件ID
        
        返回: 包含success状态、完整邮件数据和错误信息的字典
        包含信息: 主题、发件人、收件人、抄送人、正文(HTML和文本)、附件、重要性、已读状态等
        使用场景: 当需要获取邮件的所有详细信息，包括完整正文内容时使用
        注意: 返回的数据量较大，包含完整的HTML正文
        """
        print(f"[MCP DEBUG] get_email_full_content 被调用，参数: token={token[:20]}..., message_id={message_id[:20]}...", file=sys.stderr)
        
        params = {
            '$select': 'subject,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients,bccRecipients,body,bodyPreview,attachments,importance,isRead,flag,categories,internetMessageId'
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint=f'/me/messages/{message_id}',
                params=params
            )
            print(f"[MCP] 邮件完整内容获取成功，ID: {message_id}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 邮件完整内容获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 邮件完整内容获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_attachment_content(
        token: str, 
        message_id: str,
        attachment_id: str
    ) -> Dict[str, Any]:
        """
        获取邮件附件的具体内容
        
        参数说明:
        - token (str): Microsoft Graph API访问令牌，用于身份验证
        - message_id (str): 邮件的唯一标识符ID，附件所属的邮件ID
        - attachment_id (str): 附件的唯一标识符ID，从get_email_attachments方法中获取
        
        返回: 包含success状态、附件内容数据和错误信息的字典
        附件内容包括: 文件名、内容类型、大小、base64编码的文件内容等
        使用场景: 当需要下载或查看具体附件内容时使用
        注意: 大文件可能需要特殊处理，返回的内容是base64编码格式
        """
        print(f"[MCP DEBUG] get_attachment_content 被调用，参数: token={token[:20]}..., message_id={message_id[:20]}..., attachment_id={attachment_id[:20]}...", file=sys.stderr)
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint=f'/me/messages/{message_id}/attachments/{attachment_id}'
            )
            print(f"[MCP] 附件内容获取成功，attachment_id: {attachment_id}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 附件内容获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 附件内容获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
    @mcp_instance.tool
    async def send_email(
        token: str, 
        to_email: str, 
        subject: str, 
        content: str, 
        content_type: str = "Text"
    ) -> Dict[str, Any]:
        """发送邮件"""
        print(f"[MCP] 发送邮件到: {to_email}, 主题: {subject}", file=sys.stderr)
        
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": content_type,
                    "content": content
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            }
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint='/me/sendMail',
                method='POST',
                data=email_data
            )
            print(f"[MCP] 邮件发送成功", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 邮件发送失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 邮件发送异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_files(
        token: str, 
        top: int = 20, 
        select: str = "name,size,lastModifiedDateTime,webUrl,folder"
    ) -> Dict[str, Any]:
        """获取用户文件列表"""
        print(f"[MCP] 获取文件列表: top={top}", file=sys.stderr)
        
        params = {
            '$top': top,
            '$select': select
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint='/me/drive/root/children',
                params=params
            )
            print(f"[MCP] 文件获取成功，数量: {len(result.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 文件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 文件获取异常: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_calendar(
        token: str, 
        top: int = 10, 
        select: str = "subject,start,end,organizer", 
        orderby: str = "start/dateTime"
    ) -> Dict[str, Any]:
        """获取用户日历事件"""
        print(f"[MCP] 获取日历事件: top={top}", file=sys.stderr)
        
        params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
        try:
            result = await make_graph_request(
                token=token,
                endpoint='/me/events',
                params=params
            )
            print(f"[MCP] 日历事件获取成功，数量: {len(result.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
        except GraphAPIError as e:
            print(f"[MCP] 日历事件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Error {e.status_code}: {e.message}"
            }
        except Exception as e:
            print(f"[MCP] 日历事件获取异常: {e}", file=sys.stderr)
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
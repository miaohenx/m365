"""
OneDrive 邮件工具模块
纯粹的工具函数，依赖服务层
"""
from typing import Any, Dict, List, Optional
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_mail_tools(mcp_instance):
    """注册 OneDrive 邮件工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def list_emails(
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id",
        filter: Optional[str] = None,
        skip: Optional[int] = None,
        orderby: str = "receivedDateTime desc",
        search: Optional[str] = None,
        expand: Optional[str] = None,
        count: bool = False,
        folder: str = "inbox" 
    ) -> Dict[str, Any]:
        """
        List emails from a OneDrive/Outlook mailbox.
        
        参数说明:
        - top (int): Maximum number of emails to return (1-1000, default: 10)
        - select (str): Comma-separated list of properties to return
        - filter (str): OData filter expression for conditional filtering
        - skip (int): Number of emails to skip for pagination
        - orderby (str): Sort order specification (default: "receivedDateTime desc")
        - search (str): Search query string for content-based filtering
        - expand (str): Expand related properties (e.g., "attachments")
        - count (bool): Include total count in response (default: False)
        - folder (str): Mail folder name (default: "inbox")
                    Options: "inbox", "sentitems", "deleteditems", "drafts"
                    Use None to get all emails
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 构建 params 字典
        params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
        # 添加可选参数
        if filter:
            params['$filter'] = filter
        if skip is not None:
            params['$skip'] = skip
        if search:
            params['$search'] = search
        if expand:
            params['$expand'] = expand
        if count:
            params['$count'] = 'true'
            
        print(f"[MCP DEBUG] list_emails 被调用，token: {token[:20]}..., params: {params}, folder: {folder}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取邮件列表，传入 folder 参数
            mail_generator = onedrive.get_mail_with_filter(lambda: params, folder=folder)  # 👈 添加 folder 参数
            mail_data = next(mail_generator)  # 获取第一页数据
            
            print(f"[MCP] 邮件列表获取成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件列表获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件列表获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
            
    @mcp_instance.tool
    async def send_email(
        to: List[str], 
        subject: str, 
        body: str,
        cc: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        Send an email
        
        参数说明:
        - to (List[str]): List of recipient email addresses (required)
        - subject (str): The subject of the email (required)
        - body (str): The body content of the email (required)
        - cc (List[str]): List of CC email addresses (optional)
        
        Returns:
            dict: Dictionary containing success status, send result, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 处理可选的cc参数
        cc = cc or []
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 发送邮件
            result = onedrive.send_mail(to, cc, subject, body)
            
            print(f"[MCP] 邮件发送成功", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件发送失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件发送失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def list_mail_folders() -> Dict[str, Any]:
        """
        List mail folders
        
        Returns:
            dict: Dictionary containing success status, folder data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取邮件文件夹
            folders = onedrive.get_mail_folders()
            
            print(f"[MCP] 邮件文件夹获取成功，数量: {len(folders.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": folders,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件文件夹获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件文件夹获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def read_mails_in_folder(
        folder_id: str,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id",
        filter: Optional[str] = None,
        skip: Optional[int] = None,
        orderby: str = "receivedDateTime desc",
        search: Optional[str] = None,
        expand: Optional[str] = None,
        count: bool = False
    ) -> Dict[str, Any]:
        """
        Read mails in a specific folder
        
        参数说明:
        - folder_id (str): The ID of the folder to read mails from (required)
        - top (int): Maximum number of emails to return (1-1000, default: 10)
        - select (str): Comma-separated list of properties to return
        - filter (str): OData filter expression for conditional filtering
        - skip (int): Number of emails to skip for pagination
        - orderby (str): Sort order specification (default: "receivedDateTime desc")
        - search (str): Search query string for content-based filtering
        - expand (str): Expand related properties (e.g., "attachments")
        - count (bool): Include total count in response (default: False)
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 构建 filter_params 字典
        filter_params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
        # 添加可选参数
        if filter:
            filter_params['$filter'] = filter
        if skip is not None:
            filter_params['$skip'] = skip
        if search:
            filter_params['$search'] = search
        if expand:
            filter_params['$expand'] = expand
        if count:
            filter_params['$count'] = 'true'
            
        print(f"[MCP DEBUG] read_mails_in_folder 被调用，token: {token[:20]}..., folder_id: {folder_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取文件夹中的邮件
            mail_generator = onedrive.get_folder_messages(folder_id, filter_params)
            mail_data = next(mail_generator)  # 获取第一页数据
            
            print(f"[MCP] 文件夹邮件获取成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 文件夹邮件获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 文件夹邮件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_mail_attachments(mail_id: str) -> Dict[str, Any]:
        """
        Get mail attachments
        
        参数说明:
        - mail_id (str): The ID of the mail to get attachments from (required)
        
        Returns:
            dict: Dictionary containing success status, attachment data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取邮件附件
            attachments = onedrive.get_mail_attachments(mail_id)
            
            print(f"[MCP] 邮件附件获取成功，数量: {len(attachments.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": attachments,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件附件获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件附件获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def download_attachment(mail_id: str, attachment_id: str) -> Dict[str, Any]:
        """
        Download mail attachment
        
        参数说明:
        - mail_id (str): The ID of the mail containing the attachment (required)
        - attachment_id (str): The ID of the attachment to download (required)
        
        Returns:
            dict: Dictionary containing success status, attachment data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 下载附件
            attachment_data = onedrive.download_attachment(mail_id, attachment_id)
            
            print(f"[MCP] 附件下载成功", file=sys.stderr)
            return {
                "success": True,
                "data": attachment_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 附件下载失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 附件下载失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def reply_email(mail_id: str, body: str) -> Dict[str, Any]:
        """
        Reply to a mail
        
        参数说明:
        - mail_id (str): The ID of the mail to reply to (required)
        - body (str): The body content of the reply (required)
        
        Returns:
            dict: Dictionary containing success status, reply result, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 回复邮件
            result = onedrive.reply_to_mail(mail_id, body)
            
            print(f"[MCP] 邮件回复成功", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件回复失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件回复失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def find_emails_by_sender(
        sender_email: str,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: Optional[str] = None,  # 改为可选，默认为 None
        folder_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        根据发件人邮箱地址查找邮件（精确匹配）
        
        参数说明:
        - sender_email (str): 发件人的邮箱地址 (required)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式，留空使用默认排序 (optional, 不建议使用以避免复杂度错误)
        - folder_id (str): 指定文件夹ID，如果不指定则搜索所有邮件 (optional)
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当需要查找特定发件人邮箱发送的所有邮件时使用
        注意: 邮件默认按接收时间倒序排列，不需要额外指定排序
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 使用正确的OData语法
        filter_expression = f"from/emailAddress/address eq '{sender_email}'"
        
        params = {
            '$top': top,
            '$select': select,
            '$filter': filter_expression
        }
        
        # 只有当用户明确指定时才添加 orderby（不推荐）
        if orderby:
            params['$orderby'] = orderby
        
        print(f"[MCP DEBUG] find_emails_by_sender 被调用，sender: {sender_email}, params: {params}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            
            if folder_id:
                mail_generator = onedrive.get_folder_messages(folder_id, params)
            else:
                mail_generator = onedrive.get_mail_with_filter(lambda: params)
            
            mail_data = next(mail_generator)
            
            print(f"[MCP] 按发件人查找邮件成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 按发件人查找邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 按发件人查找邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    
    @mcp_instance.tool
    async def search_emails_by_sender_display_name(
        sender_name: str,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: str = "receivedDateTime desc"
    ) -> Dict[str, Any]:
        """
        通过获取所有邮件然后在客户端过滤发件人姓名
        
        参数说明:
        - sender_name (str): 发件人的姓名（支持中文和特殊字符）(required)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式 (default: "receivedDateTime desc")
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当需要根据发件人姓名查找邮件，特别是包含中文或特殊字符时使用
        注意: 此方法先获取更多邮件，然后在本地过滤，适用于复杂姓名搜索
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 获取更多邮件进行本地过滤（因为服务器端搜索对中文支持有限）
        fetch_count = min(top * 10, 100)  # 获取更多邮件以便过滤
        
        params = {
            '$top': fetch_count,
            '$select': select,
            '$orderby': orderby
            # 不使用 $search 或 $filter，直接获取邮件列表
        }
        
        print(f"[MCP DEBUG] search_emails_by_sender_display_name 被调用，sender_name: {sender_name}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            # 在客户端过滤邮件
            filtered_emails = []
            emails = mail_data.get('value', [])
            
            for email in emails:
                from_info = email.get('from', {})
                email_address_info = from_info.get('emailAddress', {})
                display_name = email_address_info.get('name', '')
                # 检查发件人姓名是否包含搜索关键词（不区分大小写）
                if sender_name.lower() in display_name.lower():
                    filtered_emails.append(email)
                    # 达到请求的数量就停止
                    if len(filtered_emails) >= top:
                        break
            
            # 构建返回结果
            result_data = {
                'value': filtered_emails[:top],
                '@odata.count': len(filtered_emails)
            }
            
            print(f"[MCP] 按发件人姓名搜索邮件成功，原始数量: {len(emails)}, 过滤后数量: {len(filtered_emails)}", file=sys.stderr)
            return {
                "success": True,
                "data": result_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 按发件人姓名搜索邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 按发件人姓名搜索邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def find_emails_by_date_range(
        start_date: str,
        end_date: Optional[str] = None,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: str = "receivedDateTime desc"
    ) -> Dict[str, Any]:
        """
        根据日期范围查找邮件
        
        参数说明:
        - start_date (str): 开始日期，格式: YYYY-MM-DD 或 YYYY-MM-DDTHH:MM:SS (required)
        - end_date (str): 结束日期，格式同上，如果不提供则搜索从开始日期到现在 (optional)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式 (default: "receivedDateTime desc")
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当需要查找特定时间段内的邮件时使用
        日期格式示例: "2024-01-01", "2024-01-01T09:00:00"
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            # 处理日期格式，确保是ISO格式
            if 'T' not in start_date:
                start_date = f"{start_date}T00:00:00Z"
            elif not start_date.endswith('Z'):
                start_date = f"{start_date}Z"
                
            if end_date:
                if 'T' not in end_date:
                    end_date = f"{end_date}T23:59:59Z"
                elif not end_date.endswith('Z'):
                    end_date = f"{end_date}Z"
                
                # 构建日期范围过滤器
                filter_expression = f"receivedDateTime ge {start_date} and receivedDateTime le {end_date}"
            else:
                # 只有开始日期，搜索从该日期到现在
                filter_expression = f"receivedDateTime ge {start_date}"
            
            params = {
                '$top': top,
                '$select': select,
                '$orderby': orderby,
                '$filter': filter_expression
            }
            
            print(f"[MCP DEBUG] find_emails_by_date_range 被调用，start: {start_date}, end: {end_date}, filter: {filter_expression}", file=sys.stderr)
            
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            print(f"[MCP] 按日期范围查找邮件成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 按日期范围查找邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 按日期范围查找邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def find_emails_by_subject_keyword(
        subject_keyword: str,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: str = "receivedDateTime desc",
        exact_match: bool = False
    ) -> Dict[str, Any]:
        """
        根据主题关键词查找邮件
        
        参数说明:
        - subject_keyword (str): 主题中要搜索的关键词 (required)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式 (default: "receivedDateTime desc")
        - exact_match (bool): 是否精确匹配整个主题 (default: False, 进行包含匹配)
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当需要根据邮件主题内容查找邮件时使用
        支持中文和特殊字符，使用客户端过滤确保兼容性
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 获取更多邮件进行本地过滤（避免服务器端搜索的编码问题）
        fetch_count = min(top * 5, 200)  # 获取更多邮件以便过滤
        
        params = {
            '$top': fetch_count,
            '$select': select,
            '$orderby': orderby
        }
        
        print(f"[MCP DEBUG] find_emails_by_subject_keyword 被调用，keyword: {subject_keyword}, exact_match: {exact_match}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            # 在客户端过滤邮件主题
            filtered_emails = []
            emails = mail_data.get('value', [])
            
            for email in emails:
                subject = email.get('subject', '').lower()
                keyword_lower = subject_keyword.lower()
                
                # 根据匹配模式进行过滤
                if exact_match:
                    if subject == keyword_lower:
                        filtered_emails.append(email)
                else:
                    if keyword_lower in subject:
                        filtered_emails.append(email)
                
                # 达到请求的数量就停止
                if len(filtered_emails) >= top:
                    break
            
            # 构建返回结果
            result_data = {
                'value': filtered_emails[:top],
                '@odata.count': len(filtered_emails)
            }
            
            print(f"[MCP] 按主题关键词查找邮件成功，原始数量: {len(emails)}, 过滤后数量: {len(filtered_emails)}", file=sys.stderr)
            return {
                "success": True,
                "data": result_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 按主题关键词查找邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 按主题关键词查找邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def find_emails_by_recent_days(
        days: int,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: str = "receivedDateTime desc"
    ) -> Dict[str, Any]:
        """
        查找最近几天内的邮件
        
        参数说明:
        - days (int): 最近多少天内的邮件 (required, 例如: 1=今天, 7=最近一周)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式 (default: "receivedDateTime desc")
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当需要快速查找最近几天的邮件时使用
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            from datetime import datetime, timedelta
            
            # 计算开始日期
            start_date = datetime.now() - timedelta(days=days)
            start_date_str = start_date.strftime("%Y-%m-%dT00:00:00Z")
            
            # 构建过滤器
            filter_expression = f"receivedDateTime ge {start_date_str}"
            
            params = {
                '$top': top,
                '$select': select,
                '$orderby': orderby,
                '$filter': filter_expression
            }
            
            print(f"[MCP DEBUG] find_emails_by_recent_days 被调用，days: {days}, start_date: {start_date_str}", file=sys.stderr)
            
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            print(f"[MCP] 查找最近{days}天邮件成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 查找最近邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 查找最近邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
        
    @mcp_instance.tool
    async def find_emails_by_sender_email_only(
        sender_email: str,
        top: int = 10,
        select: str = "subject,receivedDateTime,from,id,bodyPreview",
        orderby: str = "receivedDateTime desc"
    ) -> Dict[str, Any]:
        """
        仅通过邮箱地址精确查找邮件（最可靠的方法）
        
        参数说明:
        - sender_email (str): 发件人的邮箱地址 (required)
        - top (int): 返回的最大邮件数量 (default: 10)
        - select (str): 要返回的邮件字段，逗号分隔
        - orderby (str): 排序方式 (default: "receivedDateTime desc")
        
        Returns:
            dict: 包含成功状态、邮件数据和错误信息的字典
        使用场景: 当知道确切的发件人邮箱地址时使用，最可靠的搜索方法
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 使用 $filter 进行精确的邮箱地址匹配
        filter_expression = f"from/emailAddress/address eq '{sender_email}'"
        
        params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby,
            '$filter': filter_expression
        }
        
        print(f"[MCP DEBUG] find_emails_by_sender_email_only 被调用，sender_email: {sender_email}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            print(f"[MCP] 按邮箱地址查找邮件成功，数量: {len(mail_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": mail_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 按邮箱地址查找邮件失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 按邮箱地址查找邮件失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
            
    @mcp_instance.tool
    async def forward_email(
        mail_id: str, 
        to: List[str], 
        body: str,
        cc: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        Forward a mail
        
        参数说明:
        - mail_id (str): The ID of the mail to forward (required)
        - to (List[str]): The recipients of the forwarded mail (required)
        - body (str): The body content of the forwarded mail (required)
        - cc (List[str]): The CC recipients of the forwarded mail (optional)
        
        Returns:
            dict: Dictionary containing success status, forward result, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        # 处理可选的cc参数
        cc = cc or []
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 转发邮件
            result = onedrive.forward_mail(mail_id, to, cc, body)
            
            print(f"[MCP] 邮件转发成功", file=sys.stderr)
            return {
                "success": True,
                "data": result,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 邮件转发失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 邮件转发失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
    
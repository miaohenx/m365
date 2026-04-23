"""
OneDrive Email Tools Module
Pure utility functions that depend on service layer
"""
from typing import Any, Dict, List, Optional
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_mail_tools(mcp_instance):
    """Register OneDrive email tools to MCP instance"""
    
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
        
        Args:
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
        
        params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
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
            onedrive = await create_onedrive_service(token)
            
            mail_generator = onedrive.get_mail_with_filter(lambda: params, folder=folder)
            mail_data = next(mail_generator)
            
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
        
        Args:
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
        
        cc = cc or []
        
        try:
            onedrive = await create_onedrive_service(token)
            
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
            onedrive = await create_onedrive_service(token)
            
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
        
        Args:
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
        
        filter_params = {
            '$top': top,
            '$select': select,
            '$orderby': orderby
        }
        
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
            onedrive = await create_onedrive_service(token)
            
            mail_generator = onedrive.get_folder_messages(folder_id, filter_params)
            mail_data = next(mail_generator)
            
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
        
        Args:
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
            onedrive = await create_onedrive_service(token)
            
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
        
        Args:
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
            onedrive = await create_onedrive_service(token)
            
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
        
        Args:
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
            onedrive = await create_onedrive_service(token)
            
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
        orderby: Optional[str] = None,
        folder_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Find emails by sender email address (exact match)
        
        Args:
        - sender_email (str): Sender's email address (required)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order, leave empty to use default sorting (optional, not recommended to avoid complexity errors)
        - folder_id (str): Specific folder ID, if not specified search all emails (optional)
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you need to find all emails sent by a specific sender email address
        Note: Emails are sorted by received time in descending order by default, no need to specify additional sorting
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        filter_expression = f"from/emailAddress/address eq '{sender_email}'"
        
        params = {
            '$top': top,
            '$select': select,
            '$filter': filter_expression
        }
        
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
        Search emails by sender display name using client-side filtering
        
        Args:
        - sender_name (str): Sender's display name (supports Chinese and special characters) (required)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order (default: "receivedDateTime desc")
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you need to find emails by sender's display name, especially with Chinese or special characters
        Note: This method fetches more emails first, then filters locally, suitable for complex name searches
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        fetch_count = min(top * 10, 100)
        
        params = {
            '$top': fetch_count,
            '$select': select,
            '$orderby': orderby
        }
        
        print(f"[MCP DEBUG] search_emails_by_sender_display_name 被调用，sender_name: {sender_name}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            mail_generator = onedrive.get_mail_with_filter(lambda: params)
            mail_data = next(mail_generator)
            
            filtered_emails = []
            emails = mail_data.get('value', [])
            
            for email in emails:
                from_info = email.get('from', {})
                email_address_info = from_info.get('emailAddress', {})
                display_name = email_address_info.get('name', '')
                if sender_name.lower() in display_name.lower():
                    filtered_emails.append(email)
                    if len(filtered_emails) >= top:
                        break
            
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
        Find emails by date range
        
        Args:
        - start_date (str): Start date, format: YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS (required)
        - end_date (str): End date, same format as above, if not provided search from start date to now (optional)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order (default: "receivedDateTime desc")
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you need to find emails within a specific time period
        Date format examples: "2024-01-01", "2024-01-01T09:00:00"
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        try:
            if 'T' not in start_date:
                start_date = f"{start_date}T00:00:00Z"
            elif not start_date.endswith('Z'):
                start_date = f"{start_date}Z"
                
            if end_date:
                if 'T' not in end_date:
                    end_date = f"{end_date}T23:59:59Z"
                elif not end_date.endswith('Z'):
                    end_date = f"{end_date}Z"
                
                filter_expression = f"receivedDateTime ge {start_date} and receivedDateTime le {end_date}"
            else:
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
        Find emails by subject keyword
        
        Args:
        - subject_keyword (str): Keyword to search in subject (required)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order (default: "receivedDateTime desc")
        - exact_match (bool): Whether to match the entire subject exactly (default: False, performs contains match)
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you need to find emails based on email subject content
        Supports Chinese and special characters, uses client-side filtering for compatibility
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        fetch_count = min(top * 5, 200)
        
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
            
            filtered_emails = []
            emails = mail_data.get('value', [])
            
            for email in emails:
                subject = email.get('subject', '').lower()
                keyword_lower = subject_keyword.lower()
                
                if exact_match:
                    if subject == keyword_lower:
                        filtered_emails.append(email)
                else:
                    if keyword_lower in subject:
                        filtered_emails.append(email)
                
                if len(filtered_emails) >= top:
                    break
            
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
        Find emails from recent days
        
        Args:
        - days (int): Number of recent days to search (required, e.g.: 1=today, 7=last week)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order (default: "receivedDateTime desc")
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you need to quickly find emails from the last few days
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
            
            start_date = datetime.now() - timedelta(days=days)
            start_date_str = start_date.strftime("%Y-%m-%dT00:00:00Z")
            
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
        Find emails by email address only (most reliable method)
        
        Args:
        - sender_email (str): Sender's email address (required)
        - top (int): Maximum number of emails to return (default: 10)
        - select (str): Comma-separated list of email fields to return
        - orderby (str): Sort order (default: "receivedDateTime desc")
        
        Returns:
            dict: Dictionary containing success status, email data, and error information
        Use case: When you know the exact sender email address, most reliable search method
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
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
        
        Args:
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
        
        cc = cc or []
        
        try:
            onedrive = await create_onedrive_service(token)
            
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
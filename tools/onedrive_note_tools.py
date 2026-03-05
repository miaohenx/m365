"""
OneDrive OneNote 工具模块
纯粹的工具函数，依赖服务层
"""
from typing import Any, Dict, List
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_note_tools(mcp_instance):
    """注册 OneDrive OneNote 工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def read_note_books() -> Dict[str, Any]:
        """
        List and retrieve information about all accessible OneNote notebooks.

        Returns:
            dict: Dictionary containing success status, notebooks data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_books 被调用，token: {token[:20]}...", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取笔记本列表
            notebooks_iterator = onedrive.get_notebooks()
            
            # 转换迭代器为列表并结构化数据
            notebooks = []
            for notebook in notebooks_iterator:
                notebook_data = {
                    'id': notebook.get('id', ''),
                    'displayName': notebook.get('displayName', ''),
                    'createdDateTime': notebook.get('createdDateTime', ''),
                    'lastModifiedDateTime': notebook.get('lastModifiedDateTime', ''),
                    'links': notebook.get('links', {}),
                    'isDefault': notebook.get('isDefault', False),
                    'userRole': notebook.get('userRole', ''),
                    'isShared': notebook.get('isShared', False),
                    'sectionsUrl': notebook.get('sectionsUrl', ''),
                    'sectionGroupsUrl': notebook.get('sectionGroupsUrl', '')
                }
                notebooks.append(notebook_data)
            
            print(f"[MCP] 笔记本获取成功，数量: {len(notebooks)}", file=sys.stderr)
            return {
                "success": True,
                "data": notebooks,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 笔记本获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 笔记本获取失败: {e}", file=sys.stderr)
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'auth' in error_msg or 'unauthorized' in error_msg or 'token' in error_msg:
                error_detail = f"Authentication failed: Invalid or expired token. {str(e)}"
            elif 'network' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
                error_detail = f"Network connectivity issue: Unable to connect to OneNote service. {str(e)}"
            elif 'permission' in error_msg or 'access' in error_msg or 'forbidden' in error_msg:
                error_detail = f"Access permission error: Insufficient permissions to access OneNote notebooks. {str(e)}"
            elif 'service' in error_msg or 'unavailable' in error_msg:
                error_detail = f"Service unavailability: OneNote service is currently unavailable. {str(e)}"
            else:
                error_detail = f"Error reading OneNote notebooks: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def read_note_sections(notebook_id: str) -> Dict[str, Any]:
        """
        List and retrieve information about sections within a specific notebook.

        Args:
            notebook_id (str): Unique identifier of the target notebook

        Returns:
            dict: Dictionary containing success status, sections data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_sections 被调用，token: {token[:20]}..., notebook_id: {notebook_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取笔记本章节
            sections_iterator = onedrive.get_sections(notebook_id)
            
            # 转换迭代器为列表并结构化数据
            sections = []
            for section in sections_iterator:
                section_data = {
                    'id': section.get('id', ''),
                    'displayName': section.get('displayName', ''),
                    'createdDateTime': section.get('createdDateTime', ''),
                    'lastModifiedDateTime': section.get('lastModifiedDateTime', ''),
                    'pagesUrl': section.get('pagesUrl', ''),
                    'parentNotebook': section.get('parentNotebook', {}),
                    'isDefault': section.get('isDefault', False)
                }
                sections.append(section_data)
            
            print(f"[MCP] 章节获取成功，数量: {len(sections)}", file=sys.stderr)
            return {
                "success": True,
                "data": sections,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 章节获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 章节获取失败: {e}", file=sys.stderr)
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'auth' in error_msg or 'unauthorized' in error_msg or 'token' in error_msg:
                error_detail = f"Authentication failed: Invalid or expired token. {str(e)}"
            elif 'network' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
                error_detail = f"Network connectivity issue: Unable to connect to OneNote service. {str(e)}"
            elif 'permission' in error_msg or 'access' in error_msg or 'forbidden' in error_msg:
                error_detail = f"Access permission error: Insufficient permissions to access notebook. {str(e)}"
            elif 'service' in error_msg or 'unavailable' in error_msg:
                error_detail = f"Service unavailability: OneNote service is currently unavailable. {str(e)}"
            elif 'not found' in error_msg or 'invalid' in error_msg or '404' in error_msg:
                error_detail = f"Invalid notebook ID: Notebook '{notebook_id}' not found or not accessible. {str(e)}"
            else:
                error_detail = f"Error reading notebook sections: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def read_note_pages(section_id: str) -> Dict[str, Any]:
        """
        List and retrieve information about pages within a specific notebook section.

        Args:
            section_id (str): Unique identifier of the target section

        Returns:
            dict: Dictionary containing success status, pages data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_pages 被调用，token: {token[:20]}..., section_id: {section_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取章节页面
            pages_iterator = onedrive.get_pages(section_id)
            
            # 转换迭代器为列表并结构化数据
            pages = []
            for page in pages_iterator:
                page_data = {
                    'id': page.get('id', ''),
                    'title': page.get('title', ''),
                    'createdDateTime': page.get('createdDateTime', ''),
                    'lastModifiedDateTime': page.get('lastModifiedDateTime', ''),
                    'level': page.get('level', 0),
                    'order': page.get('order', 0),
                    'links': page.get('links', {}),
                    'parentSection': page.get('parentSection', {}),
                    'contentUrl': page.get('contentUrl', '')
                }
                pages.append(page_data)
            
            print(f"[MCP] 页面获取成功，数量: {len(pages)}", file=sys.stderr)
            return {
                "success": True,
                "data": pages,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 页面获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 页面获取失败: {e}", file=sys.stderr)
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'auth' in error_msg or 'unauthorized' in error_msg or 'token' in error_msg:
                error_detail = f"Authentication failed: Invalid or expired token. {str(e)}"
            elif 'network' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
                error_detail = f"Network connectivity issue: Unable to connect to OneNote service. {str(e)}"
            elif 'permission' in error_msg or 'access' in error_msg or 'forbidden' in error_msg:
                error_detail = f"Access permission error: Insufficient permissions to access section. {str(e)}"
            elif 'service' in error_msg or 'unavailable' in error_msg:
                error_detail = f"Service unavailability: OneNote service is currently unavailable. {str(e)}"
            elif 'not found' in error_msg or 'invalid' in error_msg or '404' in error_msg:
                error_detail = f"Invalid section ID: Section '{section_id}' not found or not accessible. {str(e)}"
            else:
                error_detail = f"Error reading section pages: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def read_note_page_content(page_id: str, content_format: str = "html") -> Dict[str, Any]:
        """
        Retrieve the actual content of a specific OneNote page.

        Args:
            page_id (str): Unique identifier of the target page
            content_format (str): Content format - "html" (default), "text", or "json"

        Returns:
            dict: Dictionary containing success status, page content, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_page_content 被调用，token: {token[:20]}..., page_id: {page_id}, format: {content_format}", file=sys.stderr)
        
        # 验证格式参数
        supported_formats = ["html", "text", "json"]
        if content_format not in supported_formats:
            return {
                "success": False,
                "data": None,
                "error": f"Unsupported format '{content_format}'. Supported formats are: {', '.join(supported_formats)}"
            }
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 读取页面内容
            content_bytes = onedrive.get_page_content(page_id)
            
            # 解码内容为UTF-8字符串
            if isinstance(content_bytes, bytes):
                content = content_bytes.decode('utf-8')
            else:
                content = str(content_bytes)
            
            # 根据请求的格式处理内容
            if content_format == "html":
                # 返回HTML内容
                processed_content = content
            elif content_format == "text":
                # 对于文本格式，理想情况下应该去除HTML标签
                # 现在先返回原始内容（API可能根据实现直接返回文本）
                processed_content = content
            elif content_format == "json":
                # 对于JSON格式，理想情况下应该解析并结构化内容
                # 现在先返回原始内容（API可能根据实现直接返回JSON）
                processed_content = content
            
            print(f"[MCP] 页面内容获取成功，长度: {len(processed_content)}", file=sys.stderr)
            return {
                "success": True,
                "data": {
                    "content": processed_content,
                    "format": content_format,
                    "page_id": page_id
                },
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 页面内容获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 页面内容获取失败: {e}", file=sys.stderr)
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'auth' in error_msg or 'unauthorized' in error_msg or 'token' in error_msg:
                error_detail = f"Authentication failed: Invalid or expired token. {str(e)}"
            elif 'network' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
                error_detail = f"Network connectivity issue: Unable to connect to OneNote service. {str(e)}"
            elif 'permission' in error_msg or 'access' in error_msg or 'forbidden' in error_msg:
                error_detail = f"Access permission error: Insufficient permissions to access page. {str(e)}"
            elif 'service' in error_msg or 'unavailable' in error_msg:
                error_detail = f"Service unavailability: OneNote service is currently unavailable. {str(e)}"
            elif 'not found' in error_msg or 'invalid' in error_msg or '404' in error_msg:
                error_detail = f"Invalid page ID: Page '{page_id}' not found or not accessible. {str(e)}"
            elif 'parse' in error_msg or 'encoding' in error_msg:
                error_detail = f"Content parsing error: Unable to parse page content. {str(e)}"
            elif 'size' in error_msg or 'large' in error_msg or 'limit' in error_msg:
                error_detail = f"Large content handling: Page content too large to process. {str(e)}"
            else:
                error_detail = f"Error reading page content: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }
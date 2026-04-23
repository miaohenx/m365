"""
OneDrive OneNote Tools Module
Pure utility functions that depend on service layer
"""
from typing import Any, Dict, List
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_note_tools(mcp_instance):
    """Register OneDrive OneNote tools to MCP instance"""
    
    @mcp_instance.tool
    async def read_note_books() -> Dict[str, Any]:
        """
        **STEP 1 - START HERE**: List all accessible OneNote notebooks and their IDs.
        
        This is the FIRST step in the OneNote workflow. You MUST call this tool first
        to get valid notebook IDs before calling any other OneNote tools.
        
        WORKFLOW:
        1. Call read_note_books() → Get notebook IDs
        2. Use notebook ID from step 1 → Call read_note_sections(notebook_id)
        3. Use section ID from step 2 → Call read_note_pages(section_id)
        4. Use page ID from step 3 → Call read_note_page_content(page_id)
        
        ⚠️ WARNING: Never guess or fabricate IDs. Always use IDs from API responses.

        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (list): List of notebooks with their IDs and metadata
                    - id (str): Notebook ID - USE THIS for read_note_sections()
                    - displayName (str): Notebook name
                    - createdDateTime (str): Creation timestamp
                    - lastModifiedDateTime (str): Last modification timestamp
                    - Other metadata fields
                - error (str|None): Error message if failed
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
            onedrive = await create_onedrive_service(token)
            notebooks_iterator = onedrive.get_notebooks()
            
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
        **STEP 2**: List all sections within a specific notebook.
        
        ⚠️ IMPORTANT: You MUST first call read_note_books() to get a valid notebook_id.
        DO NOT guess, fabricate, or use cached IDs. Always use fresh IDs from API responses.
        
        WORKFLOW REMINDER:
        1. ✅ Call read_note_books() first
        2. ➡️ YOU ARE HERE: Use the 'id' field from step 1 results
        3. Next: Use section IDs from this response → Call read_note_pages(section_id)
        
        Automatically handles large notebooks (>5000 sections) by switching to 
        filtered query method that returns recently modified sections.

        Args:
            notebook_id (str): **REQUIRED** - The exact 'id' value from read_note_books() response.
                Example: "1-e6cc806f-1039-4cc6-94a6-5d9c25fa012e"
                
                ⚠️ HOW TO GET THIS ID:
                1. Call read_note_books()
                2. Find your target notebook in the 'data' array
                3. Copy the EXACT 'id' field value
                4. Pass it to this function
                
                ❌ DO NOT: Use old IDs, guess IDs, or fabricate ID patterns
                ✅ DO: Always get fresh IDs from read_note_books() response

        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (list): List of sections with:
                    - id (str): Section ID - USE THIS for read_note_pages()
                    - displayName (str): Section name
                    - Other metadata fields
                - metadata (dict): Query metadata including count and notes
                - error (str|None): Error message if failed
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "metadata": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_sections 被调用，token: {token[:20]}..., notebook_id: {notebook_id}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            sections_iterator = onedrive.get_sections(notebook_id)
            
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
            
            metadata = {
                'total_count': len(sections),
                'notebook_id': notebook_id,
                'sorted_by': 'lastModifiedDateTime (desc)',
                'note': None
            }
            
            if len(sections) >= 100:
                metadata['note'] = 'Large notebook detected. Results may be limited to recently modified sections due to API limitations (>5000 items).'
            
            print(f"[MCP] 章节获取成功，数量: {len(sections)}", file=sys.stderr)
            return {
                "success": True,
                "data": sections,
                "metadata": metadata,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 章节获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "metadata": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 章节获取失败: {e}", file=sys.stderr)
            
            error_msg = str(e).lower()
            
            if 'auth' in error_msg or 'unauthorized' in error_msg or 'token' in error_msg:
                error_detail = f"Authentication failed: Invalid or expired token. {str(e)}"
            elif 'network' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
                error_detail = f"Network connectivity issue: Unable to connect to OneNote service. {str(e)}"
            elif 'permission' in error_msg or 'access' in error_msg or 'forbidden' in error_msg:
                if '403' in str(e) or 'forbidden' in error_msg:
                    error_detail = (
                        f"Large notebook handling failed: This notebook contains too many sections (>5000). "
                        f"The automatic fallback method also encountered issues. "
                        f"Error details: {str(e)}"
                    )
                else:
                    error_detail = f"Access permission error: Insufficient permissions to access notebook. {str(e)}"
            elif 'service' in error_msg or 'unavailable' in error_msg:
                error_detail = f"Service unavailability: OneNote service is currently unavailable. {str(e)}"
            elif 'not found' in error_msg or 'invalid' in error_msg or '404' in error_msg:
                error_detail = (
                    f"Invalid notebook ID: Notebook '{notebook_id}' not found or not accessible. "
                    f"Please call read_note_books() first to get valid notebook IDs. {str(e)}"
                )
            elif '所有方法都失败' in str(e) or 'all methods failed' in error_msg:
                error_detail = (
                    f"Multiple method failure: Unable to retrieve sections using any available method. "
                    f"This typically occurs with very large notebooks (>5000 sections) or API limitations. "
                    f"Details: {str(e)}"
                )
            else:
                error_detail = f"Error reading notebook sections: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "metadata": {
                    'notebook_id': notebook_id,
                    'attempted_methods': ['direct_fetch', 'filtered_query']
                },
                "error": error_detail
            }

    @mcp_instance.tool
    async def read_note_pages(section_id: str) -> Dict[str, Any]:
        """
        **STEP 3**: List all pages within a specific section.
        
        ⚠️ IMPORTANT: You MUST first call read_note_sections() to get a valid section_id.
        DO NOT guess, fabricate, or use cached IDs. Always use fresh IDs from API responses.
        
        WORKFLOW REMINDER:
        1. ✅ Call read_note_books() first
        2. ✅ Call read_note_sections(notebook_id) second
        3. ➡️ YOU ARE HERE: Use the 'id' field from step 2 results
        4. Next: Use page IDs from this response → Call read_note_page_content(page_id)

        Args:
            section_id (str): **REQUIRED** - The exact 'id' value from read_note_sections() response.
                Example: "1-e230f970-f123-4a3b-aa71-c7d060d8e6cb"
                
                ⚠️ HOW TO GET THIS ID:
                1. Call read_note_sections(notebook_id)
                2. Find your target section in the 'data' array
                3. Copy the EXACT 'id' field value
                4. Pass it to this function
                
                ❌ DO NOT: Use old IDs, guess IDs, or fabricate ID patterns
                ✅ DO: Always get fresh IDs from read_note_sections() response

        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (list): List of pages with:
                    - id (str): Page ID - USE THIS for read_note_page_content()
                    - title (str): Page title
                    - Other metadata fields
                - error (str|None): Error message if failed
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
            onedrive = await create_onedrive_service(token)
            pages_iterator = onedrive.get_pages(section_id)
            
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
                error_detail = (
                    f"Invalid section ID: Section '{section_id}' not found or not accessible. "
                    f"Please call read_note_sections(notebook_id) first to get valid section IDs. {str(e)}"
                )
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
        **STEP 4 - FINAL STEP**: Retrieve the actual content of a specific OneNote page.
        
        ⚠️ IMPORTANT: You MUST first call read_note_pages() to get a valid page_id.
        DO NOT guess, fabricate, or use cached IDs. Always use fresh IDs from API responses.
        
        WORKFLOW REMINDER:
        1. ✅ Call read_note_books() first
        2. ✅ Call read_note_sections(notebook_id) second
        3. ✅ Call read_note_pages(section_id) third
        4. ➡️ YOU ARE HERE: Use the 'id' field from step 3 results

        Args:
            page_id (str): **REQUIRED** - The exact 'id' value from read_note_pages() response.
                Example: "1-7ab23781-6dd3-49a0-9cc1-f4257da00ef9"
                
                ⚠️ HOW TO GET THIS ID:
                1. Call read_note_pages(section_id)
                2. Find your target page in the 'data' array
                3. Copy the EXACT 'id' field value
                4. Pass it to this function
                
                ❌ DO NOT: Use old IDs, guess IDs, or fabricate ID patterns
                ✅ DO: Always get fresh IDs from read_note_pages() response
                
            content_format (str): Content format - "html" (default), "text", or "json"

        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (dict): Page content with:
                    - content (str): The actual page content
                    - format (str): Content format used
                    - page_id (str): The page ID
                - error (str|None): Error message if failed
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_note_page_content 被调用，token: {token[:20]}..., page_id: {page_id}, format: {content_format}", file=sys.stderr)
        
        supported_formats = ["html", "text", "json"]
        if content_format not in supported_formats:
            return {
                "success": False,
                "data": None,
                "error": f"Unsupported format '{content_format}'. Supported formats are: {', '.join(supported_formats)}"
            }
        
        try:
            onedrive = await create_onedrive_service(token)
            content_bytes = onedrive.get_page_content(page_id)
            
            if isinstance(content_bytes, bytes):
                content = content_bytes.decode('utf-8')
            else:
                content = str(content_bytes)
            
            if content_format == "html":
                processed_content = content
            elif content_format == "text":
                processed_content = content
            elif content_format == "json":
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
                error_detail = (
                    f"Invalid page ID: Page '{page_id}' not found or not accessible. "
                    f"Please call read_note_pages(section_id) first to get valid page IDs. {str(e)}"
                )
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
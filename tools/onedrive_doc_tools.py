"""
OneDrive Tools Module
Pure utility functions that depend on service layer
Optimized AI call guidance to ensure correct tool usage workflow
"""
from typing import Any, Dict
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_doc_tools(mcp_instance):
    """Register OneDrive tools to MCP instance"""
    

    @mcp_instance.tool
    async def list_doc_from_onedrive(url: str) -> Dict[str, Any]:
        """
        List files in a folder from a OneDrive SHARED LINK (from other people).
        
        ⚠️ IMPORTANT: This is ONLY for shared folder links from other people.
        For YOUR OWN OneDrive folders, use list_files() instead.
        
        Args:
            url: The OneDrive sharing URL of a folder
            
        Returns:
            List of files and folders in the shared folder
            
        Example:
            url = "https://1drv.ms/f/s!AkR4bF..."
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
            onedrive.get_driveitem(url)
            path = '/' + url.split('/')[-1]
            dir_content = onedrive.listdir(path)
            
            print(f"[MCP] 目录列表获取成功，项目数量: {len(dir_content.json_data.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": dir_content.json_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 目录列表获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 目录列表获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_file_info(url: str) -> Dict[str, Any]:
        """
        Get file/folder metadata from a OneDrive SHARED LINK (from other people).
        
        ⚠️ IMPORTANT: This is ONLY for shared links from other people.
        For YOUR OWN OneDrive files, use get_file_or_folder_info() instead.
        
        Args:
            url: The OneDrive sharing URL
            
        Returns:
            File/folder metadata including name, size, type, dates, etc.
            
        Example:
            url = "https://1drv.ms/w/s!AkR4bF..."
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] get_file_info 被调用，token: {token[:20]}..., url: {url}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            onedrive.get_driveitem(url)
            
            print(f"[MCP] 文件信息获取成功: {onedrive.driveitem.get('name', 'N/A')}", file=sys.stderr)
            return {
                "success": True,
                "data": onedrive.driveitem,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 文件信息获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 文件信息获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def search_files(query: str, folder_url: str) -> Dict[str, Any]:
        """
        Search files in a SHARED OneDrive folder (from a sharing link).
        
        ⚠️ IMPORTANT: This is ONLY for shared folder links from other people.
        For searching YOUR OWN OneDrive, this function is not supported yet.
        
        Args:
            query: The search keyword
            folder_url: The OneDrive sharing URL of the folder to search in
            
        Returns:
            List of files matching the search query
            
        Example:
            query = "report"
            folder_url = "https://1drv.ms/f/s!AkR4bF..."
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
            onedrive.get_driveitem(folder_url)
            
            search_results = onedrive.search_files(query)
            
            print(f"[MCP] 文件搜索成功，结果数量: {len(search_results.get('value', []))}", file=sys.stderr)
            return {
                "success": True,
                "data": search_results,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 文件搜索失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 文件搜索失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
    
    # ==================== User's own OneDrive file operations ====================
    
    @mcp_instance.tool
    async def list_files(path: str = '/') -> Dict[str, Any]:
        """
        📂 List files and folders in YOUR OWN OneDrive account.
        
        🎯 PRIMARY USE CASE: This is your STARTING POINT for browsing OneDrive.
        Always call this FIRST to explore the folder structure.
        
        ⚠️ IMPORTANT WORKFLOW:
        1. Start with list_files(path='/') to see root folders
        2. Navigate into folders: list_files(path='/Documents')
        3. Find your target file's full path
        4. Then use read_file_from_onedrive() with the EXACT path
        
        Args:
            path: The folder path in your OneDrive (case-sensitive)
                Default: '/' (root directory)
                
        Path Examples:
            ✅ '/' - Root directory (DEFAULT, use this to start)
            ✅ '/Documents' - Documents folder
            ✅ '/Pictures/Vacation' - Nested folder
            ✅ '/Projects/Python' - Deep nested folder
            ❌ 'Documents' - Missing leading slash (WRONG)
            ❌ '/documents' - Wrong case (WRONG if actual folder is 'Documents')
            
        Returns:
            {
                "success": true,
                "data": {
                    "value": [
                        {
                            "name": "Documents",
                            "folder": {...},  // This is a folder
                            "size": 0,
                            ...
                        },
                        {
                            "name": "func.py",
                            "file": {...},    // This is a file
                            "size": 1234,
                            ...
                        }
                    ],
                    "@odata.nextLink": "..."  // Use list_files_next_page() if present
                }
            }
            
        Response Indicators:
            - "folder" key present = It's a folder (can navigate into it)
            - "file" key present = It's a file (can read or download it)
            
        Next Steps After Getting Results:
            - Found a folder? Call list_files(path='/FolderName') to browse it
            - Found your file? Use read_file_from_onedrive(path='/folder/file.txt')
            - Too many results? Use list_files_next_page() with @odata.nextLink
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] list_files 被调用，token: {token[:20]}..., path: {path}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            files_data = onedrive.list_my_drive_items(path)
            
            items_count = len(files_data.get('value', []))
            print(f"[MCP] 文件列表获取成功，路径: {path}, 项目数量: {items_count}", file=sys.stderr)
            for item in files_data.get('value', [])[:5]: 
                item_type = "folder" if 'folder' in item else "file"
                print(f"[MCP]   - {item.get('name', 'unknown')} ({item_type})", file=sys.stderr)
            
            if items_count > 5:
                print(f"[MCP]   ... and {items_count - 5} more items", file=sys.stderr)
            
            return {
                "success": True,
                "data": files_data,  
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 文件列表获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 文件列表获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def get_file_or_folder_info(path: str = '/') -> Dict[str, Any]:
        """
        📋 Get detailed metadata about a specific file or folder in YOUR OWN OneDrive.
        
        🎯 USE THIS WHEN: You need detailed information about a specific item.
        
        ⚠️ IMPORTANT: You MUST know the EXACT full path first.
        If you don't know the path, use list_files() to browse and find it.
        
        Args:
            path: The EXACT path to your file or folder (case-sensitive)
                Default: '/' (root info)
                
        Path Examples:
            ✅ '/Documents/report.pdf' - File in Documents folder
            ✅ '/Photos' - Folder in root
            ✅ '/notes.txt' - File in root
            ✅ '/Projects/Python/app.py' - Deeply nested file
            ❌ '/func.py' - DON'T guess! Use list_files() first
            
        Returns:
            Detailed metadata including:
            - name: File/folder name
            - size: Size in bytes
            - createdDateTime: When created
            - lastModifiedDateTime: When last modified
            - folder: Present if it's a folder
            - file: Present if it's a file (includes mimeType)
            - @microsoft.graph.downloadUrl: Direct download link
            
        Use Cases:
            - Check file size before reading
            - Get file modification date
            - Verify file exists at expected path
            - Get download URL
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_file_or_folder_info 被调用，token: {token[:20]}..., path: {path}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            item_info = onedrive.get_my_drive_item(path)
            
            item_name = item_info.get('name', 'N/A')
            item_type = "folder" if 'folder' in item_info else "file"
            item_size = item_info.get('size', 0)
            
            print(f"[MCP] 项目信息获取成功: {item_name} ({item_type}, {item_size} bytes)", file=sys.stderr)
            
            return {
                "success": True,
                "data": item_info,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 项目信息获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 项目信息获取失败: {e}", file=sys.stderr)
            # 404 usually means wrong path
            if "404" in str(e):
                return {
                    "success": False,
                    "data": None,
                    "error": f"File or folder not found at path '{path}'. Did you use list_files() to verify the path first?"
                }
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
    
    @mcp_instance.tool
    async def list_files_next_page(next_link: str) -> Dict[str, Any]:
        """
        📄 Load the next page of results when browsing YOUR OWN OneDrive files.
        
        🎯 USE THIS WHEN: list_files() returns an "@odata.nextLink" field.
        This means there are more items to load beyond the first page.
        
        Args:
            next_link: The exact "@odata.nextLink" URL from the previous list_files() response
            
        Example Workflow:
            1. response = list_files(path='/Documents')
            2. Check if response["data"]["@odata.nextLink"] exists
            3. If yes: next_response = list_files_next_page(next_link=response["data"]["@odata.nextLink"])
            4. Repeat until no more "@odata.nextLink"
            
        Returns:
            Same format as list_files(), may include another @odata.nextLink for more pages
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        if not next_link:
            return {
                "success": False,
                "data": None,
                "error": "nextLink parameter is required. Get it from the @odata.nextLink field in list_files() response."
            }
        
        print(f"[MCP DEBUG] list_files_next_page 被调用，token: {token[:20]}...", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            result = onedrive._make_request('GET', next_link)
            files_data = result.json()
            
            items_count = len(files_data.get('value', []))
            print(f"[MCP] 下一页获取成功，项目数量: {items_count}", file=sys.stderr)
            
            return {
                "success": True,
                "data": files_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 下一页获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 下一页获取失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def read_file_from_onedrive(path: str) -> Dict[str, Any]:
        """
        📖 Read file content from YOUR OWN OneDrive.
        
        🎯 USE THIS WHEN: You want to read the actual content of a file.
        
        ⚠️ CRITICAL REQUIREMENTS:
        
        1. You MUST know the EXACT FULL PATH of the file
        2. You CANNOT guess file paths
        3. You MUST use list_files() FIRST to find the file
        
        🔴 COMMON MISTAKES TO AVOID:
        ❌ DON'T call read_file_from_onedrive(path='/func.py') without checking first
        ❌ DON'T assume files are in root directory
        ❌ DON'T guess folder names or file locations
        
        ✅ CORRECT WORKFLOW:
        
        Step 1: Browse to find the file
            response = list_files(path='/')
            # User says file is "func.py"
            # Check each item in response["data"]["value"]
            # Look for item where item["name"] == "func.py"
            
        Step 2: If not found in root, check common folders
            response = list_files(path='/Documents')
            response = list_files(path='/Desktop')
            response = list_files(path='/Projects')
            
        Step 3: Once found, construct the EXACT path
            # If found in /Documents/Projects/
            path = '/Documents/Projects/func.py'
            
        Step 4: NOW read the file
            content = read_file_from_onedrive(path=path)
        
        Args:
            path: The EXACT FULL PATH to your file (case-sensitive, must include all folders)
            
        ✅ VALID Path Examples (after confirming with list_files):
            '/Documents/readme.txt'           # Text file
            '/Projects/Python/app.py'         # Code file
            '/Reports/sales.xlsx'             # Excel file
            '/Data/users.csv'                 # CSV file
            
        ❌ INVALID Path Examples (DO NOT USE):
            '/func.py'                        # Unless VERIFIED with list_files()
            'func.py'                         # Missing leading slash
            '/FUNC.py'                        # Wrong case
            
        Smart File Handling:
            
            📄 Text files (< 1MB): 
                → Returns full content
                → Supports: .txt, .md, .py, .js, .json, .csv, .xml, etc.
                
            📊 Excel files (.xlsx, < 5MB):
                → Converts to CSV format
                → Returns ALL worksheets
                → Each sheet as separate CSV string
                
            📈 Large text files (≥ 1MB):
                → Returns download link
                
            📁 Large Excel files (≥ 5MB):
                → Returns download link
                
            🖼️ Binary files (images, PDFs, etc.):
                → Returns download link
                
        Supported Formats:
            Text: .txt, .md, .csv, .json, .xml, .yaml, .py, .js, .html, etc.
            Excel: .xlsx (converted to CSV)
            Binary: All other formats (download link provided)
            
        Returns:
            Success response types:
            
            1. Text content (small text files):
            {
                "success": true,
                "data": {
                    "type": "text",
                    "name": "func.py",
                    "size": 1234,
                    "content": "def hello():\n    print('world')",
                    "encoding": "utf-8",
                    "char_count": 156,
                    "line_count": 10
                }
            }
            
            2. Excel content (xlsx files < 5MB):
            {
                "success": true,
                "data": {
                    "type": "xlsx",
                    "name": "sales.xlsx",
                    "size": 51200,
                    "size_mb": 0.05,
                    "sheets": {
                        "Sheet1": "Name,Age,City\nJohn,25,NY\nJane,30,LA\n",
                        "Sheet2": "Product,Price\nApple,1.5\nBanana,0.8\n",
                        "Summary": "Total,Sales\n100,5000\n"
                    },
                    "sheet_count": 3,
                    "total_rows": 25,
                    "total_cells": 150,
                    "sheet_names": ["Sheet1", "Sheet2", "Summary"]
                }
            }
            
            3. Download link (large/binary files):
            {
                "success": true,
                "data": {
                    "type": "binary" or "text_too_large" or "xlsx_too_large",
                    "name": "large_file.pdf",
                    "size": 5242880,
                    "download_link": "https://...",
                    "message": "Please use the download link"
                }
            }
            
            4. File not found:
            {
                "success": false,
                "error": "File not found at path '/func.py'. Did you use list_files() first?"
            }
            
        Working with Excel Results:
            When you get an xlsx file, the data structure is:
            
            # Access all sheets
            sheets = result["data"]["sheets"]
            
            # List all sheet names
            sheet_names = result["data"]["sheet_names"]
            
            # Access specific sheet CSV content
            sheet1_csv = sheets["Sheet1"]
            
            # Parse CSV if needed (it's already in CSV format)
            # Each sheet is a string with rows separated by \n
            # and columns separated by commas
            
        Error Prevention Tips:
            1. Always start with list_files(path='/') to see root contents
            2. Navigate folder by folder using list_files(path='/FolderName')
            3. Verify file exists before trying to read it
            4. Pay attention to case sensitivity in paths
            5. Include the full path with all parent folders
            6. For xlsx files, check size first (must be < 5MB)
            
        Example Complete Workflow:
            # User asks: "Read sales data from report.xlsx"
            
            # Step 1: Find the file
            root = list_files(path='/')
            # Not in root, check Documents
            docs = list_files(path='/Documents')
            # Found it in /Documents/Reports/
            
            # Step 2: Read the Excel file
            result = read_file_from_onedrive(path='/Documents/Reports/report.xlsx')
            
            # Step 3: Process the Excel data
            if result["success"]:
                data = result["data"]
                if data["type"] == "xlsx":
                    # Access each sheet
                    for sheet_name in data["sheet_names"]:
                        csv_content = data["sheets"][sheet_name]
                        print(f"Sheet: {sheet_name}")
                        print(csv_content)
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] read_file_from_onedrive 被调用，token: {token[:20]}..., path: {path}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            result = onedrive.get_file_content(path, max_size_mb=1.0)
            result_type = result.get("type", "unknown")
            result_name = result.get('name', 'unknown')
            
            if result.get("success") or result_type in ["binary", "text_too_large", "decode_error"]:
                print(f"[MCP] 文件读取完成: {result_name}, 类型: {result_type}", file=sys.stderr)
                if result_type == "text":
                    print(f"[MCP] ✅ 文本内容已返回，共 {result.get('char_count', 0)} 字符", file=sys.stderr)
                elif result_type == "text_too_large":
                    print(f"[MCP] ⚠️ 文件过大 ({result.get('size_mb', 0)}MB)，已返回下载链接", file=sys.stderr)
                elif result_type == "binary":
                    print(f"[MCP] ℹ️ 二进制文件，已返回下载链接", file=sys.stderr)
                elif result_type == "decode_error":
                    print(f"[MCP] ⚠️ 无法解码文本，已返回下载链接", file=sys.stderr)
                
                return {
                    "success": True,
                    "data": result,
                    "error": None
                }
            else:
                error_msg = result.get('error', 'Unknown error')
                print(f"[MCP] 文件读取失败: {error_msg}", file=sys.stderr)
                
                if "404" in error_msg or "not found" in error_msg.lower():
                    friendly_error = (
                        f"File not found at path '{path}'. "
                        f"IMPORTANT: You must use list_files() first to find the correct path. "
                        f"Do not guess file locations!"
                    )
                    print(f"[MCP] ❌ {friendly_error}", file=sys.stderr)
                    return {
                        "success": False,
                        "data": None,
                        "error": friendly_error
                    }
                
                return {
                    "success": False,
                    "data": None,
                    "error": error_msg
                }
            
        except MongoDBError as e:
            print(f"[MCP] 文件读取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            error_msg = str(e)
            print(f"[MCP] 文件读取失败: {error_msg}", file=sys.stderr)
            
            if "404" in error_msg:
                friendly_error = (
                    f"File not found at path '{path}'. "
                    f"Did you use list_files() to browse and verify the path first? "
                    f"Remember: paths are case-sensitive and must include all parent folders."
                )
                return {
                    "success": False,
                    "data": None,
                    "error": friendly_error
                }
            
            return {
                "success": False,
                "data": None,
                "error": error_msg
            }

    @mcp_instance.tool
    async def get_file_download_link(path: str) -> Dict[str, Any]:
        """
        🔗 Get a temporary direct download link for any file in YOUR OWN OneDrive.
        
        🎯 USE THIS WHEN:
        - You need to download a binary file (image, PDF, video, etc.)
        - You want a direct URL to access the file
        - The file is too large to read directly
        
        ⚠️ IMPORTANT: You MUST know the EXACT full path first.
        If you don't know the path, use list_files() to browse and find it.
        
        Args:
            path: The EXACT path to your file (case-sensitive)
            
        Path Examples:
            ✅ '/Documents/report.pdf' - PDF in Documents
            ✅ '/Pictures/vacation.jpg' - Image in Pictures
            ✅ '/Videos/tutorial.mp4' - Video file
            ✅ '/Archive/backup.zip' - Archive file
            ❌ '/report.pdf' - Don't guess! Use list_files() first
            
        Workflow:
            1. Use list_files() to find your file
            2. Note the exact path from the results
            3. Call get_file_download_link(path=exact_path)
            4. Use the returned download_link URL
            
        Returns:
            {
                "success": true,
                "data": {
                    "name": "report.pdf",
                    "size": 1234567,
                    "size_mb": 1.18,
                    "download_link": "https://..."  // Temporary URL, use soon
                }
            }
            
        Note: The download link is temporary and expires after a short time.
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        print(f"[MCP DEBUG] get_file_download_link 被调用，token: {token[:20]}..., path: {path}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            result = onedrive.get_download_link(path)
            
            if result.get("success"):
                result_name = result.get('name', 'unknown')
                result_size = result.get('size_mb', 0)
                print(f"[MCP] ✅ 下载链接获取成功: {result_name} ({result_size}MB)", file=sys.stderr)
                return {
                    "success": True,
                    "data": result,
                    "error": None
                }
            else:
                error_msg = result.get('error', 'Unknown error')
                print(f"[MCP] 下载链接获取失败: {error_msg}", file=sys.stderr)
                
                if "404" in error_msg or "not found" in error_msg.lower():
                    friendly_error = (
                        f"File not found at path '{path}'. "
                        f"Use list_files() first to find the correct path."
                    )
                    return {
                        "success": False,
                        "data": None,
                        "error": friendly_error
                    }
                
                return {
                    "success": False,
                    "data": None,
                    "error": error_msg
                }
            
        except MongoDBError as e:
            print(f"[MCP] 下载链接获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            error_msg = str(e)
            print(f"[MCP] 下载链接获取失败: {error_msg}", file=sys.stderr)
            
            if "404" in error_msg:
                friendly_error = (
                    f"File not found at path '{path}'. "
                    f"Did you verify the path with list_files() first?"
                )
                return {
                    "success": False,
                    "data": None,
                    "error": friendly_error
                }
            
            return {
                "success": False,
                "data": None,
                "error": error_msg
            }
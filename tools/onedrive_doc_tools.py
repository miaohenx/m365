"""
OneDrive 工具模块
纯粹的工具函数，依赖服务层
"""
from typing import Any, Dict
from datetime import datetime
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_doc_tools(mcp_instance):
    """注册 OneDrive 工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def download_doc_from_onedrive(url: str) -> Dict[str, Any]:
        """
        Download documents from onedrive
        
        url: the url of one drive file/folder
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] download_doc_from_onedrive 被调用，token: {token[:20]}..., url: {url}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            
            # 获取驱动器项目信息
            onedrive.get_driveitem(url)
            
            # 下载文件
            filename = '/' + url.split('/')[-1]
            onedrive.downloadfile(filename)
            
            print(f"[MCP] 文件下载成功: {filename}", file=sys.stderr)
            return {
                "success": True,
                "data": {"message": f"Download successfully: {filename}"},
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 文件下载失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 文件下载失败: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }

    @mcp_instance.tool
    async def list_doc_from_onedrive(url: str) -> Dict[str, Any]:
        """
        List documents from one drive
        
        url: the url of one drive folder
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
            
            # 获取驱动器项目信息
            onedrive.get_driveitem(url)
            
            # 列出目录内容
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
        Get file info from one drive
        
        url: the url of one drive file/folder
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
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取驱动器项目信息
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
        Search files in OneDrive
        
        query: the search query
        folder_url: the url of the folder to search in
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
            
            # 设置搜索的根目录
            onedrive.get_driveitem(folder_url)
            
            # 执行搜索
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
    @mcp_instance.tool
    async def list_files(path: str = '/') -> Dict[str, Any]:
        """
        List files and folders from user's OneDrive
        
        path: the path of the folder to list, default is root '/'
            examples: '/', '/Documents', '/Pictures/Vacation'
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
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 列出文件
            files_data = onedrive.list_my_drive_items(path)
            
            print(f"[MCP] 文件列表获取成功，项目数量: {len(files_data.get('value', []))}", file=sys.stderr)
            
            # 🔥 返回完整数据，包括 nextLink
            return {
                "success": True,
                "data": files_data,  # 包含 value 和 @odata.nextLink
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
        Get information about a file or folder from user's OneDrive
        
        path: the path of the file or folder, default is root '/'
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
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 获取项目信息
            item_info = onedrive.get_my_drive_item(path)
            
            print(f"[MCP] 项目信息获取成功: {item_info.get('name', 'N/A')}", file=sys.stderr)
            
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
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
    
    @mcp_instance.tool
    async def list_files_next_page(next_link: str) -> Dict[str, Any]:
        """
        Load next page of files using the nextLink
        
        next_link: the @odata.nextLink from previous response
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
                "error": "nextLink is required"
            }
        
        print(f"[MCP DEBUG] list_files_next_page 被调用，token: {token[:20]}...", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 🔥 直接请求 nextLink
            result = onedrive._make_request('GET', next_link)
            files_data = result.json()
            
            print(f"[MCP] 下一页获取成功，项目数量: {len(files_data.get('value', []))}", file=sys.stderr)
            
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
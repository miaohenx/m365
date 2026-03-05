"""
OneDrive Teams 工具模块
纯粹的工具函数，依赖服务层
"""
from typing import Any, Dict, List, Optional
from datetime import datetime, timedelta, timezone
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_teams_tools(mcp_instance):
    """注册 OneDrive Teams 工具到 MCP 实例"""
    
    @mcp_instance.tool
    async def read_team_chats(
        team_id: Optional[str] = None, 
        chat_type: str = "all", 
        days_filter: int = 5, 
        max_results: Optional[int] = 50,  # 修改默认值
        next_link: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        List and retrieve chat information within a specific Microsoft Team, including both channel conversations and private chats.
        If team_id is not provided, returns all chats the user has access to.

        Args:
            team_id (str, optional): Unique identifier of the target team (if None, returns all chats)
            chat_type (str): Type of chats to retrieve ("all", "oneOnOne", "group", "meeting") - default "all"
            days_filter (int): Number of days to look back for chat activity - default 30 (applied client-side)
            max_results (int, optional): Maximum number of chats to return - default 50
            next_link (str, optional): Direct URL for next page of results (if provided, other filters are ignored)

        Returns:
            dict: Dictionary containing success status, chats data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_team_chats 被调用，token: {token[:20]}..., chat_type: {chat_type}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 如果提供了next_link，直接使用它
            if next_link:
                response = onedrive._make_request('GET', next_link)
                response_data = response.json()
                chats_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            else:
                # 准备API参数（注意：/me/chats 不支持 $filter）
                params = {
                    '$expand': 'members',  # 展开成员信息
                }
                
                # 设置分页大小
                if max_results:
                    params['$top'] = min(max_results, 50)  # API最大支持50
                
                # 获取聊天数据
                api_url = f"{onedrive.BASE_URL}/me/chats"
                response = onedrive._make_request('GET', api_url, params=params)
                response_data = response.json()
                chats_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            
            # 计算时间过滤阈值（客户端过滤）
            now = datetime.now(timezone.utc)
            target_time = now - timedelta(days=days_filter)
            
            # 结构化并过滤数据
            chats_list = []
            for chat in chats_data:
                # 解析最后更新时间
                last_updated_str = chat.get('lastUpdatedDateTime', '')
                try:
                    # 处理多种时间格式
                    if last_updated_str:
                        # 移除微秒和Z
                        last_updated_str = last_updated_str.split('.')[0].rstrip('Z')
                        last_updated = datetime.strptime(last_updated_str, '%Y-%m-%dT%H:%M:%S')
                        last_updated = last_updated.replace(tzinfo=timezone.utc)
                    else:
                        last_updated = datetime.min.replace(tzinfo=timezone.utc)
                except Exception as e:
                    print(f"[MCP] 时间解析错误: {last_updated_str}, {e}", file=sys.stderr)
                    last_updated = datetime.min.replace(tzinfo=timezone.utc)
                
                # 应用时间过滤（客户端）
                if last_updated < target_time:
                    continue
                
                # 应用聊天类型过滤
                current_chat_type = chat.get('chatType', '')
                if chat_type != "all" and current_chat_type != chat_type:
                    continue
                
                # 构建返回数据
                chat_data = {
                    'id': chat.get('id', ''),
                    'topic': chat.get('topic'),
                    'chatType': current_chat_type,
                    'createdDateTime': chat.get('createdDateTime', ''),
                    'lastUpdatedDateTime': chat.get('lastUpdatedDateTime', ''),
                    'webUrl': chat.get('webUrl', ''),
                    'viewpoint': chat.get('viewpoint', {}),
                    'isHiddenForAllMembers': chat.get('isHiddenForAllMembers', False),
                    'onlineMeetingInfo': chat.get('onlineMeetingInfo'),
                    'members': chat.get('members', []),
                    'installedApps': chat.get('installedApps', [])
                }
                chats_list.append(chat_data)
            
            # 按最后更新时间排序
            chats_list.sort(key=lambda x: x.get('lastUpdatedDateTime', ''), reverse=True)
            
            print(f"[MCP] 团队聊天获取成功，总数: {len(chats_list)}, 原始数据: {len(chats_data)}", file=sys.stderr)
            return {
                "success": True,
                "data": {
                    'chats': chats_list,
                    'next_link': next_page_link,
                    'total_count': len(chats_list),
                    'filtered_count': len(chats_data) - len(chats_list)  # 被过滤掉的数量
                },
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 团队聊天获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 团队聊天获取失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'authentication' in error_msg or 'unauthorized' in error_msg:
                error_detail = f"Authentication failed: {str(e)}. Please check your token."
            elif 'network' in error_msg or 'connection' in error_msg:
                error_detail = f"Network connectivity issue: {str(e)}. Please check your internet connection."
            elif 'not found' in error_msg:
                error_detail = f"No chats found for user: {str(e)}."
            elif 'rate limit' in error_msg or 'throttle' in error_msg:
                error_detail = f"API rate limit exceeded: {str(e)}. Please try again later."
            else:
                error_detail = f"Failed to retrieve chats: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def read_team_chat_messages(
        chat_id: str, 
        days_filter: int = 7, 
        message_limit: int = 50, 
        include_system_messages: bool = False, 
        next_link: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Retrieve messages from a specific chat within a Microsoft Team, with filtering and pagination support.

        Args:
            chat_id (str): Unique identifier of the target chat
            days_filter (int): Number of days to look back for messages - default 7 (applied client-side)
            message_limit (int): Maximum number of messages to retrieve - default 50
            include_system_messages (bool): Whether to include system messages - default False
            next_link (str, optional): Direct URL for next page of results (if provided, other filters are ignored)

        Returns:
            dict: Dictionary containing success status, messages data, and error information
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
            
        print(f"[MCP DEBUG] read_team_chat_messages 被调用，chat_id: {chat_id}", file=sys.stderr)
        
        try:
            # 创建并认证OneDrive服务
            onedrive = await create_onedrive_service(token)
            
            # 如果提供了next_link，直接使用它
            if next_link:
                response = onedrive._make_request('GET', next_link)
                response_data = response.json()
                messages_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            else:
                # 准备API参数
                params = {
                    '$top': min(message_limit, 50),  # API最大支持50
                    '$orderby': 'createdDateTime desc'
                }
                
                # 构建API端点
                api_url = f"{onedrive.BASE_URL}/chats/{chat_id}/messages"
                response = onedrive._make_request('GET', api_url, params=params)
                response_data = response.json()
                messages_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            
            # 计算时间过滤阈值
            now = datetime.now(timezone.utc)
            target_time = now - timedelta(days=days_filter)
            
            # 结构化并过滤返回数据
            messages_list = []
            for message in messages_data:
                # 解析创建时间
                created_str = message.get('createdDateTime', '')
                try:
                    if created_str:
                        created_str = created_str.split('.')[0].rstrip('Z')
                        created_time = datetime.strptime(created_str, '%Y-%m-%dT%H:%M:%S')
                        created_time = created_time.replace(tzinfo=timezone.utc)
                    else:
                        created_time = datetime.min.replace(tzinfo=timezone.utc)
                except Exception as e:
                    print(f"[MCP] 消息时间解析错误: {created_str}, {e}", file=sys.stderr)
                    created_time = datetime.min.replace(tzinfo=timezone.utc)
                
                # 应用时间过滤
                if created_time < target_time:
                    continue
                
                # 过滤系统消息
                message_type = message.get('messageType', '')
                if not include_system_messages and message_type in ['systemEventMessage', 'chatEvent', 'unknownFutureValue']:
                    continue
                
                message_data = {
                    'id': message.get('id', ''),
                    'messageType': message_type,
                    'createdDateTime': message.get('createdDateTime', ''),
                    'lastModifiedDateTime': message.get('lastModifiedDateTime', ''),
                    'lastEditedDateTime': message.get('lastEditedDateTime'),
                    'deletedDateTime': message.get('deletedDateTime'),
                    'subject': message.get('subject'),
                    'summary': message.get('summary'),
                    'importance': message.get('importance', 'normal'),
                    'locale': message.get('locale'),
                    'from': message.get('from', {}),
                    'body': message.get('body', {}),
                    'attachments': message.get('attachments', []),
                    'mentions': message.get('mentions', []),
                    'reactions': message.get('reactions', []),
                    'replies': message.get('replies'),
                    'webUrl': message.get('webUrl'),
                    'channelIdentity': message.get('channelIdentity'),
                    'policyViolation': message.get('policyViolation'),
                    'eventDetail': message.get('eventDetail', '')
                }
                messages_list.append(message_data)
            
            print(f"[MCP] 聊天消息获取成功，数量: {len(messages_list)}, 原始: {len(messages_data)}", file=sys.stderr)
            return {
                "success": True,
                "data": {
                    'messages': messages_list,
                    'next_link': next_page_link,
                    'total_count': len(messages_list),
                    'filtered_count': len(messages_data) - len(messages_list)
                },
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 聊天消息获取失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 聊天消息获取失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            
            # 处理各种错误情况
            error_msg = str(e).lower()
            if 'authentication' in error_msg or 'unauthorized' in error_msg:
                error_detail = f"Authentication failed: {str(e)}. Please check your token."
            elif 'network' in error_msg or 'connection' in error_msg:
                error_detail = f"Network connectivity issue: {str(e)}. Please check your internet connection."
            elif 'not found' in error_msg:
                error_detail = f"Chat not found or user has no access: {str(e)}."
            elif 'rate limit' in error_msg or 'throttle' in error_msg:
                error_detail = f"API rate limit exceeded: {str(e)}. Please try again later."
            else:
                error_detail = f"Failed to retrieve chat messages: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }
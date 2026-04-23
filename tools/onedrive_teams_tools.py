"""
OneDrive Teams Tools Module
Pure utility functions that depend on service layer
"""
from typing import Any, Dict, List, Optional
from datetime import datetime, timedelta, timezone
import sys

from services.mongo_service import MongoTokenService
from services.onedrive_service import create_onedrive_service
from exceptions import MongoDBError
from utils import get_token_from_context

def register_teams_tools(mcp_instance):
    """Register OneDrive Teams tools to MCP instance"""
    
    @mcp_instance.tool
    async def read_team_chats(
        team_id: Optional[str] = None, 
        chat_type: str = "all", 
        days_filter: int = 5, 
        max_results: Optional[int] = 50,  # Modified default value
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
            onedrive = await create_onedrive_service(token)
            
            if next_link:
                response = onedrive._make_request('GET', next_link)
                response_data = response.json()
                chats_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            else:
                params = {
                    '$expand': 'members',  
                }
                
                if max_results:
                    params['$top'] = min(max_results, 50)  
                
                api_url = f"{onedrive.BASE_URL}/me/chats"
                response = onedrive._make_request('GET', api_url, params=params)
                response_data = response.json()
                chats_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')

            now = datetime.now(timezone.utc)
            target_time = now - timedelta(days=days_filter)
            
            chats_list = []
            for chat in chats_data:
                last_updated_str = chat.get('lastUpdatedDateTime', '')
                try:
                    if last_updated_str:
                        last_updated_str = last_updated_str.split('.')[0].rstrip('Z')
                        last_updated = datetime.strptime(last_updated_str, '%Y-%m-%dT%H:%M:%S')
                        last_updated = last_updated.replace(tzinfo=timezone.utc)
                    else:
                        last_updated = datetime.min.replace(tzinfo=timezone.utc)
                except Exception as e:
                    print(f"[MCP] 时间解析错误: {last_updated_str}, {e}", file=sys.stderr)
                    last_updated = datetime.min.replace(tzinfo=timezone.utc)
                
                if last_updated < target_time:
                    continue
                
                current_chat_type = chat.get('chatType', '')
                if chat_type != "all" and current_chat_type != chat_type:
                    continue
                
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
            
            chats_list.sort(key=lambda x: x.get('lastUpdatedDateTime', ''), reverse=True)
            
            print(f"[MCP] 团队聊天获取成功，总数: {len(chats_list)}, 原始数据: {len(chats_data)}", file=sys.stderr)
            return {
                "success": True,
                "data": {
                    'chats': chats_list,
                    'next_link': next_page_link,
                    'total_count': len(chats_list),
                    'filtered_count': len(chats_data) - len(chats_list)
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
            onedrive = await create_onedrive_service(token)
            
            if next_link:
                response = onedrive._make_request('GET', next_link)
                response_data = response.json()
                messages_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            else:
                params = {
                    '$top': min(message_limit, 50), 
                    '$orderby': 'createdDateTime desc'
                }
                

                api_url = f"{onedrive.BASE_URL}/chats/{chat_id}/messages"
                response = onedrive._make_request('GET', api_url, params=params)
                response_data = response.json()
                messages_data = response_data.get('value', [])
                next_page_link = response_data.get('@odata.nextLink')
            
            now = datetime.now(timezone.utc)
            target_time = now - timedelta(days=days_filter)
            
            messages_list = []
            for message in messages_data:
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
                
                if created_time < target_time:
                    continue
                
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

    @mcp_instance.tool
    async def send_team_chat_message(
        chat_id: str,
        content: str,
        content_type: str = "text"
    ) -> Dict[str, Any]:
        """
        Send a message to a specific Teams chat.
        
        ⚠️ IMPORTANT: You MUST first call read_team_chats() to get a valid chat_id.
        DO NOT guess or fabricate chat IDs.
        
        Args:
            chat_id (str): **REQUIRED** - The exact 'id' value from read_team_chats() response.
                Example: "19:meeting_MjdhNjM4YzUtYzExZi00OTY4LTkzYWUtNTVlNmZmMDhkNGU2@thread.v2"
                
                ⚠️ HOW TO GET THIS ID:
                1. Call read_team_chats() first
                2. Find your target chat in the 'chats' array
                3. Copy the EXACT 'id' field value
                4. Pass it to this function
                
            content (str): **REQUIRED** - The message content to send.
                Cannot be empty or whitespace only.
                
            content_type (str): Content format - "text" (default) or "html"
                - "text": Plain text message
                - "html": HTML formatted message (supports rich formatting)
        
        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (dict): Sent message data including:
                    - id (str): Message ID
                    - createdDateTime (str): When message was sent
                    - from (dict): Sender information
                    - body (dict): Message body with content and contentType
                - error (str|None): Error message if failed
                
        Example Usage:
            # Step 1: Get chat ID
            chats = read_team_chats()
            target_chat_id = chats['data']['chats'][0]['id']
            
            # Step 2: Send message
            result = send_team_chat_message(
                chat_id=target_chat_id,
                content="Hello from MCP!",
                content_type="text"
            )
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        if not content or not content.strip():
            return {
                "success": False,
                "data": None,
                "error": "Message content cannot be empty"
            }
        
        if content_type not in ["text", "html"]:
            return {
                "success": False,
                "data": None,
                "error": f"Invalid content_type '{content_type}'. Must be 'text' or 'html'"
            }
            
        print(f"[MCP DEBUG] send_team_chat_message 被调用，chat_id: {chat_id}, content_type: {content_type}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            
            message_data = onedrive.send_chat_message(chat_id, content, content_type)
            
            print(f"[MCP] 聊天消息发送成功 - 消息ID: {message_data.get('id', 'Unknown')}", file=sys.stderr)
            return {
                "success": True,
                "data": message_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 消息发送失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except Exception as e:
            print(f"[MCP] 消息发送失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            
            error_msg = str(e).lower()
            if 'authentication' in error_msg or 'unauthorized' in error_msg:
                error_detail = f"Authentication failed: {str(e)}. Please check your token."
            elif 'not found' in error_msg or '404' in error_msg:
                error_detail = f"Chat not found: Chat ID '{chat_id}' does not exist or you don't have access. Please call read_team_chats() first to get valid chat IDs."
            elif 'permission' in error_msg or 'forbidden' in error_msg or '403' in error_msg:
                error_detail = f"Permission denied: You don't have permission to send messages to this chat. Error: {str(e)}"
            elif 'rate limit' in error_msg or 'throttle' in error_msg:
                error_detail = f"API rate limit exceeded: {str(e)}. Please try again later."
            else:
                error_detail = f"Failed to send message: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def create_team_chat(
        chat_type: str,
        members: List[Dict[str, Any]],
        topic: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create a new Teams chat (one-on-one or group chat).
        
        ⚠️ IMPORTANT - MUST READ BEFORE CALLING:
        Microsoft Graph API REQUIRES the caller (current user) to be included in the members list.
        
        MANDATORY STEPS:
        1. First, identify the current user's ID (from previous API calls, token context, or by calling
           read_team_chats which returns member information including the current user)
        2. Always include the current user in the members list with roles ["owner"]
        3. Then add the other target members
        
        If the current user is NOT in the members list, Microsoft will return:
        400 BadRequest: "The caller must be one of the members specified in request body."
        
        Args:
            chat_type (str): **REQUIRED** - Type of chat to create:
                - "oneOnOne": Direct chat between two people
                - "group": Group chat with multiple people
                
            members (List[Dict]): **REQUIRED** - List of chat members.
                Each member must be a dictionary with:
                - "user_id" (str): User's email address or Azure AD user ID
                - "roles" (List[str], optional): User roles, default ["owner"]
                  Options: ["owner"] or ["guest"]
                  
                ⚠️ MUST include the current caller's user_id in this list!
                  
                Example:
                [
                    {"user_id": "current_user@example.com", "roles": ["owner"]},  # ← MUST include self
                    {"user_id": "other_user@example.com", "roles": ["owner"]}
                ]
                
            topic (str, optional): Chat topic/name.
                - **REQUIRED** for group chats
                - Optional for oneOnOne chats (usually not used)
        
        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (dict): Created chat data including:
                    - id (str): New chat ID (use this for sending messages)
                    - chatType (str): Type of chat created
                    - topic (str): Chat topic (if applicable)
                    - createdDateTime (str): When chat was created
                    - members (list): List of chat members
                - error (str|None): Error message if failed
                
        Example Usage:
            # Create a one-on-one chat (current user MUST be included)
            result = create_team_chat(
                chat_type="oneOnOne",
                members=[
                    {"user_id": "current_user@example.com"},   # ← caller/self
                    {"user_id": "other_user@example.com"}
                ]
            )
            
            # Create a group chat (current user MUST be included)
            result = create_team_chat(
                chat_type="group",
                members=[
                    {"user_id": "current_user@example.com", "roles": ["owner"]},  # ← caller/self
                    {"user_id": "user2@example.com", "roles": ["owner"]},
                    {"user_id": "user3@example.com", "roles": ["owner"]}
                ],
                topic="Project Discussion"
            )
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        if chat_type not in ["oneOnOne", "group"]:
            return {
                "success": False,
                "data": None,
                "error": f"Invalid chat_type '{chat_type}'. Must be 'oneOnOne' or 'group'"
            }
        
        if not members or len(members) < 1:
            return {
                "success": False,
                "data": None,
                "error": "At least one member is required"
            }
        
        if chat_type == "group" and not topic:
            return {
                "success": False,
                "data": None,
                "error": "Topic is required for group chats"
            }
            
        print(f"[MCP DEBUG] create_team_chat 被调用，chat_type: {chat_type}, members: {len(members)}", file=sys.stderr)
        if topic:
            print(f"[MCP DEBUG] 主题: {topic}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            
            chat_data = onedrive.create_chat(chat_type, members, topic)
            
            print(f"[MCP] 聊天创建成功 - 聊天ID: {chat_data.get('id', 'Unknown')}", file=sys.stderr)
            return {
                "success": True,
                "data": chat_data,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 聊天创建失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except ValueError as e:
            print(f"[MCP] 聊天创建失败 - 参数错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
        except Exception as e:
            print(f"[MCP] 聊天创建失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            
            error_msg = str(e).lower()
            if 'authentication' in error_msg or 'unauthorized' in error_msg:
                error_detail = f"Authentication failed: {str(e)}. Please check your token."
            elif 'not found' in error_msg or '404' in error_msg:
                error_detail = f"User not found: One or more user IDs are invalid. Error: {str(e)}"
            elif 'permission' in error_msg or 'forbidden' in error_msg or '403' in error_msg:
                error_detail = f"Permission denied: You don't have permission to create chats. Error: {str(e)}"
            elif 'rate limit' in error_msg or 'throttle' in error_msg:
                error_detail = f"API rate limit exceeded: {str(e)}. Please try again later."
            elif 'caller must be one of the members' in error_msg or 'badrequest' in error_msg:
                error_detail = f"Current user is not in members list: You must include yourself (the caller) in the members list. Error: {str(e)}"
            else:
                error_detail = f"Failed to create chat: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }

    @mcp_instance.tool
    async def update_team_chat_topic(
        chat_id: str,
        topic: str
    ) -> Dict[str, Any]:
        """
        Update the topic/name of a Teams group chat.
        
        ⚠️ IMPORTANT: 
        - This only works for GROUP chats, not one-on-one chats
        - You MUST first call read_team_chats() to get a valid chat_id
        - DO NOT guess or fabricate chat IDs
        
        Args:
            chat_id (str): **REQUIRED** - The exact 'id' value from read_team_chats() response.
                Must be a group chat ID (chatType: "group")
                
                ⚠️ HOW TO GET THIS ID:
                1. Call read_team_chats() first
                2. Find your target GROUP chat (check chatType field)
                3. Copy the EXACT 'id' field value
                4. Pass it to this function
                
            topic (str): **REQUIRED** - New topic/name for the chat.
                Cannot be empty or whitespace only.
        
        Returns:
            dict: Dictionary containing:
                - success (bool): Operation status
                - data (dict): Update result including:
                    - status (str): "updated"
                    - chat_id (str): The chat ID that was updated
                    - new_topic (str): The new topic that was set
                - error (str|None): Error message if failed
                
        Example Usage:
            # Step 1: Get group chat ID
            chats = read_team_chats(chat_type="group")
            target_chat_id = chats['data']['chats'][0]['id']
            
            # Step 2: Update topic
            result = update_team_chat_topic(
                chat_id=target_chat_id,
                topic="New Project Name - Q2 2024"
            )
        """
        token = get_token_from_context()
        
        if not token:
            return {
                "success": False,
                "data": None,
                "error": "No Authorization token found in request headers"
            }
        
        if not topic or not topic.strip():
            return {
                "success": False,
                "data": None,
                "error": "Topic cannot be empty"
            }
            
        print(f"[MCP DEBUG] update_team_chat_topic 被调用，chat_id: {chat_id}, new_topic: {topic}", file=sys.stderr)
        
        try:
            onedrive = await create_onedrive_service(token)
            
            update_result = onedrive.update_chat_topic(chat_id, topic)
            
            print(f"[MCP] 聊天主题更新成功", file=sys.stderr)
            return {
                "success": True,
                "data": update_result,
                "error": None
            }
            
        except MongoDBError as e:
            print(f"[MCP] 主题更新失败 - MongoDB错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": f"Token authentication failed: {str(e)}"
            }
        except ValueError as e:
            print(f"[MCP] 主题更新失败 - 参数错误: {e}", file=sys.stderr)
            return {
                "success": False,
                "data": None,
                "error": str(e)
            }
        except Exception as e:
            print(f"[MCP] 主题更新失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            
            error_msg = str(e).lower()
            if 'authentication' in error_msg or 'unauthorized' in error_msg:
                error_detail = f"Authentication failed: {str(e)}. Please check your token."
            elif 'not found' in error_msg or '404' in error_msg:
                error_detail = f"Chat not found: Chat ID '{chat_id}' does not exist or you don't have access. Please call read_team_chats() first to get valid chat IDs."
            elif 'permission' in error_msg or 'forbidden' in error_msg or '403' in error_msg:
                error_detail = f"Permission denied: You don't have permission to update this chat, or this is a one-on-one chat (topics can only be updated for group chats). Error: {str(e)}"
            elif 'rate limit' in error_msg or 'throttle' in error_msg:
                error_detail = f"API rate limit exceeded: {str(e)}. Please try again later."
            else:
                error_detail = f"Failed to update chat topic: {str(e)}"
            
            return {
                "success": False,
                "data": None,
                "error": error_detail
            }
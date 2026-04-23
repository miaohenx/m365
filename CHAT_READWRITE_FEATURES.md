# Chat.ReadWrite 权限新增功能文档

## 📋 概述

本文档记录了为 **Chat.ReadWrite** 权限新增的3个写入功能工具。这些工具允许用户通过MCP服务器与Microsoft Teams聊天进行交互。

**权限说明**:
- ✅ **Chat.ReadWrite** - 新增写入功能(发送消息、创建聊天、更新主题)
- ✅ **Files.Read.All** - 已有完整读取功能
- ✅ **Notes.Read.All** - 已有完整读取功能

---

## 🆕 新增工具列表

### 1️⃣ send_team_chat_message
**发送Teams聊天消息**

**位置**: [`tools/onedrive_teams_tools.py:298`](tools/onedrive_teams_tools.py:298)

**功能**: 向指定的Teams聊天发送文本或HTML格式的消息

**参数**:
- `chat_id` (str, 必需) - 聊天ID,必须从 `read_team_chats()` 获取
- `content` (str, 必需) - 消息内容,不能为空
- `content_type` (str, 可选) - 内容格式: "text"(默认) 或 "html"

**返回值**:
```json
{
  "success": true,
  "data": {
    "id": "消息ID",
    "createdDateTime": "2024-01-01T12:00:00Z",
    "from": { "user": {...} },
    "body": {
      "contentType": "text",
      "content": "Hello from MCP!"
    }
  },
  "error": null
}
```

**使用示例**:
```python
# 步骤1: 获取聊天ID
chats = read_team_chats()
target_chat_id = chats['data']['chats'][0]['id']

# 步骤2: 发送消息
result = send_team_chat_message(
    chat_id=target_chat_id,
    content="Hello from MCP!",
    content_type="text"
)
```

**错误处理**:
- ❌ 404: 聊天不存在或无权访问
- ❌ 403: 无权限发送消息
- ❌ 401: Token认证失败
- ❌ 429: API速率限制

---

### 2️⃣ create_team_chat
**创建新的Teams聊天**

**位置**: [`tools/onedrive_teams_tools.py:398`](tools/onedrive_teams_tools.py:398)

**功能**: 创建一对一聊天或群聊

**参数**:
- `chat_type` (str, 必需) - 聊天类型:
  - `"oneOnOne"` - 一对一聊天
  - `"group"` - 群聊
- `members` (List[Dict], 必需) - 成员列表,每个成员包含:
  - `user_id` (str) - 用户邮箱或Azure AD ID
  - `roles` (List[str], 可选) - 角色列表,默认 `["owner"]`
- `topic` (str, 可选) - 聊天主题(群聊必需)

**返回值**:
```json
{
  "success": true,
  "data": {
    "id": "新聊天ID",
    "chatType": "group",
    "topic": "Project Discussion",
    "createdDateTime": "2024-01-01T12:00:00Z",
    "members": [...]
  },
  "error": null
}
```

**使用示例**:
```python
# 创建一对一聊天
result = create_team_chat(
    chat_type="oneOnOne",
    members=[
        {"user_id": "user1@example.com"},
        {"user_id": "user2@example.com"}
    ]
)

# 创建群聊
result = create_team_chat(
    chat_type="group",
    members=[
        {"user_id": "user1@example.com", "roles": ["owner"]},
        {"user_id": "user2@example.com", "roles": ["owner"]},
        {"user_id": "user3@example.com", "roles": ["owner"]}
    ],
    topic="Project Discussion"
)
```

**错误处理**:
- ❌ 404: 用户ID无效
- ❌ 403: 无权限创建聊天
- ❌ 400: 参数错误(如群聊缺少topic)
- ❌ 401: Token认证失败

---

### 3️⃣ update_team_chat_topic
**更新群聊主题**

**位置**: [`tools/onedrive_teams_tools.py:498`](tools/onedrive_teams_tools.py:498)

**功能**: 更新群聊的主题/名称(仅适用于群聊)

**参数**:
- `chat_id` (str, 必需) - 群聊ID,必须从 `read_team_chats()` 获取
- `topic` (str, 必需) - 新主题,不能为空

**返回值**:
```json
{
  "success": true,
  "data": {
    "status": "updated",
    "chat_id": "聊天ID",
    "new_topic": "New Project Name - Q2 2024"
  },
  "error": null
}
```

**使用示例**:
```python
# 步骤1: 获取群聊ID
chats = read_team_chats(chat_type="group")
target_chat_id = chats['data']['chats'][0]['id']

# 步骤2: 更新主题
result = update_team_chat_topic(
    chat_id=target_chat_id,
    topic="New Project Name - Q2 2024"
)
```

**错误处理**:
- ❌ 404: 聊天不存在或无权访问
- ❌ 403: 无权限更新或尝试更新一对一聊天
- ❌ 401: Token认证失败
- ❌ 400: 主题为空

---

## 🔧 服务层实现

### 新增服务方法

**位置**: [`services/onedrive_service.py:883`](services/onedrive_service.py:883)

#### 1. `send_chat_message(chat_id, content, content_type)`
- **API端点**: `POST /chats/{chat_id}/messages`
- **功能**: 发送聊天消息
- **日志**: 详细的请求/响应日志

#### 2. `create_chat(chat_type, members, topic)`
- **API端点**: `POST /chats`
- **功能**: 创建新聊天
- **成员格式**: 自动转换为Graph API所需的格式

#### 3. `update_chat_topic(chat_id, topic)`
- **API端点**: `PATCH /chats/{chat_id}`
- **功能**: 更新聊天主题
- **限制**: 仅适用于群聊

---

## 📊 工作流程图

### 发送消息工作流
```
1. read_team_chats()          → 获取聊天列表
   ↓
2. 选择目标聊天ID
   ↓
3. send_team_chat_message()   → 发送消息
   ↓
4. 返回消息数据(包含消息ID)
```

### 创建聊天工作流
```
1. 准备成员列表(user_id + roles)
   ↓
2. create_team_chat()         → 创建聊天
   ↓
3. 返回新聊天数据(包含chat_id)
   ↓
4. 使用chat_id发送消息
```

### 更新主题工作流
```
1. read_team_chats(chat_type="group")  → 获取群聊列表
   ↓
2. 选择目标群聊ID
   ↓
3. update_team_chat_topic()            → 更新主题
   ↓
4. 返回更新状态
```

---

## ⚠️ 重要注意事项

### 1. ID验证
- ❌ **禁止猜测或伪造ID**
- ✅ **必须从API响应中获取真实ID**
- ✅ **使用 `read_team_chats()` 获取有效的chat_id**

### 2. 权限要求
- 需要 **Chat.ReadWrite** 权限
- 用户必须是聊天成员才能发送消息
- 创建聊天需要所有成员的有效ID

### 3. 聊天类型限制
- `update_team_chat_topic()` **仅适用于群聊**
- 一对一聊天无法更新主题
- 群聊创建时**必须提供topic**

### 4. 错误处理
- 所有工具都返回统一的 `{success, data, error}` 格式
- 详细的错误信息帮助诊断问题
- 自动区分认证、权限、网络等错误类型

---

## 🧪 测试建议

### 测试场景1: 发送消息
```python
# 1. 获取聊天列表
chats = read_team_chats(top=5)

# 2. 选择第一个聊天
if chats['success'] and chats['data']['chats']:
    chat_id = chats['data']['chats'][0]['id']
    
    # 3. 发送测试消息
    result = send_team_chat_message(
        chat_id=chat_id,
        content="Test message from MCP",
        content_type="text"
    )
    
    print(f"Message sent: {result['success']}")
```

### 测试场景2: 创建群聊并发送消息
```python
# 1. 创建群聊
chat_result = create_team_chat(
    chat_type="group",
    members=[
        {"user_id": "user1@example.com"},
        {"user_id": "user2@example.com"}
    ],
    topic="Test Group Chat"
)

if chat_result['success']:
    new_chat_id = chat_result['data']['id']
    
    # 2. 向新聊天发送消息
    msg_result = send_team_chat_message(
        chat_id=new_chat_id,
        content="Welcome to the new chat!"
    )
    
    print(f"Chat created and message sent: {msg_result['success']}")
```

### 测试场景3: 更新群聊主题
```python
# 1. 获取群聊
chats = read_team_chats(chat_type="group", top=5)

if chats['success'] and chats['data']['chats']:
    group_chat_id = chats['data']['chats'][0]['id']
    
    # 2. 更新主题
    result = update_team_chat_topic(
        chat_id=group_chat_id,
        topic="Updated Topic - " + datetime.now().strftime("%Y-%m-%d")
    )
    
    print(f"Topic updated: {result['success']}")
```

---

## 📈 统计信息

- **新增工具数量**: 3个
- **新增服务方法**: 3个
- **代码行数**: 
  - 工具层: ~400行
  - 服务层: ~150行
- **支持的操作**: 发送消息、创建聊天、更新主题
- **API端点**: 3个Graph API端点

---

## 🔗 相关文档

- [Microsoft Graph API - Chat Resource](https://learn.microsoft.com/en-us/graph/api/resources/chat)
- [Microsoft Graph API - Send Message](https://learn.microsoft.com/en-us/graph/api/chat-post-messages)
- [Microsoft Graph API - Create Chat](https://learn.microsoft.com/en-us/graph/api/chat-post)
- [Chat.ReadWrite Permission](https://learn.microsoft.com/en-us/graph/permissions-reference#chatreadwrite)

---

## ✅ 完成状态

- [x] 分析Chat.ReadWrite权限可用的写入API
- [x] 在services/onedrive_service.py中添加对应的服务方法
- [x] 在onedrive_teams_tools.py中添加发送聊天消息工具
- [x] 添加创建新聊天工具
- [x] 添加更新聊天属性工具(如topic)
- [x] 创建功能总结文档

**实现日期**: 2026-04-22
**实现者**: Kilo Code (Claude Sonnet 4.5)

"""
OneDrive服务模块 - 统一的Microsoft Graph API客户端
整合了文件操作、邮件、OneNote、Teams等功能
"""
import requests
import json
import os
import base64
import time
import html2text
import bs4
from email import policy
from email.parser import BytesParser
from typing import Optional, Dict, Any, List, Generator
import urllib3
import sys
from datetime import datetime

from .mongo_service import MongoTokenService
from exceptions import MongoDBError

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def log(message: str, level: str = "INFO"):
    """统一的日志输出函数"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    print(f"[{timestamp}] [{level}] [OneDrive] {message}", file=sys.stderr)


# HTML处理工具类
class Mail2Text(BytesParser):
    """邮件HTML转文本工具"""
    def __init__(self, html_email_bytes):
        super().__init__(policy=policy.default)
        msg = self.parsebytes(html_email_bytes)
        self.text = msg.get_body(preferencelist=('plain', 'html'))
        if self.text is None:
            self.text = msg.get_body(preferencelist=('html',)).get_content()
        else:
            self.text = self.text.get_content()


class BeautifulSoup(bs4.BeautifulSoup):
    """自定义BeautifulSoup，优化HTML处理"""
    def __init__(self, html_content):
        super().__init__(html_content, 'html.parser')
        for br in self.find_all('br'): 
            br.replace_with('\n')
        for li in self.find_all('li'): 
            li.insert_before('• ')

    def get_text(self):
        return super().get_text(separator='\n', strip=True)


class HTML2Text(html2text.HTML2Text):
    """自定义HTML2Text转换器"""
    def __init__(self):
        super().__init__()
        self.body_width = 0
        self.ul_item_mark = '-'
        self.emphasis_mark = '*'
        self.wrap_links = False


# 数据容器类
class Dir:
    """目录数据容器"""
    def __init__(self, json_data):
        self.json_data = json_data
    
    def __getitem__(self, index):
        return Dict(self.json_data['value'][index])

    def __repr__(self):
        return 'Dir{0}'.format([i['name'] for i in self.json_data['value']])


class Dict(dict):
    """文件/文件夹数据容器"""
    def __repr__(self):
        return 'Dict({0})'.format(self['name'])


# Teams相关类
class Base:
    """Teams基础类"""
    def __init__(self, value, onedrive):
        self.value = value
        self.onedrive = onedrive

    def is_less_days_by_now(self, date_str, days=0.5):
        """检查日期是否在指定天数内"""
        date_time = time.strptime(date_str.split('.')[0].rstrip('Z'), '%Y-%m-%dT%H:%M:%S')
        if date_time.tm_year >= 2020:
            return (time.time() - time.mktime(date_time)) < days * 24 * 3600
        else: 
            return False


class Chat(Base):
    """Teams聊天类"""
    def is_less_days_by_now(self, days=0.5):
        return super().is_less_days_by_now(self.value['viewpoint']['lastMessageReadDateTime'], days)

    def read_messages(self):
        chat_id = self.value['id']
        self.messages = self.onedrive.call_rest_api(f'/chats/{chat_id}/messages', lambda v,o: Message(v,o))
        return self.messages


class Message(Base):
    """Teams消息类"""
    def is_less_days_by_now(self, days=0.5):
        return super().is_less_days_by_now(self.value['lastModifiedDateTime'], days)

    def read_content(self):
        return self.value['body']['content']


# 主要的OneDrive服务类
class OneDriverService:
    """
    统一的Microsoft Graph API客户端
    整合了文件操作、邮件、OneNote、Teams等所有功能
    """
    BASE_URL = os.environ.get("BASE_URL")
    def __init__(self, token: str):
        """
        初始化OneDrive服务
        
        Args:
            token: unique_token（直接用于后端认证）
        """
        log(f"==================== 初始化 OneDriverService ====================")
        log(f"收到 token: {token[:20]}...{token[-10:]}")
        log(f"Token 长度: {len(token)}")
        log(f"BASE_URL: {self.BASE_URL}")
        
        self.token = token
        self.headers = {'Authorization': f'Bearer {self.token}'}
        self.requests = requests
        
        log(f"认证头已设置: Authorization: Bearer {token[:20]}...{token[-10:]}")
        log(f"初始化完成")
        log(f"=================================================================")
    
    def _make_request(self, method: str, url: str, **kwargs) -> requests.Response:
        """统一的请求方法，包含详细日志"""
        request_id = f"{int(time.time() * 1000)}"
        
        log(f"-------------------- 请求开始 [{request_id}] --------------------")
        log(f"请求方法: {method}")
        log(f"请求URL: {url}")
        log(f"请求头: {kwargs.get('headers', {})}")
        
        if 'params' in kwargs and kwargs['params']:
            log(f"查询参数: {json.dumps(kwargs['params'], ensure_ascii=False)}")
        
        if 'json' in kwargs and kwargs['json']:
            log(f"JSON数据: {json.dumps(kwargs['json'], ensure_ascii=False, indent=2)[:500]}...")
        
        if 'data' in kwargs and kwargs['data']:
            data_preview = str(kwargs['data'])[:200]
            log(f"请求体数据 (前200字符): {data_preview}...")
        
        kwargs.setdefault('headers', {}).update(self.headers)
        kwargs.setdefault('verify', False)
        
        log(f"最终请求头: {json.dumps(dict(kwargs['headers']), ensure_ascii=False)}")
        
        try:
            start_time = time.time()
            log(f"发送请求...")
            
            response = getattr(requests, method.lower())(url, **kwargs)
            
            elapsed_time = (time.time() - start_time) * 1000
            log(f"收到响应 - 耗时: {elapsed_time:.2f}ms")
            log(f"响应状态码: {response.status_code}")
            log(f"响应头: {dict(response.headers)}")
            
            # 尝试解析响应内容
            content_type = response.headers.get('Content-Type', '')
            if 'application/json' in content_type:
                try:
                    response_json = response.json()
                    log(f"响应JSON数据 (前500字符): {json.dumps(response_json, ensure_ascii=False, indent=2)[:500]}...")
                    
                    # 如果有错误信息，详细记录
                    if 'error' in response_json:
                        log(f"❌ 响应包含错误: {json.dumps(response_json['error'], ensure_ascii=False, indent=2)}", "ERROR")
                except json.JSONDecodeError as e:
                    log(f"⚠️ 无法解析JSON响应: {e}", "WARN")
                    log(f"响应文本 (前500字符): {response.text[:500]}...", "WARN")
            else:
                log(f"响应内容类型: {content_type}")
                if len(response.content) < 1000:
                    log(f"响应内容: {response.text[:500]}...")
                else:
                    log(f"响应内容大小: {len(response.content)} bytes")
            
            response.raise_for_status()
            
            log(f"✅ 请求成功 [{request_id}]")
            log(f"---------------------------------------------------------------\n")
            
            return response
            
        except requests.HTTPError as e:
            log(f"❌ HTTP错误 [{request_id}]: {e}", "ERROR")
            log(f"状态码: {e.response.status_code}", "ERROR")
            
            try:
                error_detail = e.response.json()
                log(f"错误详情: {json.dumps(error_detail, ensure_ascii=False, indent=2)}", "ERROR")
            except:
                log(f"错误响应文本: {e.response.text[:500]}...", "ERROR")
            
            log(f"---------------------------------------------------------------\n")
            raise
            
        except requests.ConnectionError as e:
            log(f"❌ 连接错误 [{request_id}]: {e}", "ERROR")
            log(f"无法连接到: {url}", "ERROR")
            log(f"---------------------------------------------------------------\n")
            raise
            
        except requests.Timeout as e:
            log(f"❌ 请求超时 [{request_id}]: {e}", "ERROR")
            log(f"---------------------------------------------------------------\n")
            raise
            
        except requests.RequestException as e:
            log(f"❌ 请求异常 [{request_id}]: {e}", "ERROR")
            log(f"异常类型: {type(e).__name__}", "ERROR")
            log(f"---------------------------------------------------------------\n")
            raise
        
        except Exception as e:
            log(f"❌ 未预期的错误 [{request_id}]: {e}", "ERROR")
            log(f"异常类型: {type(e).__name__}", "ERROR")
            import traceback
            log(f"堆栈跟踪:\n{traceback.format_exc()}", "ERROR")
            log(f"---------------------------------------------------------------\n")
            raise

    def call_rest_api(self, api: str, init_func) -> Generator:
        """通用的REST API调用方法，支持分页"""
        log(f"📄 调用分页API: {api}")
        url = self.BASE_URL + api
        page_count = 0
        
        while url:
            page_count += 1
            log(f"获取第 {page_count} 页数据: {url}")
            
            result = self._make_request('GET', url)
            data = result.json()
            
            items_count = len(data.get('value', []))
            log(f"第 {page_count} 页包含 {items_count} 个项目")
            
            for v in data.get('value', []):
                yield init_func(v, self)
            
            url = data.get('@odata.nextLink')
            if url:
                log(f"存在下一页: {url}")
            else:
                log(f"已到达最后一页")

    # ================= 文件操作相关方法 =================

    def list_my_drive_items(self, path: str = '/', top: int = 100) -> Dict:
        """列出用户 OneDrive 中的文件和文件夹"""
        log(f"📁 列出文件 - 路径: {path}, 数量: {top}")
        
        if path == '/' or path == '':
            url = f"{self.BASE_URL}/me/drive/root/children"
            log(f"使用根目录路径")
        else:
            clean_path = path.lstrip('/')
            url = f"{self.BASE_URL}/me/drive/root:/{clean_path}:/children"
            log(f"使用清理后的路径: {clean_path}")
        
        params = {'$top': min(top, 200)}
        log(f"查询参数: {params}")
        
        result = self._make_request('GET', url, params=params)
        response_data = result.json()
        
        items_count = len(response_data.get('value', []))
        log(f"✅ 成功获取 {items_count} 个项目")
        
        return response_data

    def get_my_drive_item(self, path: str = '/') -> Dict:
        """获取用户 OneDrive 中指定路径的项目信息"""
        log(f"📄 获取项目信息 - 路径: {path}")
        
        if path == '/' or path == '':
            url = f"{self.BASE_URL}/me/drive/root"
            log(f"获取根目录信息")
        else:
            clean_path = path.lstrip('/')
            url = f"{self.BASE_URL}/me/drive/root:/{clean_path}"
            log(f"获取路径信息: {clean_path}")
        
        result = self._make_request('GET', url)
        item_data = result.json()
        
        log(f"✅ 成功获取项目: {item_data.get('name', 'Unknown')}")
        
        return item_data

    def url_to_base64(self, url: str) -> str:
        """将分享URL转换为base64编码"""
        log(f"🔐 转换分享URL为base64: {url[:50]}...")
        
        encoded = base64.b64encode(url.encode())
        encoded = b'/shares/u!' + encoded.strip(b'=').replace(b'/',b'_').replace(b'+',b'-')
        result = encoded.decode()
        
        log(f"✅ Base64编码结果: {result[:50]}...")
        
        return result

    def get_driveitem(self, share_path: str):
        """获取驱动器项目信息"""
        log(f"🔗 获取共享驱动器项目: {share_path}")
        
        self.url_root = self.BASE_URL + self.url_to_base64(share_path) + '/driveItem'
        log(f"构建URL: {self.url_root}")
        
        result = self._make_request('GET', self.url_root)
        self.driveitem = result.json()
        self.root = self.BASE_URL + self.driveitem['parentReference']['path']
        
        log(f"✅ 成功获取驱动器项目，根路径: {self.root}")

    def listdir(self, path: str) -> Dir:
        """列出目录内容"""
        log(f"📂 列出目录: {path}")
        
        result = self._make_request('GET', f"{self.root}{path}:/children")
        dir_data = Dir(result.json())
        
        log(f"✅ 目录内容: {dir_data}")
        
        return dir_data

    def downloadfile(self, file: str):
        """下载文件"""
        log(f"⬇️ 下载文件: {file}")
        
        result = self._make_request('GET', f"{self.root}{file}:/content")
        filename = os.path.split(file)[1]
        
        log(f"文件名: {filename}, 大小: {len(result.content)} bytes")
        
        with open(filename, 'wb') as f:
            f.write(result.content)
        
        log(f"✅ 文件已保存: {filename}")

    def search_files(self, query: str) -> Dict:
        """搜索文件"""
        log(f"🔍 搜索文件: {query}")
        
        if not hasattr(self, 'root'):
            error_msg = "请先调用get_driveitem()设置根路径"
            log(f"❌ {error_msg}", "ERROR")
            raise ValueError(error_msg)
        
        drive_root = self.root.split('/root')[0]
        url = f"{drive_root}/root/search(q='{{{query}}}')"
        log(f"搜索URL: {url}")
        
        result = self._make_request('GET', url)
        search_results = result.json()
        
        results_count = len(search_results.get('value', []))
        log(f"✅ 找到 {results_count} 个结果")
        
        return search_results

    # ================= 邮件相关方法 =================
    
    def get_me_email(self) -> str:
        """获取当前用户邮箱地址"""
        log(f"📧 获取用户邮箱")
        
        result = self._make_request('GET', f"{self.BASE_URL}/me/?$select=mail")
        email = result.json().get('mail')
        
        log(f"✅ 邮箱地址: {email}")
        
        return email

    def get_mail_with_filter(self, filter_func, folder: str = "inbox") -> Generator[Dict, None, None]:
        """获取邮件（支持过滤）"""
        log(f"📬 获取邮件 - 文件夹: {folder}")
        
        if folder:
            url = f"{self.BASE_URL}/me/mailFolders/{folder}/messages"
        else:
            url = f"{self.BASE_URL}/me/messages"
        
        params = filter_func() if callable(filter_func) else filter_func
        log(f"过滤参数: {params}")
        
        page_count = 0
        while url:
            page_count += 1
            log(f"获取第 {page_count} 页邮件")
            
            result = self._make_request('GET', url, params=params if page_count == 1 else None)
            data = result.json()
            
            mail_count = len(data.get('value', []))
            log(f"第 {page_count} 页包含 {mail_count} 封邮件")
            
            yield data
            
            url = data.get('@odata.nextLink')
            if url:
                log(f"存在下一页")
            else:
                log(f"已到达最后一页")

    def send_mail(self, to: List[str], cc: List[str], subject: str, body: str) -> Dict:
        """发送邮件"""
        log(f"📤 发送邮件")
        log(f"收件人: {to}")
        log(f"抄送: {cc}")
        log(f"主题: {subject}")
        log(f"内容长度: {len(body)} 字符")
        
        if not to:
            error_msg = "至少需要一个收件人"
            log(f"❌ {error_msg}", "ERROR")
            raise ValueError(error_msg)
        if not subject:
            error_msg = "主题不能为空"
            log(f"❌ {error_msg}", "ERROR")
            raise ValueError(error_msg)
        if not body:
            error_msg = "邮件内容不能为空"
            log(f"❌ {error_msg}", "ERROR")
            raise ValueError(error_msg)
        
        to_recipients = [{"emailAddress": {"address": addr}} for addr in to]
        cc_recipients = [{"emailAddress": {"address": addr}} for addr in cc] if cc else []
        
        content_type = "HTML" if any(tag in body.lower() for tag in ['<html>', '<div>', '<p>', '<br>', '<span>']) else "Text"
        log(f"检测到内容类型: {content_type}")
        
        message = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": content_type,
                    "content": body
                },
                "toRecipients": to_recipients,
                "ccRecipients": cc_recipients
            },
            "saveToSentItems": True
        }
        
        result = self._make_request('POST', f"{self.BASE_URL}/me/sendMail", json=message)
        
        log(f"✅ 邮件发送成功")
        
        return {"status": "sent", "status_code": result.status_code}

    def get_single_mail(self, message_id: str, select_fields: Optional[List[str]] = None) -> Dict:
        """获取单个邮件"""
        log(f"📩 获取邮件详情 - ID: {message_id}")
        
        url = f"{self.BASE_URL}/me/messages/{message_id}"
        params = {}
        if select_fields:
            params['$select'] = ','.join(select_fields)
            log(f"选择字段: {select_fields}")
        
        result = self._make_request('GET', url, params=params)
        mail_data = result.json()
        
        log(f"✅ 邮件主题: {mail_data.get('subject', 'No subject')}")
        
        return mail_data

    def reply_to_mail(self, message_id: str, body: str, reply_all: bool = False) -> Dict:
        """回复邮件"""
        action = "全部回复" if reply_all else "回复"
        log(f"↩️ {action}邮件 - ID: {message_id}")
        log(f"回复内容长度: {len(body)} 字符")
        
        endpoint = 'replyAll' if reply_all else 'reply'
        url = f"{self.BASE_URL}/me/messages/{message_id}/{endpoint}"
        data = {"comment": body}
        
        result = self._make_request('POST', url, json=data)
        
        log(f"✅ {action}成功")
        
        return {"status": f"{'reply_all' if reply_all else 'reply'}_sent", "status_code": result.status_code}

    def forward_mail(self, message_id: str, to_recipients: List[str], cc_recipients: List[str] = None, body: Optional[str] = None) -> Dict:
        """转发邮件"""
        log(f"➡️ 转发邮件 - ID: {message_id}")
        log(f"转发给: {to_recipients}")
        if cc_recipients:
            log(f"抄送: {cc_recipients}")
        if body:
            log(f"附加说明长度: {len(body)} 字符")
        
        url = f"{self.BASE_URL}/me/messages/{message_id}/forward"
        
        to_list = [{"emailAddress": {"address": addr}} for addr in to_recipients]
        data = {"toRecipients": to_list}
        
        if cc_recipients:
            cc_list = [{"emailAddress": {"address": addr}} for addr in cc_recipients]
            data["ccRecipients"] = cc_list

        if body:
            data["comment"] = body
        
        result = self._make_request('POST', url, json=data)
        
        log(f"✅ 转发成功")
        
        return {"status": "forwarded", "status_code": result.status_code}

    def get_mail_folders(self) -> Dict:
        """获取所有邮件文件夹"""
        log(f"📁 获取邮件文件夹列表")
        
        result = self._make_request('GET', f"{self.BASE_URL}/me/mailFolders")
        folders_data = result.json()
        
        folder_count = len(folders_data.get('value', []))
        log(f"✅ 找到 {folder_count} 个文件夹")
        
        return folders_data

    def get_folder_messages(self, folder_id: str, filter_params=None) -> Generator[Dict, None, None]:
        """获取文件夹中的邮件"""
        log(f"📂 获取文件夹邮件 - 文件夹ID: {folder_id}")
        if filter_params:
            log(f"过滤参数: {filter_params}")
        
        url = f"{self.BASE_URL}/me/mailFolders/{folder_id}/messages"
        params = filter_params() if callable(filter_params) else (filter_params or {})
        
        page_count = 0
        while url:
            page_count += 1
            log(f"获取第 {page_count} 页")
            
            result = self._make_request('GET', url, params=params if page_count == 1 else None)
            data = result.json()
            
            mail_count = len(data.get('value', []))
            log(f"第 {page_count} 页包含 {mail_count} 封邮件")
            
            yield data
            
            url = data.get('@odata.nextLink')

    def get_mail_attachments(self, message_id: str) -> Dict:
        """获取邮件附件"""
        log(f"📎 获取邮件附件列表 - 邮件ID: {message_id}")
        
        result = self._make_request('GET', f"{self.BASE_URL}/me/messages/{message_id}/attachments")
        attachments_data = result.json()
        
        attachment_count = len(attachments_data.get('value', []))
        log(f"✅ 找到 {attachment_count} 个附件")
        
        return attachments_data

    def download_attachment(self, message_id: str, attachment_id: str) -> Dict:
        """下载附件"""
        log(f"⬇️ 下载附件 - 邮件ID: {message_id}, 附件ID: {attachment_id}")
        
        result = self._make_request('GET', f"{self.BASE_URL}/me/messages/{message_id}/attachments/{attachment_id}")
        attachment_data = result.json()
        
        log(f"✅ 附件名称: {attachment_data.get('name', 'Unknown')}")
        
        return attachment_data

    def search_mail(self, search_query: str, folder_id: Optional[str] = None) -> Dict:
        """搜索邮件"""
        log(f"🔍 搜索邮件 - 关键词: {search_query}")
        if folder_id:
            log(f"搜索文件夹: {folder_id}")
        
        if folder_id:
            url = f"{self.BASE_URL}/me/mailFolders/{folder_id}/messages"
        else:
            url = f"{self.BASE_URL}/me/messages"
        
        params = {"$search": f'"{search_query}"'}
        
        result = self._make_request('GET', url, params=params)
        total_result = result.json()
        
        # 处理分页
        page_count = 1
        data = result.json()
        while '@odata.nextLink' in data:
            page_count += 1
            log(f"获取搜索结果第 {page_count} 页")
            
            url = data['@odata.nextLink']
            result = self._make_request('GET', url)
            data = result.json()
            total_result['value'].extend(data.get('value', []))
        
        total_result.pop('@odata.nextLink', None)
        
        total_count = len(total_result.get('value', []))
        log(f"✅ 搜索完成，共找到 {total_count} 封邮件")
        
        return total_result

    def get_single_mail_folder(self, folder_id: str) -> Dict:
        """获取单个邮件文件夹信息"""
        log(f"📁 获取文件夹信息 - ID: {folder_id}")
        
        result = self._make_request('GET', f"{self.BASE_URL}/me/mailFolders/{folder_id}")
        folder_data = result.json()
        
        log(f"✅ 文件夹名称: {folder_data.get('displayName', 'Unknown')}")
        
        return folder_data

    def save_each_mail_as_markdown(self, mail_data: Dict, saved_dir: str = 'mail'):
        """将邮件保存为Markdown文件"""
        log(f"💾 保存邮件为Markdown - 目录: {saved_dir}")
        
        os.makedirs(saved_dir, exist_ok=True)
        
        saved_count = 0
        for mail in mail_data.get('value', []):
            fname = f'{time.time()} {mail["subject"]}'[:255]
            for schar in ("?", "/", ":", "*", "<", ">", "|", "\\", "\""): 
                fname = fname.replace(schar, "")
            
            markdown = HTML2Text().handle(mail["body"]["content"])
            filepath = f'{saved_dir}/{fname}.md'
            
            with open(filepath, 'w', encoding='utf8') as f:
                f.write(markdown)
            
            saved_count += 1
            log(f"已保存: {fname}.md")
        
        log(f"✅ 共保存 {saved_count} 封邮件")

    def save_attachments(self, message_id: str, save_dir: str = 'attachments') -> List[Dict]:
        """保存邮件的所有附件"""
        log(f"💾 保存附件 - 邮件ID: {message_id}, 目录: {save_dir}")
        
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        attachments = self.get_mail_attachments(message_id)
        saved_files = []
        
        for attachment in attachments.get('value', []):
            if attachment.get('contentBytes'):
                filename = attachment.get('name', f"attachment_{attachment['id']}")
                filepath = os.path.join(save_dir, filename)
                
                content = base64.b64decode(attachment['contentBytes'])
                with open(filepath, 'wb') as f:
                    f.write(content)
                
                file_info = {
                    "filename": filename,
                    "filepath": filepath,
                    "size": len(content),
                    "content_type": attachment.get('contentType', 'unknown')
                }
                saved_files.append(file_info)
                
                log(f"已保存附件: {filename} ({len(content)} bytes)")
        
        log(f"✅ 共保存 {len(saved_files)} 个附件")
        
        return saved_files

    def get_unread_count(self, folder_id: Optional[str] = None) -> int:
        """获取未读邮件数量"""
        log(f"📊 获取未读邮件数量")
        
        if folder_id:
            log(f"指定文件夹: {folder_id}")
            folder_info = self.get_single_mail_folder(folder_id)
        else:
            log(f"获取收件箱")
            folders = self.get_mail_folders()
            inbox = next((f for f in folders.get('value', []) if f['displayName'] == 'Inbox'), None)
            folder_info = inbox
        
        unread_count = folder_info.get('unreadItemCount', 0) if folder_info else 0
        log(f"✅ 未读邮件数: {unread_count}")
        
        return unread_count

    # ================= OneNote相关方法 =================
    
    def get_notebooks(self, user_email: Optional[str] = None) -> Generator:
        """获取笔记本"""
        log(f"📓 获取笔记本列表")
        if user_email:
            log(f"指定用户: {user_email}")
        
        if user_email: 
            url = f'/users/{user_email}/onenote/notebooks'
        else: 
            url = '/me/onenote/notebooks'
        
        return self.call_rest_api(url, lambda v, o: v)

    def get_sections(self, notebook_id: str, user_email: Optional[str] = None) -> Generator:
        """获取笔记本章节"""
        log(f"📑 获取章节 - 笔记本ID: {notebook_id}")
        
        if user_email: 
            url = f'/users/{user_email}/onenote/notebooks'
        else: 
            url = '/me/onenote/notebooks'
        
        return self.call_rest_api(f'{url}/{notebook_id}/sections', lambda v, o: v)

    def get_pages(self, section_id: str, user_email: Optional[str] = None) -> Generator:
        """获取章节页面"""
        log(f"📄 获取页面 - 章节ID: {section_id}")
        
        if user_email: 
            url = f'/users/{user_email}/onenote/sections'
        else: 
            url = f'/me/onenote/sections'
        
        return self.call_rest_api(f'{url}/{section_id}/pages', lambda v, o: v)

    def get_page_content(self, page_id: str, user_email: Optional[str] = None) -> bytes:
        """获取页面内容"""
        log(f"📝 获取页面内容 - 页面ID: {page_id}")
        
        if user_email: 
            url = f'/users/{user_email}/onenote/pages'
        else: 
            url = '/me/onenote/pages'
        
        full_url = f"{self.BASE_URL}{url}/{page_id}/content"
        result = self._make_request('GET', full_url)
        
        content_size = len(result.content)
        log(f"✅ 页面内容大小: {content_size} bytes")
        
        return result.content

    # ================= Teams相关方法 =================
    
    def get_chats(self) -> Generator[Chat, None, None]:
        """获取Teams聊天"""
        log(f"💬 获取Teams聊天列表")
        return self.call_rest_api('/me/chats', lambda v, o: Chat(v, o))

    def get_chat_messages(self, chat_id: str) -> Generator[Message, None, None]:
        """获取聊天消息"""
        log(f"💬 获取聊天消息 - 聊天ID: {chat_id}")
        return self.call_rest_api(f'/chats/{chat_id}/messages', lambda v, o: Message(v, o))


# 便捷的工厂函数
async def create_onedrive_service(token: str) -> OneDriverService:
    """
    创建并认证OneDrive服务实例
    
    Args:
        token: unique_token
        
    Returns:
        已认证的OneDriverService实例
    """
    log(f"🏭 工厂函数: 创建 OneDrive 服务")
    log(f"Token: {token[:20]}...{token[-10:]}")
    
    service = OneDriverService(token)
    
    log(f"✅ 服务创建完成")
    
    return service
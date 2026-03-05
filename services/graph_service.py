"""
Microsoft Graph API服务模块
"""
import aiohttp
from typing import Any, Dict, Optional
import sys
import os
from exceptions import GraphAPIError

class GraphAPIService:
    """Microsoft Graph API服务类"""
    
    @staticmethod
    async def make_request(
        token: str,
        endpoint: str,
        method: str = "GET",
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        content: Optional[bytes] = None
    ) -> Dict[str, Any]:
        """统一的 Graph API 请求方法"""
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json' if data else 'application/octet-stream'
        }
        base_url = os.environ.get("BASE_URL")
        url = f'{base_url}{endpoint}'
        
        async with aiohttp.ClientSession() as session:
            try:
                if method == 'GET':
                    async with session.get(url, headers=headers, params=params) as response:
                        return await GraphAPIService._handle_response(response)
                elif method == 'POST':
                    if content:
                        headers['Content-Type'] = 'application/octet-stream'
                        async with session.post(url, headers=headers, data=content) as response:
                            return await GraphAPIService._handle_response(response)
                    else:
                        async with session.post(url, headers=headers, json=data) as response:
                            return await GraphAPIService._handle_response(response)
                elif method == 'PUT':
                    if content:
                        headers['Content-Type'] = 'application/octet-stream'
                        async with session.put(url, headers=headers, data=content) as response:
                            return await GraphAPIService._handle_response(response)
                    else:
                        async with session.put(url, headers=headers, json=data) as response:
                            return await GraphAPIService._handle_response(response)
                else:
                    raise GraphAPIError(400, f"Unsupported HTTP method: {method}")
                    
            except aiohttp.ClientError as e:
                raise GraphAPIError(500, f"Request failed: {str(e)}")

    @staticmethod
    async def _handle_response(response: aiohttp.ClientResponse) -> Dict[str, Any]:
        """处理 API 响应"""
        if response.status in [200, 201, 202]:
            if response.content_length == 0:
                return {}
            try:
                return await response.json()
            except:
                return {"status": "success"}
        elif response.status == 401:
            raise GraphAPIError(401, "Unauthorized - Token may be expired")
        elif response.status == 403:
            raise GraphAPIError(403, "Forbidden - Insufficient permissions")
        elif response.status == 404:
            raise GraphAPIError(404, "Resource not found")
        else:
            error_text = await response.text()
            raise GraphAPIError(response.status, error_text)
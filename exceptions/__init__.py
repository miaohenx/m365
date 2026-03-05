"""
异常类模块
"""

class GraphAPIError(Exception):
    """Graph API 错误类"""
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        self.message = message
        super().__init__(f"GraphAPI Error {status_code}: {message}")

class MongoDBError(Exception):
    """MongoDB 错误类"""
    def __init__(self, message: str):
        self.message = message
        super().__init__(f"MongoDB Error: {message}")
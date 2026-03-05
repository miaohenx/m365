"""
MongoDB服务模块
"""
from motor.motor_asyncio import AsyncIOMotorClient
from dotenv import load_dotenv
from typing import Optional, Dict, Any
import os
import sys
from exceptions import MongoDBError
load_dotenv()

class MongoTokenService:
    """MongoDB Token服务类"""
    _client = None
    _db = None
    
    @classmethod
    async def get_client(cls):
        if cls._client is None:
            try:
                # 直接使用完整的MONGO_URI，包含认证信息和数据库名
                mongo_uri = os.environ.get("MONGO_URI")
                
                if not mongo_uri:
                    raise MongoDBError("MONGO_URI 环境变量未设置")
                
                print(f"[MCP] 连接MongoDB: {mongo_uri.split('@')[0]}@***", file=sys.stderr)  # 隐藏密码部分
                
                # 创建客户端，URI中已包含所有连接信息
                cls._client = AsyncIOMotorClient(mongo_uri)
                
                # 测试连接
                await cls._client.admin.command('ping')
                print("[MCP] MongoDB连接成功", file=sys.stderr)
                
            except Exception as e:
                print(f"[MCP] MongoDB连接失败: {e}", file=sys.stderr)
                raise MongoDBError(f"连接失败: {e}")
        return cls._client
    
    @classmethod
    async def get_db(cls):
        if cls._db is None:
            client = await cls.get_client()
            
            # 从MONGO_URI中提取数据库名，或使用环境变量
            mongo_uri = os.environ.get("MONGO_URI")
            
            # 尝试从URI中解析数据库名
            db_name = None
            if mongo_uri and '/' in mongo_uri:
                # 解析URI格式: mongodb://user:pass@host:port/dbname?options
                try:
                    # 提取数据库名部分
                    uri_parts = mongo_uri.split('/')
                    if len(uri_parts) > 3:
                        db_part = uri_parts[3]  # dbname?options
                        db_name = db_part.split('?')[0]  # 去掉查询参数
                except Exception as e:
                    print(f"[MCP] 解析数据库名失败: {e}", file=sys.stderr)
            
            # 如果无法从URI解析，使用环境变量或默认值
            if not db_name:
                db_name = os.environ.get("MONGO_DB_NAME", "fivcopilot_rag_mongodb")
            
            print(f"[MCP] 使用数据库: {db_name}", file=sys.stderr)
            cls._db = client[db_name]
            
        return cls._db
    
    @classmethod
    async def get_collection(cls, collection_name: str):
        db = await cls.get_db()
        return db[collection_name]
    
    @classmethod
    async def get_az_token_by_token(cls, token: str, collection_name: str = "user_tokens") -> Optional[str]:
        """根据unique_token查找对应的access_token"""
        try:
            collection = await cls.get_collection(collection_name)
            
            result = await collection.find_one(
                {"unique_token": token},
                {"access_token": 1, "_id": 0}
            )
            
            if result and result.get("access_token"):
                return result["access_token"]
            else:
                return None
                
        except Exception as e:
            print(f"[MCP] 查询access_token失败: {e}", file=sys.stderr)
            raise MongoDBError(f"查询失败: {e}")
    
    @classmethod
    async def get_user_by_token(cls, token: str, collection_name: str = "user_tokens") -> Optional[Dict[str, Any]]:
        """根据token获取用户信息"""
        try:
            collection = await cls.get_collection(collection_name)
            
            result = await collection.find_one(
                {"token": token},
                {"user_id": 1, "az_token": 1, "expires_at": 1, "_id": 0}
            )
            
            if result:
                print(f"[MCP] 找到用户信息: user_id={result.get('user_id')}", file=sys.stderr)
            else:
                print(f"[MCP] 未找到用户信息: {token[:20]}...", file=sys.stderr)
            
            return result
                
        except Exception as e:
            print(f"[MCP] 查询用户信息失败: {e}", file=sys.stderr)
            raise MongoDBError(f"查询失败: {e}")
    
    @classmethod
    async def test_connection(cls) -> bool:
        """测试MongoDB连接"""
        try:
            client = await cls.get_client()
            db = await cls.get_db()
            
            # 测试数据库访问
            collections = await db.list_collection_names()
            print(f"[MCP] 数据库连接测试成功，找到 {len(collections)} 个集合", file=sys.stderr)
            print(f"[MCP] 集合列表: {collections[:5]}{'...' if len(collections) > 5 else ''}", file=sys.stderr)
            
            return True
        except Exception as e:
            print(f"[MCP] 数据库连接测试失败: {e}", file=sys.stderr)
            return False
def get_token_from_context():
    """从上下文变量中获取Authorization token"""
    try:
        from contextvars import copy_context
        
        ctx = copy_context()
        for var, value in ctx.items():
            if hasattr(var, 'name') and 'request' in var.name.lower():
                if hasattr(value, 'headers'):
                    auth_header = value.headers.get('Authorization', '')
                    if auth_header.startswith('Bearer '):
                        return auth_header[7:]
        return None
    except:
        return None
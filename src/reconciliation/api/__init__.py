"""API 層：FastAPI routers 與資料傳輸物件。

只處理 HTTP。所有業務規則都在 ``domain/`` 與 ``services/``。
"""

from .app import create_app

__all__ = ["create_app"]

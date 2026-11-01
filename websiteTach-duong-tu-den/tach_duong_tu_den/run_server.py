"""Khởi động bản server. Đổi cổng/địa chỉ bằng biến môi trường HOST, PORT.

Mặc định chỉ mở ở máy này (127.0.0.1). Muốn cả mạng nội bộ truy cập:
    HOST=0.0.0.0 python run_server.py     (Windows PowerShell: $env:HOST="0.0.0.0")
"""

import os

import uvicorn

if __name__ == "__main__":
    uvicorn.run(
        "server.app:app",
        host=os.getenv("HOST", "127.0.0.1"),
        port=int(os.getenv("PORT", "8000")),
    )

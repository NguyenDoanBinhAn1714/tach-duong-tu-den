"""Cấu hình xác thực riêng cho bộ test; không dùng các giá trị này khi triển khai."""
import os

os.environ.setdefault("APP_ENV", "test")
os.environ.setdefault("APP_USERNAME", "test-admin")
os.environ.setdefault("APP_PASSWORD", "TestPassword!234")
os.environ.setdefault("SESSION_SECRET", "test-only-session-secret-please-do-not-use")
os.environ.pop("GEMINI_API_KEY", None)

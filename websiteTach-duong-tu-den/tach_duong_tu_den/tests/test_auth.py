"""Kiểm thử đăng nhập và bảo vệ các đường dẫn nhạy cảm."""
from fastapi.testclient import TestClient

from server.app import APP_PASSWORD, APP_USERNAME, app


def test_trang_chu_va_api_can_dang_nhap():
    with TestClient(app, follow_redirects=False) as client:
        page = client.get("/")
        assert page.status_code == 303
        assert page.headers["location"].startswith("/login?next=/")
        api = client.post("/api/split-text", json={"text": "Đường A từ X đến Y"})
        assert api.status_code == 401
        assert "đăng nhập" in api.json()["detail"].lower()


def test_dang_nhap_thanh_cong_va_dang_xuat():
    with TestClient(app, follow_redirects=False) as client:
        r = client.post("/login", data={"username": APP_USERNAME, "password": APP_PASSWORD})
        assert r.status_code == 303 and r.headers["location"] == "/"
        home = client.get("/")
        assert home.status_code == 200 and "Tách Đường" in home.text
        out = client.post("/logout")
        assert out.status_code == 303 and out.headers["location"] == "/login"
        assert client.get("/", follow_redirects=False).status_code == 303


def test_sai_mat_khau_khong_duoc_cap_phien():
    with TestClient(app, follow_redirects=False) as client:
        r = client.post("/login", data={"username": APP_USERNAME, "password": "mat-khau-sai-khong-dung"})
        assert r.status_code == 401
        assert "không chính xác" in r.text
        assert client.get("/", follow_redirects=False).status_code == 303


def test_dang_ky_va_dang_nhap_tai_khoan_moi(tmp_path, monkeypatch):
    import server.app as srv
    monkeypatch.setattr(srv, "USER_STORE_PATH", tmp_path / "users.json")
    with TestClient(app, follow_redirects=False) as client:
        page = client.get("/register")
        assert page.status_code == 200 and "Đăng ký tài khoản" in page.text
        registered = client.post("/register", data={
            "username": "nguoiDungMoi",
            "password": "MatKhauMoi123",
            "confirm_password": "MatKhauMoi123",
        })
        assert registered.status_code == 303 and registered.headers["location"] == "/"
        assert client.get("/", follow_redirects=False).status_code == 200

        client.post("/logout")
        login = client.post("/login", data={
            "username": "nguoiDungMoi",
            "password": "MatKhauMoi123",
        })
        assert login.status_code == 303 and login.headers["location"] == "/"


def test_dang_ky_tu_choi_mau_khau_khong_khop(tmp_path, monkeypatch):
    import server.app as srv
    monkeypatch.setattr(srv, "USER_STORE_PATH", tmp_path / "users.json")
    with TestClient(app, follow_redirects=False) as client:
        r = client.post("/register", data={
            "username": "nguoiDung",
            "password": "MatKhauMoi123",
            "confirm_password": "MatKhauKhac123",
        })
        assert r.status_code == 400
        assert "không khớp" in r.text

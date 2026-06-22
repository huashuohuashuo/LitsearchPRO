import base64
import http.client
import importlib.util
import json
import shutil
import tempfile
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SERVER_FILE = ROOT / "LitSearchPro_Generic_Server.py"


def load_server_module():
    spec = importlib.util.spec_from_file_location("lit_server_v22120", SERVER_FILE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class DummyApp:
    def __init__(self, db, config):
        self.db = db
        self.config = config
        self.logs = []
        self.refresh_requested = threading.Event()

    def app_log(self, message):
        self.logs.append(message)

    def log(self, message):
        self.logs.append(message)

    def refresh_async(self):
        pass


def post(port, path, payload, token=None, expect_status=200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json", "Content-Length": str(len(body))}
    if token:
        headers["Authorization"] = "Bearer " + token
    conn = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
    conn.request("POST", path, body=body, headers=headers)
    resp = conn.getresponse()
    raw = resp.read()
    conn.close()
    data = json.loads(raw.decode("utf-8"))
    if resp.status != expect_status:
        raise AssertionError(f"{path} expected {expect_status}, got {resp.status}: {data}")
    return data


def seed_user(module, conn, group, username, display_name, role, team, active=1, password="pw", mentor_id=None, mentor_username=""):
    salt, digest = module.hash_password(password)
    return conn.execute(
        """INSERT INTO users(group_code,username,role,display_name,team_name,mentor_user_id,
           mentor_username,salt,password_hash,active,created_at)
           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
        (group, username, role, display_name, team, mentor_id, mentor_username, salt, digest, active, module.now()),
    ).lastrowid


def main():
    module = load_server_module()
    tmp = tempfile.mkdtemp(prefix="litsearch_v22120_")
    httpd = None
    try:
        module.apply_data_dir({"data_dir": tmp})
        conn = module.connect_db()
        admin_id = seed_user(module, conn, "admin", "admin", "超级管理员", "超级管理员", "校级")
        mentor_a = seed_user(module, conn, "team-a", "mentor_a", "李老师", "导师", "课题组A")
        mentor_b = seed_user(module, conn, "team-b", "mentor_b", "李老师", "导师", "课题组B")
        student_a = seed_user(module, conn, "team-a", "student_a", "学生甲", "学生", "课题组A", active=0, mentor_id=mentor_a, mentor_username="mentor_a")
        student_b = seed_user(module, conn, "team-b", "student_b", "学生乙", "学生", "课题组B", active=0, mentor_id=mentor_b, mentor_username="mentor_b")
        own_inactive_mentor = seed_user(module, conn, "team-a", "mentor_pending", "待批导师", "导师", "课题组A", active=0)
        active_student = seed_user(module, conn, "team-a", "student_active", "已激活学生", "学生", "课题组A", active=1, mentor_id=mentor_a, mentor_username="mentor_a")
        conn.commit()

        app = DummyApp(conn, {"group_code": "admin", "max_concurrent_requests": 16, "public_url": ""})
        httpd = module.ManagedHTTPServer(("127.0.0.1", 0), module.ApiHandler, app)
        port = httpd.server_address[1]
        threading.Thread(target=httpd.serve_forever, daemon=True).start()

        admin_token = post(port, "/api/login", {"username": "admin", "password": "pw", "group_code": "admin"})["token"]
        mentor_token = post(port, "/api/login", {"username": "mentor_a", "password": "pw", "group_code": "team-a"})["token"]
        student_token = post(port, "/api/login", {"username": "student_active", "password": "pw", "group_code": "team-a"})["token"]

        mentor_accounts = post(port, "/api/users/list", {}, mentor_token)["items"]
        mentor_ids = {x["id"] for x in mentor_accounts}
        assert student_a in mentor_ids
        assert student_b not in mentor_ids
        assert mentor_a not in mentor_ids
        assert own_inactive_mentor not in mentor_ids
        assert all(x["role"] == "学生" for x in mentor_accounts)

        student_accounts = post(port, "/api/users/list", {}, student_token)["items"]
        assert student_accounts == []
        post(port, "/api/users/approve", {"id": student_a, "active": 1}, student_token, expect_status=403)
        post(port, "/api/users/approve", {"id": mentor_a, "active": 0}, mentor_token, expect_status=403)
        post(port, "/api/users/approve", {"id": student_b, "active": 1}, mentor_token, expect_status=403)
        post(port, "/api/users/approve", {"id": student_a, "active": 1}, mentor_token)
        assert conn.execute("SELECT active FROM users WHERE id=?", (student_a,)).fetchone()["active"] == 1

        csv_text = "器材名称,品牌,类别,型号,学生管理员1,学生管理员2\r\n旋涂仪,Laurell,工艺设备,WS-650,,\r\n"
        content_b64 = base64.b64encode(csv_text.encode("gbk")).decode("ascii")
        imported = post(port, "/api/equipment/import-csv", {"file_name": "equipment.csv", "content_b64": content_b64}, mentor_token)
        assert imported["imported"] == 1
        equipment = post(port, "/api/equipment/list", {}, mentor_token)["items"]
        assert any(x["name"] == "旋涂仪" and x["model"] == "WS-650" for x in equipment)

        admin_accounts = post(port, "/api/users/list", {}, admin_token)["items"]
        assert any(x["id"] == mentor_b for x in admin_accounts)
        print("v22.1.20 account approval and web equipment CSV import OK")
    finally:
        if httpd:
            httpd.shutdown()
            httpd.server_close()
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()

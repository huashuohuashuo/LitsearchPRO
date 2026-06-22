import http.client
import importlib.util
import json
import shutil
import tempfile
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SERVER_FILE = ROOT / "LitSearchPro_Generic_Server.py"


def load_module():
    spec = importlib.util.spec_from_file_location("lit_server_v22121", SERVER_FILE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class DummyApp:
    def __init__(self, db, config):
        self.db = db
        self.config = config
        self.refresh_requested = threading.Event()

    def app_log(self, _message):
        pass

    def log(self, _message):
        pass

    def refresh_async(self):
        pass


def post(port, path, payload, token=None, expected=200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json", "Content-Length": str(len(body))}
    if token:
        headers["Authorization"] = "Bearer " + token
    conn = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
    conn.request("POST", path, body=body, headers=headers)
    response = conn.getresponse()
    data = json.loads(response.read().decode("utf-8"))
    conn.close()
    assert response.status == expected, (path, response.status, data)
    return data


def seed_user(module, conn, group, username, name, role, team):
    salt, digest = module.hash_password("pw")
    return conn.execute(
        """INSERT INTO users(group_code,username,role,display_name,team_name,salt,password_hash,active,created_at)
           VALUES(?,?,?,?,?,?,?,?,?)""",
        (group, username, role, name, team, salt, digest, 1, module.now()),
    ).lastrowid


def main():
    module = load_module()
    temp_dir = tempfile.mkdtemp(prefix="litsearch_v22121_")
    server = None
    try:
        module.apply_data_dir({"data_dir": temp_dir})
        conn = module.connect_db()
        seed_user(module, conn, "team-a", "manager_a", "甲导师", "导师", "团队A")
        seed_user(module, conn, "team-b", "manager_b", "乙导师", "导师", "团队B")
        lab_id = conn.execute(
            """INSERT INTO laboratories(name,college,address,lab_type,team_name,managers_json,
               commitment_text,active,created_by,created_at,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            ("实验室A", "学院", "A101", "公共实验室", "", json.dumps(["manager_a"]), "", 1, "manager_a", module.now(), module.now()),
        ).lastrowid
        warehouse_id = conn.execute(
            """INSERT INTO chemical_warehouses(name,college,address,managers_json,commitment_text,
               active,created_by,created_at,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            ("危化品库房A", "学院", "B101", json.dumps(["manager_a", "keeper_a"]), "", 1, "manager_a", module.now(), module.now()),
        ).lastrowid
        conn.commit()

        app = DummyApp(conn, {"group_code": "team-a", "max_concurrent_requests": 16, "public_url": ""})
        server = module.ManagedHTTPServer(("127.0.0.1", 0), module.ApiHandler, app)
        port = server.server_address[1]
        threading.Thread(target=server.serve_forever, daemon=True).start()
        token = post(port, "/api/login", {"username": "manager_b", "password": "pw", "group_code": "team-b"})["token"]

        lab_error = post(port, "/api/laboratory/channel", {"laboratory_id": lab_id}, token, expected=403)
        assert lab_error["error"] == "您没有该实验室的管理权限"
        warehouse_error = post(port, "/api/warehouse/channel", {"warehouse_id": warehouse_id}, token, expected=403)
        assert warehouse_error["error"] == "您没有该危险化学品库房的管理权限"
        print("v22.1.21 management permission prompts OK")
    finally:
        if server:
            server.shutdown()
            server.server_close()
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()

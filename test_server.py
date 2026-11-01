#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Backend-only unit/integration tests for server.py (stdlib unittest).

Run:  python -m unittest -v test_server
Each test uses its own temporary data directory; production data/ is untouched.
"""

from __future__ import annotations

import io
import json
import os
import shutil
import sqlite3
import sys
import threading
import unittest
import uuid
import zipfile
from http.client import HTTPConnection
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import server  # noqa: E402

try:
    from openpyxl import Workbook, load_workbook
except ImportError:  # pragma: no cover
    Workbook = None
    load_workbook = None


SEED = {
    "tests": [
        {"id": "T-001", "priority": "P0", "module": "AFE", "title": "输入参考噪声",
         "purpose": "验证噪声", "stimulus": "暗场", "metrics": "nArms", "conditions": "25C",
         "calculation": "RMS", "criteria": "<40", "instruments": "示波器",
         "source_result": "37", "source_unit": "nArms", "source_status": "Pass",
         "source_date": "2025-01-01", "source_notes": "论文"},
        {"id": "T-002", "priority": "P1", "module": "TIA", "title": "跨阻增益",
         "purpose": "验证增益", "source_result": "96", "source_unit": "dB"},
    ],
    "summary": [{"id": "s1", "text": "汇总"}],
    "instructions": [{"id": "i1", "text": "说明"}],
    "source": "unit-test-seed",
}


def _xlsx_bytes(text="ok"):
    wb = Workbook()
    wb.active.append([text])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _jpg_bytes():
    return b"\xff\xd8\xff\xe0" + b"\x00" * 32 + b"\xff\xd9"


def _project_workbook():
    """A workbook shaped like AFE table."""
    wb = Workbook()
    ws = wb.active
    ws.title = "测试计划"
    ws.append(["序号", "优先级", "模块", "测试项目", "测试目的", "测试连接/激励", "需记录的指标",
               "建议测试条件", "计算/处理方法", "论文参考/建议判据", "仪器", "实测结果", "单位",
               "Pass/Fail", "测试日期", "备注"])
    ws.append(["X-01", "P0", "TIA", "增益", "测增益", "正弦", "ZT", "常温", "ZT=V/I", ">=90dB",
               "VNA", "", "", "", "", ""])
    ws.append(["X-02", "P1", "TIA", "带宽", "测带宽", "扫频", "-3dB", "常温", "f", ">=10MHz",
               "VNA", "", "", "", "", ""])
    ws.append([None, "P1", "TIA", "无序号项", "自动编号", "", "", "", "", "", "", "", "", "", "", ""])
    ws.append([None] * 16)
    summary = wb.create_sheet("指标汇总")
    summary.append(["模块", "建议优先记录的核心指标", "结果", "单位", "目标/判据", "备注"])
    summary.append(["TIA", "增益", "", "dB", ">=90", ""])
    guide = wb.create_sheet("填写说明")
    guide.append(["项目", "说明"])
    guide.append(["使用方法", "先做 P0"])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


class BackendTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        server._QUIET = True
        cls.temp_root = APP_DIR / ".test-tmp"
        cls.temp_root.mkdir(exist_ok=True)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.temp_root, ignore_errors=True)

    def setUp(self):
        # Avoid tempfile: mode 0o700 creates an inaccessible ACL on this host.
        self.root = self.temp_root / f"afe-test-{uuid.uuid4().hex}"
        self.root.mkdir()
        (self.root / "public").mkdir()
        (self.root / "public" / "index.html").write_text(
            "<!doctype html><html><body>TestFlow</body></html>", encoding="utf-8")
        (self.root / "public" / "app.js").write_text("console.log('afe');", encoding="utf-8")
        (self.root / "seed.json").write_text(
            json.dumps(SEED, ensure_ascii=False), encoding="utf-8")
        server.CONFIG = server.Config(base_dir=self.root, port=0)
        server.init_storage()
        server.init_db()
        server.seed_default_project()

        self.httpd = server.TestFlowServer(("127.0.0.1", 0), server.Handler)
        self.port = self.httpd.server_address[1]
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.httpd.shutdown()
        self.httpd.server_close()
        self.thread.join(timeout=5)
        shutil.rmtree(self.root, ignore_errors=True)

    # ---- helpers ---------------------------------------------------------
    def request(self, method, path, body=None, headers=None, raw_body=None):
        conn = HTTPConnection("127.0.0.1", self.port, timeout=15)
        headers = dict(headers or {})
        payload = raw_body
        if body is not None:
            payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers.setdefault("Content-Type", "application/json")
        conn.request(method, path, body=payload, headers=headers)
        response = conn.getresponse()
        data = response.read()
        result = (response.status, dict(response.getheaders()), data)
        conn.close()
        return result

    def json_request(self, method, path, body=None, headers=None, raw_body=None):
        status, resp_headers, data = self.request(method, path, body, headers, raw_body)
        try:
            parsed = json.loads(data.decode("utf-8")) if data else None
        except json.JSONDecodeError:
            parsed = None
        return status, resp_headers, parsed

    def create_record(self, test_id="T-001", **overrides):
        body = {"mode": "simulation", "result": "0.5", "unit": "nArms",
                "status": "Pass", "test_date": "2025-01-02", "notes": "note"}
        body.update(overrides)
        status, _, parsed = self.json_request("POST", f"/api/projects/afe/tests/{test_id}/records", body)
        self.assertEqual(status, 201, parsed)
        return parsed["record"]

    def multipart(self, fields, files):
        boundary = "----AFETest" + uuid.uuid4().hex
        chunks = []
        for key, value in fields.items():
            chunks.append(f"--{boundary}\r\n".encode())
            chunks.append(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
            chunks.append(str(value).encode("utf-8"))
            chunks.append(b"\r\n")
        for field, filename, payload in files:
            chunks.append(f"--{boundary}\r\n".encode())
            chunks.append(
                f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode())
            chunks.append(b"Content-Type: application/octet-stream\r\n\r\n")
            chunks.append(payload)
            chunks.append(b"\r\n")
        chunks.append(f"--{boundary}--\r\n".encode())
        return b"".join(chunks), f"multipart/form-data; boundary={boundary}"

    def upload(self, record_id, files, kind="data"):
        body, ctype = self.multipart({"kind": kind}, files)
        return self.json_request("POST", f"/api/records/{record_id}/attachments",
                                 raw_body=body, headers={"Content-Type": ctype})

    # ---- health / security ----------------------------------------------
    def test_health(self):
        status, headers, parsed = self.json_request("GET", "/api/health")
        self.assertEqual(status, 200)
        self.assertEqual(parsed, {"ok": True})
        self.assertEqual(headers.get("X-Content-Type-Options"), "nosniff")

    def test_host_header_rejected(self):
        import socket
        sock = socket.create_connection(("127.0.0.1", self.port), timeout=5)
        try:
            sock.sendall(b"GET /api/health HTTP/1.1\r\nHost: evil.example\r\n"
                         b"Connection: close\r\n\r\n")
            data = b""
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                data += chunk
        finally:
            sock.close()
        self.assertTrue(data.startswith(b"HTTP/1.1 403"), data[:80])
        self.assertIn(b"localhost", data)

    def test_lan_host_and_origin_allowed(self):
        """LAN access (host browser -> VM) must pass Host/Origin checks."""
        lan_host = "192.0.2.10"     # TEST-NET-1, representative private network address
        added = lan_host not in server.ALLOWED_HOSTS
        if added:
            server.ALLOWED_HOSTS.add(lan_host)
        try:
            import socket
            sock = socket.create_connection(("127.0.0.1", self.port), timeout=5)
            try:
                payload = (
                    f"GET /api/health HTTP/1.1\r\nHost: {lan_host}:{self.port}\r\n"
                    "Connection: close\r\n\r\n"
                ).encode("ascii")
                sock.sendall(payload)
                data = b""
                while True:
                    chunk = sock.recv(4096)
                    if not chunk:
                        break
                    data += chunk
            finally:
                sock.close()
            self.assertTrue(data.startswith(b"HTTP/1.1 200"), data[:80])
            status, _, _ = self.json_request(
                "POST", "/api/projects/afe/tests/T-001/records", {"mode": "simulation"},
                headers={"Origin": f"http://{lan_host}:{self.port}"})
            self.assertEqual(status, 201)
        finally:
            if added:
                server.ALLOWED_HOSTS.discard(lan_host)

    def test_configure_allowed_hosts_from_env(self):
        old = os.environ.get("AFE_ALLOWED_HOSTS")
        os.environ["AFE_ALLOWED_HOSTS"] = "10.1.2.3, test-flow.local"
        try:
            allowed = server.configure_allowed_hosts("0.0.0.0")
            self.assertIn("10.1.2.3", allowed)
            self.assertIn("test-flow.local", allowed)
            self.assertIn("127.0.0.1", allowed)
        finally:
            if old is None:
                os.environ.pop("AFE_ALLOWED_HOSTS", None)
            else:
                os.environ["AFE_ALLOWED_HOSTS"] = old
            server.ALLOWED_HOSTS.discard("10.1.2.3")
            server.ALLOWED_HOSTS.discard("test-flow.local")

    def test_origin_rejected_for_writes(self):
        body = {"mode": "simulation"}
        status, _, parsed = self.json_request(
            "POST", "/api/projects/afe/tests/T-001/records", body,
            headers={"Origin": "http://evil.example"})
        self.assertEqual(status, 403)
        self.assertIn("error", parsed)

    def test_origin_allowed_for_writes(self):
        status, _, _ = self.json_request(
            "POST", "/api/projects/afe/tests/T-001/records", {"mode": "simulation"},
            headers={"Origin": f"http://127.0.0.1:{self.port}"})
        self.assertEqual(status, 201)

    def test_unknown_api(self):
        status, _, parsed = self.json_request("GET", "/api/nope")
        self.assertEqual(status, 404)
        self.assertIn("error", parsed)

    def test_method_not_allowed(self):
        status, _, parsed = self.json_request("POST", "/api/health", {})
        self.assertEqual(status, 405)
        self.assertIn("error", parsed)

    # ---- bootstrap -------------------------------------------------------
    def test_bootstrap(self):
        status, _, parsed = self.json_request("GET", "/api/bootstrap")
        self.assertEqual(status, 200)
        self.assertEqual(set(parsed), {"project", "tests", "summary", "instructions", "source"})
        self.assertEqual([t["id"] for t in parsed["tests"]], ["T-001", "T-002"])
        self.assertEqual(parsed["tests"][0]["source_result"], "37")
        self.assertEqual(parsed["summary"], SEED["summary"])
        self.assertEqual(parsed["instructions"], SEED["instructions"])
        self.assertEqual(parsed["source"], "unit-test-seed")

    # ---- records ---------------------------------------------------------
    def test_create_record_validation(self):
        for body in ({}, {"mode": "bad"}, {"mode": "simulation", "status": "OK"},
                     {"mode": "simulation", "test_date": "2025-02-30"},
                     {"mode": "simulation", "test_date": "2025-1-2"},
                     {"mode": "simulation", "result": 5}):
            status, _, parsed = self.json_request(
                "POST", "/api/projects/afe/tests/T-001/records", body)
            self.assertEqual(status, 400, (body, parsed))
            self.assertIn("error", parsed)
        status, _, _ = self.json_request("POST", "/api/projects/afe/tests/NOPE/records",
                                         {"mode": "simulation"})
        self.assertEqual(status, 404)
        status, _, _ = self.json_request(
            "POST", "/api/projects/afe/tests/T-001/records",
            raw_body=b"{bad json", headers={"Content-Type": "application/json"})
        self.assertEqual(status, 400)
        status, _, _ = self.json_request(
            "POST", "/api/projects/afe/tests/T-001/records", raw_body=b"x",
            headers={"Content-Type": "text/plain"})
        self.assertEqual(status, 415)

    def test_create_and_list_records(self):
        record = self.create_record()
        self.assertEqual(record["test_id"], "T-001")
        self.assertEqual(record["mode"], "simulation")
        self.assertEqual(record["attachments"], [])
        self.assertIn("created_at", record)
        second = self.create_record(mode="measurement", status="", test_date="")
        self.assertEqual(second["status"], "")
        status, _, parsed = self.json_request("GET", "/api/projects/afe/tests/T-001/records")
        self.assertEqual(status, 200)
        self.assertEqual([r["id"] for r in parsed["records"]], [record["id"], second["id"]])
        self.assertEqual(parsed["records"][1]["mode"], "measurement")
        status, _, _ = self.json_request("GET", "/api/projects/afe/tests/NOPE/records")
        self.assertEqual(status, 404)

    def test_update_record(self):
        record = self.create_record()
        status, _, parsed = self.json_request(
            "PUT", f"/api/records/{record['id']}",
            {"mode": "measurement", "status": "Fail", "test_date": "2025-03-04"})
        self.assertEqual(status, 200)
        updated = parsed["record"]
        self.assertEqual(updated["mode"], "measurement")
        self.assertEqual(updated["status"], "Fail")
        self.assertEqual(updated["test_date"], "2025-03-04")
        self.assertEqual(updated["notes"], "note")  # omitted fields preserved
        self.assertEqual(updated["created_at"], record["created_at"])
        status, _, _ = self.json_request("PUT", "/api/records/999999",
                                         {"mode": "simulation"})
        self.assertEqual(status, 404)

    def test_restart_keeps_records_and_seed_is_not_reapplied(self):
        record = self.create_record()
        changed = json.loads(json.dumps(SEED, ensure_ascii=False))
        changed["tests"][0]["title"] = "更新后的标题"
        (self.root / "seed.json").write_text(
            json.dumps(changed, ensure_ascii=False), encoding="utf-8")
        server.init_db()
        server.seed_default_project()
        server.seed_default_project()
        conn = sqlite3.connect(str(server.CONFIG.db_path))
        try:
            count = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
            title = conn.execute(
                "SELECT title FROM tests WHERE project_id='afe' AND id='T-001'").fetchone()[0]
            tests = conn.execute("SELECT COUNT(*) FROM tests").fetchone()[0]
            projects = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(count, 1)
        self.assertEqual(projects, 1)
        self.assertEqual(tests, 2)
        self.assertEqual(title, "输入参考噪声")
        status, _, parsed = self.json_request("GET", "/api/projects/afe/tests/T-001/records")
        self.assertEqual(status, 200)
        self.assertEqual(parsed["records"][0]["id"], record["id"])

    # ---- editable guidance ------------------------------------------------
    def test_edit_test_guidance(self):
        status, _, parsed = self.json_request(
            "PUT", "/api/projects/afe/tests/T-001",
            {"purpose": "新的目的", "instruments": "新仪器"})
        self.assertEqual(status, 200, parsed)
        test = parsed["test"]
        self.assertEqual(test["purpose"], "新的目的")
        self.assertEqual(test["instruments"], "新仪器")
        self.assertEqual(test["title"], "输入参考噪声")
        self.assertEqual(test["stimulus"], "暗场")

        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        stored = [t for t in boot["tests"] if t["id"] == "T-001"][0]
        self.assertEqual(stored["purpose"], "新的目的")

        # partial update keeps other edited fields
        status, _, parsed = self.json_request(
            "PUT", "/api/projects/afe/tests/T-001", {"metrics": "新的指标"})
        self.assertEqual(parsed["test"]["purpose"], "新的目的")
        self.assertEqual(parsed["test"]["metrics"], "新的指标")

        # values are trimmed and can be cleared
        status, _, parsed = self.json_request(
            "PUT", "/api/projects/afe/tests/T-001", {"metrics": "   "})
        self.assertEqual(parsed["test"]["metrics"], "")

        # read-only columns are ignored, not applied
        status, _, parsed = self.json_request(
            "PUT", "/api/projects/afe/tests/T-001", {"purpose": "保留", "title": "hack"})
        self.assertEqual(status, 200)
        self.assertEqual(parsed["test"]["title"], "输入参考噪声")

        for body in ({"purpose": 5}, {}, {"purpose": "x" * 4001}):
            status, _, parsed = self.json_request(
                "PUT", "/api/projects/afe/tests/T-001", body)
            self.assertEqual(status, 400, (body, parsed))

        status, _, _ = self.json_request(
            "PUT", "/api/projects/afe/tests/NOPE", {"purpose": "x"})
        self.assertEqual(status, 404)
        status, _, _ = self.json_request(
            "PUT", "/api/projects/nope/tests/T-001", {"purpose": "x"})
        self.assertEqual(status, 404)
        status, _, _ = self.json_request(
            "PUT", "/api/projects/afe/tests/T-001", raw_body=b"{bad",
            headers={"Content-Type": "application/json"})
        self.assertEqual(status, 400)

    # ---- add / delete test items ------------------------------------------
    def test_create_test_item(self):
        status, _, parsed = self.json_request(
            "POST", "/api/projects/afe/tests", {"title": "新测试项"})
        self.assertEqual(status, 201, parsed)
        test = parsed["test"]
        self.assertEqual(test["title"], "新测试项")
        self.assertEqual(test["priority"], "P0")
        self.assertEqual(test["module"], "未分组")
        self.assertEqual(test["id"], "T-003")   # suggested after T-001 / T-002

        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        self.assertEqual([item["id"] for item in boot["tests"]], ["T-001", "T-002", "T-003"])

        status, _, parsed = self.json_request(
            "POST", "/api/projects/afe/tests",
            {"id": "X-9", "title": "带指导", "priority": "P1", "module": "模块A",
             "purpose": "目的", "instruments": "示波器"})
        self.assertEqual(status, 201, parsed)
        self.assertEqual(parsed["test"]["id"], "X-9")
        self.assertEqual(parsed["test"]["purpose"], "目的")
        self.assertEqual(parsed["test"]["instruments"], "示波器")

        status, _, parsed = self.json_request(
            "POST", "/api/projects/afe/tests", {"id": "X-9", "title": "重复"})
        self.assertEqual(status, 409, parsed)

        for body in ({"title": ""}, {}, {"id": "bad id!", "title": "t"},
                     {"title": "x" * 201}, {"title": 5}):
            status, _, parsed = self.json_request("POST", "/api/projects/afe/tests", body)
            self.assertEqual(status, 400, (body, parsed))

        status, _, _ = self.json_request("POST", "/api/projects/nope/tests", {"title": "t"})
        self.assertEqual(status, 404)

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_delete_test_item_cascades(self):
        kept = self.create_record(test_id="T-002", mode="measurement")
        record = self.create_record()
        status, _, parsed = self.upload(record["id"], [("files", "a.xlsx", _xlsx_bytes())])
        self.assertEqual(status, 200, parsed)

        status, _, parsed = self.json_request("DELETE", "/api/projects/afe/tests/T-001")
        self.assertEqual(status, 200, parsed)
        self.assertEqual(parsed["id"], "T-001")
        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        self.assertEqual([item["id"] for item in boot["tests"]], ["T-002"])

        conn = sqlite3.connect(str(server.CONFIG.db_path))
        try:
            records = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
            attachments = conn.execute("SELECT COUNT(*) FROM attachments").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(records, 1)          # only T-002's record survives
        self.assertEqual(attachments, 0)
        self.assertEqual(list(server.CONFIG.uploads_dir.iterdir()), [])

        status, _, kept_records = self.json_request("GET", "/api/projects/afe/tests/T-002/records")
        self.assertEqual([item["id"] for item in kept_records["records"]], [kept["id"]])

        status, _, _ = self.json_request("DELETE", "/api/projects/afe/tests/T-001")
        self.assertEqual(status, 404)

    # ---- clipboard (copy / cut / paste) -----------------------------------
    def _new_project(self, name="第二项目"):
        data = _project_workbook()
        body, ctype = self.multipart({"name": name}, [("file", "p2.xlsx", data)])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 201, parsed)
        return parsed["project"]["id"]

    def test_clipboard_copy_and_paste_within_project(self):
        status, _, parsed = self.json_request(
            "POST", "/api/clipboard", {"project_id": "afe", "test_ids": ["T-001"], "op": "copy"})
        self.assertEqual(status, 200, parsed)
        self.assertEqual(parsed["count"], 1)
        self.assertEqual(parsed["items"][0]["test_id"], "T-001")
        self.assertEqual(parsed["items"][0]["test"]["title"], "输入参考噪声")

        status, _, parsed = self.json_request("POST", "/api/clipboard/paste", {"module": "复制模块"})
        self.assertEqual(status, 200, parsed)
        entry = parsed["pasted"][0]
        self.assertEqual(entry["mode"], "copy")
        self.assertEqual(entry["test_id"], "T-003")      # T-001 taken -> next id
        self.assertEqual(entry["project_id"], "afe")
        self.assertEqual(parsed["clipboard"]["count"], 1)  # copy stays on the clipboard

        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        copied = [item for item in boot["tests"] if item["id"] == "T-003"][0]
        self.assertEqual(copied["module"], "复制模块")
        self.assertEqual(copied["title"], "输入参考噪声")
        self.assertEqual(copied["purpose"], "验证噪声")
        self.assertEqual(len(boot["tests"]), 3)

        # pasting again makes another copy
        status, _, parsed = self.json_request("POST", "/api/clipboard/paste", {"module": "复制模块"})
        self.assertEqual(parsed["pasted"][0]["test_id"], "T-004")

    def test_clipboard_cut_changes_module_only(self):
        status, _, parsed = self.json_request(
            "POST", "/api/clipboard", {"project_id": "afe", "test_ids": ["T-002"], "op": "cut"})
        self.assertEqual(status, 200, parsed)
        status, _, parsed = self.json_request("POST", "/api/clipboard/paste", {"module": "搬移模块"})
        self.assertEqual(status, 200, parsed)
        entry = parsed["pasted"][0]
        self.assertEqual(entry["mode"], "move-module")
        self.assertEqual(entry["test_id"], "T-002")
        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        moved = [item for item in boot["tests"] if item["id"] == "T-002"][0]
        self.assertEqual(moved["module"], "搬移模块")
        self.assertEqual(parsed["clipboard"]["count"], 0)   # cut is consumed

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_clipboard_cut_moves_records_to_other_project(self):
        record = self.create_record()
        status, _, parsed = self.upload(record["id"], [("files", "a.xlsx", _xlsx_bytes())])
        self.assertEqual(status, 200, parsed)
        other = self._new_project()

        status, _, parsed = self.json_request(
            "POST", "/api/clipboard", {"project_id": "afe", "test_ids": ["T-001"], "op": "cut"})
        self.assertEqual(status, 200, parsed)
        status, _, parsed = self.json_request(
            "POST", "/api/clipboard/paste", {"project_id": other, "module": "搬入"})
        self.assertEqual(status, 200, parsed)
        entry = parsed["pasted"][0]
        self.assertEqual(entry["mode"], "move")
        self.assertEqual(entry["records"], 1)

        status, _, boot = self.json_request("GET", f"/api/projects/{other}/bootstrap")
        moved = [item for item in boot["tests"] if item["id"] == "T-001"][0]
        self.assertEqual(moved["module"], "搬入")
        status, _, records = self.json_request("GET", f"/api/projects/{other}/tests/T-001/records")
        self.assertEqual([item["id"] for item in records["records"]], [record["id"]])
        self.assertEqual(len(records["records"][0]["attachments"]), 1)

        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        self.assertEqual([item["id"] for item in boot["tests"]], ["T-002"])
        self.assertEqual(len(list(server.CONFIG.uploads_dir.iterdir())), 1)  # file stayed

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_clipboard_cross_project_copy(self):
        other = self._new_project()
        status, _, parsed = self.json_request(
            "POST", "/api/clipboard", {"project_id": "afe", "test_ids": ["T-001", "T-002"], "op": "copy"})
        self.assertEqual(parsed["count"], 2)
        status, _, parsed = self.json_request(
            "POST", "/api/clipboard/paste", {"project_id": other, "module": "外部"})
        self.assertEqual(status, 200, parsed)
        self.assertEqual([item["test_id"] for item in parsed["pasted"]], ["T-001", "T-002"])
        status, _, boot = self.json_request("GET", f"/api/projects/{other}/bootstrap")
        self.assertEqual([item["id"] for item in boot["tests"]], ["X-01", "X-02", "P03", "T-001", "T-002"])
        copied = [item for item in boot["tests"] if item["id"] == "T-001"][0]
        self.assertEqual(copied["module"], "外部")
        self.assertEqual(copied["title"], "输入参考噪声")
        # source project untouched
        status, _, boot = self.json_request("GET", "/api/projects/afe/bootstrap")
        self.assertEqual(len(boot["tests"]), 2)

    def test_clipboard_validation_and_clear(self):
        status, _, parsed = self.json_request("POST", "/api/clipboard/paste", {})
        self.assertEqual(status, 400, parsed)          # empty clipboard
        self.assertIn("空", parsed["error"])

        for body in ({"project_id": "afe", "test_ids": [], "op": "copy"},
                     {"project_id": "afe", "test_ids": ["T-001"], "op": "move"},
                     {"project_id": "afe", "test_ids": "T-001", "op": "copy"}):
            status, _, parsed = self.json_request("POST", "/api/clipboard", body)
            self.assertEqual(status, 400, (body, parsed))

        for body in ({"project_id": "nope", "test_ids": ["T-001"], "op": "copy"},
                     {"project_id": "afe", "test_ids": ["NOPE"], "op": "copy"}):
            status, _, parsed = self.json_request("POST", "/api/clipboard", body)
            self.assertEqual(status, 404, (body, parsed))

        status, _, parsed = self.json_request(
            "POST", "/api/clipboard", {"project_id": "afe", "test_ids": ["T-001"], "op": "copy"})
        self.assertEqual(parsed["count"], 1)

        status, _, parsed = self.json_request("POST", "/api/clipboard/paste",
                                              {"project_id": "nope"})
        self.assertEqual(status, 404, parsed)          # unknown target project

        status, _, parsed = self.json_request("DELETE", "/api/clipboard")
        self.assertEqual((status, parsed["count"]), (200, 0))
        status, _, parsed = self.json_request("GET", "/api/clipboard")
        self.assertEqual(parsed["count"], 0)

    # ---- fresh install (no seed) -----------------------------------------
    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_fresh_install_without_seed(self):
        (self.root / "seed.json").unlink()
        for suffix in ("", "-wal", "-shm"):
            leftover = Path(str(server.CONFIG.db_path) + suffix)
            if leftover.exists():
                leftover.unlink()
        server.init_db()
        server.seed_default_project()          # no seed.json -> no project is created

        status, _, parsed = self.json_request("GET", "/api/projects")
        self.assertEqual(status, 200)
        self.assertEqual(parsed["projects"], [])
        status, _, _ = self.json_request("GET", "/api/bootstrap")
        self.assertEqual(status, 404)

        body, ctype = self.multipart({"name": "第一个项目"},
                                     [("file", "first.xlsx", _project_workbook())])
        status, _, created = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 201, created)
        pid = created["project"]["id"]
        status, _, boot = self.json_request("GET", f"/api/projects/{pid}/bootstrap")
        self.assertEqual(status, 200)
        self.assertEqual(len(boot["tests"]), 3)

    # ---- attachments -----------------------------------------------------
    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_upload_download_delete(self):
        record = self.create_record()
        data = _xlsx_bytes("hello")
        status, _, parsed = self.upload(
            record["id"], [("files", "../evil.xlsx", data),
                           ("files", "second.xlsx", data)])
        self.assertEqual(status, 200, parsed)
        attachments = parsed["attachments"]
        self.assertEqual(len(attachments), 2)
        self.assertEqual(attachments[0]["name"], "evil.xlsx")
        self.assertEqual(attachments[0]["kind"], "data")
        self.assertEqual(attachments[0]["size"], len(data))
        self.assertEqual(attachments[0]["url"],
                         f"/api/attachments/{attachments[0]['id']}/file")

        status, headers, body = self.request(
            "GET", attachments[0]["url"])
        self.assertEqual(status, 200)
        self.assertEqual(body, data)
        self.assertEqual(headers.get("Content-Type"), "application/octet-stream")
        self.assertTrue(headers.get("Content-Disposition", "").startswith("attachment;"))
        self.assertEqual(headers.get("X-Content-Type-Options"), "nosniff")
        self.assertIn("sandbox", headers.get("Content-Security-Policy", ""))

        status, _, parsed = self.json_request("GET", "/api/projects/afe/tests/T-001/records")
        self.assertEqual(len(parsed["records"][0]["attachments"]), 2)

        status, _, parsed = self.json_request(
            "DELETE", f"/api/attachments/{attachments[0]['id']}")
        self.assertEqual((status, parsed), (200, {"ok": True}))
        status, _, _ = self.request("GET", attachments[0]["url"])
        self.assertEqual(status, 404)
        status, _, _ = self.json_request(
            "DELETE", f"/api/attachments/{attachments[0]['id']}")
        self.assertEqual(status, 404)

    def test_upload_image_inline(self):
        record = self.create_record()
        status, _, parsed = self.upload(
            record["id"], [("files", "shot.jpg", _jpg_bytes())], kind="image")
        self.assertEqual(status, 200, parsed)
        attachment = parsed["attachments"][0]
        self.assertEqual(attachment["kind"], "image")
        status, headers, body = self.request("GET", attachment["url"])
        self.assertEqual(status, 200)
        self.assertEqual(headers.get("Content-Type"), "image/jpeg")
        self.assertTrue(headers.get("Content-Disposition", "").startswith("inline;"))
        self.assertEqual(body, _jpg_bytes())

    def test_upload_validation(self):
        record = self.create_record()
        cases = [
            ([("files", "bad.txt", b"hello")], "data"),
            ([("files", "bad.xlsx", b"not a zip")], "data"),
            ([("files", "image.png", b"\x89PNG\r\n\x1a\n")], "image"),
            ([("files", "data.xlsx", _xlsx_bytes())], "image"),
        ]
        for files, kind in cases:
            status, _, parsed = self.upload(record["id"], files, kind=kind)
            self.assertEqual(status, 400, (files[0][1], kind, parsed))
            self.assertIn("error", parsed)
        status, _, parsed = self.upload(record["id"], [])
        self.assertEqual(status, 400)
        status, _, _ = self.upload(999999, [("files", "a.xlsx", _xlsx_bytes())])
        self.assertEqual(status, 404)

    def test_per_file_limit(self):
        record = self.create_record()
        original = server.MAX_FILE_BYTES
        server.MAX_FILE_BYTES = 8
        try:
            status, _, parsed = self.upload(
                record["id"], [("files", "big.jpg", b"\xff\xd8\xff" + b"0" * 8)],
                kind="image")
        finally:
            server.MAX_FILE_BYTES = original
        self.assertEqual(status, 400)
        self.assertIn("50 MiB", parsed["error"])

    def test_attachment_not_found_and_traversal(self):
        for path in ("/api/attachments/999999/file",
                     "/api/attachments/1%2F..%2F..%2Fseed.json/file",
                     "/api/attachments/1/file/../../seed.json"):
            status, _, _ = self.json_request("GET", path)
            self.assertEqual(status, 404, path)

    # ---- projects --------------------------------------------------------
    def test_projects_list_and_default_bootstrap(self):
        status, _, parsed = self.json_request("GET", "/api/projects")
        self.assertEqual(status, 200)
        self.assertEqual(len(parsed["projects"]), 1)
        project = parsed["projects"][0]
        self.assertEqual(project["id"], "afe")
        self.assertEqual(project["test_count"], 2)
        self.assertEqual(project["record_count"], 0)
        self.assertIn("name", project)

        status, _, parsed = self.json_request("GET", "/api/projects/afe/bootstrap")
        self.assertEqual(status, 200)
        self.assertEqual(parsed["project"]["id"], "afe")
        self.assertEqual([t["id"] for t in parsed["tests"]], ["T-001", "T-002"])

        status, _, _ = self.json_request("GET", "/api/projects/nope/bootstrap")
        self.assertEqual(status, 404)

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_create_project_from_workbook(self):
        data = _project_workbook()
        body, ctype = self.multipart({"name": "新项目 Alpha"},
                                     [("file", "alpha.xlsx", data)])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 201, parsed)
        project = parsed["project"]
        self.assertEqual(project["name"], "新项目 Alpha")
        self.assertEqual(parsed["tests"], 3)
        self.assertEqual(parsed["summary"], 1)
        self.assertEqual(parsed["instructions"], 1)

        status, _, boot = self.json_request("GET", f"/api/projects/{project['id']}/bootstrap")
        self.assertEqual(status, 200)
        ids = [t["id"] for t in boot["tests"]]
        self.assertEqual(ids[:2], ["X-01", "X-02"])
        self.assertEqual(len(ids), 3)          # row without 序号 gets a generated id
        self.assertEqual(boot["tests"][0]["title"], "增益")
        self.assertEqual(boot["tests"][0]["stimulus"], "正弦")
        self.assertEqual(boot["tests"][0]["criteria"], ">=90dB")
        self.assertEqual(boot["summary"][0]["模块"], "TIA")
        self.assertEqual(boot["instructions"][0]["项目"], "使用方法")

        # without a name field the file stem becomes the project name
        body, ctype = self.multipart({}, [("file", "beta.xlsx", data)])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 201, parsed)
        self.assertEqual(parsed["project"]["name"], "beta")

        status, _, listed = self.json_request("GET", "/api/projects")
        self.assertEqual(len(listed["projects"]), 3)

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_create_project_rejects_bad_input(self):
        wb = Workbook()
        wb.active["A1"] = "hello"
        buffer = io.BytesIO()
        wb.save(buffer)
        body, ctype = self.multipart({}, [("file", "bad.xlsx", buffer.getvalue())])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 400, parsed)
        self.assertIn("error", parsed)

        body, ctype = self.multipart({}, [("file", "old.xls", b"\xd0\xcf\x11\xe0" + b"\x00" * 64)])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 400)
        self.assertIn(".xlsx", parsed["error"])

        body, ctype = self.multipart({}, [("file", "notes.txt", b"hello")])
        status, _, _ = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 400)

        body, ctype = self.multipart({}, [])
        status, _, _ = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 400)

        status, _, _ = self.json_request(
            "POST", "/api/projects", raw_body=b"{}",
            headers={"Content-Type": "application/json"})
        self.assertEqual(status, 415)

    @unittest.skipIf(Workbook is None, "openpyxl not installed")
    def test_project_isolation_and_delete(self):
        data = _project_workbook()
        body, ctype = self.multipart({"name": "隔离测试"},
                                     [("file", "iso.xlsx", data)])
        status, _, parsed = self.json_request(
            "POST", "/api/projects", raw_body=body, headers={"Content-Type": ctype})
        self.assertEqual(status, 201, parsed)
        other = parsed["project"]["id"]

        afe_record = self.create_record()
        status, _, created = self.json_request(
            "POST", f"/api/projects/{other}/tests/X-01/records",
            {"mode": "measurement", "status": "Fail"})
        self.assertEqual(status, 201, created)

        status, _, listed = self.json_request("GET", f"/api/projects/{other}/tests/X-01/records")
        self.assertEqual(len(listed["records"]), 1)
        status, _, listed = self.json_request("GET", "/api/projects/afe/tests/T-001/records")
        self.assertEqual([r["id"] for r in listed["records"]], [afe_record["id"]])
        # same test id in another project does not resolve across projects
        status, _, _ = self.json_request("GET", f"/api/projects/{other}/tests/T-001/records")
        self.assertEqual(status, 404)

        status, _, parsed = self.json_request("DELETE", f"/api/projects/{other}")
        self.assertEqual((status, parsed["ok"]), (200, True))
        status, _, _ = self.json_request("GET", f"/api/projects/{other}/bootstrap")
        self.assertEqual(status, 404)
        conn = sqlite3.connect(str(server.CONFIG.db_path))
        try:
            left = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(left, 1)

        status, _, parsed = self.json_request("DELETE", "/api/projects/afe")
        self.assertEqual(status, 400)
        self.assertIn("至少", parsed["error"])

    # ---- static ----------------------------------------------------------
    def test_static_files_and_traversal(self):
        status, headers, body = self.request("GET", "/")
        self.assertEqual(status, 200)
        self.assertIn(b"TestFlow", body)
        self.assertIn(b"default-src", headers.get("Content-Security-Policy", "").encode())
        status, headers, body = self.request("GET", "/app.js")
        self.assertEqual(status, 200)
        self.assertIn("javascript", headers.get("Content-Type", ""))
        for path in ("/../seed.json", "/%2e%2e/seed.json", "/..%2fseed.json"):
            status, _, _ = self.request("GET", path)
            self.assertEqual(status, 404, path)
        status, _, _ = self.request("GET", "/missing.txt")
        self.assertEqual(status, 404)
        status, _, _ = self.request("GET", "/some/route")
        self.assertEqual(status, 200)  # SPA fallback to index.html

    # ---- export / backup -------------------------------------------------
    @unittest.skipIf(load_workbook is None, "openpyxl not installed")
    def test_export_xlsx(self):
        self.create_record()
        self.create_record(mode="measurement", status="Fail")
        status, headers, body = self.request("GET", "/api/projects/afe/export.xlsx")
        self.assertEqual(status, 200)
        self.assertIn("spreadsheetml", headers.get("Content-Type", ""))
        self.assertIn("attachment", headers.get("Content-Disposition", ""))
        wb = load_workbook(io.BytesIO(body))
        ws = wb.active
        self.assertEqual(ws.max_row, 3)
        self.assertEqual(ws.cell(row=1, column=1).value, "记录ID")
        modes = {ws.cell(row=2, column=6).value, ws.cell(row=3, column=6).value}
        self.assertEqual(modes, {"仿真", "实测"})

    def test_backup_zip(self):
        record = self.create_record()
        status, _, parsed = self.upload(
            record["id"], [("files", "data.xlsx", _xlsx_bytes())])
        self.assertEqual(status, 200, parsed)
        status, headers, body = self.request("GET", "/api/backup.zip")
        self.assertEqual(status, 200)
        self.assertIn("application/zip", headers.get("Content-Type", ""))
        with zipfile.ZipFile(io.BytesIO(body)) as zf:
            names = zf.namelist()
            self.assertIn("data/afe-testflow.db", names)
            self.assertIn("seed.json", names)
            self.assertIn("backup-info.json", names)
            uploads = [n for n in names if n.startswith("data/uploads/")]
            self.assertEqual(len(uploads), 1)
            db_bytes = zf.read("data/afe-testflow.db")
            self.assertTrue(db_bytes.startswith(b"SQLite format 3"))
            info = json.loads(zf.read("backup-info.json").decode("utf-8"))
            self.assertIn("created_at", info)

    # ---- limits ----------------------------------------------------------
    def test_request_size_limit(self):
        import socket
        sock = socket.create_connection(("127.0.0.1", self.port), timeout=5)
        try:
            sock.sendall(
                b"POST /api/projects/afe/tests/T-001/records HTTP/1.1\r\n"
                + f"Host: 127.0.0.1:{self.port}\r\n".encode()
                + b"Content-Type: application/json\r\n"
                + f"Content-Length: {server.MAX_REQUEST_BYTES + 1}\r\n".encode()
                + b"Connection: close\r\n\r\n")
            data = b""
            while True:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                data += chunk
        finally:
            sock.close()
        self.assertTrue(data.startswith(b"HTTP/1.1 413"), data[:80])
        self.assertIn(b"error", data)

    def test_sanitize_filename(self):
        self.assertEqual(server._sanitize_filename("../../x.xlsx", "f.xlsx"), "x.xlsx")
        self.assertEqual(server._sanitize_filename("..\\evil.jpg", "f.jpg"), "evil.jpg")
        self.assertEqual(server._sanitize_filename("", "fallback.xlsx"), "fallback.xlsx")
        name = server._sanitize_filename("长" * 300 + ".xlsx", "f.xlsx")
        self.assertLessEqual(len(name), server.MAX_FILENAME_CHARS)
        self.assertTrue(name.endswith(".xlsx"))


if __name__ == "__main__":
    unittest.main(verbosity=2)

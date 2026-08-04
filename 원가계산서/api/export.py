# -*- coding: utf-8 -*-
"""
Vercel 서버리스 함수: 정규화 레코드 -> 스타일이 적용된 표준 엑셀 다운로드.

POST /api/export   body: {"records": [ {정규화레코드}, ... ]}
반환: xlsx 바이너리 (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)

파일명은 클라이언트가 Blob 다운로드 시 지정하므로, 헤더에는 ASCII fallback 만 둔다.
"""
import os
import sys
import io
import json
import traceback
from http.server import BaseHTTPRequestHandler

sys.path.append(os.path.dirname(__file__))
from _lib import exporter  # noqa: E402


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            body = self.rfile.read(length) if length else b""
            data = json.loads(body.decode("utf-8")) if body else {}
            records = data.get("records") or []
            if not records:
                return self._json(400, {"ok": False, "error": "정규화할 데이터가 없습니다."})

            wb = exporter.build_workbook(records)
            buf = io.BytesIO()
            wb.save(buf)
            xlsx = buf.getvalue()

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="jebiyul.xlsx"')
            self.send_header("Content-Length", str(len(xlsx)))
            self.end_headers()
            self.wfile.write(xlsx)
        except Exception as e:  # noqa
            self._json(500, {"ok": False, "error": str(e),
                             "trace": traceback.format_exc()})

    def _json(self, code, payload):
        b = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

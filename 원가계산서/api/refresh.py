# -*- coding: utf-8 -*-
"""
Vercel 서버리스 함수: 제비율 최신화(크롤링) + 정규화 파싱.

POST /api/refresh   body: {"years": 3, "ref_date": "YYYY-MM-DD"}
GET  /api/refresh?years=3&ref_date=YYYY-MM-DD   (편의)

반환(JSON):
  {ok, downloaded[], posts, files[], count, columns[], records[], preview[], errors[], logs[]}

- 크롤링 결과 원본 엑셀은 서버리스 임시경로(/tmp)에 저장 후 즉시 파싱하여
  정규화 레코드(records)로 반환한다. (서버리스는 상태를 보존하지 않으므로 파일은 응답에 담지 않고
  정규화 결과만 클라이언트로 넘긴다. 엑셀 다운로드는 /api/export 가 records 를 받아 생성.)
"""
import os
import sys
import json
import tempfile
import datetime
import traceback
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

sys.path.append(os.path.dirname(__file__))          # api/ 를 import 경로에 추가
from _lib import crawler, parser as pparser, exporter  # noqa: E402


def _run(years, ref):
    try:
        ref_date = datetime.date.fromisoformat(ref) if ref else datetime.date.today()
    except (ValueError, TypeError):
        ref_date = datetime.date.today()

    years = max(1, min(int(years or 3), 10))
    logs = []
    dest = tempfile.mkdtemp(prefix="jebiyul_")

    res = crawler.crawl(dest, years=years, ref_date=ref_date, max_pages=6,
                        log=lambda *a: logs.append(" ".join(map(str, a))))

    files_meta, parse_targets = [], []
    for fn in sorted(os.listdir(dest)):
        if not fn.lower().endswith((".xlsx", ".xls")) or fn.startswith("~$"):
            continue
        g, d = pparser.parse_filename(fn)
        recognized = bool(g and d)
        files_meta.append({
            "file": fn,
            "공종": g or "(미인식)",
            "적용기준일": d.isoformat() if d else "(미인식)",
            "recognized": recognized,
        })
        if recognized:
            parse_targets.append(os.path.join(dest, fn))

    records, errors = pparser.parse_files(parse_targets)
    return {
        "ok": True,
        "downloaded": [m["file"] for m in files_meta],
        "posts": res.get("posts", 0),
        "files": files_meta,
        "count": len(records),
        "columns": [c[0] for c in exporter.COLUMNS],
        "records": records,
        "preview": records[:300],
        "errors": errors,
        "logs": logs,
    }


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            body = self.rfile.read(length) if length else b""
            data = json.loads(body.decode("utf-8")) if body else {}
            self._json(200, _run(data.get("years", 3), data.get("ref_date")))
        except Exception as e:  # noqa
            self._json(500, {"ok": False, "error": str(e),
                             "trace": traceback.format_exc()})

    def do_GET(self):
        try:
            q = parse_qs(urlparse(self.path).query)
            self._json(200, _run(q.get("years", ["3"])[0], q.get("ref_date", [None])[0]))
        except Exception as e:  # noqa
            self._json(500, {"ok": False, "error": str(e),
                             "trace": traceback.format_exc()})

    def _json(self, code, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

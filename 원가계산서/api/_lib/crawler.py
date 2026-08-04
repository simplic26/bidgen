# -*- coding: utf-8 -*-
"""
조달청(pps.go.kr) 시설공사 게시판에서 '제비율' 게시물의
간접공사비 적용기준 엑셀 파일(토목/건축)을 자동 수집한다.

동작 확인된 엔드포인트
  - 목록/검색 : POST https://www.pps.go.kr/kor/bbs/list.do?key=00038
                data = {key, pageIndex, orderBy, sc=BBS_SJ, sw=제비율}
  - 상세      : GET  https://www.pps.go.kr/kor/bbs/view.do?bbsSn=<bbsSn>&key=00038
  - 다운로드  : GET  https://www.pps.go.kr/common/fileDown.do?key=<그룹키>&sn=<번호>
"""
import os
import re
import time
import datetime
from urllib.parse import unquote

import requests
from bs4 import BeautifulSoup

BASE = "https://www.pps.go.kr"
LIST_URL = BASE + "/kor/bbs/list.do"
VIEW_URL = BASE + "/kor/bbs/view.do"
DOWN_URL = BASE + "/common/fileDown.do"
BOARD_KEY = "00038"           # 조달업무 > 업무별자료 > 시설공사
SEARCH_WORD = "제비율"

# 수집 대상 공종 (파일명 접두어). 국가유산 등은 제외.
TARGET_PREFIXES = ("토목공사", "건축공사")

# 서버리스(Vercel)에서 함수 실행시간 제한을 지키기 위한 요청 간 대기(초).
# 로컬 실행에서는 0.2~0.3 였으나, 클라우드에서는 최소화한다.
SLEEP = float(os.environ.get("CRAWL_SLEEP", "0.05"))

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


def _new_session():
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Referer": LIST_URL + "?key=" + BOARD_KEY})
    # 세션 쿠키(JSESSIONID) 획득
    s.get(LIST_URL, params={"key": BOARD_KEY}, timeout=20)
    return s


def _decode_filename(content_disposition):
    """Content-Disposition 헤더에서 한글 파일명 복원.

    조달청 서버는 UA에 따라 아래 3가지 형태를 섞어 내려준다.
      1) filename*=UTF-8''%ED%86...        (RFC 5987)
      2) filename="%ED%86..."              (퍼센트 인코딩된 UTF-8)
      3) filename="í\x86\xa0..."           (UTF-8 바이트를 latin-1로 오인코딩)
    """
    if not content_disposition:
        return None
    cd = content_disposition

    # 1) RFC 5987 : filename*=charset''percent-encoded
    m = re.search(r"filename\*\s*=\s*([^;]+)", cd)
    if m:
        raw = m.group(1).strip().strip('"')
        mm = re.match(r"([\w\-]+)''(.*)", raw)
        if mm:
            charset = mm.group(1) or "utf-8"
            try:
                return unquote(mm.group(2), encoding=charset)
            except Exception:  # noqa
                return unquote(mm.group(2))
        return unquote(raw)

    # 2)/3) 일반 filename=
    m = re.search(r'filename\s*=\s*"?([^";]+)"?', cd)
    if not m:
        return None
    raw = m.group(1).strip()
    if "%" in raw:                       # 퍼센트 인코딩
        try:
            return unquote(raw, encoding="utf-8")
        except Exception:  # noqa
            pass
    try:                                  # latin-1 오인코딩 복원
        return raw.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return raw


def _parse_date(s):
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", s)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def search_posts(session, max_pages=10, log=print):
    """'제비율' 제목검색 결과 게시물 목록 [{bbsSn, title, date}] 반환."""
    posts = []
    seen = set()
    for page in range(1, max_pages + 1):
        data = {
            "key": BOARD_KEY,
            "pageIndex": str(page),
            "orderBy": "bbsOrdr desc",
            "sc": "BBS_SJ",
            "sw": SEARCH_WORD,
            "bbsSn": "",
        }
        r = session.post(LIST_URL, params={"key": BOARD_KEY}, data=data, timeout=20)
        r.encoding = "utf-8"
        soup = BeautifulSoup(r.text, "html.parser")
        rows = soup.select("table tbody tr")
        page_found = 0
        for tr in rows:
            a = tr.select_one('a[onclick^="goView"]')
            if not a:
                continue
            m = re.search(r"goView\('(\d+)'", a.get("onclick", ""))
            if not m:
                continue
            bbs_sn = m.group(1)
            if bbs_sn in seen:
                continue
            seen.add(bbs_sn)
            title = a.get_text(strip=True)
            # 등록일: 행의 날짜 형태 셀 탐색
            date = None
            for td in tr.select("td"):
                d = _parse_date(td.get_text(strip=True))
                if d:
                    date = d
                    break
            posts.append({"bbsSn": bbs_sn, "title": title, "date": date})
            page_found += 1
        log("  검색 %d페이지: 게시물 %d건" % (page, page_found))
        if page_found == 0:
            break
        time.sleep(SLEEP)
    return posts


def get_attachments(session, bbs_sn):
    """상세페이지에서 첨부파일 목록 [{name, key, sn}] 반환."""
    r = session.get(VIEW_URL, params={"bbsSn": bbs_sn, "key": BOARD_KEY}, timeout=20)
    r.encoding = "utf-8"
    soup = BeautifulSoup(r.text, "html.parser")
    files = []
    for a in soup.select('a[href*="fileDown"]'):
        href = a.get("href", "")
        m = re.search(r"key=(\d+)&sn=(\d+)", href)
        if not m:
            continue
        files.append({
            "name": a.get_text(strip=True),
            "key": m.group(1),
            "sn": m.group(2),
        })
    return files


def download_file(session, file_key, sn, dest_dir, fallback_name=None):
    """첨부파일 1개 다운로드. 저장된 절대경로 반환."""
    r = session.get(DOWN_URL, params={"key": file_key, "sn": sn}, timeout=60)
    r.raise_for_status()
    fname = _decode_filename(r.headers.get("Content-Disposition")) or fallback_name
    if not fname:
        fname = "%s_%s.xlsx" % (file_key, sn)
    # 확장자 보정
    if not os.path.splitext(fname)[1]:
        fname += ".xlsx"
    fname = re.sub(r'[\\/:*?"<>|]', "_", fname)   # 윈도우 금지문자 제거
    path = os.path.join(dest_dir, fname)
    with open(path, "wb") as f:
        f.write(r.content)
    return path


def crawl(dest_dir, years=3, ref_date=None, max_pages=10, log=print):
    """
    조회기준일(ref_date, 기본 오늘)로부터 `years`년 이내의 '제비율' 게시물에서
    토목/건축 간접공사비 적용기준 파일을 dest_dir에 내려받는다.

    반환: {"downloaded": [경로...], "skipped_posts": [...], "posts": N}
    """
    os.makedirs(dest_dir, exist_ok=True)
    ref_date = ref_date or datetime.date.today()
    cutoff = ref_date.replace(year=ref_date.year - years)
    log("조회기준일 %s / 수집범위 %s 이후" % (ref_date, cutoff))

    session = _new_session()
    posts = search_posts(session, max_pages=max_pages, log=log)
    log("전체 '제비율' 게시물 %d건" % len(posts))

    downloaded, details = [], []
    for p in posts:
        d = p["date"]
        if d and d < cutoff:
            log("  건너뜀(기간초과 %s): %s" % (d, p["title"][:30]))
            continue
        atts = get_attachments(session, p["bbsSn"])
        targets = [a for a in atts
                   if any(a["name"].startswith(pref) for pref in TARGET_PREFIXES)]
        log("  [%s] %s → 대상첨부 %d/%d" %
            (d, p["title"][:28], len(targets), len(atts)))
        for a in targets:
            try:
                path = download_file(session, a["key"], a["sn"], dest_dir,
                                     fallback_name=a["name"] + ".xlsx")
                downloaded.append(path)
                details.append({"post_date": d.isoformat() if d else "",
                                "file": os.path.basename(path)})
                log("      ↓ %s" % os.path.basename(path))
                time.sleep(SLEEP)
            except Exception as e:  # noqa
                log("      ! 다운로드 실패: %s (%s)" % (a["name"], e))
        time.sleep(SLEEP)

    return {"downloaded": downloaded, "details": details, "posts": len(posts)}


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "downloads"
    res = crawl(out, years=3)
    print("\n완료: %d개 파일 다운로드" % len(res["downloaded"]))

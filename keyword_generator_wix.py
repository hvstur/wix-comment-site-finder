# WIX + COMMENTS ANALYZER v7
# Один Excel-файл: wix_analysis.xlsx
# Внутри:
# SUMMARY
# WIX_BLOG
# WIX_GROUP
# WIX_GROUP_NO_DISCUSSION
# NO_POST_SIGNAL
# NOT_WIX
# FETCH_FAILED
# SKIPPED
# NEW_QUERIES

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse, parse_qs, unquote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================
# CONFIG
# =========================

INPUT_FILE = "serp-google.xlsx"
OUTPUT_FILE = "wix_analysis.xlsx"

WORKERS = 12
CONNECT_TIMEOUT = 8
READ_TIMEOUT = 12
MAX_BYTES = 8 * 1024 * 1024

ALLOWED_INPUT_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg", ".ico",
    ".css", ".js", ".json", ".xml", ".pdf", ".zip", ".rar",
    ".mp3", ".mp4", ".avi", ".mov", ".webm", ".woff", ".woff2",
    ".ttf", ".eot", ".map"
}

BLACKLIST_HOSTS = {
    "reddit.com", "facebook.com", "stackoverflow.com", "linkedin.com",
    "github.com", "youtube.com", "twitter.com", "instagram.com",
    "pinterest.com", "quora.com", "wikipedia.org", "tiktok.com",
    "amazon.com", "ebay.com", "tumblr.com", "medium.com", "kaggle.com",
    "huggingface.co", "steamcommunity.com", "blogspot.com", "blogger.com",
    "sourceforge.net", "flutterflow.io", "strikingly.com", "openjdk.org",
    "avast.com", "codementor.io", "fark.com", "studocu.com",
    "researchgate.net", "threads.net", "kemono.su", "contra.com",
    "cloudynights.com", "sketchucation.com", "derpibooroo.org",
    "hybrid-analysis.com", "checkphish.ai", "fiverr.com",
    "forum.wixstudio.com", "dev.wix.com", "wix.com",
    "support.google.com", "apps.apple.com"
}

WIX_HOST_MARKERS = (
    "wixsite.com",
    "wixstatic.com",
    "parastorage.com",
)

WIX_HTML_MARKERS = (
    "wix-thunderbolt",
    "wix-image",
    "wix-image-component",
    "wix-code",
    "wixapps.net",
    "wixpress.com",
    "x-wix-",
    "wix-bi-session",
    "wixstatic.com",
    "parastorage.com",
)

BLOG_PATH_RE = re.compile(r"/(?:post|single-post)/", re.I)

GROUP_DISCUSSION_RE = re.compile(
    r"/(?:group|group-page)/[^/]+/discussion(?:/|$)",
    re.I
)

GROUP_BASE_RE = re.compile(
    r"/(?:group|group-page)/[^/]+(?:/|$)",
    re.I
)

OTHER_GROUP_DISCUSSION_RE = re.compile(
    r"/groups?/[^/]+/discussion(?:/|$)",
    re.I
)

COMMENT_PATTERNS = (
    "write a comment",
    "leave a comment",
    "post a comment",
    "submit comment",
    "comment on this post",
    "comments",
    "comment form",
    "comment-form",
    "wix-comments",
    "reply",
    "post a reply",
    "login to comment",
)

SKIP_PATTERNS = (
    "static.wixstatic.com",
    "static.parastorage.com",
)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
]


# =========================
# HELPERS
# =========================

def normalize_url(raw):
    if raw is None:
        return ""

    u = str(raw).strip()
    if not u:
        return ""

    # Google redirect URLs
    try:
        p = urlparse(u)
        qs = parse_qs(p.query)
        for key in ("url", "q", "u"):
            if key in qs and qs[key]:
                candidate = unquote(qs[key][0])
                if candidate.startswith(("http://", "https://")):
                    u = candidate
                    break
    except Exception:
        pass

    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u

    try:
        p = urlparse(u)
        host = (p.hostname or "").lower().strip(".")
        path = p.path or "/"
        return f"https://{host}{path}"
    except Exception:
        return u


def host_of(url):
    try:
        return (urlparse(url).hostname or "").lower().strip(".")
    except Exception:
        return ""


def root_domain(host):
    parts = host.split(".")
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return host


def is_blacklisted_host(host):
    host = host.lower().strip(".")
    for blocked in BLACKLIST_HOSTS:
        if host == blocked or host.endswith("." + blocked):
            return True
    return False


def is_static_or_invalid(url):
    try:
        p = urlparse(url)
        host = (p.hostname or "").lower()
        path = p.path.lower()

        if any(x in host for x in SKIP_PATTERNS):
            return True

        if any(path.endswith(ext) for ext in ALLOWED_INPUT_EXTENSIONS):
            return True

        if host.startswith("cdn-"):
            return True

        return False
    except Exception:
        return True


def path_type(url):
    """
    Строгая классификация:
      GROUP_DISCUSSION
      GROUP_NO_DISCUSSION
      GROUP_DISCUSSION_OTHER
      BLOG_POST
      BLOG_SINGLE
      OTHER
    """
    try:
        path = urlparse(url).path
    except Exception:
        return "OTHER"

    if GROUP_DISCUSSION_RE.search(path):
        return "GROUP_DISCUSSION"

    if OTHER_GROUP_DISCUSSION_RE.search(path):
        return "GROUP_DISCUSSION_OTHER"

    if GROUP_BASE_RE.search(path):
        return "GROUP_NO_DISCUSSION"

    if re.search(r"/post/", path, re.I):
        return "BLOG_POST"

    if re.search(r"/single-post/", path, re.I):
        return "BLOG_SINGLE"

    return "OTHER"


def detect_wix(host, headers, html):
    h = host.lower()

    if any(marker in h for marker in WIX_HOST_MARKERS):
        return True

    header_text = " ".join(
        f"{k}: {v}" for k, v in headers.items()
    ).lower()

    if any(marker in header_text for marker in WIX_HTML_MARKERS):
        return True

    low = html.lower()
    return any(marker in low for marker in WIX_HTML_MARKERS)


def detect_comments(html):
    low = html.lower()

    hits = []
    for pattern in COMMENT_PATTERNS:
        if pattern in low:
            hits.append(pattern)

    # Уникальные сигналы
    hits = list(dict.fromkeys(hits))

    return bool(hits), ", ".join(hits)


def make_session():
    session = requests.Session()
    session.trust_env = False

    retry = Retry(
        total=0,
        connect=0,
        read=0,
        redirect=2,
        status=0,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=WORKERS,
        pool_maxsize=WORKERS,
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


# =========================
# FETCH
# =========================

def fetch_url(url, query, row_number):
    result = {
        "row": row_number,
        "query": query,
        "source_url": url,
        "final_url": "",
        "domain": "",
        "status": "",
        "classification": "",
        "comments": "",
        "comment_signals": "",
        "http_status": "",
        "error": "",
    }

    if not url:
        result["classification"] = "SKIPPED"
        result["status"] = "empty_url"
        return result

    normalized = normalize_url(url)

    if is_static_or_invalid(normalized):
        result["classification"] = "SKIPPED"
        result["status"] = "static_or_invalid"
        result["final_url"] = normalized
        result["domain"] = root_domain(host_of(normalized))
        return result

    host = host_of(normalized)

    if is_blacklisted_host(host):
        result["classification"] = "SKIPPED"
        result["status"] = "blacklisted_host"
        result["final_url"] = normalized
        result["domain"] = root_domain(host)
        return result

    session = make_session()

    try:
        headers = {
            "User-Agent": USER_AGENTS[row_number % len(USER_AGENTS)],
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.8",
            "Cache-Control": "no-cache",
        }

        response = session.get(
            normalized,
            headers=headers,
            timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            allow_redirects=True,
            stream=True,
        )

        result["http_status"] = response.status_code
        result["final_url"] = normalize_url(response.url)

        final_host = host_of(result["final_url"])
        result["domain"] = root_domain(final_host)

        chunks = []
        total = 0

        for chunk in response.iter_content(chunk_size=65536):
            if not chunk:
                continue

            total += len(chunk)

            if total > MAX_BYTES:
                break

            chunks.append(chunk)

        content = b"".join(chunks)

        try:
            html = content.decode(
                response.encoding or "utf-8",
                errors="ignore"
            )
        except Exception:
            html = content.decode("utf-8", errors="ignore")

        if response.status_code >= 400:
            result["classification"] = "FETCH_FAILED"
            result["status"] = f"http_{response.status_code}"
            return result

        if not detect_wix(final_host, response.headers, html):
            result["classification"] = "NOT_WIX"
            result["status"] = "no_wix_signal"
            return result

        ptype = path_type(result["final_url"])

        has_comments, comment_signals = detect_comments(html)
        result["comments"] = "YES" if has_comments else "NO"
        result["comment_signals"] = comment_signals

        if ptype in ("BLOG_POST", "BLOG_SINGLE"):
            result["classification"] = "WIX_BLOG"
            result["status"] = ptype
            return result

        if ptype in ("GROUP_DISCUSSION", "GROUP_DISCUSSION_OTHER"):
            result["classification"] = "WIX_GROUP"
            result["status"] = ptype
            return result

        if ptype == "GROUP_NO_DISCUSSION":
            result["classification"] = "WIX_GROUP_NO_DISCUSSION"
            result["status"] = ptype
            return result

        # Wix detected, but URL does not contain a recognized post/group path.
        result["classification"] = "NO_POST_SIGNAL"
        result["status"] = "wix_but_unknown_page_type"

        return result

    except requests.RequestException as e:
        result["classification"] = "FETCH_FAILED"
        result["status"] = "request_error"
        result["error"] = str(e)[:500]
        return result

    except Exception as e:
        result["classification"] = "FETCH_FAILED"
        result["status"] = "unexpected_error"
        result["error"] = str(e)[:500]
        return result

    finally:
        session.close()


# =========================
# INPUT
# =========================

def load_input():
    wb = load_workbook(INPUT_FILE, read_only=True, data_only=True)

    ws = wb.active

    rows = []

    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if not row:
            continue

        raw_url = row[0] if len(row) >= 1 else ""
        query = row[1] if len(row) >= 2 else ""

        if i == 1 and str(raw_url).strip().lower() in {
            "url", "source url", "ссылка", "ссылка на сайт"
        }:
            continue

        url = normalize_url(raw_url)

        if not url:
            continue

        rows.append({
            "row": i,
            "url": url,
            "query": str(query or "").strip(),
        })

    wb.close()
    return rows


# =========================
# NEW QUERIES
# =========================

def build_new_queries(input_rows):
    """
    Сохраняем уникальные запросы из исходного SERP-файла.
    Здесь НЕ генерируются искусственные дорки:
    NEW_QUERIES отражает реальные уникальные query из входных данных.
    """
    seen = set()
    queries = []

    for item in input_rows:
        q = item["query"].strip()
        if q and q not in seen:
            seen.add(q)
            queries.append(q)

    return queries


# =========================
# EXCEL
# =========================

HEADERS = [
    "Source URL",
    "Final URL",
    "Domain",
    "Query",
    "Status",
    "Comments",
    "Comment Signals",
    "HTTP Status",
    "Error",
    "Path Type",
]


def row_from_result(r):
    return [
        r["source_url"],
        r["final_url"],
        r["domain"],
        r["query"],
        r["status"],
        r["comments"],
        r["comment_signals"],
        r["http_status"],
        r["error"],
        r["status"],
    ]


def write_sheet(ws, results):
    ws.append(HEADERS)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for r in results:
        ws.append(row_from_result(r))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 48,
        "B": 48,
        "C": 28,
        "D": 70,
        "E": 30,
        "F": 12,
        "G": 45,
        "H": 12,
        "I": 45,
        "J": 30,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def write_summary(ws, all_results, unique_input_queries):
    categories = [
        "TOTAL",
        "WIX_BLOG",
        "WIX_GROUP",
        "WIX_GROUP_NO_DISCUSSION",
        "NO_POST_SIGNAL",
        "NOT_WIX",
        "FETCH_FAILED",
        "SKIPPED",
    ]

    ws.append(["Metric", "Value"])

    for c in categories:
        if c == "TOTAL":
            value = len(all_results)
        else:
            value = sum(
                1 for r in all_results
                if r["classification"] == c
            )

        ws.append([c, value])

    group_valid = [
        r for r in all_results
        if r["classification"] == "WIX_GROUP"
    ]

    blog_valid = [
        r for r in all_results
        if r["classification"] == "WIX_BLOG"
    ]

    group_urls = {
        r["final_url"] for r in group_valid
        if r["final_url"]
    }

    group_domains = {
        r["domain"] for r in group_valid
        if r["domain"]
    }

    blog_urls = {
        r["final_url"] for r in blog_valid
        if r["final_url"]
    }

    blog_domains = {
        r["domain"] for r in blog_valid
        if r["domain"]
    }

    ws.append([])
    ws.append(["VALID GROUP DISCUSSION URLS", len(group_urls)])
    ws.append(["VALID GROUP DOMAINS", len(group_domains)])
    ws.append(["VALID BLOG URLS", len(blog_urls)])
    ws.append(["VALID BLOG DOMAINS", len(blog_domains)])
    ws.append([])
    ws.append(["UNIQUE INPUT QUERIES", len(unique_input_queries)])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 18


def write_new_queries(ws, queries):
    ws.append(["Query"])

    for q in queries:
        ws.append([q])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions["A"].width = 100

    ws["A1"].font = Font(bold=True)


# =========================
# MAIN
# =========================

def main():
    start = time.time()

    input_rows = load_input()

    print(f"Loaded URLs: {len(input_rows)}")

    results = []

    with ThreadPoolExecutor(max_workers=WORKERS) as executor:
        futures = [
            executor.submit(
                fetch_url,
                item["url"],
                item["query"],
                item["row"]
            )
            for item in input_rows
        ]

        for n, future in enumerate(as_completed(futures), start=1):
            try:
                results.append(future.result())
            except Exception as e:
                results.append({
                    "row": 0,
                    "query": "",
                    "source_url": "",
                    "final_url": "",
                    "domain": "",
                    "status": "worker_error",
                    "classification": "FETCH_FAILED",
                    "comments": "",
                    "comment_signals": "",
                    "http_status": "",
                    "error": str(e)[:500],
                })

            if n % 50 == 0 or n == len(futures):
                print(f"Processed: {n}/{len(futures)}")

    # Стабильный порядок по исходной строке
    results.sort(key=lambda x: x["row"])

    grouped = {
        "WIX_BLOG": [],
        "WIX_GROUP": [],
        "WIX_GROUP_NO_DISCUSSION": [],
        "NO_POST_SIGNAL": [],
        "NOT_WIX": [],
        "FETCH_FAILED": [],
        "SKIPPED": [],
    }

    for r in results:
        cls = r["classification"]
        if cls in grouped:
            grouped[cls].append(r)

    unique_queries = build_new_queries(input_rows)

    # Один workbook — все результаты внутри одного файла
    wb = Workbook()

    # Удаляем стандартный лист
    default_ws = wb.active
    wb.remove(default_ws)

    # Листы в нужном порядке
    ws_summary = wb.create_sheet("SUMMARY")
    ws_blog = wb.create_sheet("WIX_BLOG")
    ws_group = wb.create_sheet("WIX_GROUP")
    ws_group_no = wb.create_sheet("WIX_GROUP_NO_DISCUSSION")
    ws_no_post = wb.create_sheet("NO_POST_SIGNAL")
    ws_not_wix = wb.create_sheet("NOT_WIX")
    ws_failed = wb.create_sheet("FETCH_FAILED")
    ws_skipped = wb.create_sheet("SKIPPED")
    ws_queries = wb.create_sheet("NEW_QUERIES")

    write_summary(ws_summary, results, unique_queries)
    write_sheet(ws_blog, grouped["WIX_BLOG"])
    write_sheet(ws_group, grouped["WIX_GROUP"])
    write_sheet(ws_group_no, grouped["WIX_GROUP_NO_DISCUSSION"])
    write_sheet(ws_no_post, grouped["NO_POST_SIGNAL"])
    write_sheet(ws_not_wix, grouped["NOT_WIX"])
    write_sheet(ws_failed, grouped["FETCH_FAILED"])
    write_sheet(ws_skipped, grouped["SKIPPED"])
    write_new_queries(ws_queries, unique_queries)

    # Акцент для главных листов
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="D9EAF7"
            )

    wb.save(OUTPUT_FILE)

    elapsed = time.time() - start

    print()
    print("=" * 60)
    print("ГОТОВО")
    print(f"Файл: {OUTPUT_FILE}")
    print(f"Всего URL: {len(results)}")
    print(f"WIX_BLOG: {len(grouped['WIX_BLOG'])}")
    print(f"WIX_GROUP: {len(grouped['WIX_GROUP'])}")
    print(
        "WIX_GROUP_NO_DISCUSSION:",
        len(grouped["WIX_GROUP_NO_DISCUSSION"])
    )
    print(f"NO_POST_SIGNAL: {len(grouped['NO_POST_SIGNAL'])}")
    print(f"NOT_WIX: {len(grouped['NOT_WIX'])}")
    print(f"FETCH_FAILED: {len(grouped['FETCH_FAILED'])}")
    print(f"SKIPPED: {len(grouped['SKIPPED'])}")
    print(f"Время: {elapsed:.1f} сек.")
    print("=" * 60)


if __name__ == "__main__":
    main()

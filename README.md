# Wix Comment Site Finder (Python)

Python script that takes a SERP export (URLs and the queries they came from), fetches each page in parallel, detects pages built on Wix, classifies them (blog post vs group discussion), checks for comment signals, and writes everything into a single multi-sheet Excel report.

[Описание на русском](#на-русском)

## What it does

1. Reads `serp-google.xlsx` (URL and query columns). Google redirect links are unwrapped and URLs normalized; static files and blacklisted hosts are skipped.
2. Downloads pages in 12 threads (`ThreadPoolExecutor`) with connect/read timeouts, retries and an 8 MB per-page cap, rotating User-Agent.
3. Detects Wix by host markers, response headers and HTML signatures.
4. Classifies the path: blog post (`/post/`), single post (`/single-post/`), group discussion, group without discussion, or other.
5. Checks the page for comment signals (a list of comment/reply patterns) and records which signals matched.
6. Groups every URL into a category and writes one workbook.

## Output

A single file `wix_analysis.xlsx` with these sheets:

| Sheet | Contents |
| --- | --- |
| `SUMMARY` | Totals per category and the list of unique input queries |
| `WIX_BLOG` | Wix blog posts (`/post/`, `/single-post/`) |
| `WIX_GROUP` | Wix group pages with a discussion |
| `WIX_GROUP_NO_DISCUSSION` | Wix group pages without a discussion |
| `NO_POST_SIGNAL` | Wix pages with no post/discussion signal |
| `NOT_WIX` | Pages that are not built on Wix |
| `FETCH_FAILED` | Fetch errors (timeout, SSL, connection) |
| `SKIPPED` | Empty rows, static files, blacklisted hosts |
| `NEW_QUERIES` | Unique queries taken from the input file |

Result sheets have these columns: Source URL, Final URL, Domain, Query, Status, Comments, Comment Signals, HTTP Status, Error, Path Type.

`NEW_QUERIES` preserves the unique queries from the input SERP file (it does not invent new search footprints).

## Usage

```bash
pip install -r requirements.txt
python keyword_generator_wix.py
```

Put `serp-google.xlsx` next to the script. Tested with Python 3.11.

## Stack

Python · requests (retries via urllib3) · openpyxl · concurrent.futures (ThreadPoolExecutor) · regular expressions

## На русском

Python-скрипт для анализа выгрузки из поисковой выдачи (`serp-google.xlsx`: URL и запрос). В 12 потоков скачивает страницы, определяет Wix по маркерам в хосте, заголовках и HTML, классифицирует путь (пост блога, обсуждение группы, группа без обсуждения и т.д.) и проверяет сигналы формы комментариев.

Результат — один файл `wix_analysis.xlsx` с листами: `SUMMARY`, `WIX_BLOG`, `WIX_GROUP`, `WIX_GROUP_NO_DISCUSSION`, `NO_POST_SIGNAL`, `NOT_WIX`, `FETCH_FAILED`, `SKIPPED`, `NEW_QUERIES`. Лист `NEW_QUERIES` хранит уникальные запросы из входного файла.

Запуск: `pip install -r requirements.txt`, затем `python keyword_generator_wix.py`, входной файл рядом со скриптом.

# -*- coding: utf-8 -*-
"""
TDnet(指定日) PDF/ZIP(XBRL) を GitHub へアップロード

✅ 取り違え防止: 「表題リンク由来のID(pdf_id)」のみを採用
✅ 英字入りticker対応: 268A0 -> 268A / 331A0 -> 331A 等
✅ 除外: JPX上場銘柄リスト(data_j.xls)のD列で以下を取得対象外:
       ETF・ETN / PRO Market / プライム（外国株式） /
       スタンダード（外国株式） / グロース（外国株式） /
       REIT・ベンチャーファンド・カントリーファンド・インフラファンド / 出資証券
✅ 収納先: PDF -> tekigikaizi/yyyymmdd/ , ZIP(XBRL) -> XBRL/yyyymmdd/
✅ 1フォルダ1000件超過時に自動分轄: yyyymmdd_Part1 / yyyymmdd_Part2 …
✅ manifest.json をそれぞれ保存
✅ GitHub 404/権限事故対策:
   - token種別に応じて Authorization を自動切替 (github_pat_ -> Bearer / それ以外 -> token)
   - 実行前に preflight_write_check() で「書き込み可能」を検証してから進む
   - tokenに非ASCII混入（日本語等）があれば即停止
"""

# Colab環境では !pip install、通常環境では事前に pip install してください
try:
    get_ipython  # type: ignore
    import subprocess
    subprocess.run(["pip", "-q", "install", "beautifulsoup4", "lxml", "tqdm", "xlrd"], check=False)
except NameError:
    pass  # 非Colab環境: pip install beautifulsoup4 lxml tqdm xlrd を別途実行

import io
import os
import re
import sys
import time
import json
import base64
import hashlib
from urllib.parse import urljoin
from pathlib import Path
from typing import Dict, FrozenSet, List, Optional, Tuple

import requests
from bs4 import BeautifulSoup

try:
    from tqdm.auto import tqdm
    def _progress(it, **kw):
        return tqdm(it, **kw)
except ImportError:
    def _progress(it, **kw):
        return it

# ─────────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────────
GITHUB_OWNER = "yukizi1113"
GITHUB_REPO  = "tdnet"

TARGET_DATE = "20260226"   # yyyymmdd（ここを変える）

PDF_DIR_IN_REPO = "tekigikaizi"
ZIP_DIR_IN_REPO = "XBRL"

# 既に同名がある場合に上書きするか
OVERWRITE = True

# アクセス間隔（秒）
SLEEP_SEC = 1.0

# 1フォルダに収めるファイル数の上限（超過時に _Part1 / _Part2 … に分轄）
PART_SIZE = 1000

# JPX上場銘柄リスト URL
JPX_LISTED_URL = (
    "https://www.jpx.co.jp/markets/statistics-equities/misc/"
    "tvdivq0000001vg2-att/data_j.xls"
)

# 市場・商品区分のうち取得対象から除外するカテゴリ
JPX_EXCLUDE_TYPES: FrozenSet[str] = frozenset([
    "ETF・ETN",
    "PRO Market",
    "プライム（外国株式）",
    "スタンダード（外国株式）",
    "グロース（外国株式）",
    "REIT・ベンチャーファンド・カントリーファンド・インフラファンド",
    "出資証券",
])

# ─────────────────────────────────────────────────
# GitHub token（環境変数 GITHUB_TOKEN → 対話入力）
# ─────────────────────────────────────────────────
def clean_token(tok: str) -> str:
    tok = (tok or "").strip()
    tok = tok.replace("\u200b", "").replace("\ufeff", "")
    return tok

GITHUB_TOKEN = clean_token(os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN") or "")
if not GITHUB_TOKEN:
    from getpass import getpass
    GITHUB_TOKEN = clean_token(getpass("Paste your GitHub token (input hidden): "))

bad = [(i, ch, ord(ch)) for i, ch in enumerate(GITHUB_TOKEN) if ord(ch) > 127]
if bad:
    print("ERROR: GitHub token に非ASCII文字が混入しています。")
    print("混入文字（先頭10件）:", bad[:10])
    raise ValueError("Token contains non-ASCII characters.")

# ─────────────────────────────────────────────────
# 基本定数
# ─────────────────────────────────────────────────
TDNET_BASE = "https://www.release.tdnet.info/inbs/"
MAIN_URL   = urljoin(TDNET_BASE, "I_main_00.html")
UA         = "Mozilla/5.0 (compatible; tdnet-downloader/5.0-jpxfilter; +https://example.invalid)"

TDNET_HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.8,en;q=0.6",
    "Accept-Encoding": "gzip, deflate",   # Brotliを要求しない（HTML破損を回避）
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Referer": MAIN_URL,
}

FW2HW_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
FW2HW_ALPHA  = str.maketrans(
    "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
)

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": UA})

# ─────────────────────────────────────────────────
# JPX ticker フィルター
# ─────────────────────────────────────────────────

def _normalize_jpx_ticker(v) -> Optional[str]:
    """XLSのコード値（float 1301.0 / string '130A'）を4文字tickerに正規化。"""
    if v is None:
        return None
    s = str(v).strip().upper()
    if not s or s in ("NAN", "NONE", ""):
        return None
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.fullmatch(r"\d{4}", s):
        return s
    if re.fullmatch(r"\d{3}[A-Z0-9]", s):
        return s
    if re.fullmatch(r"\d{4}[A-Z]", s):
        return s
    return None


def build_excluded_tickers(
    url: str = JPX_LISTED_URL,
    exclude_types: FrozenSet[str] = JPX_EXCLUDE_TYPES,
) -> FrozenSet[str]:
    """JPX上場銘柄リスト(data_j.xls)をダウンロードし、除外対象ticker集合を返す。

    - Column B (index 1): 証券コード
    - Column D (index 3): 市場・商品区分
    Row 0 はヘッダ行、Row 1 以降がデータ。

    xlrd が未インストールの場合は空の frozenset を返して警告。
    """
    try:
        import xlrd  # noqa: F401
    except ImportError:
        print(
            "[warn] JPX filter: xlrd not installed. "
            "Run `pip install xlrd` to enable. Proceeding without filter.",
            file=sys.stderr,
        )
        return frozenset()

    try:
        import xlrd as _xlrd
        r = SESSION.get(url, timeout=60)
        r.raise_for_status()
        wb = _xlrd.open_workbook(file_contents=r.content)
        ws = wb.sheet_by_index(0)

        excluded: set = set()
        for i in range(1, ws.nrows):  # row 0 = header
            raw_code = ws.cell_value(i, 1)   # Column B: 証券コード
            raw_type = ws.cell_value(i, 3)   # Column D: 市場・商品区分
            mtype = str(raw_type).strip() if raw_type else ""
            if mtype not in exclude_types:
                continue
            t = _normalize_jpx_ticker(raw_code)
            if t:
                excluded.add(t)

        result = frozenset(excluded)
        print(
            f"[JPX filter] {ws.nrows - 1} listings → "
            f"{len(result)} tickers excluded"
        )
        return result

    except Exception as e:
        print(f"[warn] JPX filter: failed ({e}). Proceeding without filter.", file=sys.stderr)
        return frozenset()

# ─────────────────────────────────────────────────
# GitHub API
# ─────────────────────────────────────────────────

def gh_headers(token: str) -> dict:
    token = clean_token(token)
    scheme = "Bearer" if token.startswith("github_pat_") else "token"
    return {
        "Authorization": f"{scheme} {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "tdnet-to-github",
    }


def gh_get_content(owner: str, repo: str, path: str, token: str):
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    return SESSION.get(url, headers=gh_headers(token), timeout=60)


def gh_put_file(
    owner: str,
    repo: str,
    path: str,
    token: str,
    content_bytes: bytes,
    message: str,
    overwrite: bool = False,
) -> dict:
    """GitHub Contents API: PUT /repos/{owner}/{repo}/contents/{path}"""
    if len(content_bytes) > 100 * 1024 * 1024:
        raise RuntimeError("File too large (>100MB). GitHub Contents API limit.")

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"

    sha = None
    r0 = gh_get_content(owner, repo, path, token)
    if r0.status_code == 200:
        if not overwrite:
            return {"skipped": True, "reason": "exists", "path": path}
        sha = r0.json().get("sha")
    elif r0.status_code != 404:
        raise RuntimeError(f"GitHub precheck failed: {r0.status_code} {r0.text[:300]}")

    payload: Dict = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
    }
    if sha:
        payload["sha"] = sha

    r = SESSION.put(url, headers=gh_headers(token), data=json.dumps(payload), timeout=120)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub upload failed: {r.status_code} {r.text[:500]}")
    return {"skipped": False, "path": path, "status": r.status_code}


def preflight_write_check(owner: str, repo: str, token: str) -> None:
    """実際に小さなファイルを書き込んで権限を検証。失敗時は例外を上げる。"""
    path = "__preflight__.txt"
    url  = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"

    payload: Dict = {
        "message": "preflight write check",
        "content": base64.b64encode(b"ok").decode("utf-8"),
    }
    r0 = SESSION.get(url, headers=gh_headers(token), timeout=30)
    if r0.status_code == 200:
        payload["sha"] = r0.json().get("sha")
    elif r0.status_code != 404:
        raise RuntimeError(f"Preflight GET failed: {r0.status_code} {r0.text[:200]}")

    r = SESSION.put(url, headers=gh_headers(token), data=json.dumps(payload), timeout=30)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Preflight PUT failed: {r.status_code} {r.text[:200]}")
    print("[preflight] write check passed.")

# ─────────────────────────────────────────────────
# 文字ユーティリティ
# ─────────────────────────────────────────────────

def normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def normalize_text(s: str) -> str:
    return normalize_ws(
        (s or "").translate(FW2HW_DIGITS).translate(FW2HW_ALPHA).replace("　", " ")
    )


def sanitize_filename(name: str, max_len: int = 180) -> str:
    name = normalize_text(name)
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = normalize_ws(name)
    if len(name) <= max_len:
        return name
    h = hashlib.sha1(name.encode("utf-8")).hexdigest()[:8]
    keep = max_len - (len(h) + 2)
    return f"{name[:keep]}__{h}"


def sniff_ext_from_bytes(b: bytes) -> str:
    if b.startswith(b"%PDF"):
        return ".pdf"
    if b[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"):
        return ".zip"
    return ""

# ─────────────────────────────────────────────────
# TDnet HTML取得
# ─────────────────────────────────────────────────

def decode_html_bytes(b: bytes) -> str:
    head = b[:5000].lower()
    m = re.search(rb"charset\s*=\s*['\"]?([a-z0-9_\-]+)", head)
    enc = m.group(1).decode("ascii", errors="ignore") if m else None
    for e in [enc, "utf-8", "cp932"]:
        if not e:
            continue
        try:
            return b.decode(e, errors="replace")
        except Exception:
            pass
    return b.decode("utf-8", errors="replace")


def fetch_html(url: str, sleep_sec: float = SLEEP_SEC, retry: int = 2) -> Optional[str]:
    last = None
    for k in range(retry + 1):
        try:
            time.sleep(sleep_sec)
            r = SESSION.get(url, headers=TDNET_HEADERS, timeout=60)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return decode_html_bytes(r.content)
        except Exception as e:
            last = e
            if k == retry:
                break
            time.sleep(0.5 * (k + 1))
    raise RuntimeError(f"fetch_html failed: {url} ({last})")


def iter_list_pages_for_date(date_yyyymmdd: str, max_pages: int = 200):
    for i in range(1, max_pages + 1):
        url = urljoin(TDNET_BASE, f"I_list_{i:03d}_{date_yyyymmdd}.html")
        html = fetch_html(url)
        if html is None:
            break
        yield url, html


def unwrap_to_content_html(html: str, max_depth: int = 5) -> str:
    cur = html
    for _ in range(max_depth):
        soup = BeautifulSoup(cur, "lxml")
        node = soup.find("iframe", src=True) or soup.find("frame", src=True)
        if not node:
            return cur
        src = (node.get("src") or "").strip()
        if not src:
            return cur
        nxt_url = urljoin(TDNET_BASE, src)
        nxt = fetch_html(nxt_url)
        if not nxt:
            return cur
        cur = nxt
    return cur

# ─────────────────────────────────────────────────
# ID抽出（リンクからのみ）
# ─────────────────────────────────────────────────

_RE_PDF_ID = re.compile(r"(\d{18})\.pdf", re.IGNORECASE)
_RE_ZIP_ID = re.compile(r"(\d{18})\.zip", re.IGNORECASE)
_RE_ID18   = re.compile(r"\b(\d{18})\b")


def extract_pdf_id_from_anchor(a) -> Optional[str]:
    if not a:
        return None
    for attr in ("href", "onclick"):
        s = a.get(attr) or ""
        m = _RE_PDF_ID.search(s)
        if m:
            return m.group(1)
        ids = [x for x in _RE_ID18.findall(s) if not x.startswith("08")]
        if len(ids) == 1:
            return ids[0]
    return None


def extract_xbrl_zip_id_from_anchor(a) -> Optional[str]:
    if not a:
        return None
    for attr in ("href", "onclick"):
        s = a.get(attr) or ""
        m = _RE_ZIP_ID.search(s)
        if m:
            return m.group(1)
        ids = [x for x in _RE_ID18.findall(s) if x.startswith("08")]
        if len(ids) == 1:
            return ids[0]
    return None


def pick_title_anchor_and_title(tr) -> Tuple[Optional[object], str]:
    for a in tr.find_all("a"):
        txt = normalize_ws(a.get_text(" ", strip=True))
        if not txt:
            continue
        if txt.upper() == "XBRL":
            continue
        return a, txt
    return None, ""


def pick_xbrl_anchor(tr):
    for a in tr.find_all("a"):
        if normalize_ws(a.get_text(" ", strip=True)).upper() == "XBRL":
            return a
    return None

# ─────────────────────────────────────────────────
# ticker抽出（35930→3593 / 268A0→268A）
# ─────────────────────────────────────────────────

def parse_tdnet_code_to_ticker(code_cell_text: str) -> Optional[str]:
    s = normalize_text(code_cell_text)
    if not s:
        return None
    tok = s.split()[0].strip("()（）[]【】")

    patterns = [
        (r"^(\d{4})0$",           lambda m: m.group(1)),
        (r"^(\d{3})([A-Za-z])0$", lambda m: (m.group(1) + m.group(2)).upper()),
        (r"^(\d{4})([A-Za-z])0$", lambda m: (m.group(1) + m.group(2)).upper()),
        (r"^(\d{4})$",            lambda m: m.group(1)),
        (r"^(\d{3})([A-Za-z])$",  lambda m: (m.group(1) + m.group(2)).upper()),
        (r"^(\d{4})([A-Za-z])$",  lambda m: (m.group(1) + m.group(2)).upper()),
        (r"^(\d{3})([A-Za-z])0\D",lambda m: (m.group(1) + m.group(2)).upper()),
        (r"^(\d{4})0\D",          lambda m: m.group(1)),
    ]
    for pat, fn in patterns:
        m = re.match(pat, tok)
        if m:
            return fn(m)
    return None


def parse_time_from_tds(td_texts: List[str]) -> str:
    for t in td_texts[:10]:
        tt = normalize_text(t)
        if re.fullmatch(r"\d{1,2}:\d{2}", tt):
            return tt
    return ""


def parse_row_by_td_positions(tr) -> Optional[dict]:
    tds = tr.find_all("td")
    if not tds:
        return None

    title_a, title = pick_title_anchor_and_title(tr)
    if not title:
        return None

    pdf_id  = extract_pdf_id_from_anchor(title_a)
    xbrl_a  = pick_xbrl_anchor(tr)
    xbrl_id = extract_xbrl_zip_id_from_anchor(xbrl_a)

    if not pdf_id and not xbrl_id:
        return None

    # title が入っているtdの位置を推定
    title_td_idx = None
    for i, td in enumerate(tds):
        if title_a is not None:
            hit = td.find(lambda tag: getattr(tag, "name", None) == "a" and tag is title_a)
            if hit is not None:
                title_td_idx = i
                break
        txt = normalize_ws(td.get_text(" ", strip=True))
        if txt and normalize_ws(title) == txt and txt.upper() != "XBRL":
            title_td_idx = i

    if title_td_idx is None:
        return None

    td_text_list = [normalize_ws(td.get_text(" ", strip=True)) for td in tds]
    tm = parse_time_from_tds(td_text_list)

    # title の左側から company / code を拾う
    company = ""
    found: List[str] = []
    for j in range(title_td_idx - 1, -1, -1):
        txt = normalize_ws(td_text_list[j])
        if not txt:
            continue
        if re.fullmatch(r"\d{1,2}:\d{2}", normalize_text(txt)):
            continue
        found.append(txt)
        if len(found) >= 2:
            break

    if len(found) >= 1:
        company = found[0]
    code_text = found[1] if len(found) >= 2 else ""

    ticker = parse_tdnet_code_to_ticker(code_text)
    if not ticker:
        for j in range(title_td_idx - 1, -1, -1):
            cand = parse_tdnet_code_to_ticker(normalize_ws(td_text_list[j]))
            if cand:
                ticker = cand
                if j + 1 < len(td_text_list) and normalize_ws(td_text_list[j + 1]):
                    company = normalize_ws(td_text_list[j + 1])
                break

    if not ticker or not company:
        return None

    return {
        "ticker":  ticker,
        "company": company,
        "title":   title,
        "time":    tm,
        "pdf_id":  pdf_id,
        "xbrl_id": xbrl_id,
    }


def parse_list_page(html: str) -> List[dict]:
    html = unwrap_to_content_html(html, max_depth=5)
    soup = BeautifulSoup(html, "lxml")
    out = []
    for tr in soup.find_all("tr"):
        r = parse_row_by_td_positions(tr)
        if r:
            out.append(r)
    return out

# ─────────────────────────────────────────────────
# Download / Upload
# ─────────────────────────────────────────────────

def download_bytes(url: str, sleep_sec: float = SLEEP_SEC, retry: int = 2) -> bytes:
    last = None
    for k in range(retry + 1):
        try:
            time.sleep(sleep_sec)
            r = SESSION.get(url, headers=TDNET_HEADERS, timeout=120, allow_redirects=True)
            r.raise_for_status()
            return r.content
        except Exception as e:
            last = e
            if k == retry:
                break
            time.sleep(0.5 * (k + 1))
    raise RuntimeError(f"download failed: {url} ({last})")


def build_base_name(ticker: str, company: str, title: str) -> str:
    return sanitize_filename(f"{ticker}_{company}_{title}")


def _folder_name(date_yyyymmdd: str, use_parts: bool, part_idx: int) -> str:
    """フォルダ名を返す。use_parts=False なら日付のみ、True なら _Part1 / _Part2 … を付与。"""
    if not use_parts:
        return date_yyyymmdd
    return f"{date_yyyymmdd}_Part{part_idx + 1}"


def _split_into_chunks(items: list, chunk_size: int) -> List[list]:
    """リストを chunk_size ごとに分割して返す。空リストの場合は [[]] を返す。"""
    if not items:
        return [[]]
    return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]

# ─────────────────────────────────────────────────
# メイン実行
# ─────────────────────────────────────────────────

def run(date_yyyymmdd: str, debug_sample: int = 25) -> None:
    # ── Step 1: GitHub書き込み確認 ──
    preflight_write_check(GITHUB_OWNER, GITHUB_REPO, GITHUB_TOKEN)

    # ── Step 2: JPX ticker フィルターを構築 ──
    excluded_tickers = build_excluded_tickers()

    # ── Step 3: TDNet一覧ページを全ページ取得・パース ──
    all_rows: List[dict] = []
    page_count = 0
    for url, html in iter_list_pages_for_date(date_yyyymmdd, max_pages=200):
        page_count += 1
        all_rows.extend(parse_list_page(html))

    # pdf_id をキーに統合（同一PDFにxbrl_idが付く/付かない揺れを吸収）
    merged: Dict[str, dict] = {}
    for r in all_rows:
        key = (
            r.get("pdf_id")
            or f"NO_PDF__{r['ticker']}__{r['company']}__{r['title']}__{r.get('time','')}"
        )
        if key not in merged:
            merged[key] = dict(r)
        else:
            if not merged[key].get("xbrl_id") and r.get("xbrl_id"):
                merged[key]["xbrl_id"] = r["xbrl_id"]
            if not merged[key].get("pdf_id") and r.get("pdf_id"):
                merged[key]["pdf_id"] = r["pdf_id"]

    rows = list(merged.values())
    print(f"Pages fetched: {page_count}")
    print(f"Rows parsed (before JPX filter): {len(rows)}")

    # ── Step 4: JPX フィルター適用 ──
    if excluded_tickers:
        before = len(rows)
        rows = [r for r in rows if r["ticker"] not in excluded_tickers]
        print(f"[JPX filter] {before} → {len(rows)} rows (excluded {before - len(rows)})")

    if debug_sample > 0:
        print(f"DEBUG sample rows (first {debug_sample}):")
        for r in rows[:debug_sample]:
            print(
                f"  {r['ticker']}  {r['company']}  {r['title'][:60]}"
                f"  (pdf={r.get('pdf_id')}, xbrl={r.get('xbrl_id')})"
            )

    # ── Step 5: タスクリスト作成 ──
    pdf_tasks: List[Tuple] = []
    zip_tasks: List[Tuple] = []
    seen_task: set = set()

    for r in rows:
        if r.get("pdf_id"):
            k = ("pdf", r["pdf_id"])
            if k not in seen_task:
                pdf_tasks.append((r, r["pdf_id"]))
                seen_task.add(k)
        if r.get("xbrl_id"):
            k = ("zip", r["xbrl_id"])
            if k not in seen_task:
                zip_tasks.append((r, r["xbrl_id"]))
                seen_task.add(k)

    total_files = len(pdf_tasks) + len(zip_tasks)
    print(f"Files to fetch: PDF={len(pdf_tasks)}, XBRL={len(zip_tasks)}, total={total_files}")

    # ── Step 6: Part分轄の決定 ──
    pdf_chunks = _split_into_chunks(pdf_tasks, PART_SIZE)
    zip_chunks = _split_into_chunks(zip_tasks, PART_SIZE)
    pdf_use_parts = len(pdf_chunks) > 1
    zip_use_parts = len(zip_chunks) > 1

    if pdf_use_parts:
        print(f"[split] PDF: {len(pdf_tasks)} files → {len(pdf_chunks)} parts (PART_SIZE={PART_SIZE})")
    if zip_use_parts:
        print(f"[split] XBRL: {len(zip_tasks)} files → {len(zip_chunks)} parts (PART_SIZE={PART_SIZE})")

    # ── Step 7: アップロード ──
    uploaded = skipped = failed = 0
    all_manifests: Dict[str, dict] = {}  # folder_path -> manifest

    def get_or_init_manifest(repo_dir: str, folder: str) -> dict:
        key = f"{repo_dir}/{folder}"
        if key not in all_manifests:
            all_manifests[key] = {
                "date": date_yyyymmdd,
                "folder": folder,
                "root": f"{repo_dir}/{folder}/",
                "items": [],
            }
        return all_manifests[key]

    def process_chunk(chunk: list, kind: str, part_idx: int) -> None:
        nonlocal uploaded, skipped, failed

        repo_dir = PDF_DIR_IN_REPO if kind == "pdf" else ZIP_DIR_IN_REPO
        use_parts = pdf_use_parts if kind == "pdf" else zip_use_parts
        folder = _folder_name(date_yyyymmdd, use_parts, part_idx)
        manifest = get_or_init_manifest(repo_dir, folder)

        ext = ".pdf" if kind == "pdf" else ".zip"

        for r, fid in _progress(chunk, desc=f"{kind.upper()} {folder}"):
            ticker  = r["ticker"]
            company = r["company"]
            title   = r["title"]
            tm      = r.get("time", "")

            url     = urljoin(TDNET_BASE, f"{fid}{ext}")
            base    = build_base_name(ticker, company, title)
            filename = f"{base}{ext}" if kind == "pdf" else f"{base}__XBRL{ext}"
            gh_path  = f"{repo_dir}/{folder}/{filename}"

            # Download
            try:
                data = download_bytes(url, sleep_sec=SLEEP_SEC, retry=2)
            except Exception as e:
                failed += 1
                manifest["items"].append({
                    "ticker": ticker, "company": company, "title": title, "time": tm,
                    "file_id": fid, "source_url": url, "github_path": gh_path,
                    "status": "download_failed", "error": str(e)[:250],
                })
                continue

            # Magic byte check
            real_ext = sniff_ext_from_bytes(data)
            if (kind == "pdf" and real_ext != ".pdf") or (kind == "zip" and real_ext != ".zip"):
                failed += 1
                manifest["items"].append({
                    "ticker": ticker, "company": company, "title": title, "time": tm,
                    "file_id": fid, "source_url": url, "github_path": gh_path,
                    "status": "type_mismatch",
                    "error": f"expected {ext} but magic={real_ext or 'unknown'}",
                })
                continue

            # Upload to GitHub
            try:
                res = gh_put_file(
                    GITHUB_OWNER, GITHUB_REPO, gh_path, GITHUB_TOKEN, data,
                    message=f"TDnet {date_yyyymmdd} {ticker} {company}: {title} ({fid})",
                    overwrite=OVERWRITE,
                )
                time.sleep(0.25)
            except Exception as e:
                failed += 1
                manifest["items"].append({
                    "ticker": ticker, "company": company, "title": title, "time": tm,
                    "file_id": fid, "source_url": url, "github_path": gh_path,
                    "status": "upload_failed", "error": str(e)[:250],
                })
                continue

            if res.get("skipped"):
                skipped += 1
                status = "skipped_exists"
            else:
                uploaded += 1
                status = "uploaded"

            manifest["items"].append({
                "ticker": ticker, "company": company, "title": title, "time": tm,
                "file_id": fid, "source_url": url, "github_path": gh_path,
                "status": status,
            })

    # PDF chunks
    for part_idx, chunk in enumerate(pdf_chunks):
        if chunk:
            process_chunk(chunk, "pdf", part_idx)

    # ZIP chunks
    for part_idx, chunk in enumerate(zip_chunks):
        if chunk:
            process_chunk(chunk, "zip", part_idx)

    # ── Step 8: manifest.json をフォルダごとに保存 ──
    for folder_path, manifest in all_manifests.items():
        manifest_path = f"{folder_path}/manifest.json"
        try:
            gh_put_file(
                GITHUB_OWNER, GITHUB_REPO, manifest_path, GITHUB_TOKEN,
                json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8"),
                message=f"Add manifest {manifest['folder']}",
                overwrite=True,
            )
            print(f"[manifest] saved: {manifest_path}")
        except Exception as e:
            print(f"[manifest] failed to save {manifest_path}: {e}", file=sys.stderr)

    print("\n=== Done ===")
    print(f"uploaded={uploaded}, skipped={skipped}, failed={failed}")


# ─────────────────────────────────────────────────
# エントリポイント
# ─────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="TDNet PDF/XBRL → GitHub アップローダー")
    p.add_argument("--date", default=TARGET_DATE,
                   help="対象日 yyyymmdd (default: %(default)s)")
    p.add_argument("--debug-sample", type=int, default=25,
                   help="デバッグ用先頭N件表示 (0で無効)")
    p.add_argument("--no-jpx-filter", action="store_true",
                   help="JPX tickerフィルターを無効化")
    p.add_argument("--part-size", type=int, default=PART_SIZE,
                   help=f"1フォルダの最大ファイル数 (default: {PART_SIZE})")
    args = p.parse_args()

    if args.no_jpx_filter:
        # フィルターを無効化: exclude_types を空にする
        import builtins
        _orig_build = build_excluded_tickers
        build_excluded_tickers = lambda **kw: frozenset()  # noqa: E731

    PART_SIZE = args.part_size
    run(args.date, debug_sample=args.debug_sample)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SCIDownload —— 按 DOI 清单批量下载论文（正文 + 补充材料）。

支持 ScienceDirect（主力）以及已实测通过的其他出版商：
Wiley 系（10.1002 / 10.1111 / 10.2134 / 10.2136）、
Canadian Science（10.1139 / 10.4141）、CSIRO（10.1071）、Copernicus（10.5194）。

原理（一句话）：不假装成脚本去请求，而是**驱动一个真浏览器**，
用你所在的机构 IP（或已登录的会话）把 PDF 取下来。

    python 3.SCIDownload.py dois.txt --out ./papers --si

只用 Python 标准库；浏览器用系统已装的 Chrome / Edge / Chromium。
"""
from __future__ import annotations

import argparse
import atexit
import csv
import json
import math
import os
import random
import re
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import SCIDownload_cdp as cdp  # noqa: E402

OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
VERSION = "1.6.1"

# 从文章页抽取：本文章的 PDF 入口、补充材料链接、授权状态
ARTICLE_JS = r"""
(() => {
  const m = location.pathname.match(/\/article\/pii\/([A-Za-z0-9]+)/);
  const pii = m ? m[1].toUpperCase() : null;
  const links = [...document.querySelectorAll('a[href]')].map(a => a.href);
  // 多数期刊 /pdfft?md5=…；Pedosphere、JIA 等中国期刊英文版是 /pdf?md5=…
  const isPdf = h => /\/(pdf|pdfft)\?/i.test(h);
  let pdf = null;
  if (pii) pdf = links.find(h => isPdf(h) && h.toUpperCase().includes(pii)) || null;
  if (!pdf) pdf = links.find(isPdf) || null;
  // 补充材料（ScienceDirect）：只认 -mmc<数字>.，要排除 -gr/-fx 插图
  const si = [...new Set(links.filter(h => /ars\.els-cdn\.com/.test(h) && /-mmc\d+\./i.test(h)))];
  // 补充材料（Wiley 系，含 ACSESS 子域）：统一走 /action/downloadSupplement?doi=…&file=…
  // 实测 2026-09-19：onlinelibrary.wiley.com 与 acsess.onlinelibrary.wiley.com 同构；
  // 文件名在 &file= 参数里（如 gcb15410-sup-0001-supinfo.docx）。
  const si_wiley = [...new Set(links.filter(h => /\/action\/downloadSupplement\?/i.test(h)))];
  const body = (document.body ? document.body.innerText : '') || '';
  const url = location.href;
  // 判定「停在登录页」必须严格：先确认不是在文章页，再匹配明确的认证路径。
  // 早期版本用了裸 sso，会把正常页误判成登录页。
  const isArticlePage = /sciencedirect\.com\/science\/article/.test(url);
  const loginish = !isArticlePage && /authserver|\/login|carsi|shibboleth|openathens|ezproxy|\/wayf|\/idp\/|\/sso\/|login\?/i.test(url);
  // 区分「页面没加载完」和「DOI 本身解析不到」—— 后者不该反复重试白等。
  // 实测：错误的 DOI 会停在 doi.org 上，标题为 "Error: DOI Not Found"。
  const onSd = /(^|\.)sciencedirect\.com$/.test(location.host);
  const notfound = /doi not found|this doi cannot be found/i.test(document.title + ' ' + body)
                 || (!onSd && /^error\s*:/i.test(document.title));
  return JSON.stringify({
    pii, pdf, si, si_wiley, url, loginish, notfound, on_sd: onSd,
    host: location.host,
    access: /Full text access|Open access/i.test(body) ? 'granted'
          : (/Get access|Purchase PDF/i.test(body) ? 'denied' : 'unknown'),
    title: document.title,
    ready: document.readyState,
    body_len: body.length,
    brought_by: (body.match(/Brought to you by\s*:?\s*([^\n]{0,60})/i) || [])[1] || null,
    // 非 ScienceDirect 出版商用：页面自报的 PDF 直链（Wiley/CSIRO/Springer 等都会给）
    meta_pdf: (() => {
      const m = document.querySelector('meta[name="citation_pdf_url" i]');
      return m ? m.getAttribute('content') : null;
    })(),
    // 页面里所有像 PDF 的链接（供非 SD 出版商回退）
    pdf_links: [...new Set(links.filter(h => /pdf|pdfdirect|epdf|article-pdf/i.test(h)))].slice(0, 8),
    // 机构登录页（Wiley ssostart / Canadian ssostart 等）—— 登录态失效的信号
    sso_start: /\/action\/ssostart|\/institutional[-_]?login|\/shibboleth/i.test(url),
  });
})()
"""


# --------------------------------------------------------------------- 输入
_WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def safe_filename_component(value: str) -> str:
    """把不可信文本变成单个跨平台安全文件名组件。"""
    s = str(value or "").strip().replace("/", "_").replace("\\", "_").replace(":", "-")
    s = re.sub(r"[^A-Za-z0-9._()\-]+", "_", s)
    s = re.sub(r"_+", "_", s).strip(" ._")
    if not s:
        s = "document"
    if s.split(".", 1)[0].upper() in _WINDOWS_RESERVED:
        s = "_" + s
    return s[:200].rstrip(" .") or "document"


def safe_doi(doi: str) -> str:
    return safe_filename_component(str(doi or "").strip().lower())


def mask_ip(value: str | None) -> str:
    """日志默认只显示脱敏后的 IP，避免截图/日志意外泄露完整公网地址。"""
    s = str(value or "").strip()
    if not s:
        return "-"
    try:
        import ipaddress
        ip = ipaddress.ip_address(s)
        if ip.version == 4:
            a, b, _, _ = s.split(".")
            return f"{a}.{b}.x.x"
        parts = ip.exploded.split(":")
        return ":".join(parts[:3]) + ":…"
    except ValueError:
        return "(已隐藏)"


def redact_local_path(path: str | None) -> str:
    """绝对本地路径只保留末级目录；相对路径原样保留。"""
    if path is None:
        return "-"
    raw = str(path)
    expanded = os.path.expanduser(raw)
    if not os.path.isabs(expanded):
        return raw
    base = os.path.basename(os.path.normpath(expanded)) or "(root)"
    return f"…{os.sep}{base}"


def redact_url(url: str | None) -> str:
    """日志中的 URL 去掉 query/fragment，避免临时签名参数进入公开日志。"""
    try:
        u = urllib.parse.urlsplit(str(url or ""))
        if not u.scheme or not u.netloc:
            return str(url or "")[:120]
        return urllib.parse.urlunsplit((u.scheme, u.netloc, u.path, "", ""))
    except Exception:
        return "(URL 已隐藏)"


def contained_path(directory: str, filename: str) -> str:
    """返回受限于 directory 内的绝对路径，拒绝路径逃逸。"""
    base = os.path.abspath(directory)
    candidate = os.path.abspath(os.path.join(base, filename))
    try:
        inside = os.path.commonpath((base, candidate)) == base
    except ValueError:
        inside = False
    if not inside:
        raise ValueError(f"输出路径越界: {filename}")
    return candidate


def is_valid_pdf(path: str, expected_size: int | None = None,
                 min_size: int = 1024) -> tuple[bool, str]:
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            head = f.read(5)
    except OSError as e:
        return False, str(e)
    if expected_size is not None and size != expected_size:
        return False, f"length mismatch: expected {expected_size}, got {size}"
    if size < min_size:
        return False, f"file too small: {size}B"
    if head != b"%PDF-":
        return False, f"invalid PDF header: {head!r}"
    return True, f"{size}B"


def is_valid_saved_file(path: str, min_size: int = 1024) -> tuple[bool, str]:
    """校验已经落盘的正文/SI，供断点续跑决定是否真的可以跳过。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return is_valid_pdf(path, min_size=min_size)
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            head = f.read(512)
    except OSError as e:
        return False, str(e)
    if size < min_size:
        return False, f"file too small: {size}B"
    low = head.lstrip().lower()
    if low.startswith((b"<!doctype html", b"<html", b"<head", b"<body")):
        return False, "HTML response is not a downloadable file"
    if ext in (".zip", ".docx", ".xlsx", ".pptx") and not head.startswith(b"PK"):
        return False, f"invalid ZIP/Office signature: {head[:4]!r}"
    return True, f"{size}B"


def write_verified_bytes(out_path: str, data: bytes, expect_pdf: bool = False,
                         expected_size: int | None = None) -> tuple[bool, str]:
    """验证完整性后原子写入，失败不破坏已有文件。"""
    size = len(data)
    if expected_size is not None and size != expected_size:
        return False, f"length mismatch: expected {expected_size}, got {size}"
    if size < 1024:
        return False, f"file too small: {size}B"
    head = data[:512].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html", b"<head", b"<body")):
        return False, "HTML response is not a downloadable file"
    ext = os.path.splitext(out_path)[1].lower()
    if expect_pdf or ext == ".pdf":
        if data[:5] != b"%PDF-":
            return False, f"invalid PDF header: {data[:5]!r}"
    elif ext in (".zip", ".docx", ".xlsx", ".pptx") and not data.startswith(b"PK"):
        return False, f"invalid ZIP/Office signature: {data[:4]!r}"

    directory = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(directory, exist_ok=True)
    fd, part = tempfile.mkstemp(prefix=".scidownload-", suffix=".part", dir=directory)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(part, out_path)
    except Exception:
        try:
            os.remove(part)
        except OSError:
            pass
        raise
    return True, f"{size}B"


# Elsevier 对「没有独立补充材料」的文章会返回一个极小的**合法 zip**，
# 里面只有一个 `Data Profile.xml`（数据档案）。它是终态，不是抓取失败 ——
# 当成可重试问题会让整篇白等几轮长休息。
_SI_PLACEHOLDER_HINT = re.compile(r"data profile\.xml", re.I)


def classify_si_payload(data: bytes, ext: str = "") -> tuple[str, str]:
    """SI 载荷三分类：`ok` / `placeholder`（终态）/ `invalid`（可回退重试）。

    只有 `placeholder` 是新增的语义 —— 它让调用方能把「出版商确实没放 SI」
    和「这次没抓到」区分开，避免无意义的补跑。
    """
    if not data:
        return "invalid", "empty response"
    if data[:2] != b"PK":
        return "ok", ""            # 非 PK 容器交给 write_verified_bytes 常规校验
    try:
        import io
        import zipfile
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            infos = [i for i in z.infolist() if not i.is_dir()]
            names = [i.filename for i in infos]
            total = sum(i.file_size for i in infos)
    except Exception:
        return "invalid", "broken ZIP container"
    if not names:
        return "placeholder", "ZIP 内没有任何文件"
    if any(_SI_PLACEHOLDER_HINT.search(n) for n in names):
        return "placeholder", f"Elsevier 数据档案占位包（{', '.join(names[:3])}）"
    # 兜底：仅 1 个 xml 且**原始载荷 < 1024B**。
    # 阈值刻意与 write_verified_bytes 的下限对齐 —— 这样兜底只会把
    # 「原本就会因过小被拒、然后白跑重试」的情形改判为终态，
    # **不可能跳掉本来能拿到的 SI**。
    if len(data) < 1024 and len(names) == 1 and names[0].lower().endswith(".xml"):
        return "placeholder", f"仅含 1 个 xml（{names[0]}，原始 {len(data)}B）"
    return "ok", ""


def normalize_num(num) -> str:
    """Excel 数字单元格常给 145 或 145.0 —— 统一成 145。"""
    if num is None:
        return ""
    s = str(num).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def make_stem(row) -> str:
    """文件名主干。Num 与 DOI 都按单个安全文件名组件处理。"""
    raw_num = normalize_num(row.get("num"))
    num = safe_filename_component(raw_num) if raw_num else ""
    doi = safe_doi(row.get("doi", ""))
    return f"{num}_{doi}" if num else doi


# DOI 的特征：10.<4-9 位数字>/<后缀>。允许括号（Pedosphere / JIA 的 DOI 里就有括号）
DOI_RE = re.compile(r"10\.\d{4,9}/\S+")


def clean_doi(raw: str) -> str:
    """把从杂乱文本里抠出来的 DOI 修剪干净。

    注意：不能简单地把右括号一律砍掉 —— `10.1016/s1002-0160(17)60410-7`
    的括号是 DOI 本体的一部分。只有**括号不配对**时才砍。
    """
    s = str(raw).strip().strip("\"'<>（）").strip()
    s = s.rstrip("。，,;；::")
    while s.endswith(")") and s.count("(") < s.count(")"):
        s = s[:-1]
    while s.endswith("]") and s.count("[") < s.count("]"):
        s = s[:-1]
    return s.rstrip(".,;:").lower()


def extract_doi(text: str):
    """从任意一段文字里抠出第一个 DOI，抠不到返回 None。"""
    if not text:
        return None
    m = DOI_RE.search(str(text))
    return clean_doi(m.group(0)) if m else None


# -------------------------------------------------- xlsx：零依赖读取
def _xlsx_shared_strings(z):
    try:
        xml = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    out = []
    for si in root.findall(f"{NS}si"):
        # 富文本会拆成多个 <r><t>，要拼接
        out.append("".join(t.text or "" for t in si.iter(f"{NS}t")))
    return out


def _col_index(ref: str) -> int:
    """A1 -> 0, B1 -> 1, AA1 -> 26"""
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx_stdlib(path: str):
    """用标准库读 .xlsx 的第一张工作表，返回 list[list[str]]。

    不依赖 openpyxl —— 分享给别人时"只要 Python 就行"这句才成立。
    """
    import xml.etree.ElementTree as ET
    import zipfile
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    with zipfile.ZipFile(path) as z:
        shared = _xlsx_shared_strings(z)

        # 找到「第一张工作表」的真实文件路径（别硬编码 sheet1.xml）
        sheet_target = None
        try:
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rid_map = {r.get("Id"): r.get("Target") for r in rels}
            sheet = wb.find(f"{NS}sheets/{NS}sheet")
            rid = sheet.get(f"{RNS}id") if sheet is not None else None
            target = rid_map.get(rid)
            if target:
                # rels 里的 Target 可能是 `worksheets/sheet1.xml`
                # 也可能是 `/xl/worksheets/sheet1.xml`（绝对路径）——
                # 必须先剥掉开头的斜杠，否则会拼成 xl/xl/...
                t = str(target).lstrip("/")
                sheet_target = t if t.startswith("xl/") else "xl/" + t
        except Exception:
            sheet_target = None
        if not sheet_target:
            cands = [n for n in z.namelist()
                     if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            if not cands:
                raise SystemExit(f"这个 xlsx 里找不到工作表：{path}")
            sheet_target = sorted(cands)[0]

        root = ET.fromstring(z.read(sheet_target))

    rows = []
    for row in root.iter(f"{NS}row"):
        cells = {}
        for c in row.findall(f"{NS}c"):
            ref = c.get("r") or ""
            idx = _col_index(ref) if ref else len(cells)
            typ = c.get("t")
            if typ == "inlineStr":
                is_el = c.find(f"{NS}is")
                val = "".join(t.text or "" for t in is_el.iter(f"{NS}t")) if is_el is not None else ""
            else:
                v = c.find(f"{NS}v")
                val = v.text or "" if v is not None else ""
                if typ == "s":
                    try:
                        val = shared[int(val)]
                    except Exception:
                        val = ""
            cells[idx] = val
        if cells:
            width = max(cells) + 1
            rows.append([cells.get(i, "") for i in range(width)])
    return rows


DOI_HEADERS = ("doi", "doi号", "文献doi", "文献 doi")
NUM_HEADERS = ("num", "序号", "编号", "id", "no", "no.")
TITLE_HEADERS = ("title", "title1", "题名", "标题", "文献题名", "篇名", "题目")


def _find_header_row(rows, max_scan: int = 15) -> int:
    """表头不一定在第一行 —— Excel 清单常有标题行、说明行、空行。

    在前 max_scan 行里找**含 DOI 表头**的那一行；找不到返回 -1
    （表示"没有表头"，那就整表扫 DOI）。
    """
    for i, r in enumerate(rows[:max_scan]):
        cells = [str(x).strip().lower() for x in (r or [])]
        if any(h in DOI_HEADERS for h in cells if h):
            return i
    return -1


def _rows_to_dois(rows):
    """从表格里取 DOI / Num / Title。

    优先找 DOI 表头列；找不到就在所有单元格里正则扫 DOI。
    """
    if not rows:
        return []
    hi = _find_header_row(rows)
    i_doi = i_num = i_title = None
    if hi >= 0:
        header = [str(x).strip().lower() for x in rows[hi]]
        i_doi = next((i for i, h in enumerate(header) if h in DOI_HEADERS), None)
        i_num = next((i for i, h in enumerate(header) if h in NUM_HEADERS), None)
        i_title = next((i for i, h in enumerate(header) if h in TITLE_HEADERS), None)
        body = rows[hi + 1:]
    else:
        body = rows

    def cell(r, i):
        return r[i] if (i is not None and i < len(r)) else None

    out = []
    if i_doi is not None:
        for r in body:
            r = r or []
            if i_doi < len(r):
                # 即使列名写着 DOI，也只接受真正符合 DOI 形状的内容；
                # N/A、not-a-doi 等占位文本一律跳过。
                d = extract_doi(r[i_doi])
                if d:
                    out.append({"doi": d, "num": cell(r, i_num), "title": cell(r, i_title)})
        return out

    # 没有 DOI 表头：把每个单元格都扫一遍
    # （能处理"题名+DOI 混在一列"，也能处理"表头不在第一行"）
    for r in body:
        for c in (r or []):
            found = extract_doi(c)
            if found:
                out.append({"doi": found, "num": cell(r, i_num), "title": None})
                break
    return out


def read_inputs(path: str):
    """支持：纯文本（每行一个 DOI）/ csv / tsv / xlsx —— 全部零依赖。

    任何格式都容忍"DOI 混在文字里"：会先用正则抠 DOI，抠不到才退回按列取。
    """
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        return _rows_to_dois(read_xlsx_stdlib(path))
    if ext == ".xls":
        raise SystemExit(
            "旧版 .xls（二进制格式）标准库读不了。\n"
            "  用 Excel 打开后「另存为」→ 选 .xlsx 或 .csv 即可。")

    text = open(path, encoding="utf-8-sig", errors="replace").read()
    lines = text.splitlines()
    first = lines[0] if lines else ""

    # 分隔符与表头都不能只看第一行 —— 清单前面常有标题行 / 空行 / 说明行
    sample = "\n".join(lines[:15])
    delim = ""
    for d in ("\t", ",", ";"):
        if d in sample:
            delim = d
            break

    has_doi_header = False
    for ln in lines[:15]:
        cells = [c.strip().lower() for c in re.split(r"[\t,;]", ln)]
        if any(c in DOI_HEADERS for c in cells if c):
            has_doi_header = True
            break

    first_cells = [c.strip().lower() for c in re.split(r"[\t,;]", first)] if first else []

    # 表格类：要么表头明写 DOI，要么明显是多列（>=3 列）
    if delim and (has_doi_header or len(first_cells) >= 3):
        rows = _rows_to_dois(list(csv.reader(lines, delimiter=delim)))
        if rows:
            return rows

    # 兜底：逐行抠 DOI。**只认 DOI 形状**，不能把普通单词当 DOI
    out = []
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        d = extract_doi(line)
        if d:
            out.append({"doi": d, "num": None, "title": None})
    return out


def dedupe_rows(rows):
    """按规范化 DOI 去重，稳定保留第一次出现的元数据。"""
    seen = set()
    out = []
    for source in rows:
        row = dict(source)
        doi = clean_doi(row.get("doi", ""))
        if not doi or doi in seen:
            continue
        seen.add(doi)
        row["doi"] = doi
        out.append(row)
    return out


def filter_rows_by_only(rows, only: str):
    if not only:
        return list(rows)
    wanted = {clean_doi(value) for value in only.split(",") if value.strip()}
    return [row for row in rows if clean_doi(row.get("doi", "")) in wanted]


def validate_runtime_options(*, pace, batch_size, batch_rest, backoff,
                             retry_failed, port, login_wait, limit):
    """在创建目录或启动浏览器前拒绝危险/无意义的运行参数。"""
    for name, value in {
        "pace": pace,
        "batch-rest": batch_rest,
        "backoff": backoff,
        "login-wait": login_wait,
    }.items():
        if not math.isfinite(value) or value < 0:
            raise ValueError(f"--{name} 必须是非负有限数")
    for name, value in {
        "batch-size": batch_size,
        "retry-failed": retry_failed,
        "limit": limit,
    }.items():
        if value < 0:
            raise ValueError(f"--{name} 必须是非负整数")
    if not 1 <= port <= 65535:
        raise ValueError("--port 必须在 1..65535 之间")


def _normalize_manifest_record(record):
    """把 1.0 旧记录补成新版字段；不改写磁盘上的历史记录。"""
    rec = dict(record)
    doi = clean_doi(rec.get("doi", ""))
    rec["doi"] = doi
    if "pdf_status" not in rec:
        rec["pdf_status"] = (
            "downloaded" if rec.get("status") == "downloaded" and rec.get("pdf")
            else "not_downloaded"
        )
    if "si_requested" not in rec:
        rec["si_requested"] = bool(rec.get("si"))
    if "si_status" not in rec:
        if rec["si_requested"]:
            rec["si_status"] = "downloaded"
        else:
            rec["si_status"] = "not_requested"
    return rec


def load_manifest_latest(path: str):
    """读取每个 DOI 的最后一条有效记录，后续失败也会覆盖先前成功。"""
    latest = {}
    if not os.path.exists(path):
        return latest
    with open(path, encoding="utf-8", errors="replace") as f:
        for line in f:
            try:
                rec = _normalize_manifest_record(json.loads(line))
            except (ValueError, TypeError, json.JSONDecodeError):
                continue
            if rec.get("doi"):
                latest[rec["doi"]] = rec
    return latest


SI_COMPLETE = {"downloaded", "none_found"}
SI_TERMINAL = SI_COMPLETE | {"not_supported"}

# 只有这两类失败才值得怀疑是站点验证/网络问题、才触发退避。
# no_entitlement（当前会话无全文权限）、needs_manual_login（登录态失效）、bad_doi、
# unsupported_publisher 重试都不会变好；"PDF 已下到、仅 SI 缺"更不是反爬信号。
# 把它们算进"连续失败"会把正常的批量跑法误判成"疑似限流/验证失败"白等 900 秒
# （2026-09-19 实测教训：会长的 v1.0 在混合清单上连续撞 900 秒退避）。
PDF_RETRYABLE_STATUSES = {"challenge_not_resolved", "error"}


def interruptible_sleep(seconds: float, log=None, announce: str = "") -> None:
    """分片睡眠：Ctrl+C 随时能在 ≤0.5 秒内被响应，等待期间每 30 秒报剩余时间。"""
    if seconds <= 0:
        return
    if announce and log:
        log(announce)
    end = time.time() + seconds
    next_hint = time.time() + 30.0
    while True:
        remain = end - time.time()
        if remain <= 0:
            break
        if log and remain > 5 and time.time() >= next_hint:
            log(f"    …等待中，还剩 {remain:.0f} 秒（Ctrl+C 可随时退出，"
                "已下载的会自动跳过）")
            next_hint = time.time() + 30.0
        time.sleep(min(0.5, remain))


def resume_action(record, pdfdir: str, want_si: bool, sidir: str | None = None) -> str:
    """返回 full / si_only / skip_all；正文和已声明下载的 SI 都以磁盘实物为准。"""
    if not record or record.get("pdf_status") != "downloaded" or not record.get("pdf"):
        return "full"
    try:
        pdf_name = str(record["pdf"])
        if os.path.basename(pdf_name) != pdf_name:
            return "full"
        pdf_path = contained_path(pdfdir, pdf_name)
    except (TypeError, ValueError):
        return "full"
    valid, _ = is_valid_pdf(pdf_path)
    if not valid:
        return "full"
    if not want_si:
        return "skip_all"

    if not record.get("si_requested"):
        return "si_only"
    status = record.get("si_status")
    if status in {"none_found", "not_supported"}:
        return "skip_all"
    if status != "downloaded":
        return "si_only"

    names = list(record.get("si") or [])
    if not names:
        return "si_only"
    sidir = sidir or os.path.join(os.path.dirname(os.path.abspath(pdfdir)),
                                  "SupportingInformation")
    for name in names:
        name = str(name or "")
        if not name or os.path.basename(name) != name:
            return "si_only"
        try:
            path = contained_path(sidir, name)
        except ValueError:
            return "si_only"
        ok, _ = is_valid_saved_file(path)
        if not ok:
            return "si_only"
    return "skip_all"


def record_complete(record, pdfdir: str, want_si: bool, sidir: str | None = None) -> bool:
    return resume_action(record, pdfdir, want_si, sidir) == "skip_all"


# --------------------------------------------------------------------- 抓取
# 非 ScienceDirect 出版商的 PDF 入口规则。
# 已有专门适配：Wiley 系 / Canadian Science / CSIRO / Copernicus；
# v1.5.2 新增 PeerJ / Nature（含 Scientific Reports）；v1.6.0 新增 MDPI attachment 捕获。
#
# 对没有专门规则的出版商，主流程还会尝试页面标准元数据
# `citation_pdf_url` 与页面中明确的 PDF 链接；只有专门规则和通用候选都
# 无法取得有效 PDF 时，才会标记 unsupported_publisher。
def _wiley_pdf_candidates(info, doi: str) -> list:
    """Wiley 系 PDF 候选：按文章实际落地的子域优先，再补通用的 www。

    为什么不能只给 www：Wiley 的子域是按刊物分的
    （`onlinelibrary` / `acsess.onlinelibrary` / `scijournals.onlinelibrary` …），
    `pdfdirect` 路径要和子域配套。只试 www 时，落在其它子域的文章会
    一路失败到超时（实测 10.1002/ghg.1892 白等 7 分钟）。
    """
    cands = []
    url = (info or {}).get("url") or ""
    m = re.match(r"https?://([^/]+)/", url)
    host = m.group(1) if m else ""
    if host.endswith("onlinelibrary.wiley.com"):
        cands.append(f"https://{host}/doi/pdfdirect/{doi}")
    cands.append(f"https://onlinelibrary.wiley.com/doi/pdfdirect/{doi}")
    if host.endswith("onlinelibrary.wiley.com"):
        cands.append(f"https://{host}/doi/pdf/{doi}")
    # 去重保序
    seen, out = set(), []
    for u in cands:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def _peerj_pdf_candidates(info, doi: str) -> list:
    """PeerJ 正文候选。

    标准 PeerJ DOI `10.7717/peerj.<article_id>` 可映射到
    `https://peerj.com/articles/<article_id>.pdf`。其它 10.7717 DOI 不臆测
    路径，交给页面 `citation_pdf_url` / 通用 PDF 链接兜底。
    """
    d = (doi or "").strip().lower()
    m = re.fullmatch(r"10\.7717/peerj\.(\d+)", d)
    return [f"https://peerj.com/articles/{m.group(1)}.pdf"] if m else []


def _nature_pdf_candidates(info, doi: str) -> list:
    """Nature Portfolio 正文候选（含 Scientific Reports）。"""
    d = (doi or "").strip()
    m = re.fullmatch(r"10\.1038/(.+)", d, re.I)
    if not m:
        return []
    article_id = urllib.parse.quote(m.group(1), safe="-._~")
    return [f"https://www.nature.com/articles/{article_id}.pdf"]


def generic_pdf_candidates(info: dict) -> list[str]:
    """从文章页标准元数据/链接提取安全的通用 PDF 候选。

    只接受 http(s) URL；去重保序。候选是否真的为 PDF 仍由
    `fetch_pdf_other()` + `%PDF-` 文件验证决定，因此普通 HTML 不会被保存。
    """
    raw = []
    if (info or {}).get("meta_pdf"):
        raw.append(info["meta_pdf"])
    raw.extend(list((info or {}).get("pdf_links") or []))
    seen, out = set(), []
    for u in raw:
        if not isinstance(u, str):
            continue
        u = u.strip()
        try:
            parsed = urllib.parse.urlparse(u)
        except Exception:
            continue
        if parsed.scheme not in ("http", "https") or not parsed.netloc:
            continue
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def non_sd_pdf_candidates(info: dict, doi: str, rule_fn=None) -> list[str]:
    """合并“出版商专门规则 + 页面通用候选”，去重保序。"""
    cands = []
    if rule_fn is not None:
        try:
            cands.extend([u for u in rule_fn(info, doi) if u])
        except Exception:
            pass
    cands.extend(generic_pdf_candidates(info))
    seen, out = set(), []
    for u in cands:
        if u and u not in seen:
            seen.add(u)
            out.append(u)
    return out


PUBLISHER_RULES = (
    # (标签, 正则, 入口构造函数)
    ("peerj", r"^10\.7717/",
     lambda info, doi: _peerj_pdf_candidates(info, doi)),
    ("nature", r"^10\.1038/",
     lambda info, doi: _nature_pdf_candidates(info, doi)),
    # MDPI：正文 PDF 常以 Content-Disposition: attachment 触发浏览器原生下载。
    # v1.6.0 起由 CDP 把浏览器下载定向到临时目录并接管验证/落盘。
    ("mdpi", r"^10\.3390/",
     lambda info, doi: [info["meta_pdf"]] if info.get("meta_pdf") else []),
    # Wiley 系（含 ACSESS / SSSAJ / scijournals 等子域）。**必须用 pdfdirect**：
    # 页面 meta 与链接都写 /doi/pdf/，但那个地址会 302 回落地页（拿到 HTML）。
    # 注意子域不一样：多数在 www，但 ACSESS / scijournals 等有自己的子域，
    # 必须按文章实际落地的域名构造，否则拿不到东西（2026-09-19 实测
    # 10.1002/ghg.1892 落在 scijournals 子域，只给 www 会连试 7 分钟）。
    ("wiley", r"^10\.(1002|1111|2134|2136)/",
     lambda info, doi: _wiley_pdf_candidates(info, doi)),
    # Canadian Science Publishing：/doi/pdf/{doi} 可直接取字节
    ("cdnscience", r"^10\.(1139|4141)/",
     lambda info, doi: [f"https://cdnsciencepub.com/doi/pdf/{doi}"]),
    # CSIRO（connectsci.au）：用页面 meta，会先签名跳转到 silverchair CDN
    ("csiro", r"^10\.1071/",
     lambda info, doi: [info["meta_pdf"]] if info.get("meta_pdf") else []),
    # Copernicus（EGU，OA）：页面 meta 即真 PDF 直链
    ("copernicus", r"^10\.5194/",
     lambda info, doi: [info["meta_pdf"]] if info.get("meta_pdf") else []),
)

# 已知不支持 / 需要别的途径的出版商 —— 明确报出来，
# 免得用户以为是"没权限"或"程序坏了"，反复重试。
KNOWN_UNSUPPORTED = (
    (r"^10\.1080/", "Taylor & Francis",
     "当前版本尚未适配该出版商；不同机构的订阅情况可能不同"),
    (r"^10\.1007/", "Springer",
     "Springer 的 PDF 需要先过查看器页，暂未实现"),
    (r"^10\.1017/", "Cambridge University Press", "未实现"),
    (r"^10\.3724/", "Science Press / sciengine", "页面给的是查看器页而非 PDF，未实现"),
    (r"^10\.5846/", "生态学报", "站点有滑块反爬，未实现"),
    (r"^10\.2139/", "SSRN", "预印本平台，入口特殊，未实现"),
)


def publisher_rule(doi: str):
    """返回 (标签, 入口构造函数)；没有匹配则 (None, None)。"""
    d = (doi or "").strip().lower()
    for label, pat, fn in PUBLISHER_RULES:
        if re.search(pat, d):
            return label, fn
    return None, None


def known_unsupported(doi: str):
    """返回 (标签, 说明)；不是已知的不支持出版商则 (None, None)。"""
    d = (doi or "").strip().lower()
    for pat, label, why in KNOWN_UNSUPPORTED:
        if re.search(pat, d):
            return label, why
    return None, None


def is_sciencedirect_info(info: dict) -> bool:
    """不只依赖 PII；页面已经落在 ScienceDirect 也应按 SD 分支处理。"""
    host = str((info or {}).get("host") or "").lower()
    return bool((info or {}).get("pii")) or (info or {}).get("on_sd") is True \
        or host == "sciencedirect.com" or host.endswith(".sciencedirect.com")


def _parse_complete(info: dict) -> bool:
    """页面是否已经渲染到可以判断的程度。

    两条路：
      - **ScienceDirect**：靠 `pii` + 授权信号判定。
        全新 profile 的第一次访问要建 DNS/TLS/缓存、还要过 ScienceDirect 的首次
        cookie 流程，7 秒常常不够 —— 那时解析出来就是 机构=None / access=unknown /
        pdf=无，看着像"没权限"或"没有 PDF 入口"，其实只是没加载完。
      - **非 ScienceDirect**：这些站没有 `pii`，且**几乎永远不会**
        `readyState == complete` —— Wiley 这类站点会持续挂着一两个第三方
        （广告 / 统计）请求，`ready` 长期停在 `interactive`（实测 15s 仍不 complete）。
        所以不能拿 `complete` 当闸门，改为：**已落到出版商域名 + 解析出标题** 即算就绪。
    """
    if is_sciencedirect_info(info):
        if info.get("pii") and (info.get("brought_by") or info.get("access") != "unknown"):
            return True
        # 极少数页面已明确给出授权结论但 PII 尚未从路径解析出来；
        # 这时也可以安全结束等待，后续仍按 SD 分支处理，不会误判成其它出版商。
        if info.get("access") in {"granted", "denied"} and info.get("title"):
            return True
        return False
    # 非 SD：已离开 doi.org / SD 落地页，且取到了标题 —— 足够判断出版商与入口
    if info.get("on_sd") is False:
        host = info.get("host") or ""
        landed = bool(host) and "doi.org" not in host and "sciencedirect" not in host
        if landed and info.get("title"):
            return True
        # 兜底：某些站点拿不到 host 但 readyState 已 complete
        if info.get("ready") == "complete" and info.get("url") and info.get("title"):
            return True
    # 还没确认是不是 SD（落地页尚未就绪）—— 继续等
    return False



def resolve_article(tab, doi, log, login_wait=0.0, attempts=3):
    """第 ①② 步：打开 DOI，等页面**真正**就绪，取出 PDF 入口 / SI 链接 / 授权状态。

    用「轮询就绪」而不是「固定死等」：页面 2 秒渲染完就 2 秒走，
    实测单篇能省 4~5 秒，而安全性不变（等的是同一个就绪信号，只是更灵敏）。

    间隔按站点分流（见 `_parse_complete`）：
      - ScienceDirect：1.2s 常规间隔；
      - 非 SD（Wiley / CSIRO / Canadian…）：0.5s 快探 —— 这些站会持续挂着
        第三方请求永不 `complete`，短间隔才能一落地就走。
    """
    info = {}
    for k in range(1, attempts + 1):
        try:
            tab.goto(f"https://doi.org/{doi}", settle=0.35 if k == 1 else 1.5)
        except Exception:
            pass
        deadline = time.time() + (6.0 if k == 1 else 12.0)
        while True:
            try:
                info = json.loads(tab.js(ARTICLE_JS, timeout=40) or "{}")
            except Exception:
                info = {}
            # 有结论就立刻返回：解析完整 / 停在登录页 / DOI 打不开
            if _parse_complete(info) or info.get("loginish") or info.get("notfound"):
                return _finish_login(tab, doi, info, log, login_wait)
            if time.time() >= deadline:
                break
            # 非 SD 站点不会 complete，用固定短间隔快探（0.5s）；
            # SD 页面加载完就转常规间隔，避免高频打扰。
            time.sleep(0.5 if info.get("on_sd") is False else 1.2)
        if k < attempts:
            log(f"    (页面未渲染完整：pii={info.get('pii')} "
                f"机构={info.get('brought_by')} access={info.get('access')}"
                f"，重试 {k + 1}/{attempts})")
    return _finish_login(tab, doi, info, log, login_wait)


def _finish_login(tab, doi, info, log, login_wait):
    """停在机构登录页时，按需等用户手工登录一次。"""
    if info.get("loginish"):
        if login_wait > 0:
            log(f"    ↳ 页面停在登录/机构认证（{redact_url(info.get('url'))[:70]}）")
            log(f"    ↳ 请在弹出的浏览器窗口里完成登录，最多等 {login_wait:.0f} 秒…")
            deadline = time.time() + login_wait
            while time.time() < deadline:
                time.sleep(3)
                try:
                    cur = tab.js("location.href") or ""
                    if not re.search(r"login|authserver|carsi|shibboleth|ezproxy", cur, re.I):
                        break
                except Exception:
                    pass
            try:
                tab.goto(f"https://doi.org/{doi}", settle=2.0)
                info = json.loads(tab.js(ARTICLE_JS, timeout=40) or "{}")
            except Exception:
                pass
        if info.get("loginish"):
            info["_needs_login"] = True
    return info


def fetch_pdf(tab, pdf_url, out_path, attempts=3, log=print):
    """第 ③④ 步：导航到 PDF 入口 → 完成站点验证 → 在 PDF 文档上下文取字节。

    同样用轮询而非死等：PDF 文档一就绪立刻取字节（实测每篇省 2~3 秒）。
    """
    if not pdf_url:
        return False, "页面未提供 PDF 入口"
    for k in range(1, attempts + 1):
        tab.front()
        try:
            tab.goto(pdf_url, settle=1.0)
        except Exception:
            pass
        deadline = time.time() + 60 + 30 * (k - 1)
        while time.time() < deadline:
            try:
                tab.front()          # 必须反复提前台，否则挑战页定时器被节流
                # 一次取两个值，省一半 CDP 往返
                st = tab.js("JSON.stringify([document.contentType, location.href])",
                            timeout=15)
                ct, href = json.loads(st) if st else (None, "")
                if ct == "application/pdf" or "pdf.sciencedirectassets.com" in (href or ""):
                    ok, msg = cdp.fetch_bytes_via_page(tab, href, out_path)
                    return ok, (f"{msg} (第 {k} 次)" if ok else msg)
            except Exception:
                pass
            time.sleep(1.0)
        time.sleep(4)
    return False, f"站点验证连续 {attempts} 次都未完成"


def fetch_pdf_other(tab, urls, out_path, log=print, attempts=2,
                    per_url_wait: float = 20.0):
    """非 ScienceDirect 出版商：同时兼容“标签页 PDF”和“浏览器附件下载”。

    v1.6.0 的关键变化：部分出版商（实机发现 MDPI）访问 PDF URL 时不会在
    当前标签页打开 PDF，而是返回 ``Content-Disposition: attachment``，Chrome
    会直接把文件交给下载管理器。旧逻辑只检查 ``document.contentType``，因此会
    出现“文件其实已进系统 Downloads，但程序判失败并再次访问，导致下载两次”。

    现在每个候选 URL 都先通过 CDP 设置**专用临时下载目录**，然后一次导航同时监控：
      1) 若浏览器触发 attachment 下载：等待 .crdownload 完成，验证 ``%PDF-``，
         原子移动到 ``out_path``；
      2) 若 PDF 在标签页内打开：继续用页面上下文取字节；
      3) 若只是普通 HTML/摘要页：按原逻辑判断是否无权限或继续尝试其它候选。

    一旦某个 URL 已触发浏览器下载，即使该次下载失败/超时，也不会在同一轮再次
    导航这个 URL，从源头避免 ``paper.pdf`` + ``paper (1).pdf`` 式重复下载。
    """
    last = "没有可用的候选 PDF 地址"
    clean_urls = [u for u in urls if u]
    n = len(clean_urls)
    wait = per_url_wait if n > 1 else 45.0
    browser_started_urls = set()

    for k in range(1, attempts + 1):
        for u in clean_urls:
            if u in browser_started_urls:
                continue

            ok, msg, download_started, href = cdp.fetch_pdf_candidate(
                tab, u, out_path, total_timeout=wait
            )
            if ok:
                return True, msg

            last = msg
            if download_started:
                # 这个 URL 已真实触发过浏览器下载。不要在下一 attempt 再点一次，
                # 否则 Chrome 会产生“(1)”重复文件；仍可继续尝试其它候选 URL。
                browser_started_urls.add(u)

            h = str(href or "").lower()
            if "/doi/abs/" in h or "/doi/full/" in h:
                return False, "no_entitlement: PDF 入口被转回摘要页"

        if k < attempts:
            # 如果所有候选都已经触发过浏览器附件下载，就没有必要再开第二轮。
            if clean_urls and all(u in browser_started_urls for u in clean_urls):
                break
            time.sleep(2.0)
    return False, last


def si_filename(url: str) -> str:
    """从 SI 链接推出文件名（含扩展名）。

    - ScienceDirect：`.../-mmc1.pdf` → 用 URL 路径的扩展名
    - Wiley 系：`/action/downloadSupplement?doi=…&file=gcb15410-sup-0001-supinfo.docx`
      → 真实文件名在 `file=` 参数里，**不能**拿 URL 末尾当扩展名
      （那样会得到 "downloadSupplement" 这种没有扩展名的怪名字）。
    """
    try:
        q = urllib.parse.urlparse(url).query
        params = urllib.parse.parse_qs(q)
        for key in ("file", "filename", "name"):
            vals = params.get(key) or []
            if vals and vals[0]:
                cand = safe_filename_component(urllib.parse.unquote(vals[0]))
                if cand:
                    return cand
    except Exception:
        pass
    tail = os.path.basename(urllib.parse.urlparse(url).path)
    return safe_filename_component(tail) or "supplement.bin"


def download_si(tab, si_urls, outdir, stem, log=print):
    os.makedirs(outdir, exist_ok=True)
    got = []
    placeholders = []
    for u in si_urls:
        # 文件名优先取出版商给的原始名（Wiley 的 file= 参数），
        # 否则退回 URL 路径。带 stem 前缀，避免不同文章的文件撞名。
        base = si_filename(u)
        ext = os.path.splitext(base)[1] or ".bin"
        mm = re.search(r"-mmc(\d+)\.", u, re.I)
        if mm:
            name = f"{stem}_mmc{mm.group(1)}{ext}"
        elif base.lower().startswith("downloadsupplement"):
            name = f"{stem}_SI{ext}"
        else:
            name = f"{stem}_{base}" if base != "supplement.bin" else f"{stem}_SI{ext}"
        name = safe_filename_component(name)
        path = contained_path(outdir, name)
        existing_ok, _ = is_valid_saved_file(path)
        if existing_ok:
            got.append(name)
            log(f"    SI ↷ 已验证现有文件，未重复下载：{name}")
            continue
        # 这些链接通常在同一站点/公开 CDN 上，先走直连（快）
        referer = "https://www.sciencedirect.com/"
        host = urllib.parse.urlparse(u).netloc
        if host:
            referer = f"https://{host}/"
        try:
            req = urllib.request.Request(u, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36",
                "Referer": referer})
            with OPENER.open(req, timeout=120) as r:
                b = r.read()
            # 先分清「出版商占位包」和「这次真没抓到」—— 前者是终态，重试纯属白等
            verdict, why = classify_si_payload(b, ext)
            if verdict == "placeholder":
                placeholders.append(name)
                log(f"    SI ⊘ {name} 无实质内容：{why}（终态，不重试）")
                continue
            ok, msg = write_verified_bytes(path, b, expect_pdf=(ext == ".pdf"))
            if ok:
                got.append(name)
                log(f"    SI ✓ {name}  {len(b):,}B")
                continue
            log(f"    SI 直连未通过验证 {name}: {msg}，改用页面上下文")
        except Exception:
            pass
        # 回退：页面上下文（Wiley 的 downloadSupplement 需要登录 cookie，
        # 直连常拿到登录页 HTML，这时必须借页面上下文带 cookie 取）
        try:
            ok, msg = cdp.fetch_bytes_via_page(tab, u, path,
                                               expect_pdf=(ext == ".pdf"))
            if ok:
                got.append(name)
                log(f"    SI ✓ {name}  {msg}")
            else:
                log(f"    SI ✗ {name} {msg}")
        except Exception as e:
            log(f"    SI ✗ {name} {repr(e)[:80]}")
    return got, placeholders


SUPPORTED_INPUT_EXTENSIONS = {".txt", ".csv", ".tsv", ".xlsx", ".xlsm"}
AUTO_DISCOVERY_IGNORES = {
    "requirements.txt", "environment.txt", "pip-freeze.txt",
}


def discover_input_candidates(directory: str = ".") -> list[str]:
    """返回目录中可作为 DOI 清单的候选文件。

    自动发现故意保持保守：只看程序明确支持的扩展名；隐藏文件和常见
    Python 环境清单不参与。若候选不止一个，调用方必须让用户明确指定，
    不能靠文件名/修改时间猜测。
    """
    base = os.path.abspath(directory)
    found: list[str] = []
    try:
        names = sorted(os.listdir(base), key=str.lower)
    except OSError:
        return found
    for name in names:
        if not name or name.startswith("."):
            continue
        if name.lower() in AUTO_DISCOVERY_IGNORES:
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED_INPUT_EXTENSIONS:
            continue
        path = os.path.join(base, name)
        if os.path.isfile(path):
            found.append(path)
    return found


def resolve_input_path(explicit: str | None, directory: str = ".") -> tuple[str | None, str, list[str]]:
    """解析输入清单；返回 (path, mode, candidates)。

    mode 为 explicit / auto / none / ambiguous。
    """
    if explicit:
        return explicit, "explicit", []
    candidates = discover_input_candidates(directory)
    if len(candidates) == 1:
        return candidates[0], "auto", candidates
    if not candidates:
        return None, "none", candidates
    return None, "ambiguous", candidates


# --------------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(
        description="按 DOI 清单批量下载已有访问权的论文（ScienceDirect + 多出版商）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例:\n"
               "  python 3.SCIDownload.py dois.txt --out ./papers --si\n"
               "  python 3.SCIDownload.py list.csv --out ./papers --no-si --limit 5\n"
               "  python 3.SCIDownload.py --limit 3 --si   # 当前目录只有一个清单时自动发现\n")
    ap.add_argument("inputs", nargs="?",
                    help="DOI 清单：.txt（每行一个）/ .csv / .tsv / .xlsx / .xlsm；省略时自动扫描当前目录")
    ap.add_argument("--out", default="./SCIDownload-out", help="输出目录（默认 ./SCIDownload-out）")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--si", dest="si", action="store_true", help="同时下载补充材料")
    g.add_argument("--no-si", dest="si", action="store_false", help="只下正文")
    ap.set_defaults(si=False)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--only", default="", help="逗号分隔的 DOI 白名单")
    ap.add_argument("--force", action="store_true", help="忽略断点，全部重做")
    ap.add_argument("--pace", type=float, default=6.0,
                    help="篇间基础停顿秒数（默认 6，实际会乘 0.7~1.6 的随机抖动）")
    ap.add_argument("--batch-size", type=int, default=20, metavar="N",
                    help="每处理 N 篇插入一次长休息，降低累计限流/验证风险（默认 20，0 = 不休息）")
    ap.add_argument("--batch-rest", type=float, default=120.0,
                    help="长休息秒数（默认 120，实际会乘 0.8~1.4 的随机抖动）")
    ap.add_argument("--backoff", type=float, default=60.0,
                    help="连续 3 篇以上未成功时的退避基准秒数（默认 60，按 2 的幂递增）")
    ap.add_argument("--retry-failed", type=int, default=2, metavar="N",
                    help="跑完一轮后，自动补跑未下载的篇目，最多再补 N 轮（默认 2，设 0 关闭）")
    ap.add_argument("--port", type=int, default=9222)
    ap.add_argument("--profile", default=None, help="浏览器 profile 目录（持久化登录态）")
    ap.add_argument("--browser", default=None, help="浏览器可执行文件路径（默认自动探测）")
    ap.add_argument("--headless", action="store_true", help="无头模式（不推荐，部分站点兼容性较差）")
    ap.add_argument("--login-wait", type=float, default=0.0,
                    help="若页面停在机构登录页，等待手工登录的秒数（CARSI/EZproxy 机构用）")
    ap.add_argument("--keep-browser", action="store_true", help="结束后不关闭浏览器")
    ap.add_argument("--ip-check", action="store_true",
                    help="显式运行出口网络提示（默认不联系第三方 IP 查询服务）")
    ap.add_argument("--skip-ip-check", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()

    input_path, input_mode, candidates = resolve_input_path(args.inputs, os.getcwd())
    if input_mode == "auto":
        args.inputs = input_path
        print(f"未指定清单文件；自动发现唯一候选：{os.path.basename(args.inputs)}")
    elif input_mode == "ambiguous":
        print("未指定清单文件，但当前目录发现多个候选；为避免选错，请明确指定一个：")
        for p in candidates:
            print(f"  - {os.path.basename(p)}")
        print()
        print('例如：python 3.SCIDownload.py "你的清单.xlsx" --limit 3 --si')
        return 2
    elif input_mode == "none":
        print(__doc__)
        print("当前目录没有发现可用的 DOI 清单（支持 .txt/.csv/.tsv/.xlsx/.xlsm）。")
        print("把清单放到当前目录后，可直接运行：")
        print("  python 3.SCIDownload.py --limit 3 --si")
        print("也可以始终显式指定文件：")
        print("  python 3.SCIDownload.py dois.txt --limit 3 --si")
        print()
        print("提示：自动发现只在当前工作目录扫描；若有多个候选文件，程序不会自行猜测。")
        return 2

    try:
        validate_runtime_options(
            pace=args.pace, batch_size=args.batch_size,
            batch_rest=args.batch_rest, backoff=args.backoff,
            retry_failed=args.retry_failed, port=args.port,
            login_wait=args.login_wait, limit=args.limit,
        )
    except ValueError as e:
        ap.error(str(e))

    if not os.path.exists(args.inputs):
        raise SystemExit(
            f"找不到清单文件：{args.inputs}\n"
            f"  当前目录：{os.getcwd()}\n"
            f"  请确认路径，或先 cd 到清单所在目录再执行。")

    pdfdir = os.path.join(args.out, "PDFs")
    sidir = os.path.join(args.out, "SupportingInformation")
    os.makedirs(pdfdir, exist_ok=True)
    logpath = os.path.join(args.out, "run.log")
    manpath = os.path.join(args.out, "manifest.jsonl")

    def log(msg):
        line = f"[{time.strftime('%H:%M:%S')}] {msg}"
        print(line, flush=True)
        with open(logpath, "a", encoding="utf-8") as f:
            f.write(line + "\n")

    parsed_rows = read_inputs(args.inputs)
    total_parsed = len(parsed_rows)
    rows = dedupe_rows(parsed_rows)
    total_unique = len(rows)
    rows = filter_rows_by_only(rows, args.only)
    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        raise SystemExit("清单里没有解析到任何 DOI。")

    log(f"SCIDownload v{VERSION} | 解析 {total_parsed} 条、去重后 {total_unique} 篇，"
        f"本次处理 {len(rows)} 篇 | "
        f"SI={'开' if args.si else '关'} | 输出 {redact_local_path(args.out)}")
    log("停止方式：随时 Ctrl+C（长等待期间也可以）；重跑同一条命令会自动跳过"
        "已下载的篇目。若终端不响应 Ctrl+C，直接关闭终端窗口即可，"
        "已落盘的文件不受影响。")
    if total_unique != len(rows):
        log(f"  (--limit/--only 生效：{total_unique} → {len(rows)})")

    latest = {} if args.force else load_manifest_latest(manpath)
    results = {}
    pending = []
    skip_count = si_only_count = 0
    for row in rows:
        previous = latest.get(row["doi"])
        action = "full" if args.force else resume_action(previous, pdfdir, args.si, sidir)
        if action == "skip_all":
            results[row["doi"]] = previous
            skip_count += 1
            continue
        work = dict(row)
        work["_resume_action"] = action
        work["_previous"] = previous
        pending.append(work)
        if action == "si_only":
            si_only_count += 1
    if latest:
        log(f"断点续跑：跳过完整记录 {skip_count} 篇，"
            f"仅补补充材料 {si_only_count} 篇，完整重跑 {len(pending) - si_only_count} 篇")
    if args.keep_browser:
        log("⚠ --keep-browser 会让带有持久登录态的浏览器继续运行；"
            "请勿在共享电脑上使用，结束后请手工关闭。")

    # ---- 预检：出口 IP 是不是机构网段 -------------------------------------
    proc = tab = None
    if pending:
        if args.ip_check and not args.skip_ip_check:
            log("预检：探测当前 Python 进程的出口 IP（仅供参考，不作为机构权限结论）…")
            try:
                with OPENER.open("https://ipinfo.io/json", timeout=20) as r:
                    info = json.loads(r.read())
                log(f"  出口 IP {mask_ip(info.get('ip'))} | {info.get('org')} | {info.get('country')}")
                log("  注：Codex、代理或分流可能使该出口不同于浏览器；"
                    "最终以文章页实际授权结果为准。")
            except Exception as e:
                log(f"  (出口 IP 探测失败，不影响继续运行: {repr(e)[:80]})")

        profile = args.profile or cdp.default_profile_dir()
        log(f"启动浏览器（profile: {redact_local_path(profile)}）…")
        proc = cdp.launch(profile, port=args.port, browser=args.browser,
                          headless=args.headless)
        if not args.keep_browser:
            # 用户 Ctrl+C 时也要把浏览器关掉，否则会留一个窗口让人疑惑
            atexit.register(lambda: proc.terminate())
        tab = cdp.Tab(args.port, cdp.new_tab(args.port, "https://www.sciencedirect.com/"))
        # 预热：全新 profile 的首次访问要建 DNS/TLS/缓存并过 cookie 流程，
        # 先把它跑热，否则第一篇容易被误判成"没权限/没有 PDF 入口"。
        log("预热浏览器（首次访问建立会话）…")
        try:
            tab.goto("https://www.sciencedirect.com/", settle=8)
        except Exception as e:
            log(f"  (预热未完成，继续: {repr(e)[:80]})")
    else:
        log("断点续跑：所选篇目均已完整，无需启动浏览器。")

    stem_cache = {}

    def stem_for(row):
        """文件名主干 `{Num}_{DOI}`（无 Num 则 `{DOI}`）。

        按 DOI 缓存是必须的 —— `--force` 重跑和补跑轮次会重复调用，
        缓存能保证同一篇始终同名，不会生成 `xxx_2`、`xxx_3` 之类的重复文件。
        """
        d = row["doi"]
        if d not in stem_cache:
            stem_cache[d] = make_stem(row) or safe_doi(d)
        return stem_cache[d]

    def process_one(row, idx, total, pass_no):
        """处理一篇。返回结果记录。所有异常在此吞掉，绝不中断批次。"""
        doi = row["doi"]
        stem = stem_for(row)
        action = row.get("_resume_action", "full")
        previous = row.get("_previous") or {}
        rec = {"doi": doi, "num": row.get("num"), "title": row.get("title"),
               "stem": stem, "status": "failed",
               "pdf_status": "not_downloaded", "pdf": None,
               "si_requested": bool(args.si),
               "si_status": "pending" if args.si else "not_requested",
               "si": [], "access": None, "note": "", "pass": pass_no}
        if action == "si_only":
            # PDF 上一轮已下到，这一轮只为补 SI —— 但**仍要解析文章页**，
            # 否则拿不到 SI 链接（早期版本在这里直接返回，导致 SI 永远补不上）。
            rec["pdf_status"] = "downloaded"
            rec["pdf"] = previous.get("pdf")
            rec["bytes"] = previous.get("bytes")
            rec["si"] = list(previous.get("si") or [])
        tag = f"[{idx}/{total}]" if pass_no == 1 else f"[第{pass_no}轮 {idx}/{total}]"
        mode = "（仅补 SI）" if action == "si_only" else ""
        log(f"{tag} {doi}{mode}")
        try:
            # v1.5.2：不再因为 DOI 前缀“预判不支持”而跳过文章页。
            # 先解析真实落地页，再尝试专门规则 + citation_pdf_url + 页面 PDF 链接；
            # 只有这些都失败，才把没有专门适配的出版商归为 unsupported。
            pu_label, pu_why = known_unsupported(doi)
            skip_rest = False
            info = resolve_article(tab, doi, log, args.login_wait)
            rec["access"] = info.get("access")
            rec["pii"] = info.get("pii")
            rec["institution_detected"] = bool(info.get("brought_by"))

            if skip_rest:
                pass
            elif info.get("sso_start"):
                if action == "full":
                    rec["status"] = "needs_manual_login"
                rec["si_status"] = "failed" if args.si else "not_requested"
                rec["note"] = "no_institution_session"
                log("    → 被出版商弹到机构登录页（登录态失效）。"
                    "在浏览器窗口里重新登录一次，然后重跑本命令即可。")
            elif info.get("_needs_login"):
                if action == "full":
                    rec["status"] = "needs_manual_login"
                rec["si_status"] = "failed" if args.si else "not_requested"
                rec["note"] = "needs_manual_login"
                log("    → 停在机构登录页。请在浏览器窗口登录一次，"
                    "然后重跑本命令即可（profile 已保存登录态）。")
            elif info.get("notfound"):
                if action == "full":
                    rec["status"] = "bad_doi"
                rec["si_status"] = "failed" if args.si else "not_requested"
                rec["note"] = "bad_doi"
                log(f"    → DOI 打开后不是文章页（{info.get('title', '')[:60]}）——"
                    "可能 DOI 写错、已撤稿或不是该数据库收录。")
            else:
                # ---- 分支：ScienceDirect 走原路径；其它出版商走规则层 ----
                is_sd = is_sciencedirect_info(info)
                other_label, other_fn = publisher_rule(doi)

                if is_sd:
                    log(f"    授权={info.get('access')} 机构={'已识别' if info.get('brought_by') else '未识别'} "
                        f"SI候选={len(info.get('si') or [])}")
                else:
                    log(f"    ScienceDirect: 否 | 出版商规则: {other_label or '无'}"
                        f" | 落地: {redact_url(info.get('url'))[:70]}")

                generic_cands = [] if is_sd else generic_pdf_candidates(info)
                if not is_sd and not other_fn and not generic_cands:
                    # 没有专门规则，页面也没有给出任何可尝试的标准 PDF 候选。
                    if action == "full":
                        rec["status"] = "unsupported_publisher"
                        if pu_label:
                            rec["note"] = f"{pu_label}: {pu_why}"
                        else:
                            rec["note"] = f"落地于 {redact_url(info.get('url'))[:120]}；未发现 PDF 候选"
                    rec["si_status"] = "not_supported" if args.si else "not_requested"
                    if action == "si_only":
                        rec["status"] = "downloaded"
                    if pu_label:
                        log(f"    ⊘ {pu_label} 暂无专门适配，且页面未给出可用 PDF 候选 —— {pu_why}")
                    else:
                        log("    ⊘ 未发现专门规则或标准 PDF 候选。"
                            f"落地页：{redact_url(info.get('url'))[:80]}")
                elif is_sd and action == "full" and info.get("access") == "denied":
                    # `access` 的文本启发式只对 ScienceDirect 足够可靠。
                    # 其它出版商页面可能同时出现 “Get access” 与真正的 OA/PDF 按钮，
                    # 所以非 SD 一律先尝试真实 PDF 候选，再由返回内容判定。
                    rec["status"] = "no_entitlement"
                    rec["si_status"] = "failed" if args.si else "not_requested"
                    log("    → 当前 IP/会话没有这篇的全文权限。")
                elif is_sd and action == "full" and not info.get("pdf"):
                    if info.get("access") == "granted":
                        rec["status"] = "no_pdf_link"
                        rec["note"] = "ScienceDirect 页面已确认可访问，但未解析到 PDF 入口"
                        log("    → 页面已确认可访问，但没有解析到 PDF 入口；查看页面/日志后再处理。")
                    else:
                        rec["status"] = "challenge_not_resolved"
                        rec["note"] = "ScienceDirect 页面尚未渲染到可判断 PDF 入口"
                        log("    → ScienceDirect 页面尚未渲染完整，未把它误判为其它出版商。")
                    rec["si_status"] = "failed" if args.si else "not_requested"
                else:
                    if args.si:
                        # SI 链接来源：SD 用 -mmc；Wiley 系（含 ACSESS 子域）用
                        # /action/downloadSupplement?file=…（2026-09-19 实测定型）
                        if is_sd:
                            si_urls = list(info.get("si") or [])
                            si_supported = True
                        elif other_label == "wiley":
                            si_urls = list(info.get("si_wiley") or [])
                            si_supported = True
                        else:
                            # 这些出版商的正文下载已验证，但当前版本没有实现可靠的
                            # SI 枚举规则。不能把“没实现”误写成“该文没有 SI”。
                            si_urls = []
                            si_supported = False
                        if not si_supported:
                            rec["si"] = []
                            rec["si_status"] = "not_supported"
                            log("    SI － 当前版本尚未适配该出版商的补充材料枚举；不判定为 none_found。")
                        elif not si_urls:
                            rec["si"] = []
                            rec["si_status"] = "none_found"
                            log("    SI － 已检查支持的页面入口，未发现独立补充材料。")
                        else:
                            got_si, placeholders = download_si(tab, si_urls, sidir,
                                                               stem, log)
                            rec["si"] = got_si
                            if len(got_si) == len(si_urls):
                                rec["si_status"] = "downloaded"
                            elif got_si:
                                rec["si_status"] = "partial"
                            elif placeholders:
                                # 出版商只给了数据档案占位包 -> 实质没有补充材料，
                                # 这是「已检查、无 SI」的**终态**，不该进补跑轮
                                rec["si_status"] = "none_found"
                                log("    → 该文无实质补充材料（出版商只提供数据档案占位包），"
                                    "计为 none_found，不重试。")
                            else:
                                rec["si_status"] = "failed"

                    if action == "si_only":
                        rec["status"] = "downloaded"
                        log(f"    PDF ↷ 已验证现有文件，未重复下载：{rec['pdf']}")
                    else:
                        out = contained_path(pdfdir, stem + ".pdf")
                        if is_sd:
                            ok, msg = fetch_pdf(tab, info["pdf"], out, log=log)
                            if not ok:
                                rec["status"] = "challenge_not_resolved"
                        else:
                            # 非 SD：专门规则优先，再接页面标准 meta / PDF 链接兜底。
                            uniq = non_sd_pdf_candidates(info, doi, other_fn)
                            ok, msg = fetch_pdf_other(tab, uniq, out, log=log)
                            if not ok:
                                # PDF 入口被打回摘要页 = 机构没订，归 no_entitlement。
                                if (msg or "").startswith("no_entitlement:"):
                                    rec["status"] = "no_entitlement"
                                elif other_fn is None:
                                    # 通用探测也失败后，才把“尚无专门规则”的站点
                                    # 判为 unsupported；这正是 v1.5.2 的关键变化。
                                    rec["status"] = "unsupported_publisher"
                                    if pu_label:
                                        rec["note"] = f"{pu_label}: 通用 PDF 探测失败；{pu_why}; {msg}"
                                    else:
                                        rec["note"] = f"通用 PDF 探测失败: {msg}"
                                else:
                                    rec["status"] = "challenge_not_resolved"
                        if ok:
                            rec["status"] = "downloaded"
                            rec["pdf_status"] = "downloaded"
                            rec["pdf"] = os.path.basename(out)
                            rec["bytes"] = os.path.getsize(out)
                            log(f"    PDF ✓ {rec['bytes']:,}B  {rec['pdf']}")
                        else:
                            rec["note"] = msg
                            log(f"    PDF ✗ {msg}")

        except Exception as e:
            if action == "full":
                rec["status"] = "error"
            if args.si:
                rec["si_status"] = "failed"
            rec["note"] = repr(e)[:200]
            log(f"    错误 {repr(e)[:150]}")

        with open(manpath, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        return rec

    # ---- 多轮：第一轮跑完，自动把没下成的再补跑，最多 args.retry_failed 轮 ----
    max_passes = 1 + max(0, args.retry_failed)
    passes_run = 0
    processed = 0          # 本次运行累计处理的篇数（跨轮次），用于长休息
    consec_fail = 0        # 连续未成功篇数，用于失败退避
    for pass_no in range(1, max_passes + 1):
        if not pending:
            break
        passes_run = pass_no
        if pass_no > 1:
            wait = max(30.0, args.batch_rest) * random.uniform(0.8, 1.2)
            log("")
            log(f"===== 第 {pass_no} 轮：补跑上一轮未下载的 {len(pending)} 篇"
                f"（先静候 {wait:.0f}s，让会话/限速恢复）=====")
            interruptible_sleep(wait, log)
        still = []
        for i, row in enumerate(pending, 1):
            rec = process_one(row, i, len(pending), pass_no)
            results[row["doi"]] = rec
            # 进补跑轮的只有两类：
            #   a) PDF 属于"重试可能有用"的失败（站点验证未完成 / 异常）
            #   b) PDF 已下到、但补充材料没拿齐（下一轮走 si_only）
            # 不进的：bad_doi / unsupported_publisher / no_entitlement /
            # needs_manual_login —— 这些重试多少次都一样，白耗一轮长等待。
            pdf_failed_retryable = rec["status"] in PDF_RETRYABLE_STATUSES
            si_incomplete = (
                args.si
                and rec.get("pdf_status") == "downloaded"
                and rec.get("si_status") not in SI_COMPLETE
                and rec.get("si_status") != "not_supported")
            if pdf_failed_retryable or si_incomplete:
                retry_row = dict(row)
                retry_row["_previous"] = rec
                retry_row["_resume_action"] = ("full" if pdf_failed_retryable
                                               else "si_only")
                still.append(retry_row)

            # ---- 节奏控制：这是降低累计限流/验证风险的节奏控制 ----
            # consec_fail 只数「重试可能有用」的 PDF 失败；SI 缺失、机构
            # 无权限等都不代表站点验证失败，不该触发长退避。
            processed += 1
            if rec.get("pdf_status") == "downloaded":
                consec_fail = 0
            elif rec["status"] in PDF_RETRYABLE_STATUSES:
                consec_fail += 1

            if consec_fail >= 3:
                back = min(args.backoff * (2 ** (consec_fail - 3)), 900.0)
                log(f"    ⚠ 连续 {consec_fail} 篇 PDF 未成功 —— 疑似站点限流/验证未完成，"
                    f"退避 {back:.0f} 秒后再试…")
                interruptible_sleep(
                    back, log,
                    "    （不想等：按 Ctrl+C 退出后重跑同一条命令，"
                    "已下载的会自动跳过）")
            elif args.batch_size and processed % args.batch_size == 0:
                rest = args.batch_rest * random.uniform(0.8, 1.4)
                log(f"    已处理 {processed} 篇，长休息 {rest:.0f} 秒（降低累计限流风险）…")
                interruptible_sleep(rest, log)
            else:
                time.sleep(max(0.0, args.pace * random.uniform(0.7, 1.6)))
        if pass_no < max_passes and still:
            log(f"  第 {pass_no} 轮结束：仍有 {len(still)} 篇未下载，进入下一轮")
        pending = still

    if tab is not None:
        try:
            tab.close()
        except Exception:
            pass
    if proc is not None and not args.keep_browser:
        try:
            proc.terminate()
        except Exception:
            pass

    from collections import Counter
    all_recs = list(results.values())
    c = Counter(r["status"] for r in all_recs)
    # "完整" = PDF 已下到，且（没开 --si，或 SI 齐/SI 不适用）。
    # not_supported 表示该出版商本来就没有可抓的 SI，不能算未完成。
    def _complete(r):
        if r.get("pdf_status") != "downloaded":
            return False
        if not args.si:
            return True
        ss = r.get("si_status")
        return ss in SI_COMPLETE or ss == "not_supported"
    complete_count = sum(1 for r in all_recs if _complete(r))
    pdf_ok = [r for r in all_recs if r.get("pdf_status") == "downloaded"]
    si_only_missing = [r for r in pdf_ok
                       if args.si and r.get("si_status") not in SI_COMPLETE
                       and r.get("si_status") != "not_supported"]
    log("——  结果  ——")
    head = f"  完整 {complete_count} / {len(all_recs)}（PDF 已下到 {len(pdf_ok)} 篇"
    if si_only_missing:
        head += f"，其中 {len(si_only_missing)} 篇仅补充材料未齐"
    head += "）"
    log(head)
    for k, v in c.items():
        if k != "downloaded":
            log(f"  {k}: {v}")
    log(f"  SI 文件 {sum(len(r.get('si') or []) for r in all_recs)} 个")
    log(f"  正文目录 {redact_local_path(pdfdir)}")
    if args.si:
        log(f"  补充材料目录 {redact_local_path(sidir)}")
    log("")
    log("  跑完可以清一下浏览器缓存：  python 4.ClearCache.py")

    # 未完成篇目按「缺什么」分三组汇报，别把 PDF 已到手的也算成"未下载"。
    unfin = [r for r in all_recs if not record_complete(r, pdfdir, args.si, sidir)]
    bad = [r for r in unfin if r["status"] == "bad_doi"]
    unsup = [r for r in unfin if r["status"] == "unsupported_publisher"]
    noent = [r for r in unfin if r["status"] == "no_entitlement"]
    login = [r for r in unfin if r["status"] == "needs_manual_login"]
    pdf_missing = [r for r in unfin if r["status"] in PDF_RETRYABLE_STATUSES]
    no_pdf_link = [r for r in unfin if r["status"] == "no_pdf_link"]
    if unsup:
        log("")
        log(f"  以下 {len(unsup)} 篇所属出版商暂不支持（重试无用，别白等）：")
        seen_pub = {}
        for r in unsup:
            seen_pub.setdefault((r.get("note") or "").split(":")[0].strip(), []).append(r["doi"])
        for pub, dois in seen_pub.items():
            log(f"    - {pub or '(未知)'}: {len(dois)} 篇")
        log("    这类只能走馆际互借 / 文献互助，或换支持该出版商的方式。")
    if noent:
        log("")
        log(f"  以下 {len(noent)} 篇当前会话没有全文权限（可能与订阅、登录或网络路径有关）：")
        for r in noent[:10]:
            log(f"    - {r['doi']}")
    if si_only_missing:
        log("")
        log(f"  以下 {len(si_only_missing)} 篇 PDF 已下到、仅补充材料未齐 ——"
            "重跑同一条命令即可，会自动只补 SI（已下好的 PDF 不会重下）。")
        log("    python 3.SCIDownload.py <清单> --out <同目录> --si --only \""
            + ",".join(r["doi"] for r in si_only_missing[:5])
            + ("..." if len(si_only_missing) > 5 else "") + "\"")
    if pdf_missing:
        log("")
        log(f"  以下 {len(pdf_missing)} 篇 PDF 未下到，可单独重试（重跑同一条命令"
            "也会自动补这些，不需要 --force）：")
        log("    python 3.SCIDownload.py <清单> --out <同目录> --only \""
            + ",".join(r["doi"] for r in pdf_missing[:5])
            + ("..." if len(pdf_missing) > 5 else "") + "\"")
    if no_pdf_link:
        log("")
        log(f"  以下 {len(no_pdf_link)} 篇页面已打开但未解析到 PDF 入口；"
            "这类不会自动反复重试，请先人工查看文章页：")
        for r in no_pdf_link[:10]:
            log(f"    - {r['doi']}")
    if login:
        log("")
        log(f"  以下 {len(login)} 篇停在机构登录页 —— 在浏览器里登录一次后重跑即可。")
        for r in login[:10]:
            log(f"    - {r['doi']}")
    if bad:
        log("")
        log(f"  以下 {len(bad)} 篇 DOI 打不开文章页，重试无用，请核对 DOI 本身：")
        for r in bad[:10]:
            log(f"    - {r['doi']}")
    return 0 if len(pdf_ok) == len(all_recs) else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已中断（Ctrl+C）。浏览器已关闭。")
        print("下次重跑同一条命令即可 —— 已下载的会自动跳过。")
        raise SystemExit(130)

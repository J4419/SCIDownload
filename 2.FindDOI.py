#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""FindDOI —— 给「只有标题、没有 DOI」的清单补上 DOI。

    python 2.FindDOI.py 清单.xlsx
    python 2.FindDOI.py titles.txt --out 带DOI.csv --mailto you@example.com

数据源：OpenAlex + Crossref（都是公开 API，**不要 key、不要登录**）。

网络要求：**公网即可**。FindDOI 默认遵循系统代理；如需直连可加 `--direct`。
（SCIDownload 的机构授权路径与 FindDOI 的公开元数据查询是两回事。）

安全设计：**只给标题高度一致的行自动填 DOI**，其余一律留空并标注状态，
因为标题检索必然返回一堆「相关但不相同」的论文，填错就是下错文献。
输出文件带 DOI 列，**可以直接喂给 SCIDownload.py**。

明确不支持：中文文献（CNKI / 万方 不在 OpenAlex / Crossref 索引内）。
"""

import argparse
import csv
import difflib
import json
import math
import os
import re
import sys
import tempfile
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

VERSION = "1.3"
CACHE_SCHEMA = 2


def build_opener(direct=False):
    """FindDOI 默认遵循系统代理；只有显式 --direct 才强制直连。"""
    return urllib.request.build_opener(
        urllib.request.ProxyHandler({} if direct else None)
    )


OPENER = build_opener()

TITLE_HEADERS = ("title", "title1", "题名", "标题", "文献题名", "篇名", "题目", "论文题目")
DOI_HEADERS = ("doi", "doi号", "文献doi", "文献 doi")
NUM_HEADERS = ("num", "序号", "编号", "id", "no", "no.")

# 这些词在标题比对里没有区分度，去掉能提高相似度的准确率
STOP = {"a", "an", "the", "of", "on", "in", "for", "and", "to", "with", "by",
        "from", "at", "as", "is", "are", "its", "their", "under", "between",
        "using", "based", "into", "via", "during", "after", "before"}

DOI_RE = re.compile(r"10\.\d{4,9}/\S+")


# ------------------------------------------------------------------ 读输入
def read_table(path):
    """xlsx / csv / tsv / txt -> list[list[str]]"""
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        try:
            from SCIDownload import read_xlsx_stdlib
            return read_xlsx_stdlib(path)
        except Exception:
            try:
                import openpyxl
            except ImportError:
                raise SystemExit(
                    "读 .xlsx 需要同目录下的 SCIDownload.py（内置零依赖解析），"
                    "或安装 openpyxl。也可以把清单另存为 .csv 再试。")
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            return [[("" if c is None else str(c)) for c in row]
                    for row in ws.iter_rows(values_only=True)]

    lines = open(path, encoding="utf-8-sig", errors="replace").read().splitlines()
    first = lines[0] if lines else ""
    if first.count(",") >= 1 or "\t" in first:
        delim = "\t" if "\t" in first else ","
        return [[c.strip() for c in row] for row in csv.reader(lines, delimiter=delim)]

    # 纯文本：每行一个标题（可能自带 DOI）
    return [[ln.strip()] for ln in lines if ln.strip() and not ln.startswith("#")]


def find_header(rows, max_scan=15):
    """在前 max_scan 行里找表头（含 Title / DOI / Num 任一）。找不到返回 -1。"""
    for i, r in enumerate(rows[:max_scan]):
        cells = [str(c).strip().lower() for c in (r or [])]
        if any(c in TITLE_HEADERS or c in DOI_HEADERS for c in cells if c):
            return i
    return -1


def pick_title_col(rows, start):
    """没有 Title 表头时，选「文字最长」的那一列当标题列。"""
    if not rows:
        return 0
    width = max(len(r) for r in rows[start:start + 40]) if rows[start:] else 1
    best, best_len = 0, -1.0
    for c in range(width):
        vals = [str(r[c]) for r in rows[start:start + 40] if c < len(r) and r[c]]
        if not vals:
            continue
        avg = sum(len(v) for v in vals) / len(vals)
        if avg > best_len:
            best, best_len = c, avg
    return best


def load_items(path):
    """-> [{'num':..., 'title':..., 'doi':...}]"""
    rows = read_table(path)
    if not rows:
        return []
    hi = find_header(rows)
    if hi >= 0:
        head = [str(c).strip().lower() for c in rows[hi]]
        i_t = next((i for i, h in enumerate(head) if h in TITLE_HEADERS), None)
        i_d = next((i for i, h in enumerate(head) if h in DOI_HEADERS), None)
        i_n = next((i for i, h in enumerate(head) if h in NUM_HEADERS), None)
        body = rows[hi + 1:]
        if i_t is None:
            i_t = pick_title_col(rows, hi + 1)
    else:
        i_t, i_d, i_n = 0, None, None
        body = rows

    items = []
    for k, r in enumerate(body, 1):
        r = r or []

        def cell(i):
            return str(r[i]).strip() if (i is not None and i < len(r)) else ""

        title = cell(i_t)
        if not title or len(title) < 8:          # 太短的不像标题
            continue
        doi = ""
        if i_d is not None:
            m = DOI_RE.search(cell(i_d))
            doi = m.group(0).rstrip(".,;") if m else ""
        else:
            m = DOI_RE.search(title)             # 标题里自带 DOI 的情况
            if m:
                doi = m.group(0).rstrip(".,;")
        num = cell(i_n) if i_n is not None else str(k)
        items.append({"num": num or str(k), "title": title, "doi": doi.lower()})
    return items


# ------------------------------------------------------------------ 相似度
def norm(s):
    s = re.sub(r"<[^>]+>", " ", str(s or "")).lower()
    s = s.replace("\u2010", "-").replace("\u2013", "-").replace("\u2014", "-")
    s = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def toks(s):
    return {w for w in norm(s).split() if w and w not in STOP}


def title_score(a, b):
    """标题相似度（不是统计学“置信度”）：字符相似度与词集合 Jaccard 取较高者。"""
    if not norm(a) or not norm(b):
        return 0.0
    char = difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()
    ta, tb = toks(a), toks(b)
    jac = len(ta & tb) / len(ta | tb) if (ta | tb) else 0.0
    return max(char, jac)


def agreement_source_count(candidates, doi):
    """同一 DOI 获得独立数据源支持的数量，单一来源重复结果只算一次。"""
    return len({source for _, candidate, source, _ in candidates if candidate == doi})


def title_diff(sent, hit, limit=90):
    """把「输入标题」与「命中标题」的差异标出来，便于人工判断是否同一篇。

    只做词级差异：命中标题里**多出**的词用 [] 括起来。
    这样一眼就能看出「得分为什么不是 1.0」。
    """
    if not sent or not hit:
        return (hit or "")[:limit]
    sa, sb = norm(sent).split(), norm(hit).split()
    if sa == sb:
        return (hit or "")[:limit]
    sm = difflib.SequenceMatcher(None, sa, sb, autojunk=False)
    parts = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        chunk = " ".join(sb[j1:j2])
        if not chunk:
            continue
        parts.append(chunk if tag == "equal" else f"[{chunk}]")
    out = re.sub(r"\s+", " ", " ".join(parts)).strip()
    return out[:limit] + ("…" if len(out) > limit else "")


def validate_options(sources, threshold, review_threshold, sleep, limit):
    known = {"openalex", "crossref"}
    unknown = set(sources) - known
    if not sources:
        raise ValueError("--sources 至少要包含 openalex 或 crossref")
    if unknown:
        raise ValueError("未知数据源: " + ", ".join(sorted(unknown)))
    if not (math.isfinite(review_threshold) and math.isfinite(threshold)
            and 0 <= review_threshold <= threshold <= 1):
        raise ValueError("门槛必须满足 0 <= --review-threshold <= --threshold <= 1")
    if not math.isfinite(sleep) or sleep < 0:
        raise ValueError("--sleep 必须是非负有限数")
    if limit < 0:
        raise ValueError("--limit 必须是非负整数")


def load_cache(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError, TypeError):
        return {}
    if not isinstance(data, dict):
        return {}
    if data.get("schema") == CACHE_SCHEMA:
        entries = data.get("entries")
        return entries if isinstance(entries, dict) else {}
    # v1.3 起评分策略收紧。旧缓存里可能含 v1.2 的“单源 0.85 抬到 0.90”
    # 结果，不能继续复用，否则升级后仍可能自动填错 DOI。
    return {}


def save_cache(path, entries):
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    fd, part = tempfile.mkstemp(prefix=".finddoi-", suffix=".part", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump({"schema": CACHE_SCHEMA, "entries": entries}, f,
                      ensure_ascii=False, indent=1)
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(part, path)
    except Exception:
        try:
            os.remove(part)
        except OSError:
            pass
        raise


# ------------------------------------------------------------------ 检索
def http_json(url, timeout=25, retries=3):
    """请求 JSON。429（限流）单独处理。

    超时/网络抖动值得重试；但 **429 连续出现说明整个出口 IP 被限流**
    （实测：带 mailto 也无效），重试只是白等 —— 第一次 429 就立刻放弃，
    让上层用剩下的数据源出结果。
    """
    for k in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": f"FindDOI/{VERSION}"})
            with OPENER.open(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8", "replace")), None
        except Exception as e:
            code = getattr(e, "code", None) or getattr(e, "status", None)
            if code == 429 or k == retries:
                return None, f"{type(e).__name__}: {str(e)[:90]}"
            time.sleep(2 * (k + 1))


def query_openalex(title, n=5, mailto=""):
    url = ("https://api.openalex.org/works?search=" + urllib.parse.quote(title) +
           f"&per-page={n}" + (f"&mailto={mailto}" if mailto else ""))
    d, err = http_json(url)
    if err:
        return [], err
    out = []
    for w in d.get("results", []):
        doi = (w.get("doi") or "").replace("https://doi.org/", "").strip().lower()
        if doi:
            out.append({"doi": doi, "title": w.get("title") or "",
                        "year": w.get("publication_year")})
    return out, None


def query_crossref(title, n=5, mailto=""):
    url = ("https://api.crossref.org/works?query.bibliographic=" +
           urllib.parse.quote(title) + f"&rows={n}&select=DOI,title,issued" +
           (f"&mailto={mailto}" if mailto else ""))
    d, err = http_json(url)
    if err:
        return [], err
    out = []
    for it in d.get("message", {}).get("items", []):
        doi = (it.get("DOI") or "").strip().lower()
        t = it.get("title") or [""]
        if doi:
            out.append({"doi": doi, "title": t[0] if t else "", "year": None})
    return out, None


def resolve(title, mailto="", sources=("openalex", "crossref"), sleep=1.0):
    """返回 (doi, 标题相似度, 来源, 命中标题, 报错)。"""
    fns = {"openalex": query_openalex, "crossref": query_crossref}
    cands, errs = [], []
    failed = []
    for src in sources:
        rows, err = fns[src](title, mailto=mailto)
        if err:
            errs.append(f"{src}:{err}")
            failed.append(src)
        for r in rows:
            cands.append((title_score(title, r["title"]), r["doi"], src, r["title"]))
        time.sleep(max(0.0, sleep))
    # 有数据源失败时补试一次 —— 只对非 429 的失败有意义。
    # 429 意味着整个出口 IP 被限流，补试纯属白等（实测会多花几分钟）。
    retry_srcs = [s for s in failed
                  if not any(e.startswith(s + ":") and "429" in e for e in errs)]
    if retry_srcs and cands:
        for src in retry_srcs:
            time.sleep(3.0)
            rows, err = fns[src](title, mailto=mailto)
            if not err:
                errs = [e for e in errs if not e.startswith(src + ":")]
                for r in rows:
                    cands.append((title_score(title, r["title"]), r["doi"], src, r["title"]))
    if not cands:
        return None, 0.0, "-", "", "; ".join(errs)
    cands.sort(key=lambda x: -x[0])
    best = cands[0]
    agree = agreement_source_count(cands, best[1])
    score = best[0]
    # 公开版保持保守：不再因为“只有一个数据源可用”而把 0.85 强行抬到 0.90，
    # 也不把多源一致直接加到标题分数里。相似标题（只差作物、地区、年份等）
    # 很容易达到 0.90 左右；这些应进入人工确认，而不是自动写 DOI。
    srcs = sorted({src for _, doi, src, _ in cands if doi == best[1]})
    source_label = "+".join(srcs) if srcs else best[2]
    return best[1], min(score, 1.0), source_label, best[3], \
        ("; ".join(errs) if errs else "")


def classify_match(doi: str, score: float, threshold: float, review_threshold: float):
    """把标题匹配结果分为自动采用 / 待确认 / 未找到。"""
    if doi and score >= threshold:
        return "已确认", doi
    if doi and score >= review_threshold:
        return "待确认", ""
    return "未找到", ""


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser(
        description="给只有标题的文献清单补 DOI（OpenAlex + Crossref）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例:\n"
               "  python 2.FindDOI.py 我的清单.xlsx\n"
               "  python 2.FindDOI.py titles.txt --out 带DOI.csv --mailto you@example.com\n")
    ap.add_argument("inputs", nargs="?",
                    help="清单：.xlsx / .csv / .tsv（含 Title 列）或 .txt（每行一个标题）")
    ap.add_argument("--out", default=None, help="输出文件（默认 <输入名>_DOI.csv）")
    ap.add_argument("--mailto", default=os.environ.get("CROSSREF_MAILTO", ""),
                    help="可选联系邮箱；会随请求发送给 OpenAlex/Crossref，不想提供可省略")
    ap.add_argument("--threshold", type=float, default=0.97,
                    help="自动采用的标题相似度门槛，默认 0.97；相邻论文通常应进入人工确认")
    ap.add_argument("--review-threshold", type=float, default=0.80,
                    help="低于此值判为「未找到」，默认 0.80")
    ap.add_argument("--sources", default="openalex,crossref", help="数据源，逗号分隔")
    ap.add_argument("--sleep", type=float, default=1.0, help="每次请求之间的间隔秒数")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 条（试跑用）")
    ap.add_argument("--force", action="store_true", help="忽略本地缓存，全部重查")
    ap.add_argument("--direct", action="store_true",
                    help="忽略系统代理，强制直连 OpenAlex / Crossref")
    args = ap.parse_args()

    if not args.inputs:
        print(__doc__)
        print("下一步：把你的清单存成 xlsx / csv / txt（要有题名），然后：")
        print("  python 2.FindDOI.py 我的清单.xlsx")
        print("  python 2.FindDOI.py 我的清单.xlsx --out 带DOI.csv --mailto 你的邮箱")
        return 2
    if not os.path.exists(args.inputs):
        raise SystemExit(f"找不到清单文件：{args.inputs}\n  当前目录：{os.getcwd()}")

    sources = tuple(s.strip().lower() for s in args.sources.split(",") if s.strip())
    try:
        validate_options(sources, args.threshold, args.review_threshold,
                         args.sleep, args.limit)
    except ValueError as e:
        ap.error(str(e))
    global OPENER
    OPENER = build_opener(args.direct)

    items = load_items(args.inputs)
    if args.limit:
        items = items[: args.limit]
    if not items:
        raise SystemExit("清单里没有解析到任何标题。请确认有一列题名（Title / 题名 / 标题）。")

    out_path = args.out or os.path.splitext(args.inputs)[0] + "_DOI.csv"
    cache_path = os.path.splitext(out_path)[0] + ".finddoi-cache.json"
    cache = {} if args.force else load_cache(cache_path)

    with_doi = sum(1 for i in items if i["doi"])
    print(f"FindDOI v{VERSION}")
    print(f"  清单：{len(items)} 条标题"
          + (f"（其中 {with_doi} 条自带 DOI，直接沿用）" if with_doi else ""))
    print(f"  标题相似度门槛：≥{args.threshold} 自动采用 ｜ {args.review_threshold}~{args.threshold} 待人工确认"
          f" ｜ <{args.review_threshold} 判未找到")
    print(f"  数据源：{args.sources}   输出：{out_path}")
    print()

    rows_out = []
    stat = {"已有": 0, "已确认": 0, "待确认": 0, "未找到": 0, "查询失败": 0}
    need_review = []

    for n, it in enumerate(items, 1):
        if it["doi"]:
            rows_out.append([it["num"], it["title"], it["doi"], "1.000",
                             "已有", "-", it["title"]])
            stat["已有"] += 1
            print(f"[{n}/{len(items)}] 已有 DOI：{it['doi']}")
            continue

        key = norm(it["title"])
        if key in cache:
            got = cache[key]
        else:
            doi, conf, src, hit, err = resolve(
                it["title"], args.mailto,
                sources,
                args.sleep)
            got = {"doi": doi, "conf": conf, "src": src, "hit": hit, "err": err}
            cache[key] = got
            save_cache(cache_path, cache)

        conf = got.get("conf") or 0.0
        doi = got.get("doi") or ""
        status, shown = classify_match(doi, conf, args.threshold, args.review_threshold)
        stat[status] += 1
        if status == "待确认":
            # DOI 候选只供人工核对，正式 DOI 列保持空白，避免误下载。
            need_review.append((it["num"], it["title"], doi, conf, got.get("hit", "")))
        elif status == "未找到" and got.get("err"):
            stat["查询失败"] += 1

        rows_out.append([it["num"], it["title"], shown, f"{conf:.3f}",
                         status, got.get("src", "-"), got.get("hit", "")])
        flag = {"已确认": "✔", "待确认": "?", "未找到": "✘"}[status]
        print(f"[{n}/{len(items)}] {flag} {status:<4} 标题相似度 {conf:.3f}  "
              f"{shown or '(留空)'}")
        if status != "已确认":
            print(f"        输入：{it['title'][:80]}")
            if got.get("hit"):
                print(f"        命中：{got['hit'][:80]}   来源 {got.get('src')}")
            if got.get("err"):
                print(f"        错误：{got['err'][:100]}")

    # 写 CSV（utf-8-sig，Excel 双击不乱码）
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Num", "Title", "DOI", "标题相似度", "状态", "来源", "匹配到的标题"])
        w.writerows(rows_out)

    print()
    print("——  结果  ——")
    for k, v in stat.items():
        if v:
            print(f"  {k}: {v}")
    print(f"  输出：{out_path}")
    print(f"  （只有「已确认」的行填了 DOI；其余留空，SCIDownload 会自动跳过）")

    if need_review:
        print()
        print(f"  ⚠ 有 {len(need_review)} 条「待确认」—— 标题相似度未达到自动采用门槛，我没有替你决定：")
        for num, title, doi, conf, hit in need_review[:10]:
            print(f"    #{num} (标题相似度 {conf:.3f}) 可能是 {doi}")
            print(f"        输入：{title}")
            # 标出命中标题比输入多出来的词 —— 一眼就能看出差异在哪、值不值得信
            print(f"        命中：{title_diff(title, hit, limit=200)}")
        if len(need_review) > 10:
            print(f"    …另有 {len(need_review) - 10} 条，见输出文件")
        print("    确认无误后，把 DOI 手工粘进输出文件对应行即可。")
    print()
    print(f"下一步：把 {os.path.basename(out_path)} 喂给 SCIDownload：")
    print(f"  python 3.SCIDownload.py \"{out_path}\" --out ./papers --si")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已中断（Ctrl+C）。缓存已保存，重跑会接着用。")
        raise SystemExit(130)

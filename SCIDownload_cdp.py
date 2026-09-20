# -*- coding: utf-8 -*-
"""SCIDownload_cdp —— 跨平台浏览器发现 + Chrome DevTools Protocol 客户端。

纯标准库实现（含自写的 WebSocket 客户端），不依赖 Node、requests、websocket-client。

对外接口：
    find_browser()                   -> 浏览器可执行文件路径
    launch(profile_dir, port)        -> (Popen, port)
    new_tab(port, url)               -> tab_id
    Tab(port, tab_id)                -> .goto() .js() .front() .close()
"""
from __future__ import annotations

import base64
import json
import os
import platform
import secrets
import shutil
import socket
import struct
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request

_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
# 注意：必须绕开环境里的 HTTP_PROXY，否则 127.0.0.1 上的调试端口会被代理劫持成 502。

_SYSTEM = platform.system()


def find_browser(explicit: str | None = None) -> str:
    """找一个 Chromium 系浏览器（Chrome / Edge / Chromium）。"""
    if explicit:
        if os.path.exists(explicit):
            return explicit
        raise FileNotFoundError(f"--browser 指定的路径不存在: {explicit}")

    cands: list[str] = []
    if _SYSTEM == "Windows":
        local = os.environ.get("LOCALAPPDATA", "")
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        cands += [
            os.path.join(pf, r"Google\Chrome\Application\chrome.exe"),
            os.path.join(pf86, r"Google\Chrome\Application\chrome.exe"),
            os.path.join(local, r"Google\Chrome\Application\chrome.exe"),
            os.path.join(pf86, r"Microsoft\Edge\Application\msedge.exe"),
            os.path.join(pf, r"Microsoft\Edge\Application\msedge.exe"),
            os.path.join(local, r"Microsoft\Edge\Application\msedge.exe"),
        ]
    elif _SYSTEM == "Darwin":
        cands += [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ]
    else:
        for exe in ("google-chrome", "google-chrome-stable", "microsoft-edge",
                    "microsoft-edge-stable", "chromium", "chromium-browser"):
            p = shutil.which(exe)
            if p:
                cands.append(p)

    for p in cands:
        if p and os.path.exists(p):
            return p
    raise FileNotFoundError(
        "没找到 Chrome / Edge / Chromium。请安装其中之一，或用 --browser <路径> 指定。")


def default_profile_dir() -> str:
    """浏览器 profile 的默认位置（持久化保存登录态）。

    目录名跟程序名保持一致。它是浏览器 profile 的落点，
    登录态、Cookie、缓存都在里面 —— **改名等于让所有用户重新登录一次**，
    所以以后不要再动这个字符串。
    """
    if _SYSTEM == "Windows":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    elif _SYSTEM == "Darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    return os.path.join(base, "SCIDownload", "browser-profile")


def _http(port: int, path: str, method: str = "GET", timeout: float = 20.0):
    req = urllib.request.Request(f"http://127.0.0.1:{port}{path}", method=method)
    with _OPENER.open(req, timeout=timeout) as r:
        body = r.read().decode("utf-8", "replace")
    return json.loads(body) if body.strip().startswith(("{", "[")) else body


def ensure_debug_port_available(port: int):
    """启动前确认回环调试端口空闲，避免误连到别的浏览器实例。"""
    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        if _SYSTEM == "Windows" and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        probe.bind(("127.0.0.1", port))
    except OSError as e:
        raise RuntimeError(
            f"本机调试端口 {port} 已被占用。请关闭旧的 SCIDownload 浏览器，"
            f"或改用 --port。原始错误: {e}"
        ) from e
    finally:
        probe.close()


def launch(profile_dir: str, port: int = 9222, browser: str | None = None,
           url: str = "about:blank", headless: bool = False,
           wait_s: float = 40.0) -> subprocess.Popen:
    """启动浏览器并开启远程调试。

    关键：Chromium 136+ 只有在显式传入非默认 --user-data-dir 时才允许
    --remote-debugging-port，所以必须给一个独立 profile 目录。
    该 profile 是持久化的 —— 需要登录的机构只需登录一次，之后自动复用。
    """
    exe = find_browser(browser)
    ensure_debug_port_available(port)
    os.makedirs(profile_dir, exist_ok=True)
    args = [
        exe,
        f"--remote-debugging-port={port}",
        "--remote-debugging-address=127.0.0.1",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-features=msEdgeSidebarV2,msImplicitSignin",
    ]
    if headless:
        args.append("--headless=new")
    args.append(url)

    # 让浏览器**真正独立于本进程存活**。
    # 只用 DETACHED_PROCESS 是不够的：进程退出时 Chrome 会一并被带走
    # （父进程的 job object 回收），于是「开窗口 → 手工登录」这条路直接断掉。
    # CREATE_BREAKAWAY_FROM_JOB + DETACHED_PROCESS 才能脱离 job 独立运行。
    flags = 0
    if _SYSTEM == "Windows":
        DETACHED_PROCESS = 0x00000008
        CREATE_NEW_PROCESS_GROUP = 0x00000200
        CREATE_BREAKAWAY_FROM_JOB = 0x01000000
        flags = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB
    try:
        proc = subprocess.Popen(args, creationflags=flags,
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                start_new_session=(_SYSTEM != "Windows"))
    except OSError:
        # 某些环境不允许脱离 job，退回普通分离（至少在当前进程存活期间可用）
        flags = (0x00000008 | 0x00000200) if _SYSTEM == "Windows" else 0
        proc = subprocess.Popen(args, creationflags=flags,
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                start_new_session=(_SYSTEM != "Windows"))
    deadline = time.time() + wait_s
    while time.time() < deadline:
        try:
            _http(port, "/json/version", timeout=2)
            return proc
        except Exception:
            time.sleep(0.5)
    raise RuntimeError(
        f"浏览器已启动但 {wait_s:.0f}s 内没有响应调试端口 {port}。"
        f"检查是否已有同 profile 的实例在运行，或换一个 --port。")


# --------------------------------------------------------------------------
# 极简 WebSocket 客户端（RFC 6455，仅实现 CDP 需要的部分）
# --------------------------------------------------------------------------
class _WS:
    def __init__(self, url: str, timeout: float = 60.0):
        u = urllib.parse.urlparse(url)
        self.host = u.hostname or "127.0.0.1"
        self.port = u.port or 80
        self.path = u.path + (f"?{u.query}" if u.query else "")
        self.sock = socket.create_connection((self.host, self.port), timeout=timeout)
        self.sock.settimeout(timeout)
        self._buf = b""
        self._handshake()

    def _handshake(self):
        key = base64.b64encode(secrets.token_bytes(16)).decode()
        req = (f"GET {self.path} HTTP/1.1\r\n"
               f"Host: {self.host}:{self.port}\r\n"
               "Upgrade: websocket\r\n"
               "Connection: Upgrade\r\n"
               f"Sec-WebSocket-Key: {key}\r\n"
               "Sec-WebSocket-Version: 13\r\n\r\n")
        self.sock.sendall(req.encode())
        while b"\r\n\r\n" not in self._buf:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise ConnectionError("WebSocket 握手时连接被关闭")
            self._buf += chunk
        head, _, rest = self._buf.partition(b"\r\n\r\n")
        self._buf = rest
        if b" 101 " not in head.split(b"\r\n")[0]:
            raise ConnectionError(f"WebSocket 握手失败: {head[:160]!r}")

    def _recv_exact(self, n: int) -> bytes:
        while len(self._buf) < n:
            chunk = self.sock.recv(65536)
            if not chunk:
                raise ConnectionError("WebSocket 连接已关闭")
            self._buf += chunk
        out, self._buf = self._buf[:n], self._buf[n:]
        return out

    def send(self, text: str):
        payload = text.encode("utf-8")
        n = len(payload)
        header = bytearray([0x81])                       # FIN + text
        if n < 126:
            header.append(0x80 | n)
        elif n < (1 << 16):
            header.append(0x80 | 126)
            header += struct.pack(">H", n)
        else:
            header.append(0x80 | 127)
            header += struct.pack(">Q", n)
        mask = secrets.token_bytes(4)
        header += mask
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        self.sock.sendall(bytes(header) + masked)

    def recv(self) -> str:
        while True:
            b0, b1 = self._recv_exact(2)
            opcode = b0 & 0x0F
            masked = b1 & 0x80
            length = b1 & 0x7F
            if length == 126:
                length = struct.unpack(">H", self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack(">Q", self._recv_exact(8))[0]
            mask = self._recv_exact(4) if masked else None
            data = self._recv_exact(length) if length else b""
            if mask:
                data = bytes(b ^ mask[i % 4] for i, b in enumerate(data))
            if opcode == 0x8:
                raise ConnectionError("WebSocket 已被对端关闭")
            if opcode == 0x9:                            # ping -> pong
                self.sock.sendall(b"\x8a\x80" + secrets.token_bytes(4))
                continue
            if opcode in (0x1, 0x2, 0x0):
                return data.decode("utf-8", "replace")

    def close(self):
        try:
            self.sock.close()
        except Exception:
            pass


class Tab:
    """一个页面标签。所有方法都做失败重试，CDP 偶发超时不该中断整批下载。"""

    def __init__(self, port: int, tab_id: str, timeout: float = 60.0):
        self.port = port
        self.id = tab_id
        self.timeout = timeout
        self._seq = 0

    def _ws_url(self) -> str:
        for t in _http(self.port, "/json/list"):
            if t.get("id") == self.id:
                return t["webSocketDebuggerUrl"]
        raise RuntimeError(f"标签页 {self.id} 已不存在")

    def _call(self, method: str, params: dict | None = None, timeout: float | None = None):
        ws = _WS(self._ws_url(), timeout or self.timeout)
        try:
            self._seq += 1
            ws.send(json.dumps({"id": self._seq, "method": method, "params": params or {}}))
            while True:
                msg = json.loads(ws.recv())
                if msg.get("id") == self._seq:
                    if "error" in msg:
                        raise RuntimeError(msg["error"].get("message", "CDP 报错"))
                    return msg.get("result", {})
        finally:
            ws.close()

    def front(self):
        """把标签提到前台。**这一步不是可选的**：后台标签的定时器会被浏览器节流，
        导致 ScienceDirect 的反爬挑战页永远无法自行解除。"""
        try:
            self._call("Page.bringToFront", timeout=15)
        except Exception:
            pass

    def js(self, expression: str, timeout: float | None = None, await_promise: bool = True):
        r = self._call("Runtime.evaluate", {
            "expression": expression,
            "returnByValue": True,
            "awaitPromise": await_promise,
            "userGesture": True,
        }, timeout=timeout)
        if r.get("exceptionDetails"):
            raise RuntimeError(str(r["exceptionDetails"].get("exception", {}).get("description"))[:200])
        return r.get("result", {}).get("value")

    def goto(self, url: str, settle: float = 6.0) -> str:
        self.front()
        self._call("Page.navigate", {"url": url})
        deadline = time.time() + max(settle, 1.0) * 3
        while time.time() < deadline:
            time.sleep(0.5)
            try:
                if self.js("document.readyState", timeout=10) == "complete":
                    break
            except Exception:
                pass
        time.sleep(settle)
        return url

    def close(self):
        try:
            _http(self.port, f"/json/close/{self.id}", timeout=10)
        except Exception:
            pass


def new_tab(port: int, url: str = "about:blank", timeout: float = 30.0) -> str:
    enc = urllib.parse.quote(url, safe="")
    for method in ("PUT", "GET"):
        try:
            t = _http(port, f"/json/new?{enc}", method=method, timeout=timeout)
            if isinstance(t, dict) and t.get("id"):
                return t["id"]
        except urllib.error.HTTPError:
            continue
    raise RuntimeError("无法新建标签页（调试端口未就绪？）")


def fetch_bytes_via_page(tab: Tab, url: str, out_path: str,
                         total_timeout: float = 300.0, chunk: int = 300_000,
                         expect_pdf: bool = True):
    """在**页面上下文里** fetch 一个二进制 URL，分块 base64 回传后落盘。

    为什么必须在页面里发请求：像 pdf.sciencedirectassets.com 那种带签名
    (X-Amz-Security-Token) 的地址，从 Python 直连会拿到 403 反爬 HTML。
    字节只在本地流转，不进入任何模型上下文。
    """
    start = """
    (() => {
      window.__scid = {status: 'pending', len: 0};
      fetch(%s, {credentials: 'include'})
        .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
        .then(b => {
          const u8 = new Uint8Array(b);
          let s = '';
          const CH = 0x8000;
          for (let i = 0; i < u8.length; i += CH) s += String.fromCharCode.apply(null, u8.subarray(i, i + CH));
          window.__scid = {status: 'ok', len: u8.length, b64: btoa(s)};
        })
        .catch(e => { window.__scid = {status: 'err', msg: String(e && e.message || e)}; });
      return 'started';
    })()
    """ % json.dumps(url)
    tab.js(start, timeout=60)

    state, t0 = {}, time.time()
    while time.time() - t0 < total_timeout:
        try:
            raw = tab.js("JSON.stringify(window.__scid ? "
                         "{status:window.__scid.status,len:window.__scid.len,msg:window.__scid.msg} : null)")
            state = json.loads(raw) if raw else {}
        except Exception:
            state = {}
        if state.get("status") in ("ok", "err"):
            break
        time.sleep(1.5)

    if state.get("status") != "ok":
        return False, f"fetch {state.get('status', 'timeout')}: {state.get('msg', '')}"

    total = int(state.get("len") or 0)
    data = bytearray()
    pos = 0
    limit = total * 4 // 3 + 16
    while pos < limit:
        part = tab.js(f"window.__scid.b64.slice({pos},{pos + chunk})", timeout=90)
        if not part:
            break
        data += base64.b64decode(part)
        pos += chunk
    try:
        tab.js("try{delete window.__scid;}catch(e){}", timeout=15)
    except Exception:
        pass

    payload = bytes(data)
    if len(payload) != total:
        return False, f"length mismatch: expected {total}, got {len(payload)}"
    if len(payload) < 1024:
        return False, f"file too small: {len(payload)}B"
    head = payload[:512].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html", b"<head", b"<body")):
        return False, "HTML response is not a downloadable file"
    ext = os.path.splitext(out_path)[1].lower()
    if expect_pdf or ext == ".pdf":
        if payload[:5] != b"%PDF-":
            return False, f"invalid PDF header: {payload[:5]!r}"
    elif ext in (".zip", ".docx", ".xlsx", ".pptx") and not payload.startswith(b"PK"):
        return False, f"invalid ZIP/Office signature: {payload[:4]!r}"

    d = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(d, exist_ok=True)
    fd, part = tempfile.mkstemp(prefix=".scidownload-", suffix=".part", dir=d)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(payload)
            f.flush()
            os.fsync(f.fileno())
        os.replace(part, out_path)
    except Exception:
        try:
            os.remove(part)
        except OSError:
            pass
        raise
    return True, f"{len(payload)}B"



def _set_download_behavior(tab: Tab, download_dir: str) -> tuple[str | None, str]:
    """把浏览器下载临时定向到 ``download_dir``。

    优先使用当前 Chromium 仍支持的 Page.setDownloadBehavior；若浏览器只接受
    Browser.setDownloadBehavior，则自动回退。只有设置成功后才会导航可能触发
    attachment 的 URL，避免控制失败时把文件意外扔进用户系统 Downloads。
    """
    errors = []
    calls = (
        ("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir,
        }),
        ("Browser.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir,
            "eventsEnabled": False,
        }),
    )
    for method, params in calls:
        try:
            tab._call(method, params, timeout=20)
            return method, ""
        except Exception as e:
            errors.append(f"{method}: {type(e).__name__}: {str(e)[:100]}")
    return None, "; ".join(errors)


def _reset_download_behavior(tab: Tab, method: str | None):
    if not method:
        return
    try:
        tab._call(method, {"behavior": "default"}, timeout=15)
    except Exception:
        pass


def _download_dir_state(download_dir: str):
    """返回 (完成文件, 临时文件)。目录是每候选独占的，因此无需比较旧文件。"""
    completed, partial = [], []
    try:
        names = os.listdir(download_dir)
    except OSError:
        return completed, partial
    for name in names:
        p = os.path.join(download_dir, name)
        if not os.path.isfile(p):
            continue
        low = name.lower()
        if low.endswith((".crdownload", ".tmp")):
            partial.append(p)
        else:
            completed.append(p)
    return completed, partial


def _finalize_browser_pdf(src: str, out_path: str):
    """验证浏览器下载产物并原子移入正式 PDF 路径。"""
    try:
        size = os.path.getsize(src)
        with open(src, "rb") as f:
            head = f.read(512)
    except OSError as e:
        return False, str(e)
    if size < 1024:
        return False, f"browser download too small: {size}B"
    stripped = head.lstrip().lower()
    if stripped.startswith((b"<!doctype html", b"<html", b"<head", b"<body")):
        return False, "browser download is HTML, not PDF"
    if head[:5] != b"%PDF-":
        return False, f"browser download has invalid PDF header: {head[:5]!r}"

    d = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(d, exist_ok=True)
    # download_dir 被创建在正式输出目录下，因此 os.replace 通常是同文件系统原子移动。
    os.replace(src, out_path)
    return True, f"{size}B via browser download"


def fetch_pdf_candidate(tab: Tab, url: str, out_path: str,
                        total_timeout: float = 45.0):
    """一次导航兼容 inline PDF 与 ``Content-Disposition: attachment``。

    返回 ``(ok, message, browser_download_started, current_href)``。
    ``browser_download_started`` 供上层阻止对同一个 URL 重复导航。
    """
    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    download_dir = tempfile.mkdtemp(prefix=".scidownload-browser-", dir=out_dir)
    behavior_method = None
    download_started = False
    last = "PDF candidate did not become downloadable"
    href = ""

    try:
        behavior_method, err = _set_download_behavior(tab, download_dir)
        if not behavior_method:
            # 关键安全点：未能接管 Chrome 下载目录时不导航 attachment 候选，
            # 否则文件会再次落到系统默认 Downloads 而程序仍无法追踪。
            # 仍可在当前文章页上下文直接 fetch 一次；若站点允许 CORS/同源，这条
            # 路径保持对旧 Chromium 的兼容，而且不会触发系统下载管理器。
            ok, fetch_msg = fetch_bytes_via_page(
                tab, url, out_path, total_timeout=max(15.0, float(total_timeout))
            )
            if ok:
                return True, fetch_msg, False, href
            return False, (f"browser download capture unavailable: {err}; "
                           f"page fetch failed: {fetch_msg}"), False, href

        tab.front()
        try:
            nav = tab._call("Page.navigate", {"url": url}, timeout=30) or {}
            download_started = bool(nav.get("isDownload"))
        except Exception as e:
            nav = {}
            last = f"Page.navigate: {type(e).__name__}: {str(e)[:120]}"

        deadline = time.time() + max(5.0, float(total_timeout))
        stable_path = None
        stable_size = None
        stable_ticks = 0

        while time.time() < deadline:
            completed, partial = _download_dir_state(download_dir)
            if completed or partial:
                download_started = True

            # Chrome 正常会先写 *.crdownload，再原子改成最终文件名。
            # 为兼容不同 Chromium 版本，再额外要求最终文件大小连续稳定两次。
            if completed and not partial:
                src = max(completed, key=lambda x: os.path.getsize(x))
                try:
                    size = os.path.getsize(src)
                except OSError:
                    size = -1
                if src == stable_path and size == stable_size and size >= 0:
                    stable_ticks += 1
                else:
                    stable_path, stable_size, stable_ticks = src, size, 0
                if stable_ticks >= 2:
                    ok, msg = _finalize_browser_pdf(src, out_path)
                    return ok, msg, True, href

            # attachment 导航通常不会把当前页面切成 PDF；inline PDF 则会。
            try:
                st = tab.js("JSON.stringify([document.contentType, location.href])", timeout=10)
                ct, href = json.loads(st) if st else (None, "")
                h = str(href or "").lower()
                if ct == "application/pdf" or h.endswith(".pdf"):
                    ok, msg = fetch_bytes_via_page(tab, href, out_path)
                    if ok:
                        return True, msg, download_started, href
                    last = msg
            except Exception as e:
                # 下载开始后页面上下文可能保持旧页或短暂不可用，这是正常现象。
                last = f"{type(e).__name__}: {str(e)[:120]}"

            time.sleep(0.5)

        if download_started:
            completed, partial = _download_dir_state(download_dir)
            if partial:
                return False, "browser download started but did not finish before timeout", True, href
            if completed:
                src = max(completed, key=lambda x: os.path.getsize(x))
                ok, msg = _finalize_browser_pdf(src, out_path)
                return ok, msg, True, href
            return False, "browser download was signaled but no file appeared", True, href
        return False, last, False, href
    finally:
        _reset_download_behavior(tab, behavior_method)
        try:
            shutil.rmtree(download_dir)
        except OSError:
            pass


def probe_entitlement(port: int, timeout: float = 15.0) -> dict:
    """快速自检：当前出口 IP 能否被机构识别（用于 preflight）。"""
    out = {}
    try:
        tab = Tab(port, new_tab(port, "about:blank"))
        tab.goto("https://ipinfo.io/json", settle=3)
        body = tab.js("document.body ? document.body.innerText : ''") or ""
        out["ip_json"] = body[:300]
        tab.close()
    except Exception as e:
        out["ip_error"] = repr(e)[:160]
    return out

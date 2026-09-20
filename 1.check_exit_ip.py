#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机构授权网络自检。

用途：排查当前 Python 进程是否因为系统代理、本机代理或 VPN 使用了与预期不同的
出口网络。机构访问可能依赖出口 IP，也可能依赖 CARSI / EZproxy / SSO / 已登录会话，
因此本脚本只做网络路径诊断，不证明或否定某篇文章的订阅权限。

默认会脱敏公网 IP；如确实需要查看完整 IP，可显式加：

    python 1.check_exit_ip.py --show-full-ip
"""
from __future__ import annotations

import argparse
import ipaddress
import json
import os
import platform
import socket
import urllib.parse
import urllib.request

OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

KNOWN_PROXY_PORTS = [7897, 7890, 10809, 10808, 1080, 8889, 33210, 7891]
SD_HOSTS = ["sciencedirect.com", "doi.org", "pdf.sciencedirectassets.com",
            "ars.els-cdn.com"]


def header(t):
    print(f"\n{'=' * 62}\n{t}\n{'=' * 62}")


def mask_ip(value):
    s = str(value or "").strip()
    if not s:
        return "-"
    try:
        ip = ipaddress.ip_address(s)
        if ip.version == 4:
            a, b, _, _ = s.split(".")
            return f"{a}.{b}.x.x"
        return ":".join(ip.exploded.split(":")[:3]) + ":…"
    except ValueError:
        return "(已隐藏)"


def shown_ip(value, show_full=False):
    return str(value or "-") if show_full else mask_ip(value)


def redact_proxy_value(value):
    """只展示代理端点；隐藏用户名、密码、路径、查询参数和片段。"""
    try:
        parsed = urllib.parse.urlsplit(str(value))
        if not parsed.hostname:
            return "(已隐藏代理详情)"
        host = parsed.hostname
        if ":" in host and not host.startswith("["):
            host = f"[{host}]"
        port = f":{parsed.port}" if parsed.port else ""
        scheme = f"{parsed.scheme}://" if parsed.scheme else ""
        return f"{scheme}{host}{port}"
    except (TypeError, ValueError):
        return "(已隐藏代理详情)"


def proxy_env():
    items = [(k, v) for k, v in sorted(os.environ.items()) if "proxy" in k.lower()]
    if items:
        for k, v in items:
            print(f"  {k} = {redact_proxy_value(v)}")
        print("  （这些环境变量会影响使用它们的程序；本脚本自身已强制直连）")
    else:
        print("  未设置代理环境变量")
    return items


def _open(port):
    s = socket.socket()
    s.settimeout(0.5)
    try:
        s.connect(("127.0.0.1", port))
        return True
    except Exception:
        return False
    finally:
        s.close()


def listening_ports():
    ports = [p for p in KNOWN_PROXY_PORTS if _open(p)]
    print("  常见代理端口：" + (", ".join(f"{p} 在监听" for p in ports) if ports
                              else "都没在监听"))
    if ports:
        print("  ⚠️  有本机代理进程在监听；是否影响机构访问取决于你的分流规则。")
    return ports


def _fetch_ipinfo(opener, timeout=20):
    """只用 HTTPS 查询公开出口信息；失败时返回错误文本。"""
    endpoints = (
        "https://ipinfo.io/json",
        "https://ipapi.co/json/",
    )
    last = "unknown"
    for url in endpoints:
        try:
            with opener.open(url, timeout=timeout) as r:
                d = json.loads(r.read())
            ip = d.get("ip") or d.get("query")
            org = d.get("org") or d.get("as") or d.get("isp") or ""
            country = d.get("country_name") or d.get("country") or ""
            city = d.get("city") or ""
            return (ip, str(org), str(country), str(city)), None
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:90]}"
    return (None, "", "", ""), last


def exit_ip(show_full=False):
    (ip, org, country, city), err = _fetch_ipinfo(OPENER)
    if not ip:
        print(f"  ✗ 探测失败：{err}")
        return None, "", ""
    print(f"  IP   : {shown_ip(ip, show_full)}")
    print(f"  归属 : {org or '-'}")
    print(f"  国家/地区 : {country or '-'}")
    if show_full and city:
        print(f"  城市 : {city}")
    return ip, org, country


def direct_or_proxied():
    """对比直连与本机代理出口；结果内部保留原 IP，只在显示时脱敏。"""
    results = {}
    for label, use_proxy in (("直连", False), ("走本机代理", True)):
        if not use_proxy:
            op = OPENER
        else:
            found = next((p for p in KNOWN_PROXY_PORTS if _open(p)), None)
            if not found:
                continue
            op = urllib.request.build_opener(urllib.request.ProxyHandler(
                {"http": f"http://127.0.0.1:{found}",
                 "https": f"http://127.0.0.1:{found}"}))
        data, err = _fetch_ipinfo(op, timeout=25)
        if err:
            results[label] = ("失败", err, "")
        else:
            ip, org, country, _city = data
            results[label] = (ip, org, country)
    return results


def main():
    ap = argparse.ArgumentParser(description="SCIDownload 机构授权网络自检")
    ap.add_argument("--show-full-ip", action="store_true",
                    help="显示完整公网 IP（默认脱敏；公开截图前不要使用此项）")
    args = ap.parse_args()

    print(f"SCIDownload 机构授权自检 | {platform.system()} | Python {platform.python_version()}")

    header("1. 代理环境")
    proxy_env()

    header("2. 本机代理端口")
    listening_ports()

    header("3. 当前 Python 进程的出口网络（仅供参考）")
    ip, org, country = exit_ip(args.show_full_ip)

    header("4. 直连 vs 走代理（对比）")
    cmp = direct_or_proxied()
    for k, v in cmp.items():
        print(f"  {k:10} : {shown_ip(v[0], args.show_full_ip)}  |  {v[1]}  |  {v[2]}")
    if len(cmp) >= 2:
        a, b = list(cmp.values())[0], list(cmp.values())[1]
        if a[0] != "失败" and b[0] != "失败" and a[0] != b[0]:
            print("  ⚠️  两者出口不同，说明本机代理会改变网络身份；"
                  "是否需要直连请以你的机构访问方式为准。")

    header("5. 结论")
    if ip:
        up = org.upper()
        if "CERNET" in up or "EDU" in up or "UNIVERSIT" in up:
            print("  ✅ 当前出口看起来像教育/校园网络，但实际全文权限仍以文章页为准。")
        else:
            print("  ℹ️  仅凭 ASN/国家无法判断订阅权限。CARSI、EZproxy、学校 VPN、"
                  "账号会话等都可能合法提供访问。")
            print("      若文章页意外变成无权限，可检查以下域名是否被不合适的代理规则接管：")
            for h in SD_HOSTS:
                print(f"         {h}")
    print()
    print("  提示：首次需要机构登录时，可给 SCIDownload.py 加 --login-wait 300，")
    print("        在弹出的专用浏览器中完成授权登录。不要分享 browser profile。")
    print("        本检查不会证明或否定机构订阅权限，只用于排查网络路径。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

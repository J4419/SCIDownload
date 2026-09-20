#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ClearCache —— 清理 SCIDownload 积攒的浏览器缓存。

    python 4.ClearCache.py            # 只清缓存，**保留登录态与 Cookie**
    python 4.ClearCache.py --all      # 连 profile 一起删（下次要重新登录）
    python 4.ClearCache.py --dry-run  # 只看会删什么，不动手

浏览器跑久了会攒下几百 MB 缓存。跑完一批之后执行一次即可。
"""

import argparse
import os
import pathlib
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

try:
    from SCIDownload_cdp import default_profile_dir
except Exception:                                     # 单文件被拎出去跑时兜底
    def default_profile_dir():
        import platform
        if platform.system() == "Windows":
            base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        elif platform.system() == "Darwin":
            base = os.path.expanduser("~/Library/Application Support")
        else:
            base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
        return os.path.join(base, "SCIDownload", "browser-profile")


# 只删这些「纯缓存」目录 —— Cookie / Login Data / Web Data 一律不动
CACHE_DIRS = (
    "Cache", "Code Cache", "GPUCache", "ShaderCache", "GrShaderCache",
    "DawnCache", "DawnGraphiteCache", "DawnWebGPUCache",
    "Application Cache", "Media Cache", "blob_storage",
    "component_crx_cache", "extensions_crx_cache",
    "Crashpad", "BrowserMetrics", "Safe Browsing",
    "OptimizationGuide", "segmentation_platform",
    os.path.join("Service Worker", "CacheStorage"),
    os.path.join("Service Worker", "ScriptCache"),
    os.path.join("Default", "Cache"),
    os.path.join("Default", "Code Cache"),
    os.path.join("Default", "GPUCache"),
    os.path.join("Default", "Service Worker", "CacheStorage"),
    os.path.join("Default", "Service Worker", "ScriptCache"),
)


def dir_size(path: str) -> int:
    total = 0
    for r, _, fs in os.walk(path):
        for f in fs:
            try:
                total += os.path.getsize(os.path.join(r, f))
            except OSError:
                pass
    return total


def human(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{int(n)} B"
        n /= 1024
    return f"{n:.1f} GB"


def validate_profile_target(path: str, default_path: str,
                            allow_custom: bool) -> str:
    """删除前校验 profile 目标，危险路径一律拒绝。"""
    resolved = os.path.realpath(os.path.abspath(os.path.expanduser(path)))
    default_resolved = os.path.realpath(os.path.abspath(os.path.expanduser(default_path)))
    home = os.path.realpath(os.path.abspath(os.path.expanduser("~")))
    anchor = pathlib.Path(resolved).anchor
    if not resolved or os.path.normcase(resolved) == os.path.normcase(anchor):
        raise ValueError("拒绝删除文件系统根目录")
    if os.path.normcase(resolved) == os.path.normcase(home):
        raise ValueError("拒绝删除用户主目录")
    meaningful = [p for p in pathlib.Path(resolved).parts if p != anchor]
    if len(meaningful) < 2:
        raise ValueError("拒绝删除层级过浅的目录")
    is_default = os.path.normcase(resolved) == os.path.normcase(default_resolved)
    if not is_default and not allow_custom:
        raise ValueError(
            "自定义 --profile 需同时加 --confirm-custom-profile 才允许清理")
    return resolved


def try_remove(path: str, dry: bool):
    """返回 (状态, 释放字节数, 错误)；状态 None 表示不存在。"""
    if not os.path.exists(path):
        return None, 0, None
    size = dir_size(path) if os.path.isdir(path) else os.path.getsize(path)
    if dry:
        return True, size, None
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
        else:
            os.remove(path)
    except OSError as e:
        return False, 0, str(e)
    if os.path.exists(path):
        return False, 0, "删除返回后目标仍存在"
    return True, size, None


def main() -> int:
    ap = argparse.ArgumentParser(
        description="清理 SCIDownload 的浏览器缓存",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--all", action="store_true",
                    help="连整个 profile 一起删（登录态也清掉，下次需重新登录）")
    ap.add_argument("--dry-run", action="store_true", help="只显示会删什么，不实际删除")
    ap.add_argument("--profile", default=None, help="指定 profile 目录（默认用程序内置路径）")
    ap.add_argument("--confirm-custom-profile", action="store_true",
                    help="明确确认允许清理自定义 --profile 目录")
    args = ap.parse_args()

    default_profile = default_profile_dir()
    try:
        profile = validate_profile_target(
            args.profile or default_profile,
            default_profile,
            args.confirm_custom_profile,
        )
    except ValueError as e:
        print(f"拒绝执行：{e}")
        return 2
    print(f"目标 profile：{profile}")
    if not os.path.exists(profile):
        print("  没有找到 profile —— 无需清理（可能还没跑过，或已经被删了）。")
        return 0

    before = dir_size(profile)
    print(f"  当前占用：{human(before)}")
    print()

    freed = 0
    failures = 0
    if args.all:
        print("模式：--all   连 profile 一起删（登录态也会清掉）")
        ok, n, error = try_remove(profile, args.dry_run)
        if ok is True:
            freed += n
            print(f"  [{'将删' if args.dry_run else '已删'}] 整个 profile  {human(n)}")
        else:
            failures += 1
            print(f"  ✗ 删除失败：{error or '未知错误'}")
    else:
        print("模式：默认   只清缓存，**保留登录态与 Cookie**")
        hit = miss = locked = 0
        for rel in CACHE_DIRS:
            p = os.path.join(profile, rel)
            ok, n, error = try_remove(p, args.dry_run)
            if ok is None:
                miss += 1
                continue
            if ok is True:
                hit += 1
                freed += n
                print(f"  [{'将删' if args.dry_run else '已删'}] {rel:<44} {human(n)}")
            else:
                locked += 1
                failures += 1
                print(f"  ✗ {rel:<44} 删除失败：{error or '浏览器可能仍在占用'}")
        print()
        print(f"  命中 {hit} 项，跳过 {miss} 项（本来就没有），被占用 {locked} 项")

    if args.dry_run:
        print(f"\n[dry-run] 预计可释放 {human(freed)}（未实际删除）")
    else:
        after = dir_size(profile) if os.path.exists(profile) else 0
        print(f"\n已释放 {human(freed)}；profile 现占 {human(after)}")
        if not args.all:
            print("登录态已保留 —— 下次运行不用重新登录。")
        else:
            print("登录态已清除 —— 下次首次运行需要重新登录，前几篇可能因此先落空，"
                  "工具会自动补跑。")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())

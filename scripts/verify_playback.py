#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 WPS 前台放映抽样截图，验证 pptx 里的原生动画真的在播。

    python verify_playback.py --pptx 动效.pptx --frames 14 --interval 0.25 --out sheet.jpg
    python verify_playback.py --pptx 动效.pptx --frames 8 --interval 0.2 --next 2   # 第 2 页再采 8 张

只负责"看一眼"：返回的拼贴图里，如果同一页出现不同透明度/位移/旋转状态，就说明动画生效；
如果每张都完全一样，说明动画没写进 timing 或 WPS 读不到。
依赖：pywin32 + Pillow，WPS 安装在默认路径（可用 --wpp 指定）。
"""
from __future__ import annotations

import argparse
import ctypes
import subprocess
import sys
import time
from pathlib import Path

try:
    import win32con
    import win32gui
    from PIL import Image, ImageGrab
except ImportError as exc:  # pragma: no cover
    print(f"缺少依赖：{exc}（需要 pywin32 与 Pillow）", file=sys.stderr)
    raise SystemExit(2)

DEFAULT_WPP = r"C:\Users\30246\AppData\Local\Kingsoft\WPS Office\12.1.0.29690\office6\wpp.exe"

VK = {"F5": 0x74, "RIGHT": 0x27, "LEFT": 0x25, "ESC": 0x1B, "ENTER": 0x0D}


def tap(vk: int, hold: float = 0.05) -> None:
    ctypes.windll.user32.keybd_event(vk, 0, 0, 0)
    time.sleep(hold)
    ctypes.windll.user32.keybd_event(vk, 0, 2, 0)


def find_window(title_part: str, timeout: float = 40.0) -> int:
    deadline = time.time() + timeout
    while time.time() < deadline:
        found: list[int] = []

        def _cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            text = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            if title_part.lower() in text.lower() or cls.startswith("KWPP"):
                found.append(hwnd)

        win32gui.EnumWindows(_cb, None)
        if found:
            return max(found, key=lambda h: win32gui.GetWindowRect(h)[2] - win32gui.GetWindowRect(h)[0])
        time.sleep(0.5)
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--out", required=True, help="拼贴图输出路径 jpg")
    ap.add_argument("--wpp", default=DEFAULT_WPP)
    ap.add_argument("--frames", type=int, default=14)
    ap.add_argument("--interval", type=float, default=0.25)
    ap.add_argument("--lead", type=float, default=0.6, help="进入放映后的等待秒数")
    ap.add_argument("--next", type=int, default=0, help="采样前先按几次右方向键（翻到第 N 页）")
    ap.add_argument("--goto", type=int, default=0, help="放映中直接跳到第 N 页（页码 + 回车）")
    ap.add_argument("--cols", type=int, default=4)
    args = ap.parse_args(argv)

    pptx = Path(args.pptx).resolve()
    if not pptx.exists():
        print(f"文件不存在：{pptx}", file=sys.stderr)
        return 2
    if not Path(args.wpp).exists():
        print(f"找不到 WPS：{args.wpp}（用 --wpp 指定 wpp.exe）", file=sys.stderr)
        return 2

    proc = subprocess.Popen([args.wpp, str(pptx)])
    hwnd = find_window(pptx.stem)
    if not hwnd:
        print("没等到 WPS 窗口", file=sys.stderr)
        proc.terminate()
        return 3
    time.sleep(1.2)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.6)

    tap(VK["F5"])
    time.sleep(args.lead)
    if args.goto:
        for ch in str(args.goto):
            tap(0x30 + int(ch), 0.06)
        tap(VK["ENTER"])
        time.sleep(0.8)
    for _ in range(args.next):
        tap(VK["RIGHT"])
        time.sleep(0.5)

    frames = []
    for i in range(args.frames):
        frames.append(ImageGrab.grab())
        time.sleep(args.interval)

    tap(VK["ESC"])
    time.sleep(0.4)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass
    time.sleep(0.8)
    if proc.poll() is None:
        proc.terminate()

    # 裁掉黑边后拼贴
    tiles = []
    for im in frames:
        w, h = im.size
        tiles.append(im.crop((0, 0, w, h)).resize((w // 3, h // 3)))
    tw, th = tiles[0].size
    rows = (len(tiles) + args.cols - 1) // args.cols
    sheet = Image.new("RGB", (tw * args.cols, th * rows), (12, 12, 12))
    for i, im in enumerate(tiles):
        sheet.paste(im, ((i % args.cols) * tw, (i // args.cols) * th))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.out, quality=88)
    print(f"[OK] {args.out}  共 {len(frames)} 帧，{args.cols} 列")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

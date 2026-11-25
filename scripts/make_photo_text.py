#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把文字做成"照片填充 + 透明底"的 PNG：字里透出雪山山脊，用于封面巨型字横向掠过。

    python make_photo_text.py --text "BE MYSELF" --fill 雪山.jpg --out word.png \
        --font "Times New Roman" --size 260 --spacing 6 --fill-zoom 1.6

原理：Chrome headless 渲染 HTML，用 background-clip:text 把图片裁进字形，
再用 --default-background-color=00000000 导出带 alpha 的截图（不是抠图，边缘干净）。
"""
from __future__ import annotations

import argparse
import html
import os
import subprocess
import sys
import tempfile
from pathlib import Path

CHROME_CANDIDATES = [
    # Windows
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    # macOS
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    # Linux
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
]

ENV_CHROME = "CHROME_PATH"


def find_chrome(explicit: str | None) -> str:
    if explicit:
        return explicit
    env = os.environ.get(ENV_CHROME)
    if env and Path(env).exists():
        return env
    for c in CHROME_CANDIDATES:
        if Path(c).exists():
            return c
    raise SystemExit(f"找不到 Chrome，用 --chrome 指定 chrome.exe，或设环境变量 {ENV_CHROME}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--text", required=True, help="要做的文字（建议全大写英文，中文也行）")
    ap.add_argument("--fill", required=True, help="填充用图片（建议山脊/水面等有明暗层次的横构图）")
    ap.add_argument("--out", required=True, help="输出 PNG（透明底）")
    ap.add_argument("--font", default="Times New Roman")
    ap.add_argument("--size", type=int, default=260, help="字号 px")
    ap.add_argument("--weight", default="700")
    ap.add_argument("--spacing", type=float, default=4, help="字距 px")
    ap.add_argument("--pad", type=int, default=40, help="四周留白 px，避免裁到字形")
    ap.add_argument("--fill-zoom", type=float, default=1.0, help="填充图放大倍数，>1 更聚焦山脊")
    ap.add_argument("--chrome")
    ap.add_argument("--scale", type=float, default=2.0, help="渲染倍率，2 更锐")
    args = ap.parse_args(argv)

    chrome = find_chrome(args.chrome)
    fill = Path(args.fill).resolve()
    if not fill.is_file():
        print(f"填充图不存在或是目录：{fill}", file=sys.stderr)
        return 2

    # 先按字号估宽高，Chrome 会按 window-size 截图，再自动裁边
    est_w = int(args.size * max(len(args.text), 1) * 0.72 + args.spacing * len(args.text) + args.pad * 2)
    est_h = int(args.size * 1.5 + args.pad * 2)
    body = f"""
<!doctype html><html><head><meta charset="utf-8"><style>
  html,body{{margin:0;padding:0;background:transparent;}}
  .wrap{{display:inline-block;padding:{args.pad}px;}}
  .word{{
    font-family:"{args.font}";font-size:{args.size}px;font-weight:{args.weight};
    letter-spacing:{args.spacing}px;line-height:1.1;white-space:nowrap;
    background-image:url("file:///{fill.as_posix().lstrip('/')}");
    background-size:{int(args.fill_zoom * 100)}% auto;background-position:center 42%;
    -webkit-background-clip:text;background-clip:text;color:transparent;
  }}
</style></head><body><div class="wrap"><span class="word">{html.escape(args.text)}</span></div></body></html>
"""
    tmp = Path(tempfile.mkdtemp(prefix="phototext_"))
    page = tmp / "word.html"
    page.write_text(body, encoding="utf-8")

    out_png = Path(args.out).resolve()
    out_png.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        chrome,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--force-device-scale-factor=" + str(args.scale),
        "--default-background-color=00000000",
        "--virtual-time-budget=2000",
        f"--window-size={est_w},{est_h}",
        f"--screenshot={out_png}",
        page.as_uri(),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if not out_png.exists():
        print(r.stdout[-400:], r.stderr[-400:], file=sys.stderr)
        return 3

    try:
        from PIL import Image

        im = Image.open(out_png).convert("RGBA")
        bbox = im.getbbox()          # 自动去掉多余透明边
        if bbox:
            im.crop(bbox).save(out_png)
            print(f"[OK] {out_png}  {im.size} -> {Image.open(out_png).size}（已裁透明边）")
        else:
            print(f"[OK] {out_png}  {im.size}")
    except ImportError:
        print(f"[OK] {out_png}（未装 Pillow，未裁边）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

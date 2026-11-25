#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
油画切片开场（参考视频同款四拍）—— 素材生成 + 动画注入，一条命令跑完。

四拍（WPS / PowerPoint 里真播，仍是可编辑 pptx）：
    1) 一叠白边照片卡绕支点扇形张开，边转边亮边
    2) 卡片收成一条线，整叠退回画外
    3) 圆环拼贴斜着滑进画面中心，停一拍再旋出
    4) 整屏油画切成 N 条竖切片，从画面中心向左右同时展开铺满
    5) 左侧毛玻璃面板从左边擦入，然后原有文字层接着进场

用法：
    python oil_opening.py --deck 静态.pptx --out 开场版.pptx
    python oil_opening.py --deck in.pptx --out out.pptx --art 自我油画.jpg
    python oil_opening.py --deck in.pptx --out out.pptx --dry-run
    python oil_opening.py --deck in.pptx --out out.pptx --no-cards --no-ring   # 只要切片+面板

素材来源：默认取封面"面积≥80% 的那张图"当油画；`--art` 可指定别的大图。
素材与几何都在临时目录里现算，不写死任何路径。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
import tempfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter
from pptx import Presentation

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU_IN = 914400

# ---------------------------------------------------------------- 时间轴（毫秒）
T = {
    "card_in": 90,       "card_stagger": 34,  "card_dur": 760, "card_spin": 900,
    "card_out": 1000,    "card_out_stagger": 16, "card_out_dur": 360, "card_unspin": 52,
    "ring_in": 1040,     "ring_in_dur": 480,
    "ring_out": 1900,    "ring_out_dur": 420,
    "bar_in": 1520,      "bar_stagger": 22,   "bar_dur": 700, "bar_fade": 240,
    "panel_in": 2300,    "panel_dur": 900,
    "bg_fade": 1250,     "bg_fade_dur": 520,
    "scrim_in": 1950,
    "text_start": 3300,  "shift_from": 900,   "compress": 0.55,
}

# ---------------------------------------------------------------- 版面比例（相对页宽/页高）
G = {
    "pivot": (0.454, 0.420),   # 扇形支点 / 切片炸开原点
    "card_w": 0.1425,          # 卡片宽 = 14.25% 页宽
    "card_ar": 0.71,           # 卡片高宽比
    "slot_x0": 0.259,          # 第一张卡终点中心 X
    "slot_step": 0.0465,       # 卡片间距
    "card_rot0": -52.0,        # 起始角
    "card_rot_step": 15.5,     # 每张递增角
    "ring_w": 0.315,           # 圆环宽 = 31.5% 页宽
    "ring_park": -0.705,       # 圆环候场位置（页宽倍数，负=画外左侧）
    "gap_px": 9,               # 切片竖缝（按 2400px 画布）
    "edge_px": 2,              # 切片描边
    "panel_ratio": 0.655,      # 面板宽度占页宽
    "feather_ratio": 0.030,    # 面板右沿羽化
}


# ================================================================ 素材
def _cover_crop(im: Image.Image, W: int, H: int, zoom: float = 1.0, focus: float = 0.5):
    w, h = im.size
    target = W / H
    if w / h > target:
        cw, ch = int(h * target / zoom), int(h / zoom)
        x0, y0 = int((w - cw) * focus), int((h - ch) * 0.45)
    else:
        cw, ch = int(w / zoom), int(w / target / zoom)
        x0, y0 = 0, int((h - ch) * 0.42)
    return im.crop((x0, y0, x0 + cw, y0 + ch)).resize((W, H), Image.LANCZOS)


def _vignette(im: Image.Image, W: int, H: int, edge: float = 0.86):
    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).ellipse((-W * 0.18, -H * 0.22, W * 1.18, H * 1.22), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(W // 15))
    dark = Image.blend(im, Image.new("RGB", (W, H), (4, 7, 12)), 1 - edge)
    return Image.composite(im, dark, mask)


def build_assets(art: Image.Image, work: Path, bars_n: int, cards_n: int, tag: str) -> dict:
    """生成竖切片、毛玻璃面板、白边照片卡、圆环拼贴；返回几何清单。"""
    work.mkdir(parents=True, exist_ok=True)
    W, H = 2400, int(round(2400 * art.size[1] / art.size[0]))
    # 画布只保证与油画等比；贴到页面时按页宽统一缩放
    base = _cover_crop(art, W, H, zoom=1.06, focus=0.55)

    gap = G["gap_px"]
    sw = (W - gap * (bars_n - 1)) // bars_n
    split = int(W * G["panel_ratio"])
    feather = int(W * G["feather_ratio"])
    panel_w = split + feather

    # 1) 毛玻璃（左半）
    glass = base.filter(ImageFilter.GaussianBlur(46))
    glass = ImageEnhance.Brightness(glass).enhance(0.62)
    glass = ImageEnhance.Color(glass).enhance(0.62)
    glass = Image.blend(glass, Image.new("RGB", (W, H), (12, 16, 24)), 0.16)
    glass = Image.blend(glass, Image.new("RGB", (W, H), (232, 238, 246)), 0.10)
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    md.rectangle((0, 0, split, H), fill=255)
    md.rectangle((split, 0, split + feather, H), fill=150)
    mask = mask.filter(ImageFilter.GaussianBlur(max(8, feather // 3)))

    # 2) 切片终态合成（每条带 2px 浅边 + 奇偶列上下错位）
    comp = base.copy()
    for i in range(bars_n):
        x = i * (sw + gap)
        patch = Image.new("RGB", (sw, H), (6, 9, 14))
        patch.paste(base.crop((x, 0, x + sw, H)), (0, 14 if i % 2 == 0 else -12))
        comp.paste(patch, (x, 0))
        ImageDraw.Draw(comp).rectangle((x, 0, x + sw - 1, H),
                                       outline=(150, 165, 190), width=G["edge_px"])
    comp = _vignette(comp, W, H)

    # 3) 面板 = 毛玻璃压左半（含羽化右沿）
    panel = comp.crop((0, 0, panel_w, H))
    panel = Image.composite(glass.crop((0, 0, panel_w, H)), panel, mask.crop((0, 0, panel_w, H)))
    panel.save(work / f"panel_{tag}.png")

    # 4) 切片
    bars = []
    for i in range(bars_n):
        x = i * (sw + gap)
        f = work / f"bar_{tag}_{i + 1:02d}.png"
        comp.crop((x, 0, x + sw, H)).save(f)
        bars.append({"file": f.name, "x": x, "w": sw})

    # 5) 白边照片卡
    cards = []
    crops = [(0.55, 0.30, 0.55), (0.60, 0.62, 0.45), (0.50, 0.45, 0.30), (0.58, 0.72, 0.62),
             (0.52, 0.25, 0.35), (0.62, 0.50, 0.70), (0.48, 0.68, 0.28), (0.56, 0.40, 0.62),
             (0.54, 0.34, 0.48), (0.60, 0.66, 0.38)]
    cw_px, ch_px, border, shadow = 300, 200, 8, 14
    for i in range(cards_n):
        zoom, cx, cy = crops[i % len(crops)]
        src = base
        sw_, sh_ = src.size
        cw, ch = int(sw_ * zoom * 1.6), int(sh_ * zoom * 1.6)
        left = max(0, min(sw_ - cw, int(sw_ * cx - cw / 2)))
        top = max(0, min(sh_ - ch, int(sh_ * cy - ch / 2)))
        pic = src.crop((left, top, left + cw, top + ch)).resize((cw_px, ch_px), Image.LANCZOS)
        canvas = Image.new("RGBA", (cw_px + 2 * border + 2 * shadow,
                                    ch_px + 2 * border + 2 * shadow), (0, 0, 0, 0))
        sh_layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        ImageDraw.Draw(sh_layer).rectangle(
            [shadow - 2, shadow + 4, shadow + cw_px + 2 * border + 2,
             shadow + 2 * border + ch_px + 6], fill=(0, 0, 0, 150))
        canvas = Image.alpha_composite(canvas, sh_layer.filter(ImageFilter.GaussianBlur(9)))
        ImageDraw.Draw(canvas).rectangle(
            [shadow, shadow, shadow + cw_px + 2 * border, shadow + ch_px + 2 * border],
            fill=(248, 248, 246, 255))
        canvas.paste(pic, (shadow + border, shadow + border))
        f = work / f"card_{tag}_{i + 1:02d}.png"
        canvas.save(f)
        cards.append(f.name)

    # 6) 圆环拼贴
    size, r_out, r_in, seg = 1300, 620, 300, 12
    ring = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    c = size // 2
    for s in range(seg):
        a0 = s * (360 / seg) - 90
        a1 = a0 + 360 / seg + 0.6
        seg_mask = Image.new("L", (size, size), 0)
        sd = ImageDraw.Draw(seg_mask)
        sd.pieslice([c - r_out, c - r_out, c + r_out, c + r_out], a0, a1, fill=255)
        sd.pieslice([c - r_in, c - r_in, c + r_in, c + r_in], a0, a1, fill=0)
        src = base
        sw_, sh_ = src.size
        zw = int(sw_ * 0.30)
        ox = max(0, min(sw_ - zw, int(sw_ * (0.25 + 0.5 * ((s * 7) % 5) / 4)) - zw // 2))
        zh = int(zw * sh_ / sw_)
        oy = max(0, min(sh_ - zh, int(sh_ * (0.25 + 0.5 * ((s * 3) % 5) / 4)) - zh // 2))
        tile = src.crop((ox, oy, ox + zw, oy + zh)).resize((2 * r_out, 2 * r_out), Image.LANCZOS)
        ring.paste(tile, (c - r_out, c - r_out),
                   seg_mask.crop((c - r_out, c - r_out, c + r_out, c + r_out)))
    ring.save(work / f"ring_{tag}.png")

    return {"canvas": [W, H], "sw": sw, "gap": gap, "panel_w": panel_w,
            "bars": bars, "cards": cards, "ring": f"ring_{tag}.png",
            "panel": f"panel_{tag}.png"}


# ================================================================ 动画 XML
class Ids:
    def __init__(self) -> None:
        self._n = 1

    def nxt(self) -> int:
        self._n += 1
        return self._n


def _vis(i: Ids, spid: int) -> str:
    return (f'<p:set><p:cBhvr><p:cTn id="{i.nxt()}" dur="1" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>')


def eff_fade(i: Ids, spid: int, delay: int, dur: int) -> str:
    return (f'<p:par><p:cTn id="{i.nxt()}" presetID="10" presetClass="entr" presetSubtype="0" '
            f'fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/>'
            f'</p:stCondLst><p:childTnLst>{_vis(i, spid)}'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
            f'<p:cTn id="{i.nxt()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')


def eff_wipe(i: Ids, spid: int, delay: int, dur: int, direction: str = "left") -> str:
    subtype = {"left": "2", "right": "1", "up": "8", "down": "4"}.get(direction, "2")
    return (f'<p:par><p:cTn id="{i.nxt()}" presetID="22" presetClass="entr" '
            f'presetSubtype="{subtype}" fill="hold" nodeType="withEffect">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{_vis(i, spid)}'
            f'<p:animEffect transition="in" filter="wipe({direction})"><p:cBhvr>'
            f'<p:cTn id="{i.nxt()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')


def eff_move_xy(i: Ids, spid: int, delay: int, dur: int, dx: float, dy: float,
                frm: bool = False) -> str:
    """frm=True：从 (dx,dy) 偏移滑到设计位；否则从设计位滑走。"""
    path = (f"M {dx:.4f} {dy:.4f} L 0 0" if frm else f"M 0 0 L {dx:.4f} {dy:.4f}")
    return (f'<p:par><p:cTn id="{i.nxt()}" presetID="0" presetClass="path" presetSubtype="0" '
            f'fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/>'
            f'</p:stCondLst><p:childTnLst>{_vis(i, spid)}'
            f'<p:animMotion origin="layout" path="{path}" pathEditMode="relative">'
            f'<p:cBhvr><p:cTn id="{i.nxt()}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName>'
            f'</p:attrNameLst></p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>'
            f'</p:childTnLst></p:cTn></p:par>')


def eff_spin(i: Ids, spid: int, delay: int, dur: int, deg: float) -> str:
    return (f'<p:par><p:cTn id="{i.nxt()}" fill="hold" nodeType="withEffect">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{_vis(i, spid)}'
            f'<p:animRot by="{int(deg * 60000)}"><p:cBhvr>'
            f'<p:cTn id="{i.nxt()}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animRot>'
            f'</p:childTnLst></p:cTn></p:par>')


def timing_xml(blocks: list[str]) -> bytes:
    body = "".join(blocks)
    return (f'<p:timing xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">'
            f'<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
            f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            f'<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="4" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>'
            f'</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>').encode("utf-8")


# ================================================================ 主流程
def _big_pictures(slide, sw: int, sh: int) -> list:
    out = []
    for sp in slide.shapes:
        if sp.shape_id is None or sp.width is None:
            continue
        try:
            if sp.shape_type != 13:      # MSO_SHAPE_TYPE.PICTURE
                continue
        except Exception:
            continue
        if (sp.width / sw) * (sp.height / sh) >= 0.80:
            out.append(sp)
    return out


def _txt_shapes(slide) -> list:
    out = []
    for sp in slide.shapes:
        try:
            if sp.text_frame.text.strip():
                out.append(sp)
        except Exception:
            continue
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="油画切片开场（视频同款四拍）")
    ap.add_argument("--deck", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--art", help="油画大图（缺省=取封面最大的满版底图）")
    ap.add_argument("--work", default=str(Path(tempfile.gettempdir()) / "oil_opening"),
                    help="素材缓存目录")
    ap.add_argument("--bars", type=int, default=12, help="竖切片条数")
    ap.add_argument("--cards", type=int, default=8, help="照片卡张数")
    ap.add_argument("--panel", type=float, default=G["panel_ratio"], help="面板宽度占比")
    ap.add_argument("--text-start", type=int, default=T["text_start"],
                    help="原有文字层的新起点(ms)")
    ap.add_argument("--no-cards", action="store_true", help="跳过照片卡那一拍")
    ap.add_argument("--no-ring", action="store_true", help="跳过圆环那一拍")
    ap.add_argument("--no-panel", action="store_true", help="不插玻璃面板")
    ap.add_argument("--drop", default="",
                    help="要删掉的封面形状名（逗号分隔），比如会抢镜的碎三角/照片填充大字："
                         "--drop \"shard0,shard1,bemyself\"")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    G["panel_ratio"] = args.panel
    prs = Presentation(args.deck)
    sw, sh = prs.slide_width, prs.slide_height
    sw_in, sh_in = sw / EMU_IN, sh / EMU_IN
    cover = prs.slides[0]

    # 先删掉会抢开场的装饰形状（碎三角、横穿的照片填充大字等）
    dropped = []
    if args.drop:
        for nm in [n.strip() for n in args.drop.split(",") if n.strip()]:
            for sp in list(cover.shapes):
                if sp.name == nm:
                    sp.element.getparent().remove(sp.element)
                    dropped.append(nm)
        if dropped:
            print(f"[drop] 已移除 {len(dropped)} 个形状：{', '.join(dropped)}")

    bigs = _big_pictures(cover, sw, sh)
    if args.art:
        art = Image.open(args.art).convert("RGB")
        art_src = args.art
    elif bigs:
        art = Image.open(BytesIO(max(bigs, key=lambda s: s.width * s.height).image.blob)).convert("RGB")
        art_src = "封面满版底图"
    else:
        print("ERROR: 封面找不到满版底图，用 --art 指定油画", file=sys.stderr)
        return 2

    tag = hashlib.md5((art_src + str(art.size) + str(args.bars) + str(args.cards))
                      .encode()).hexdigest()[:8]
    work = Path(args.work)
    meta = build_assets(art, work, args.bars, args.cards, tag)
    CW = meta["canvas"][0]
    px = sw / CW                     # 画布像素 → EMU
    pivot_x, pivot_y = G["pivot"][0] * sw_in, G["pivot"][1] * sh_in

    if args.dry_run:
        print(f"油画来源：{art_src}  {art.size[0]}×{art.size[1]}  画布 {CW}px")
        print(f"切片 {meta['bars'].__len__()} 条 / 单条 {meta['sw']}px / 竖缝 {meta['gap']}px")
        print(f"面板 {meta['panel_w']}px（占页宽 {meta['panel_w'] / CW:.1%}）")
        print(f"照片卡 {0 if args.no_cards else args.cards} 张 / 圆环 {'关' if args.no_ring else '开'}")
        print("时间轴：" + " → ".join(
            f"{k}={v}" for k, v in T.items() if k in
            ("card_in", "card_out", "ring_in", "ring_out", "bar_in", "panel_in", "text_start")))
        return 0

    # ---- 插入形状：切片紧跟底图（不挡文字层），面板在切片之上，卡片/圆环在最上层
    tree = cover.shapes._spTree
    children = list(tree)
    bg_idx = 0
    if bigs:
        bg_el = max(bigs, key=lambda s: s.width * s.height).element
        bg_idx = children.index(bg_el)

    bars = []
    for i, b in enumerate(meta["bars"]):
        pic = cover.shapes.add_picture(BytesIO((work / b["file"]).read_bytes()),
                                       int(b["x"] * px), 0, int(b["w"] * px) - 1, int(sh))
        pic.name = f"oilslice{i + 1}"
        el = pic.element
        tree.remove(el)
        tree.insert(bg_idx + 1 + i, el)
        bars.append({"spid": pic.shape_id, "i": i,
                     "cx_in": (b["x"] + b["w"] / 2) / CW * sw_in})

    panel = None
    if not args.no_panel:
        panel = cover.shapes.add_picture(BytesIO((work / meta["panel"]).read_bytes()),
                                         0, 0, int(meta["panel_w"] * px), int(sh))
        panel.name = "glass-panel"
        el = panel.element
        tree.remove(el)
        tree.insert(bg_idx + 1 + len(bars), el)

    cards = []
    if not args.no_cards:
        cw_in = G["card_w"] * sw_in
        ch_in = cw_in * G["card_ar"]
        for i in range(args.cards):
            pic = cover.shapes.add_picture(
                BytesIO((work / meta["cards"][i]).read_bytes()),
                int(-cw_in * 1.7 * EMU_IN), int((pivot_y - ch_in / 2 * EMU_IN)),
                int(cw_in * EMU_IN), int(ch_in * EMU_IN))
            pic.name = f"opencard{i + 1}"
            el = pic.element
            tree.remove(el)
            tree.append(el)
            slot_cx = (G["slot_x0"] + i * G["slot_step"]) * sw_in
            cards.append({"spid": pic.shape_id, "i": i,
                          "dx": (slot_cx - (-cw_in * 1.7)) / sw_in,
                          "dy": ((pivot_y - 0.012 * i * EMU_IN) - pivot_y) / sh,
                          "rot": G["card_rot0"] + i * G["card_rot_step"]})

    ring = None
    if not args.no_ring:
        r_in_w = G["ring_w"] * sw_in
        ring = cover.shapes.add_picture(
            BytesIO((work / meta["ring"]).read_bytes()),
            int(G["ring_park"] * sw_in * EMU_IN), int((pivot_y - r_in_w / 2 * EMU_IN)),
            int(r_in_w * EMU_IN), int(r_in_w * EMU_IN))
        ring.name = "opening-ring"
        el = ring.element
        tree.remove(el)
        tree.append(el)

    # ---- 时间轴
    tim = cover.element.find(f"{{{P_NS}}}timing")
    old_children: list = []
    ids = Ids()
    if tim is not None:
        max_id = max([int(x) for x in re.findall(
            r'cTn id="(\d+)"', etree.tostring(tim, encoding="unicode"))], default=0)
        ids._n = max_id
        main_par = tim.find(f".//{{{P_NS}}}cTn[@id='4']")
        if main_par is not None:
            old_children = list(main_par.find(f"{{{P_NS}}}childTnLst"))

    new_blocks: list[str] = []
    for c in cards:
        d = T["card_in"] + c["i"] * T["card_stagger"]
        new_blocks.append(eff_move_xy(ids, c["spid"], d, T["card_dur"], c["dx"], c["dy"]))
        new_blocks.append(eff_spin(ids, c["spid"], d, T["card_spin"], c["rot"]))
    for c in cards:
        d = T["card_out"] + c["i"] * T["card_out_stagger"]
        new_blocks.append(eff_spin(ids, c["spid"], d, 380, T["card_unspin"]))
        new_blocks.append(eff_move_xy(ids, c["spid"], d, T["card_out_dur"],
                                      -(c["dx"] + 0.22), -c["dy"]))
    if ring is not None:
        rdx = (pivot_x - G["ring_park"] * sw_in) / sw_in
        new_blocks.append(eff_move_xy(ids, ring.shape_id, T["ring_in"], T["ring_in_dur"], rdx, 0.0))
        new_blocks.append(eff_move_xy(ids, ring.shape_id, T["ring_out"], T["ring_out_dur"],
                                      -(rdx + 0.9), 0.32))
    for b in bars:
        d = T["bar_in"] + b["i"] * T["bar_stagger"]
        new_blocks.append(eff_fade(ids, b["spid"], d, T["bar_fade"]))
        new_blocks.append(eff_move_xy(ids, b["spid"], d, T["bar_dur"],
                                      (pivot_x / sw_in) - b["cx_in"], 0.0, frm=True))
    if panel is not None:
        new_blocks.append(eff_wipe(ids, panel.shape_id, T["panel_in"], T["panel_dur"], "left"))

    # ---- 原有动画：底图/压暗层往后挪，文字层整体后移压缩；没有 timing 就给文字补淡入
    bg_spid = max(bigs, key=lambda s: s.width * s.height).shape_id if bigs else None
    scrim_spid = [s.shape_id for s in bigs if bg_spid and s.shape_id != bg_spid]

    def _retime(el):
        ctn = el.find(f".//{{{P_NS}}}cTn")
        if ctn is None:
            return
        tgt = el.find(f".//{{{P_NS}}}spTgt")
        spid = int(tgt.get("spid")) if tgt is not None and tgt.get("spid") else None
        cond = ctn.find(f"{{{P_NS}}}stCondLst/{{{P_NS}}}cond")
        if cond is None or cond.get("delay") in (None, "indefinite"):
            return
        old = int(cond.get("delay"))
        if spid == bg_spid:
            cond.set("delay", str(T["bg_fade"]))
            for t in el.iter(f"{{{P_NS}}}cTn"):
                if t.get("dur") not in (None, "1", "indefinite"):
                    t.set("dur", str(T["bg_fade_dur"]))
        elif spid in scrim_spid:
            cond.set("delay", str(T["scrim_in"]))
        elif old >= T["shift_from"]:
            cond.set("delay", str(int(args.text_start + (old - T["shift_from"]) * T["compress"])))

    for c in old_children:
        _retime(c)

    if old_children:
        body = "".join(new_blocks) + "".join(etree.tostring(c, encoding="unicode")
                                            for c in old_children)
    else:
        extra = []
        for sp in _txt_shapes(cover):
            if sp.shape_id in {b["spid"] for b in bars}:
                continue
            extra.append(eff_fade(ids, sp.shape_id, args.text_start, 700))
        body = "".join(new_blocks) + "".join(extra)

    node = etree.fromstring(timing_xml([body]))
    if tim is not None:
        cover.element.remove(tim)
    anchor = cover.element.find(f"{{{P_NS}}}extLst")
    if anchor is not None:
        anchor.addprevious(node)
    else:
        cover.element.append(node)

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.out)
    if not args.quiet:
        print(f"[OK] {args.out}")
        print(f"     切片 {len(bars)} / 卡片 {len(cards)} / 圆环 {'开' if ring else '关'} / "
              f"面板 {'开' if panel else '关'}；原有效果 {len(old_children)} 个已后移")
        print("     下一步：WPS 前台放映抽样截图核对（scripts/verify_playback.py）")
    (work / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

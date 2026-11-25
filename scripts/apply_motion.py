#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
给静态 PPTX 注入原生动画（WPS / PowerPoint 均可播放，不用生成视频）。

用法：
    python apply_motion.py --deck 静态.pptx --out 动效.pptx
    python apply_motion.py --deck in.pptx --out out.pptx --dry-run        # 只打印分层方案
    python apply_motion.py --deck in.pptx --out out.pptx --params p.json  # 调时间轴
    python apply_motion.py --deck in.pptx --out out.pptx --cover-only

三层能力：
    1) 封面 12 秒分层开场：背景淡入推近 → 巨型照片填充字横向掠过 → 标题擦写
       → 细线/正文逐条上浮 → 照片滑入 → 碎片旋转掠过
    2) 内页条带动画：背景慢推、标题擦写、正文错峰上浮、图片缩放
    3) 每页 fade 慢速转场

硬约束（详见 references/wps-timing-xml.md）：
    - 动画只能手写 p:timing，WPS COM 的 AddEffect 会崩
    - 层级必须是 par(delay=indefinite) -> par(delay=0) -> 各效果 par(nodeType=withEffect)
    - 不支持 Morph / 3D 立方转场
"""
from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx import Presentation

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PX = 12700  # 1 px(@96dpi) = 12700 EMU

# ---------------------------------------------------------------- 时间轴默认值
DEFAULTS = {
    "cover": {
        "bg_fade": 900,
        "bg_zoom_dur": 14000,
        "bg_zoom_to": 105,
        "sweep_delay": 1300,      # 照片填充大字进场
        "sweep_dur": 4200,
        "sweep_dx": -2.45,        # 以页宽为单位，负数 = 向左横穿
        "title_delay": 2600,
        "title_dur": 900,
        "title2_delay": 3450,
        "rule_delay": 3760,
        "rule_dur": 700,
        "row_start": 3960,        # 正文行首条开始时间
        "row_stagger": 150,
        "row_dur": 620,
        "rise_px": 18,
        "photo_delay": 4520,
        "photo_dur": 900,
        "photo_slide_dx": 0.55,   # 右侧滑入距离（页宽倍数）
        "photo_scale_from": 96,
        "shard_delay": 4000,
        "shard_stagger": 90,
        "shard_dur": 2400,
        "shard_spin": 18,
        "shard_dx": 0.22,
        "caption_delay": 5600,
    },
    "page": {
        "bg_fade": 600,
        "bg_zoom_dur": 14000,
        "bg_zoom_to": 105,
        "title_delay": 180,
        "title_dur": 900,
        "row_start": 360,
        "row_stagger": 80,
        "row_dur": 560,
        "rise_px": 14,
        "pic_delay": 560,
        "pic_scale_from": 94,
        "footer_extra": 140,
    },
    "transition": "fade",
    "bg_color": "06080C",   # 满版底图淡入前露出的底板色；设成 None 即不改
}

TITLE_PT = 26           # 判定"标题"的字号阈值（pt）
BIG_WORD_RATIO = 2.2    # 图片宽高比 >= 该值、且高度小 → 视为"照片填充大字"
BG_AREA_RATIO = 0.82    # 面积占比 >= 该值 → 视为背景


# ---------------------------------------------------------------- XML 片段
class Ids:
    """动画节点 id 自增器（必须全页唯一）。"""

    def __init__(self) -> None:
        self._n = 1

    def nxt(self) -> int:
        self._n += 1
        return self._n


def _set_visible(ids: Ids, spid: int) -> str:
    i = ids.nxt()
    return (
        f'<p:set><p:cBhvr><p:cTn id="{i}" dur="1" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
    )


def eff_fade(ids: Ids, spid: int, delay: int, dur: int) -> str:
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" presetID="10" presetClass="entr" presetSubtype="0" '
        f'fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'{_set_visible(ids, spid)}'
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
        f'<p:cTn id="{ids.nxt()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'</p:cBhvr></p:animEffect>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def eff_wipe(ids: Ids, spid: int, delay: int, dur: int, direction: str = "right") -> str:
    subtype = "1" if direction == "right" else "4"
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" presetID="22" presetClass="entr" presetSubtype="{subtype}" '
        f'fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'{_set_visible(ids, spid)}'
        f'<p:animEffect transition="in" filter="wipe({direction})"><p:cBhvr>'
        f'<p:cTn id="{ids.nxt()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'</p:cBhvr></p:animEffect>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def eff_move(ids: Ids, spid: int, delay: int, dur: int, dx: float, dy: float = 0.0) -> str:
    """相对位移，单位 = 页宽 / 页高。origin=layout 必须写，否则 WPS 按形状坐标算。"""
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" presetID="0" presetClass="path" presetSubtype="0" '
        f'fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'{_set_visible(ids, spid)}'
        f'<p:animMotion origin="layout" path="M 0 0 L {dx:.4f} {dy:.4f}" pathEditMode="relative">'
        f'<p:cBhvr><p:cTn id="{ids.nxt()}" dur="{dur}" fill="hold"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def eff_scale(ids: Ids, spid: int, delay: int, dur: int, start: float, end: float) -> str:
    """start/end 为百分比，如 94 → 100。"""
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" presetID="23" presetClass="entr" presetSubtype="16" '
        f'fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'{_set_visible(ids, spid)}'
        f'<p:animScale><p:cBhvr><p:cTn id="{ids.nxt()}" dur="{dur}" fill="hold"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr>'
        f'<p:from x="{int(start * 1000)}" y="{int(start * 1000)}"/>'
        f'<p:to x="{int(end * 1000)}" y="{int(end * 1000)}"/></p:animScale>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def eff_spin(ids: Ids, spid: int, delay: int, dur: int, deg: float) -> str:
    """deg*60000（1 度 = 60000）。"""
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'{_set_visible(ids, spid)}'
        f'<p:animRot by="{int(deg * 60000)}"><p:cBhvr>'
        f'<p:cTn id="{ids.nxt()}" dur="{dur}" fill="hold"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animRot>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def eff_hide(ids: Ids, spid: int) -> str:
    """先置为不可见：WPS 对"延迟进场"的形状会提前显示，加这一层才干净。"""
    i = ids.nxt()
    return (
        f'<p:par><p:cTn id="{i}" fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{ids.nxt()}" dur="1" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def timing_xml(blocks: list[str]) -> str:
    body = "".join(blocks)
    return (
        f'<p:timing xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">'
        f'<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f'<p:par><p:cTn id="3" fill="hold"><p:stCondLst>'
        f'<p:cond delay="indefinite"/>'
        f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
        f'</p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>'
        f'</p:childTnLst></p:cTn></p:par>'
        f'</p:childTnLst></p:cTn></p:seq></p:childTnLst>'
        f'</p:cTn></p:par></p:tnLst></p:timing>'
    )


# ---------------------------------------------------------------- 形状探查
def _max_pt(shape) -> float:
    """组内最大字号（pt）。无文字返回 0。"""
    if shape.shape_type == 6 or hasattr(shape, "shapes"):
        try:
            return max([_max_pt(s) for s in shape.shapes] + [0])
        except Exception:
            return 0
    if not getattr(shape, "has_text_frame", False):
        return 0
    best = 0.0
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size:
                best = max(best, run.font.size.pt)
    return best


def _text(shape) -> str:
    """文本框文字；组合形状递归取子形状文字（SVG 导出的文字常被包进 <g>）。"""
    try:
        return shape.text_frame.text.strip()
    except Exception:
        pass
    try:
        parts = [_text(s) for s in shape.shapes]
    except Exception:
        return ""
    return " / ".join(p for p in parts if p)


def _box(shape, sw: int, sh: int):
    return (
        shape.left / sw,
        shape.top / sh,
        shape.width / sw,
        shape.height / sh,
    )


def _is_picture(shape) -> bool:
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE

        return shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    except Exception:
        return False


def _is_freeform(shape) -> bool:
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE

        return shape.shape_type == MSO_SHAPE_TYPE.FREEFORM
    except Exception:
        return False


def _overlap(a: dict, b: dict) -> float:
    """相对较小者的重叠比例（0~1）。"""
    dx = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
    dy = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    if dx <= 0 or dy <= 0:
        return 0.0
    inter = dx * dy
    return inter / max(min(a["area"], b["area"]), 1e-9)


def plan_cover(slide, sw: int, sh: int, cfg: dict) -> list[dict]:
    """把封面切成有语义的层，返回带时间轴的执行计划。"""
    items = []
    for shape in slide.shapes:
        if shape.shape_id is None:
            continue
        x, y, w, h = _box(shape, sw, sh)
        txt = _text(shape)
        pt = _max_pt(shape)
        if pt == 0 and txt:
            # 字号没显式写入（继承自版式/主题）时，按框高估一个，保证仍能分出标题层级
            pt = round((h * sh / 914400) * 72 * 0.78, 1)
        items.append(
            {
                "spid": shape.shape_id,
                "name": shape.name,
                "x": x, "y": y, "w": w, "h": h,
                "area": w * h,
                "pt": pt,
                "text": txt,
                "pic": _is_picture(shape),
                "free": _is_freeform(shape),
            }
        )

    bg = next((i for i in items if i["area"] >= BG_AREA_RATIO), None)
    rest = [i for i in items if bg is None or i["spid"] != bg["spid"]]

    # 与背景一样满版的另一张图 = 蒙版/压暗层，只淡入，不推近
    scrims = [i for i in rest if i["pic"] and i["area"] >= BG_AREA_RATIO]
    rest_pics = [i for i in rest if i["pic"] and i["area"] < BG_AREA_RATIO]

    word = next(
        (i for i in rest_pics if i["h"] < 0.55 and i["w"] / max(i["h"], 1e-6) >= BIG_WORD_RATIO),
        None,
    )
    pool = [p for p in rest_pics if word is None or p["spid"] != word["spid"]]
    pool = [p for p in pool if 0.3 <= p["w"] / max(p["h"], 1e-6) <= 3.2] or pool
    photo = max(pool, key=lambda p: p["area"], default=None)

    texts = [i for i in rest if not i["pic"] and i["pt"] > 0 and i["text"]]
    rules = [
        i for i in rest
        if not i["pic"] and i["pt"] == 0 and i["h"] < 0.04 and i["area"] < 0.2
    ]
    rule_ids = {i["spid"] for i in rules}
    deco = [
        i for i in rest
        if not i["pic"] and i["pt"] == 0 and i["spid"] not in rule_ids
        and 0.06 <= i["area"] < BG_AREA_RATIO
    ]
    shards = [
        i for i in rest
        if not i["pic"] and i["pt"] == 0 and i["spid"] not in rule_ids
        and 0 < i["area"] < 0.06
    ]

    title = max(texts, key=lambda i: i["pt"], default=None)
    title2 = None
    if title is not None:
        others = [i for i in texts if i["spid"] != title["spid"]]
        close = [i for i in others if i["pt"] >= max(title["pt"] - 12, 16)]
        title2 = max(close, key=lambda i: i["pt"], default=None)
    rows = [
        i for i in texts
        if i["spid"] not in {title["spid"] if title else -1, title2["spid"] if title2 else -1}
    ]
    rule = min(rules, key=lambda i: abs(i["y"] - (title["y"] if title else 0.4)), default=None)

    plan = []
    if bg:
        plan.append({"spid": bg["spid"], "kind": "bg", "label": "背景"})
    for s in scrims:
        plan.append({"spid": s["spid"], "kind": "scrim", "label": "压暗层"})
    if word:
        plan.append({"spid": word["spid"], "kind": "word", "label": "照片填充大字"})
    if title:
        plan.append({"spid": title["spid"], "kind": "title", "label": f"标题「{title['text'][:12]}」"})
    if title2:
        plan.append({"spid": title2["spid"], "kind": "title2", "label": f"副标题「{title2['text'][:12]}」"})
    if rule:
        plan.append({"spid": rule["spid"], "kind": "rule", "label": "细线"})
    rows.sort(key=lambda i: (round(i["y"], 2), i["x"]))
    for r in rows:
        plan.append({"spid": r["spid"], "kind": "row", "label": f"正文「{r['text'][:14]}」"})
    if photo:
        plan.append({"spid": photo["spid"], "kind": "photo", "label": "照片"})
    for d in sorted(deco, key=lambda i: i["area"], reverse=True):
        if photo is not None and _overlap(d, photo) >= 0.7:
            plan.append({"spid": d["spid"], "kind": "photo_frame", "label": f"照片边框 {d['name']}"})
        else:
            plan.append({"spid": d["spid"], "kind": "deco", "label": f"装饰面 {d['name']}"})
    shards.sort(key=lambda i: i["y"])
    for s in shards:
        plan.append({"spid": s["spid"], "kind": "shard", "label": f"面片 {s['name']}"})
    return plan


def shift_offstage(slide, plan: list[dict], cfg: dict, sw: int) -> None:
    """
    WPS 会把"延迟进场"的图片提前显示（文本框不会）。所以把照片填充大字和人像
    先挪到画布右侧外面，再由路径动画滑回原位——放映时看不到提前曝光，
    编辑视图里它们在页面右侧外面。

    大字按 sweep_dx 的量挪出去，正好让"横穿"从画外真正进入、最后停在设计位置；
    人像按 photo_slide_dx 挪出去，滑回后同样落在设计位置。
    """
    word_dx = abs(cfg["cover"].get("sweep_dx", 0))
    photo_dx = cfg["cover"].get("photo_slide_dx", 0)
    if word_dx <= 0 and photo_dx <= 0:
        return
    by_id = {sp.shape_id: sp for sp in slide.shapes}
    for it in plan:
        sp = by_id.get(it["spid"])
        if sp is None:
            continue
        if it["kind"] == "word" and word_dx > 0:
            sp.left = int(sp.left + word_dx * sw)
        elif it["kind"] in ("photo", "photo_frame") and photo_dx > 0:
            sp.left = int(sp.left + photo_dx * sw)


def build_cover(plan: list[dict], cfg: dict, sh: int) -> list[str]:
    c = cfg["cover"]
    ids = Ids()
    blocks: list[str] = []
    row_i = 0
    shard_i = 0
    for item in plan:
        k = item["kind"]
        spid = item["spid"]
        if k not in ("bg", "scrim"):
            blocks.append(eff_hide(ids, spid))
        if k == "bg":
            blocks.append(eff_fade(ids, spid, 0, c["bg_fade"]))
            blocks.append(eff_scale(ids, spid, 0, c["bg_zoom_dur"], 100, c["bg_zoom_to"]))
        elif k == "word":
            blocks.append(eff_fade(ids, spid, 220, 700))
            blocks.append(eff_move(ids, spid, c["sweep_delay"], c["sweep_dur"], c["sweep_dx"], 0))
        elif k in ("title", "title2"):
            delay = c["title_delay"] if k == "title" else c["title2_delay"]
            blocks.append(eff_wipe(ids, spid, delay, c["title_dur"], "right"))
        elif k == "rule":
            blocks.append(eff_wipe(ids, spid, c["rule_delay"], c["rule_dur"], "right"))
        elif k == "row":
            d = c["row_start"] + row_i * c["row_stagger"]
            blocks.append(eff_fade(ids, spid, d, c["row_dur"]))
            blocks.append(
                eff_move(ids, spid, d, c["row_dur"], 0, -(c["rise_px"] * PX) / sh)
            )
            row_i += 1
        elif k == "photo":
            dx = -c["photo_slide_dx"]
            blocks.append(
                eff_move(ids, spid, c["photo_delay"], c["photo_dur"], dx, 0)
            )
            blocks.append(eff_fade(ids, spid, c["photo_delay"], c["photo_dur"]))
            blocks.append(
                eff_scale(ids, spid, c["photo_delay"], c["photo_dur"],
                          c["photo_scale_from"], 100)
            )
        elif k == "shard":
            d = c["shard_delay"] + shard_i * c["shard_stagger"]
            blocks.append(eff_fade(ids, spid, d, 800))
            blocks.append(eff_spin(ids, spid, d, c["shard_dur"], c["shard_spin"]))
            blocks.append(eff_move(ids, spid, d, c["shard_dur"], c["shard_dx"], 0))
            shard_i += 1
        elif k == "scrim":
            blocks.append(eff_fade(ids, spid, 120, 600))
        elif k == "photo_frame":
            d = c["photo_delay"]
            blocks.append(eff_fade(ids, spid, d, c["photo_dur"]))
            blocks.append(eff_move(ids, spid, d, c["photo_dur"], -c["photo_slide_dx"], 0))
        elif k == "deco":
            d = c["shard_delay"] - 300
            blocks.append(eff_fade(ids, spid, d, 900))
            blocks.append(eff_spin(ids, spid, d, c["shard_dur"] + 400, c["shard_spin"]))
    return blocks


def build_page(slide, sw: int, sh: int, cfg: dict) -> list[str]:
    p = cfg["page"]
    ids = Ids()
    blocks: list[str] = []
    row_i = 0
    items = []
    for shape in slide.shapes:
        x, y, w, h = _box(shape, sw, sh)
        items.append({"spid": shape.shape_id, "shape": shape, "x": x, "y": y, "w": w, "h": h,
                      "area": w * h, "pt": _max_pt(shape), "pic": _is_picture(shape)})
    if not items:
        return blocks
    bg = max(items, key=lambda i: (i["area"], -items.index(i)))
    is_bg = bg["area"] >= BG_AREA_RATIO
    for it in items:
        if is_bg and it["spid"] == bg["spid"]:
            blocks.append(eff_fade(ids, it["spid"], 0, p["bg_fade"]))
            blocks.append(eff_scale(ids, it["spid"], 0, p["bg_zoom_dur"], 100, p["bg_zoom_to"]))
            continue
        if it["area"] >= BG_AREA_RATIO:      # 满版压暗层：只淡入，不跟着推近
            blocks.append(eff_fade(ids, it["spid"], 120, p["bg_fade"]))
            continue
        if it["pic"]:
            blocks.append(eff_hide(ids, it["spid"]))
            blocks.append(eff_fade(ids, it["spid"], p["pic_delay"], 700))
            blocks.append(eff_scale(ids, it["spid"], p["pic_delay"], 800, p["pic_scale_from"], 100))
            continue
        if it["pt"] >= TITLE_PT:
            blocks.append(eff_hide(ids, it["spid"]))
            blocks.append(eff_wipe(ids, it["spid"], p["title_delay"], p["title_dur"], "right"))
            continue
        if it["y"] > 0.86:
            blocks.append(eff_hide(ids, it["spid"]))
            blocks.append(eff_fade(ids, it["spid"], p["row_start"] + 600 + p["footer_extra"], 500))
            continue
        d = p["row_start"] + row_i * p["row_stagger"]
        blocks.append(eff_hide(ids, it["spid"]))
        blocks.append(eff_fade(ids, it["spid"], d, p["row_dur"]))
        blocks.append(eff_move(ids, it["spid"], d, p["row_dur"], 0, -p["rise_px"] / sh))
        row_i += 1
    return blocks


# ---------------------------------------------------------------- 写入
def strip_timing(slide) -> None:
    sld = slide.element
    for tm in sld.findall(f"{{{P_NS}}}timing"):
        sld.remove(tm)


def apply_timing(slide, blocks: list[str]) -> int:
    if not blocks:
        return 0
    xml = timing_xml(blocks)
    node = etree.fromstring(xml.encode("utf-8"))
    sld = slide.element
    strip_timing(slide)
    # p:timing 必须位于 cSld/clrMapOvr 之后、extLst 之前
    anchor = sld.find(f"{{{P_NS}}}extLst")
    if anchor is not None:
        anchor.addprevious(node)
    else:
        sld.append(node)
    return len(blocks)


def set_transition(slide, kind: str = "fade") -> None:
    sld = slide.element
    for tr in sld.findall(f"{{{P_NS}}}transition"):
        sld.remove(tr)
    if kind != "fade":
        return
    node = etree.fromstring(
        f'<p:transition xmlns:p="{P_NS}" xmlns:a="{A_NS}" spd="slow"><p:fade/></p:transition>'.encode()
    )
    # transition 在 cSld/clrMapOvr 之后、timing 之前
    tm = sld.find(f"{{{P_NS}}}timing")
    if tm is not None:
        tm.addprevious(node)
    else:
        sld.append(node)


def set_bg_color(slide, hex_color: str) -> None:
    """满版底图淡入时会先露出白色底板，把页面底色改深可消掉白闪。"""
    try:
        from pptx.dml.color import RGBColor

        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(hex_color)
    except Exception:
        pass


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="给静态 PPTX 注入原生动画")
    ap.add_argument("--deck", required=True, help="静态 PPTX")
    ap.add_argument("--out", required=True, help="输出 PPTX")
    ap.add_argument("--params", help="JSON 参数文件（覆盖 DEFAULTS）")
    ap.add_argument("--cover-only", action="store_true", help="只处理第 1 页")
    ap.add_argument("--skip-cover", action="store_true", help="跳过第 1 页")
    ap.add_argument("--dry-run", action="store_true", help="只打印分层方案")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    cfg = deepcopy(DEFAULTS)
    if args.params:
        loaded = json.loads(Path(args.params).read_text(encoding="utf-8"))
        for k, v in loaded.items():
            if isinstance(v, dict):
                cfg.setdefault(k, {}).update(v)
            else:
                cfg[k] = v

    prs = Presentation(args.deck)
    sw, sh = prs.slide_width, prs.slide_height
    if len(prs.slides) == 0:
        print("ERROR: deck has no slide", file=sys.stderr)
        return 2

    cover = prs.slides[0]
    cover_plan = plan_cover(cover, sw, sh, cfg) if not args.skip_cover else []
    if args.dry_run:
        kinds = {it["kind"] for it in cover_plan}
        print(f"slide 1 分层（共 {len(cover_plan)} 层）：")
        for i, it in enumerate(cover_plan, 1):
            print(f"  {i:>2}. {it['kind']:<7} spid={it['spid']:<3} {it['label']}")
        missing = [k for k in ("bg", "word", "photo") if k not in kinds]
        if missing:
            print(
                f"  提示：封面缺少 {'/'.join(missing)} 层，12 秒开场只会退化为文字+装饰的时间轴；"
                "想要“雪山推近 + 照片字横穿 + 人像滑入”，先让版式带上满版底图、照片填充大字与人像。"
            )
        for idx, slide in enumerate(prs.slides, 1):
            if idx == 1:
                continue
            print(f"slide {idx}: {len(slide.shapes)} 个顶层形状（按字号/位置分条带）")
        return 0

    total = 0
    if cover_plan:
        shift_offstage(cover, cover_plan, cfg, sw)
        total += apply_timing(cover, build_cover(cover_plan, cfg, sh))

    for idx, slide in enumerate(prs.slides, 1):
        if idx == 1 or args.cover_only:
            continue
        total += apply_timing(slide, build_page(slide, sw, sh, cfg))

    for slide in prs.slides:
        set_transition(slide, cfg.get("transition", "fade"))
        if cfg.get("bg_color"):
            big = any(
                (sp.width / sw) * (sp.height / sh) >= BG_AREA_RATIO for sp in slide.shapes
            )
            if big:
                set_bg_color(slide, cfg["bg_color"])

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.out)
    if not args.quiet:
        print(f"[OK] {args.out}")
        print(f"     封面 {len(cover_plan)} 层 / 共 {len(prs.slides)} 页 / {total} 个动画组")
        print("     下一步：用 WPS 前台放映抽样截图确认（scripts/verify_playback.py）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

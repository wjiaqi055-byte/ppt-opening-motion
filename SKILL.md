---
name: ppt-opening-motion
description: 给静态 PPT 加"视频级"原生开场动效（WPS/PPT 里真播，不用生成视频）。用于用户说"PPT 加动画""做个高级开场""封面要有动态""像抖音那种 PPT 动效""PPT 开头太死板""给答辩 PPT 加动效"。核心是封面 12 秒分层开场：背景淡入推近 → 巨型照片填充字横穿画面 → 标题擦写 → 正文逐条上浮 → 人像滑入 → 玻璃面片旋转掠过；另含全册条带式动画、逐页 fade 转场。不负责从零做版面，也不做 Morph 平滑切换与 3D 立方转场（WPS 不支持）。
license: MIT
---

# PPT 原生开场动效

把已经排好版的静态 pptx，注入**手写 DrawingML 动画**，让封面像参考视频那样"有镜头感"，且文件仍是可编辑的 .pptx（不是 mp4，不是嵌视频）。

## 何时用

- 用户已有 PPT（自己做的、或 svg→pptx 出来的），嫌"太死板/没动态/不如视频"。
- 作业、答辩、路演、比赛要"高级开场"但必须现场在 WPS/PowerPoint 里播放。
- 需要"每页都有条带动画"，而不是一键淡入。

不适用：从零排版（那是版面工具的事）；要求 Morph 平滑切换、3D 立方翻页（WPS 实测做不出来，用 fade 转场替代）。

## 四步跑完

```powershell
# 先用装了 python-pptx 的 Python（哪个都行，见下方"环境要求"）
$PY='<python.exe>'
$S='<本 skill 目录>\scripts'

# 1) 先看分层方案（不改文件，只打印每层被判成什么）
& $PY "$S\apply_motion.py" --deck 静态.pptx --out 动效.pptx --dry-run

# 2) 生成动效版
& $PY "$S\apply_motion.py" --deck 静态.pptx --out 动效.pptx

# 3) 验证真的在播（WPS 前台放映 + 抽样截图拼贴）
& $PY "$S\verify_playback.py" --pptx 动效.pptx --frames 16 --interval 0.25 --out 封面抽查.jpg

# 4) 看拼贴图：同一页出现不同透明度/位移/旋转 = 动画生效；张张一样 = timing 没写进去
```

`verify_playback.py` 会自动找 WPS；找不到（比如只装了 PowerPoint）就用 `--wpp <wpp.exe>` 指定，或改用 PowerPoint 手动放映看一遍。

需要逐帧看细节时：`--goto 3` 直接跳页采样，`--next 1` 翻一页再采。

## 环境要求

- Windows + WPS（PowerPoint 也能播，但 WPS 的坑更多，脚本按 WPS 调的）。
- Python 3.9+，`pip install -r requirements.txt`（python-pptx、lxml、Pillow、pywin32）。
- `make_photo_text.py` 需要本机装有 Chrome；想用的字体要在系统里存在（`--font` 传字体名）。
- 脚本本身不含任何绝对路径，换机器只需确认 python 与 WPS 位置。

## 分层的判定规则

`apply_motion.py` 按几何与字号自动判层（`--dry-run` 可核对，误判就用 `--params` 改参数或调形状名再跑）：

| 层 | 判定 | 动效 |
|---|---|---|
| `bg` | 面积 ≥ 82% 页面的第一张图 | 淡入 0.9s + 14s 推到 105% |
| `scrim` | 与 bg 一样满版的第二张图（压暗层） | 只淡入，不跟着推近 |
| `word` | 宽高比 ≥ 2.2 且高度 < 55% 的图片（照片填充大字） | 淡入 + 从左/右横穿（`sweep_dx` 页宽倍数） |
| `title` / `title2` | 字号最大 / 次大的文本框 | 擦写（wipe right） |
| `rule` | 无文字、高度 < 4% 的细线 | 擦写 |
| `row` | 其余有字文本框 | 淡入 + 上浮 18px，错峰 150ms |
| `photo` | 面积 < 50% 且宽高比 0.3~3.2 的最大图片 | 淡入 + 从画布外滑入 + 缩放 96→100 |
| `photo_frame` | 与人像框重叠 ≥ 70% 的装饰形状 | 跟人像同步滑入（不旋转） |
| `deco` | 面积 ≥ 6% 的无文字形状（菱形等） | 淡入 + 旋转 |
| `shard` | 面积 < 6% 的无文字形状（三角形面片） | 淡入 + 旋转 18° + 侧移 0.22 页宽，错峰 90ms |

内页（第 2 页起）：满版底图淡入+慢推、字号 ≥ 26pt 的标题擦写、图片缩放、其余按纵向位置错峰上浮、页脚最后 0.6s 淡入。

## 浅色模板（没有满版底图）

`--dry-run` 会直接告诉你封面缺 `bg` / `word` / `photo`。缺这几层时**不要用默认 12 秒时间轴**，否则前 2.5 秒是一张空页。改用收紧版：

```json
{"bg_color": null,
 "cover": {"shard_delay": 600, "shard_dur": 1400, "shard_dx": 0.08,
           "title_delay": 1150, "title_dur": 650, "title2_delay": 1800,
           "rule_delay": 2000, "row_start": 2200, "row_stagger": 110, "row_dur": 700},
 "page": {"row_stagger": 70}}
```

要点：`bg_color: null` 必须给（浅色页面被改成深底会整页发黑）；装饰面提前到 0.3 秒先动，标题 1.1 秒擦出，整页 4 秒内建完。想让开场真正像参考视频，还是得先让版式带上满版底图、照片填充大字和人像——这个 Skill 不替你补素材。

## 时间轴默认值（`--params` 可覆盖）

封面：`bg_fade 900` → `sweep_delay 1300 / sweep_dur 4200 / sweep_dx -2.45` → `title_delay 2600` → `title2_delay 3450` → `rule_delay 3760` → `row_start 3960 / row_stagger 150` → `photo_delay 4520 / photo_slide_dx 0.55` → `shard_delay 4000 / shard_stagger 90`。

调参示例（`p.json`）：

```json
{"cover": {"sweep_dx": -2.0, "sweep_dur": 3600, "sweep_delay": 900, "photo_delay": 5100},
 "page": {"row_stagger": 60},
 "bg_color": "06080C"}
```

## 硬约束（照做，别试别的路）

1. **只能手写 `p:timing`**：WPS COM 的 `AddEffect` 会把 WPS 弄崩（返回 0 后 RPC 失败），Morph 与 `p14:cube` 写入后 `EntryEffect` 读回 0。
2. **层级必须三层**：`par(delay=indefinite) → par(delay=0) → 每个效果 par(nodeType="withEffect")`，少一层 WPS 就读不到。
3. **图片会提前显示**：WPS 对"延迟进场"的图片不隐藏（文本框正常）。所以人像要先挪到画布外、再由路径滑回；`photo_slide_dx` 就是挪出去的页宽倍数。
4. **满版底图淡入会先露白底**：用 `bg_color` 把页面底色改深（默认 `06080C`，设 `null` 关闭）。
5. **`alpha` 是"不透明度"**：`100000` = 完全不透明。想半透明写 `alpha val=22000`，不是 78。
6. 做完必须抽样截图验收，不要凭"应该能播"交付。

## 配套脚本

- `scripts/apply_motion.py`：核心，静态 pptx → 动效 pptx。`--cover-only` 只加封面，`--skip-cover` 只加内页，`--dry-run` 看方案。
- `scripts/make_photo_text.py`：把文字做成"照片填充 + 透明底" PNG（`background-clip:text` + Chrome headless），这是封面巨型字的关键素材。示例：
  `--text "BE MYSELF" --fill 雪山.jpg --size 240 --spacing 8 --out word.png`
- `scripts/verify_playback.py`：WPS 放映 + 定时截屏拼贴验收（`--goto N` 跳页，`--next N` 翻页）。

细节与排错见 `references/cover-storyboard.md`（分镜与参数）、`references/wps-timing-xml.md`（XML 结构与坑）、`references/visual-system.md`（深色雪山版式口径）。

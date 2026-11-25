# ppt-opening-motion

给**已经排好版**的 PPT 注入"视频级"原生开场动效 —— 在 WPS / PowerPoint 里真播，文件仍是可编辑的 `.pptx`，不是 mp4、不是嵌视频。

封面 12 秒分层开场：黑场 → 满版底图淡入并缓慢推近 → 巨型**照片填充大字**从画外横穿画面 → 中文标题擦写 → 细线擦出 → 正文逐条上浮 → 玻璃面片旋转掠过。内页：底图慢推、标题擦写、内容按阅读顺序错峰浮起、图片缩放、页脚最后进场；换页统一 fade 转场。

> 做不到的部分说清楚：WPS 没有 PowerPoint 365 的 Morph（平滑）和 3D 立方转场，所以换页只能淡入，也没有真三维空间里的翻飞面片（用平面玻璃三角 + 旋转 + 位移做出同类的视觉动作）。

## 效果一览

| 层 | 判据 | 动效 |
|---|---|---|
| `bg` | 面积 ≥ 82% 页面的第一张图 | 淡入 0.9s + 14s 推到 105% |
| `scrim` | 和 bg 一样满版的第二张图（压暗层） | 只淡入，不跟着推近 |
| `word` | 宽高比 ≥ 2.2 且高度 < 55% 的图（照片填充大字） | 置画外 → 横穿进画面，停在设计位置 |
| `title` / `title2` | 字号最大 / 次大的文本框 | 擦写（wipe right） |
| `rule` | 无文字、高度 < 4% 的细线 | 擦写 |
| `row` | 其余有字文本框 | 淡入 + 上浮 18px，错峰 150ms |
| `photo` | 面积 < 50%、宽高比 0.3~3.2 的最大图（人像） | 淡入 + 从画外滑入 + 缩放 96→100 |
| `deco` | 面积 6%–82% 的无字形状 | 淡入 + 旋转 |
| `shard` | 面积 < 6% 的无字形状（三角面片） | 淡入 + 旋转 18° + 侧移，错峰 90ms |

判层是自动的，`--dry-run` 会把每层判成什么、缺哪层直接打印出来。

## 环境要求

- Windows（脚本用 WPS COM 做放映验收；PowerPoint 能播，但验收脚本是按 WPS 写的）
- Python 3.9+
- `pip install -r requirements.txt`（python-pptx / lxml / Pillow / pywin32）
- `make_photo_text.py` 需要本机装有 **Chrome**（用 headless 渲染 `background-clip:text`）
- 想用的字体要装在系统里（例如 `Impact`、`Times New Roman`）

## 安装：当成 agent 的 Skill 用

把这个文件夹整个放进你的 agent 的 skill 目录（不同客户端路径不同，例如）：

```
<agent skill 目录>/ppt-opening-motion/
├── SKILL.md            ← agent 读这个决定怎么用
├── requirements.txt
├── assets/motion.json     默认时间轴
├── references/            分镜、XML 结构、视觉口径
└── scripts/
    ├── apply_motion.py       静态 pptx → 动效 pptx
    ├── make_photo_text.py    文字 → 照片填充 + 透明底 PNG
    └── verify_playback.py    WPS 放映抽样截图验收
```

放好后直接说"给这个 PPT 加开场动效"，agent 会读 `SKILL.md` 并按里面的四步跑。

## 手动用（不装 skill，直接跑命令）

```powershell
$PY='C:\...\python.exe'          # 装了 python-pptx 的那个
$S='<本仓库>\scripts'

# 1) 先看分层方案（不动文件）
& $PY "$S\apply_motion.py" --deck 静态.pptx --out 动效.pptx --dry-run

# 2) 生成动效版
& $PY "$S\apply_motion.py" --deck 静态.pptx --out 动效.pptx

# 3) 调时间轴（可选）
& $PY "$S\apply_motion.py" --deck 静态.pptx --out 动效.pptx --params p.json

# 4) 验收：WPS 前台放映，按间隔截图拼成一张
& $PY "$S\verify_playback.py" --pptx 动效.pptx --frames 16 --interval 0.3 --out 抽查.jpg
```

拼贴里同一页出现不同透明度 / 位移 / 旋转 = 动画真的在播；张张一样 = timing 没写进去。

### 做封面那张"照片填充大字"

```powershell
& $PY "$S\make_photo_text.py" --text "BE MYSELF" --fill 雪山.jpg --out word.png `
    --font "Impact" --size 240 --spacing 8 --fill-zoom 1.0
```

要点：填充图**取的是山脊带**（不是整张含大半天光空的图），字里才会透出山的明暗；填充图宽高比接近字的框（约 5:1）时最贴合。生成的 PNG 直接当图片放进封面，`apply_motion.py` 会按宽高比自动认成 `word` 层。

## 时间轴默认值（`--params` 覆盖）

开封面：`bg_fade 900` → `sweep_delay 1300 / sweep_dur 4200 / sweep_dx -2.45` → `title_delay 2600` → `title2_delay 3450` → `rule_delay 3760` → `row_start 3960 / stagger 150` → `photo_delay 4520` → `shard_delay 4000 / stagger 90`。

`p.json` 写法（只写要改的键，其余继承默认值）：

```json
{"cover": {"sweep_dx": -2.05, "sweep_dur": 3800, "title_delay": 4200},
 "page": {"row_stagger": 70},
 "bg_color": "06080C"}
```

- `sweep_dx` 是"横穿距离"，单位是页宽倍数；脚本会先把大字移到画外右侧同样距离，再由路径动画滑回设计位置 —— 所以它真的是从画外进来的。
- `bg_color` 给页面上底色（默认深色 `06080C`），消掉满版底图淡入时的白闪。**浅色模板必须写 `null`**，否则整页发黑。

## 浅色模板怎么办

`--dry-run` 会提示封面缺 `bg` / `word` / `photo`。这种情况别用默认 12 秒时间轴（前 2.5 秒是空页），改用收紧版：

```json
{"bg_color": null,
 "cover": {"shard_delay": 600, "shard_dur": 1400, "shard_dx": 0.08,
           "title_delay": 1150, "title_dur": 650, "title2_delay": 1800,
           "rule_delay": 2000, "row_start": 2200, "row_stagger": 110, "row_dur": 700},
 "page": {"row_stagger": 70}}
```

想让开场真像参考视频，前提是版式本身有满版底图 + 照片填充大字（有人像更好）—— 这个 Skill 只负责注动效，不替你补素材。

## 已知边界（都是实测过的）

1. 动画只能手写 `p:timing`：WPS COM 的 `AddEffect` 会把 WPS 弄崩。
2. 层级必须三层：`par(delay=indefinite) → par(delay=0) → 每个效果 par(withEffect)`，少一层 WPS 读不到。
3. WPS 会把"延迟进场"的图片提前显示（文本框正常），所以大字和人像要先移到画布外再由路径滑回；编辑视图里它们待在页面右侧外面，属正常。
4. `alpha` 是**不透明度**：`alphaModFix amt="0"` 会让图片永久不可见。
5. 不支持 Morph / `p14:cube`（写入后 `EntryEffect` 读回 0）；不支持 3D 面片，用平面玻璃三角代替。
6. 交付前必须抽样截图验收，别凭"应该能播"交付。

## 仓库结构

```
ppt-opening-motion/
├── SKILL.md                     agent 的说明书（何时用、四步流程、硬约束）
├── README.md                    给人看的（本文件）
├── LICENSE                      MIT
├── requirements.txt
├── assets/motion.json           默认时间轴
├── references/
│   ├── cover-storyboard.md      开场分镜与参数含义
│   ├── wps-timing-xml.md        timing XML 结构与坑
│   └── visual-system.md         深色雪山版式的色板与字阶口径
└── scripts/
    ├── apply_motion.py
    ├── make_photo_text.py
    └── verify_playback.py
```

## License

MIT，见 `LICENSE`。

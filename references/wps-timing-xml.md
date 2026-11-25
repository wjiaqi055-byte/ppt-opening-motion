# 手写 p:timing：结构与坑（WPS 实测）

WPS 只认 PowerPoint 原生动画结构，且对层级极其敏感。下面这套结构是在 WPS 12.1 上反复实测能播的最小骨架，`scripts/apply_motion.py` 就是按它生成的。

## 1. 必须的三层嵌套

```
p:timing
└ p:tnLst
  └ p:par
    └ p:cTn id=1 dur="indefinite" restart="never" nodeType="tmRoot"
      └ p:childTnLst
        └ p:seq concurrent="1" nextAc="seek"
          └ p:cTn id=2 dur="indefinite" nodeType="mainSeq"
            └ p:childTnLst
              └ p:par
                └ p:cTn id=3 fill="hold"                      ← 层 A：等整页开始
                  └ p:stCondLst: cond delay="indefinite" + cond evt="onBegin" delay=0 / tn val=2
                  └ p:childTnLst
                    └ p:par
                      └ p:cTn id=4 fill="hold"                ← 层 B：0 时刻分组
                        └ p:stCondLst: cond delay="0"
                        └ p:childTnLst
                          └ 各效果 p:par（nodeType="withEffect"，各自 cond delay=ms）
```

少了层 A（`delay="indefinite"` + `onBegin`）WPS 会把整页当"点击触发"，自动播放失效；少了层 B 则所有延迟算错。

## 2. 五类效果块

每个效果都是 `<p:par><p:cTn ...><p:stCondLst><p:cond delay="N"/></p:stCondLst><p:childTnLst> 行为 </p:childTnLst></p:cTn></p:par>`，差别只在 `presetID` 与行为元素。

**淡入（fade，presetID=10）**

```xml
<p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
  <p:childTnLst>
    <p:set>… attrName=style.visibility … <p:to><p:strVal val="visible"/></p:to></p:set>
    <p:animEffect transition="in" filter="fade">
      <p:cBhvr><p:cTn id="7" dur="620"/><p:tgtEl><p:spTgt spid="12"/></p:tgtEl></p:cBhvr>
    </p:animEffect>
  </p:childTnLst>
</p:cTn>
```

**擦写（wipe，presetID=22）**：`presetSubtype="1"`（right）或 `"4"`（up），`<p:animEffect transition="in" filter="wipe(right)">`，`p:cTn` 不带 `fill="hold"`。

**位移（animMotion，presetID=0 / presetClass="path"）**

```xml
<p:animMotion origin="layout" path="M 0 0 L -2.4500 0.0000" pathEditMode="relative">
  <p:cBhvr><p:cTn id="16" dur="4200" fill="hold"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
    <p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>
  </p:cBhvr><p:rCtr x="0" y="0"/>
</p:animMotion>
```

- `origin="layout"` 必须有，否则路径按形状自身坐标算，位移量会离谱。
- `L dx dy` 是**相对位移**，单位 = 页宽 / 页高。`-2.45` 就是往左移 2.45 个页宽。
- `<p:rCtr x="0" y="0"/>` 是旋转中心，缺了 WPS 偶尔不认。

**缩放（animScale，presetID=23）**：`<p:animScale>` 里 `<p:from x="96000" y="96000"/><p:to x="100000" y="100000"/>`，单位是"千分之一百分比"，100000 = 100%。

**旋转（animRot）**：无 presetID，`<p:animRot by="1080000">`，1 度 = 60000，1080000 = 18°。

## 3. WPS 特有的六个坑

1. **COM `AddEffect` 会崩**：返回 count=0 之后 RPC 直接失败，动画也丢。只能改 XML。
2. **Morph / 3D 立方做不出来**：写入 `p14:cube` 后 `SlideShowTransition.EntryEffect` 读回 0，放映无效果。用 `<p:transition spd="slow"><p:fade/></p:transition>`。
3. **延迟进场的图片会提前显示**（文本框不会）：所以人像必须先挪到画布外，再用路径滑回；`apply_motion.py` 的 `photo_slide_dx` 干这事。曾试过 `alphaModFix amt="0"` 让图片透明——结果整张图再也不出现，fade 不会把它恢复。
4. **满版图淡入先露白**：PPT 默认白底板。把页面底色改成 `06080C` 就没有白闪。
5. **`a:alpha val` 是"不透明度"**：`100000` 全不透明，`22000` 才是 22%。
6. **`p:timing` 位置**：必须在 `p:cSld` / `p:clrMapOvr` 之后、`p:extLst` 之前；`p:transition` 又在 `p:timing` 之前。

## 4. 老坑（做版面阶段）

- 文本框坐标：`Emu(top * 12700)`，忘了乘 12700 会让所有文本 top=0。
- WPS 导 PDF 时中文字体常被替换成宋体，**不能拿 PDF 观感判断字体对不对**；判断要看放映截图。
- `fitz` 渲染 SVG 会丢外链图，量 SVG 要过 Chrome headless。

## 5. 验收流程（别跳）

1. `apply_motion.py` 生成。
2. `verify_playback.py` 前台放映 + 每 0.2~0.3s 截一帧，拼 4 列图。
3. 图里要能看见：黑场 → 底图显影 → 大字从画外进入并出画 → 标题刷出 → 正文逐条出现 → 人像滑入。
4. 抽一页内页（`--goto 3`）：底图不闪白、标题擦写、条目错峰。
5. 有明显提前曝光/跳帧，就回去改对应层的时间，而不是加更多效果。

# python-pptx 原生可编辑排雷手册（告别土味与 AI 堆框感）

直接裸写 `python-pptx` 做出来的幻灯片为什么常常显得廉价？以下 7 个底层坑已在 `scripts/pptx_kit.py` 与 `scripts/render_pptx.py` 中全部封装解决，写装配脚本时请始终调用 `pptx_kit`：

---

## 坑 1：中文字体自动退化成宋体（SimSun）
- **原因**：`run.font.name = "微软雅黑"` 只写 `<a:latin typeface="..."/>`，PowerPoint/WPS 渲染中文读的是 `<a:ea typeface="..."/>`。
- **解法**：`pptx_kit.set_run_font()` 同时写入 `<a:latin>`、`<a:ea>`、`<a:cs>` 三个 XML 节点（默认 `"Microsoft YaHei"`，可用 `set_default_font()` 切换为苹方/思源黑体等）。

## 坑 2：形状自带厚重脏黑阴影
- **原因**：PowerPoint 空白模板的默认 Shape 样式会继承主题 `effectStyleLst`，导致虚线框、胶囊标签、分割线底下全带一圈模糊黑影。
- **解法**：`pptx_kit.add_box()` 默认调用 `clear_shadow()` 显式写入空 `<a:effectLst/>`；仅对需要层次的纯白内容卡开启 `shadow=True`，施加克制的柔和微影（`blur=7pt, dist=2pt, alpha=7%`）。

## 坑 3：图片硬塞进固定宽高导致图标被压扁拉长
- **原因**：直接 `add_picture(path, x, y, width=w, height=h)` 若与裁剪后 PNG 比例不一致就会变形。
- **解法**：使用 `pptx_kit.add_picture_fit(slide, path, x, y, w, h, fit="contain")`，自动读取 PNG 真实宽高，在目标盒内等比缩放并居中放置（仅可拉伸的渐变底卡才用 `fit="stretch"`）。

## 坑 4：PowerPoint 原生表格默认蓝白粗格线
- **解法**：`pptx_kit.add_table_matrix()` 自动清除默认蓝白主题样式，手动写入咨询级表格 XML：品牌主色实底白字表头、极浅衬底交替行、仅保留极细浅灰底边线（`0.75pt #E6E9EF`）、去掉粗黑竖线与外框，每格内边距舒适、垂直居中。

## 坑 5：企业底板占位标题残影或误遮挡页眉元素
- **解法**：
  - 若用户提供的企业背景图已经干净（无占位标题），直接调用 `add_baseboard(slide, bg_path)` 整张锁死在 Layer 1，绝不裁切；
  - 若底板自带示例标题，用 `prepare_baseboard.py --box x0,y0,x1,y1` 采样标题正上方真实背景像素精准无痕擦除，再叠放原生可编辑标题；
  - 内容区务必避开用户底板自带的页眉线、Logo 与页脚安全区，切勿用不透明白矩形覆盖底板装饰线。

## 坑 6：`MSO_SHAPE.BLOCK_ARC` 角度乘数陷阱（画圆环比例完全错乱）
- **原因**：OOXML 角度单位是 `1/60000` 度，而 `python-pptx` 写入 `adjustments[i]` 时统一乘以 `100000`，直接填 `270.0` 会变成 `450°` 导致圆环缺口错位。
- **解法**：直接调用 `pptx_kit.add_donut_ring(slide, cx, cy, d, pct=0.48)`（内部已自动换算 `deg * 0.6` 锁定 12 点顺时针精确占比）。

## 坑 7：多 Agent 并行调 PowerPoint COM 互相杀进程
- **解法**：`scripts/render_pptx.py` 已内置跨进程文件互斥锁（`slideforge_com_render.lock`），多个并行子 Agent 同时渲染时会自动排队安全导出，不会互相 `Quit()` 冲突。

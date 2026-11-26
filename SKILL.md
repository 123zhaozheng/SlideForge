---
name: slide-forge
description: "100% native editable PPTX hybrid production engine (AI pure-white multi-batch component sprite sheets × layered python-pptx OOXML assembly × real PowerPoint/WPS/LibreOffice visual QA). Supports both Mode A (Direct-Design from document + any user background template) and Mode B (1:1 Blueprint-Replica from any existing slide image or AI blueprint). Automatically splits large component sets into multiple consistent batches, extracts crisp RGBA transparent PNGs with checkerboard QA, builds razor-sharp native vector shapes/charts/tables/text, and visually verifies every rendered slide. Triggers: /slide-forge, editable PPTX, recreate PPT, AI to editable PowerPoint, 可编辑PPT, 真PPT复刻, 组件抠图做PPT."
---

# SlideForge · AI 组件图集 × 100% 原生可编辑 PPTX 生产引擎

**核心哲学：不做「逆向像素硬抠转 PPT」，而是「视觉解构 / 创意设计 → 正向分层重建」。**
- **AI 生图模型负责视觉灵魂**：高光渐变 Hero 底板、风格统一的扁平矢量图标/徽章、丝缎渐隐箭头、定制 3D/扁平插画。
- **原生 PowerPoint（`python-pptx`）负责文字与结构**：100% 真文字（强制写入 `<a:ea>` 东亚字体，0 错字、随时可编辑）、开放式印刷级留白排版、极细 Hairline 分割线、虚线 Bento 容器、原生线性渐变、精确比例圆环图、高级原生表格与精准网格对齐。
- **完全通用，不绑定任何固定企业底板或固定 Logo**：用户提供的背景图自带什么品牌框架（Logo、页眉线、页脚、渐变底纹），原封不动锁死在 Layer 1 即可；无背景图时直接在纯净画布上按所选设计体系构建。

---

## 两种工作模式（按场景自由切换）

| 模式 | 适用场景 | 流程特点 |
| :--- | :--- | :--- |
| **模式 A：极速直出模式（Direct-Design）** | 用户给出「文档 / 大纲 + 任意背景图（可选）」，追求高效率直接出可编辑 PPT | **跳过整页蓝图**：Agent 直接在脑中规划非对称高定版式（注意克制留白、切忌堆砌“框套框”）→ 拆出本页所需视觉组件调用生图 API 出纯白底图集 → 代码分层直出可编辑 PPTX → 真实渲染回看精修。**速度翻倍、成本减半。** |
| **模式 B：大师蓝图复刻模式（Blueprint-Replica）** | ① 用户已有现成成品图/AI 幻灯片需 1:1 复刻；② 高管/重磅汇报需极致视觉突破与去“AI 卡片感” | 先由生图模型生成整页高定视觉蓝图 `blueprint.png`（或直接使用用户提供的成品图）→ 将蓝图作为 `style_ref` 分批生成纯白底无文字组件图集 → 代码 4 层 1:1 复刻（原 PPT 内若含实拍照片/证书，直接无损提取复用）→ 真实渲染逐项对标精修。 |

---

## 完整 5 步生产流水线

```text
Step 1 定模式、锁底板与品牌色系
   ├─ 若用户提供背景母版图：
   │    • 若底板本身干净无示例标题 → 直接整张铺满 Layer 1（绝不裁切，100% 保留用户原有 Logo/页眉/边框）
   │    • 若底板带有占位示例标题   → 用 prepare_baseboard.py 按 --box 坐标无痕擦除占位标题后铺满 Layer 1
   ├─ 若用户未提供底板 → 使用纯净画布，并从 references/design-systems.md 选定匹配主题色板
   ├─ 模式 A（极速直出）→ 读 references/design-systems.md + layout-archetypes.md，规划去框化/非对称高定版式
   └─ 模式 B（蓝图复刻）→ 用已有成品图或先出 1 张整页参考蓝图作为 style_ref 锁死审美、留白与画风
   ↓
Step 2 规划「组件清单 manifest」与自动分批策略（核心！）
   ├─ 严格划界：哪些归原生 PPTX 矢量/文字，哪些交 AI 画（详见 references/component-splitting.md）
   ├─ 单批硬上限：图标/徽章最多 8 个（2×4），大型卡片/复杂插画最多 4–6 个（2×2 或 2×3）
   └─ 组件多时拆成 batch_01 / batch_02 / ... 多批生成，全部绑定同一 style_ref 与 palette_lock 锁死统一画风
   ↓
Step 3 多批次生成纯白底图集 → 自动裁切 + 透明抠图
   ├─ generate_components.py --session {session} --manifest {session}/components_manifest.json
   └─ extract_components.py  --session {session} --manifest {session}/components_manifest.json
         → 输出全套 RGBA 透明 PNG + 每批棋盘格验图 *_alpha_preview.png（Agent 必须亲眼看图验收！）
   ↓
Step 4 分层装配 100% 原生可编辑 PPTX
   └─ 使用 scripts/pptx_kit.py 分 4 层装配：
         Layer 1 用户背景底板 → Layer 2 原生矢量结构 → Layer 3 透明 AI 组件/实拍原图 → Layer 4 真文字与原生表格
   ↓
Step 5 真实渲染回导 + 视觉闭环自检
   └─ render_pptx.py --pptx {out.pptx} --out-dir {session}/rendered
         → Agent 逐页查看真实渲染 PNG，排查文字折行溢出、遮挡、对齐与视觉呼吸感，微调坐标直至完美！
```

---

## 会话目录结构（推荐位于工作区 `outputs/slide-forge/<session>/`）

```text
outputs/slide-forge/<session>/
├── baseboard.png                   # 用户提供的原始背景图（可选）
├── baseboard_clean.png             # 擦除占位标题后的干净背景图（可选）
├── style_ref.png                   # 视觉风格锁死参考图（模式 B 蓝图或品牌参考图）
├── components_manifest.json        # 多批次组件清单（每批行列、每枚组件提示词、cut_mode）
├── sheets/
│   ├── batch_01_raw.png            # 第 1 批纯白底无文字组件图集
│   ├── batch_01_alpha_preview.png  # 第 1 批透明抠图棋盘格验收图
│   └── batch_02_raw.png            # 超过 8 个组件自动拆出的第 2 批
├── components/
│   ├── hero_card.png               # 自动裁切提取的 RGBA 透明 PNG
│   ├── leap_arrow.png
│   └── leap_arrow_white.png        # 自动派生的纯白透明变体（tint: "white"）
├── build_deck.py                   # 使用 pptx_kit.py 的分层装配脚本（多页并行时每页导出 build_slide(prs)）
├── rendered/
│   └── slide-01.png ...            # PowerPoint / WPS / LibreOffice 真实渲染验收图
└── output_editable.pptx            # 最终 100% 原生可编辑成品 PPTX
```

---

## 快速命令速查

> Windows 推荐使用 `uv run --with ... python`（避免系统 Python 环境差异）；macOS/Linux 可直接使用已激活虚拟环境的 `python3` 或 `uv run`。

```bash
# 1.（可选）擦除母版底板上的占位示例标题（--box 按比例 x0,y0,x1,y1 适配任意模板）
uv run --with pillow --with numpy python "{SKILLDIR}/scripts/prepare_baseboard.py" \
  --input baseboard.png --out "{session}/baseboard_clean.png" --box 0.04,0.03,0.56,0.11

# 2. 按 manifest 多批次生成纯白底无文字组件图集（内置限流退避重试）
uv run --with requests --with pillow --with python-dotenv python "{SKILLDIR}/scripts/generate_components.py" \
  --session "{session}" --manifest "{session}/components_manifest.json" --concurrency 2

# 3. 全自动网格切分 + 泛洪/反混色抠透明 PNG + 生成棋盘格验收图
uv run --with pillow --with numpy python "{SKILLDIR}/scripts/extract_components.py" \
  --session "{session}" --manifest "{session}/components_manifest.json"

# 4. 运行分层装配脚本生成 100% 可编辑 PPTX
uv run --with python-pptx --with pillow --with numpy python "{session}/build_deck.py"

# 5. 调用本机 PowerPoint / WPS COM（支持多进程互斥锁）或 LibreOffice 渲染回 PNG 自检
uv run python "{SKILLDIR}/scripts/render_pptx.py" \
  --pptx "{session}/output_editable.pptx" --out-dir "{session}/rendered"
```

---

## 多页并行生产协议（多 Agent 提速指南）

面对 10 页以上的长汇报时，推荐启动多个并行子 Agent（每页一个独立目录 `slides/slide_XX/`）：
1. 每页的 `build_slide.py` 必须导出统一契约函数 `def build_slide(prs):`，在其内部调用 `add_blank_slide(prs)` 并返回 `slide`，所有组件路径均基于 `Path(__file__).resolve().parent`。
2. 每个子 Agent 独立完成本页组件生成、抠图、单页 `slide_XX.pptx` 构建与 `render_pptx.py` 实渲验收（`render_pptx.py` 已内置跨进程文件互斥锁，不会触发 PowerPoint COM 冲突）。
3. 全部页面验收通过后，主控脚本顺序 `import` 各页 `build_slide(prs)` 即可秒级合成整本可编辑演示文稿。

---

## 核心红线（违反即降级）

1. **图集必须纯白 `#FFFFFF` + 零文字**：严禁向生图模型索要「透明底」（极易生成假棋盘格）；严禁在组件图集里让模型画任何文字、字母、数字、网格线或外框。
2. **组件过多必须拆批生成**：图标/徽章每批 **≤ 8 个（2×4）**，大卡片/复杂插画每批 **≤ 4–6 个（2×2 / 2×3）**；跨批次绑定同一 `style_ref` 与 `palette_lock` 锁死画风一致。
3. **抠图必看棋盘格验图**：`extract_components.py` 产出 `*_alpha_preview.png` 后，必须亲眼核验无白边、无误抠内部白图形、无截断，再进入 PPTX 装配。
4. **反“AI 堆框感”原则**：优先用**开放留白 + 极细 Hairline 分割线 + 字号/色阶对比 + 原生数据条/圆环**构建杂志级呼吸感，杜绝无意义的“大框套小框再套胶囊”。
5. **底层避坑强制执行**：必须使用 `pptx_kit.py`（写入 `<a:ea>` 东亚字体、清空默认主题黑影 `<a:effectLst/>`、使用 `add_donut_ring` 避开 OOXML `BLOCK_ARC` 角度乘数坑）；若底板自带页眉线/Logo，内容区严禁遮挡底板装饰元素。
6. **交付前必做真实渲染自检**：每次装配完 `.pptx` 必须运行 `render_pptx.py` 导出 PNG 并肉眼检查排版、断行、对齐与留白，发现任何瑕疵立刻修正重跑。

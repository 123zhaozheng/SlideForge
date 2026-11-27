---
name: slide-forge
description: "100% native editable PPTX hybrid production engine with friendly interactive onboarding (env check -> brief/template intake -> Mode A/B choice -> style reference & photo/illustration slot confirmation), optional full-slide master blueprint generation (Mode B), multi-batch pure-white AI component sprite sheets, RGBA alpha extraction with checkerboard QA, layered python-pptx OOXML assembly (including native cover-cropped/rounded/placeholder image slots), and real PowerPoint/WPS/LibreOffice visual verification. Triggers: /slide-forge, editable PPTX, recreate PPT, AI to editable PowerPoint, 可编辑PPT, 真PPT复刻, 组件抠图做PPT."
---

# SlideForge · AI 视觉蓝图与组件图集 × 100% 原生可编辑 PPTX 生产引擎

**核心哲学：不做「逆向像素硬抠转 PPT」，而是「审美定调 / 视觉解构 → 正向 4 层原生重建」。**
- **AI 生图模型负责视觉灵魂**：既可在模式 B 中先生成整页高定视觉蓝图锁死顶级设计感，也负责生成纯白底无文字组件图集（高光渐变 Hero 底板、风格统一的扁平矢量图标/徽章、丝缎渐隐箭头、定制插画）。
- **原生 PowerPoint（`python-pptx`）负责文字、结构与图片槽位**：100% 真文字（强制写入 `<a:ea>` 东亚字体，0 错字、随时可编辑）、开放式印刷级留白排版、极细 Hairline 分割线、虚线 Bento 容器、原生线性渐变、精确比例圆环图、咨询级原生表格，以及 `add_image_slot()` 原生图片/插图槽位（支持居中裁切 `cover`、圆角/圆形、图注与无图虚线占位框）。
- **完全通用，不绑定任何固定企业底板或 Logo**：用户提供什么背景母版就原封不动锁死在 Layer 1；无底板时直接在纯净画布上按所选设计体系构建。

---

## Step 0 · 友好交互式启动协议（每次接到做 PPT 任务时优先执行）

> **交互原则**：不闷头乱画、不一次抛 10 个技术术语。用自然、清晰的对话节奏帮用户把关键决策一步步对齐，再稳稳跑通全流程。

### 0.1 自动检查环境配置（`.env` / API Key）
收到用户的做 PPT 需求后，**第一步先静默检查生图环境**（绝不输出完整密钥）：
```bash
uv run --with requests --with python-dotenv python "{SKILLDIR}/scripts/generate_blueprint.py" --check-env
```
- **检查通过（`ENV_STATUS: OK`）**：简短告知用户当前已就绪的模型与环境（如“生图环境已就绪，模型 `gpt-image-2.6`”），直接进入下一步。
- **未配置（`MISSING_API_KEY`）**：停下来友好提示用户复制 `{SKILLDIR}/.env.example` 为 `{SKILLDIR}/.env` 并填入 `OPENAI_API_KEY` / `OPENAI_BASE_URL` / `SLIDEFORGE_MODEL`，配好后继续。

### 0.2 接收并阅读用户的文档与背景素材
如果用户还没发齐材料，礼貌请用户提供：
1. **核心内容**：文档（Word / Markdown / PDF）、大纲或口述要点；
2. **可选背景母版**：公司/汇报专用背景图（PNG/JPG）或已有 PPT 模板（`.pptx`）。

**Agent 必做动作**：收到文档或背景图后，**先完整阅读**，然后用 2–3 句话简要复述你读到的核心主题、建议拆分页数与页面节奏，确认理解无误。

### 0.3 清晰介绍两种模式，请用户选择执行路径
读完材料后，用通俗直白的方式向用户介绍两种工作模式的区别，请用户二选一（若用户开头已明确指定某模式，直接尊重并跳过此问）：

| 模式 | 核心流程 | 适合什么场景 |
| :--- | :--- | :--- |
| **模式 A · 极速直出可编辑 PPTX** | 跳过整页参考图 → Agent 直接按大师版式规范规划排版 → 生成小组件图集 → 代码 4 层直出原生可编辑 PPTX → 实渲自检。 | **速度快、生图成本减半**；适合日常汇报、系统架构图、逻辑清晰的常规业务页。 |
| **模式 B · 高设计感蓝图 → 1:1 可编辑复刻（推荐重磅汇报）** | 先用生图模型画出杂志级/高管级**整页视觉蓝图**（先试出 1–2 页给你看风格和构图，不满意随时低成本重画），确认设计感没问题后，再全自动拆组件 1:1 重建为 **100% 原生可编辑 PPTX**。 | **视觉设计感上限最高**、彻底告别“AI 卡片堆砌感”；适合高管汇报、对外宣讲，或把已有的 AI 扁平图/成品图还原为真 PPT。 |

### 0.4 确认风格参考与图片 / 插图槽位（走模式 B 或有图文需求时必问）
在动手生图前，向用户确认两个能让成片质感翻倍的细节（一句话问清楚，避免反复返工）：
1. **审美与风格参考（可选）**：
   - 询问用户手里有没有**觉得好看的参考图**，或者**符合公司审美的历史优秀 PPT（`.pptx`）**？
   - 若用户给了历史 `.pptx`，立即运行 `scripts/render_pptx.py` 把它渲染成逐页 PNG，亲眼看图提炼配色阶梯、版式呼吸感与品牌气质，并在模式 B 中将精选页作为 `--ref` 传入（严禁照抄其中的具体业务文字或水印）。
   - 若没有参考图，则从 `references/design-systems.md` 中推荐最契合主题的 1 套风格体系让用户确认。
2. **真实图片与插图槽位确认（必问）**：
   - 询问用户本次 PPT 中**有没有需要放入的真实图片或定制插图**（例如：产品/系统界面截图、现场实拍照片、荣誉证书、团队头像、或某页想要特定主题的配图插画），并说明希望放在哪一页、大概什么位置：
     - **用户提供现成图片**：记录路径与对应页面 → 模式 B 蓝图阶段写入 `blueprint_manifest.json` 的 `slots` 自动留白避让，PPTX 装配阶段用 `pptx_kit.add_image_slot(..., fit="cover", shape="rounded")` 居中裁切圆角贴入；
     - **用户想让 AI 画定制插图**：明确插图主题与风格，归入 `components_manifest.json` 单独生成透明主视觉插画，或作为整幅插图填入图片槽位；
     - **手头暂时没图但想预留槽位**：记下页面与用途标签，蓝图阶段自动留白，PPTX 装配阶段用 `add_image_slot(..., img_path=None, placeholder_label="待放入：xxx")` 生成精致的原生虚线占位槽位，用户事后在 PowerPoint 里一键替换即可。

> 一旦 Step 0 对齐完毕，后续流程即可全自动丝滑跑通！

---

## 完整生产流水线（Mode A / Mode B 统一闭环）

```text
Step 0 环境检查 (.env) → 阅读文档/底板 → 选模式 A/B → 确认风格参考 & 图片/插图槽位
   ↓
Step 1 锁底板、定色板与版式节奏
   ├─ 若用户提供历史参考 PPT → 先用 render_pptx.py 导出 PNG 学习风格/母版
   ├─ 若用户提供母版背景图：
   │    • 底板干净无占位标题 → 直接铺满 Layer 1（100% 保留 Logo/页眉/边框）
   │    • 底板带占位示例标题 → prepare_baseboard.py 按 --box 无痕擦除后铺满 Layer 1
   └─ 读取 references/design-systems.md + layout-archetypes.md + component-library.md 规划整套节奏（相邻异构、单视觉锚点）
   ↓
Step 1.5（仅模式 B · 无现成成品图时执行）生成高设计感整页蓝图 + 槽位避让
   ├─ 按 references/blueprint-prompts.md 撰写大师级五段蓝图提示词 + blueprint_manifest.json（含可选 slots）
   ├─ generate_blueprint.py --session {session} [--baseboard ...] [--ref ...] --only 1,2
   │    → 先出 1–2 页蓝图给用户亲眼看图确认审美与版式；满意后再批量出齐其余蓝图！
   └─ （若用户本来就提供了现成扁平图/成品图，直接把该图作为 blueprint / style_ref 跳过本步）
   ↓
Step 2 规划「组件清单 manifest」与自动分批策略（核心！）
   ├─ 严格划界：哪些归原生 PPTX 矢量/文字/表格/图片槽位，哪些交 AI 画（详见 references/component-splitting.md）
   ├─ 单批硬上限：图标/徽章最多 8 个（2×4），大型卡片/复杂插画最多 4–6 个（2×2 或 2×3）
   └─ 多批全部绑定同一 style_ref（模式 B 即本页蓝图）与 palette_lock 锁死统一画风
   ↓
Step 3 多批次生成纯白底组件图集 → 自动裁切 + 透明抠图验收
   ├─ generate_components.py --session {session} --manifest {session}/components_manifest.json
   └─ extract_components.py  --session {session} --manifest {session}/components_manifest.json
         → 输出全套 RGBA 透明 PNG + 每批棋盘格验图 *_alpha_preview.png（Agent 必须亲眼看图验收！）
   ↓
Step 4 分层装配 100% 原生可编辑 PPTX
   └─ 使用 scripts/pptx_kit.py 分 4 层装配：
         Layer 1 用户背景底板
         Layer 2 原生 OOXML 矢量结构（卡片/虚线 Bento/渐变/圆环/连接线/表格）
         Layer 3 透明 AI 组件 + add_image_slot() 真实照片/截图/定制插图/虚线占位槽位
         Layer 4 真文字排版（强制 <a:ea> 东亚字体，0 错字、全可编辑）
   ↓
Step 5 真实渲染回导 + 视觉闭环自检
   └─ render_pptx.py --pptx {out.pptx} --out-dir {session}/rendered
         → Agent 逐页查看真实渲染 PNG，排查文字折行溢出、遮挡、槽位裁切与呼吸感，微调坐标直至完美！
```

---

## 会话目录结构（推荐位于工作区 `outputs/slide-forge/<session>/`）

```text
outputs/slide-forge/<session>/
├── baseboard.png                   # 用户提供的原始背景图（可选）
├── baseboard_clean.png             # 擦除占位标题后的干净背景图（可选）
├── style_refs/                     # 用户提供的风格参考图或历史 PPT 渲染图（可选）
├── blueprint_manifest.json         # 模式 B 整页蓝图任务清单（含可选 slots 声明）
├── blueprint_prompts/slide-01.txt  # 模式 B 大师级五段蓝图提示词
├── blueprints/slide-01.png ...     # 模式 B 生成的整页高定视觉蓝图
├── blueprint_metadata.json         # 蓝图记录与槽位精确英寸坐标（rect_in）
├── components_manifest.json        # 多批次纯白底组件图集清单
├── sheets/
│   ├── batch_01_raw.png            # 第 1 批纯白底无文字组件图集
│   ├── batch_01_alpha_preview.png  # 第 1 批透明抠图棋盘格验收图
│   └── batch_02_raw.png            # 超过 8 个组件自动拆出的第 2 批
├── components/                     # 自动裁切提取的 RGBA 透明 PNG 组件
├── build_deck.py                   # 使用 pptx_kit.py 的分层装配脚本
├── rendered/slide-01.png ...       # PowerPoint / WPS / LibreOffice 真实渲染验收图
└── output_editable.pptx            # 最终 100% 原生可编辑成品 PPTX
```

---

## 快速命令速查

> Windows 推荐使用 `uv run --with ... python`（避免系统 Python 环境差异）；macOS/Linux 可直接使用已激活虚拟环境的 `python3` 或 `uv run`。

```bash
# 0. 检查生图环境配置（安全脱敏输出，不打印完整密钥）
uv run --with requests --with python-dotenv python "{SKILLDIR}/scripts/generate_blueprint.py" --check-env

# 1.（可选）若用户提供了历史参考 PPT，先渲染成逐页 PNG 供看图学习风格
uv run python "{SKILLDIR}/scripts/render_pptx.py" \
  --pptx reference_deck.pptx --out-dir "{session}/style_refs"

# 1b.（可选）擦除母版底板上的占位示例标题（--box 按比例 x0,y0,x1,y1 适配任意模板）
uv run --with pillow --with numpy python "{SKILLDIR}/scripts/prepare_baseboard.py" \
  --input baseboard.png --out "{session}/baseboard_clean.png" --box 0.04,0.03,0.56,0.11

# 1.5（仅模式 B）先试出前 1–2 页整页高定视觉蓝图给用户确认风格（支持 --baseboard / --ref / slots）
uv run --with requests --with pillow --with python-dotenv python "{SKILLDIR}/scripts/generate_blueprint.py" \
  --session "{session}" --baseboard "{session}/baseboard.png" --ref "{session}/style_refs/slide-01.png" --only 1,2

# 2. 按 manifest 多批次生成纯白底无文字组件图集（内置限流退避重试）
uv run --with requests --with pillow --with python-dotenv python "{SKILLDIR}/scripts/generate_components.py" \
  --session "{session}" --manifest "{session}/components_manifest.json" --concurrency 2

# 3. 全自动网格切分 + 泛洪/反混色抠透明 PNG + 生成棋盘格验收图
uv run --with pillow --with numpy python "{SKILLDIR}/scripts/extract_components.py" \
  --session "{session}" --manifest "{session}/components_manifest.json"

# 4. 运行分层装配脚本生成 100% 可编辑 PPTX（含原生 add_image_slot 图片/插图槽位）
uv run --with python-pptx --with pillow --with numpy python "{session}/build_deck.py"

# 5. 调用本机 PowerPoint / WPS COM（支持多进程互斥锁）或 LibreOffice 渲染回 PNG 自检
uv run python "{SKILLDIR}/scripts/render_pptx.py" \
  --pptx "{session}/output_editable.pptx" --out-dir "{session}/rendered"
```

---

## 多页并行生产协议（多 Agent 提速指南）

面对 10 页以上的长汇报时，推荐启动多个并行子 Agent（每页一个独立目录 `slides/slide_XX/`）：
1. 每页的 `build_slide.py` 必须导出统一契约函数 `def build_slide(prs):`，在其内部调用 `add_blank_slide(prs)` 并返回 `slide`，所有组件与槽位图片路径均基于 `Path(__file__).resolve().parent`。
2. 每个子 Agent 独立完成本页组件生成、抠图、单页 `slide_XX.pptx` 构建与 `render_pptx.py` 实渲验收（`render_pptx.py` 已内置跨进程文件互斥锁，不会触发 PowerPoint COM 冲突）。
3. 全部页面验收通过后，主控脚本顺序 `import` 各页 `build_slide(prs)` 即可秒级合成整本可编辑演示文稿。

---

## 核心红线（违反即降级）

1. **先查环境再问需求，不闷头盲画**：始终执行 Step 0（查 `.env` → 读材料 → 说明模式 A/B 差异让用户选 → 问清楚风格参考与真实图片/插图槽位）。
2. **模式 B 先试 1–2 页蓝图定审美**：不要一口气盲画 15 页蓝图；先出 1–2 页确认构图与品味，满意后再批量出图并转可编辑。
3. **组件图集必须纯白 `#FFFFFF` + 零文字**：严禁向生图模型索要「透明底」（极易生成假棋盘格）；严禁在组件图集里让模型画任何文字、字母、数字、网格线或外框。
4. **组件过多必须拆批生成**：图标/徽章每批 **≤ 8 个（2×4）**，大卡片/复杂插画每批 **≤ 4–6 个（2×2 / 2×3）**；跨批次绑定同一 `style_ref` 与 `palette_lock` 锁死画风一致。
5. **抠图必看棋盘格验图**：`extract_components.py` 产出 `*_alpha_preview.png` 后，必须亲眼核验无白边、无误抠内部白图形、无截断，再进入 PPTX 装配。
6. **反“AI 堆框感”与真实槽位保真**：优先用**开放留白 + 极细 Hairline 分割线 + 咨询级微组件**构建呼吸感；真实照片/产品截图绝不让模型胡乱重绘，统一走「蓝图槽位避让 + `pptx_kit.add_image_slot()` 原生裁切/占位」。
7. **底层避坑与交付前必实渲自检**：必须使用 `pptx_kit.py`（`<a:ea>` 东亚字体、清空默认主题黑影 `<a:effectLst/>`、`add_donut_ring` 避开 `BLOCK_ARC` 角度乘数坑）；每次装配完 `.pptx` 必须运行 `render_pptx.py` 导出 PNG 并肉眼检查排版、断行、对齐与留白。

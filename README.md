# SlideForge

**Turn documents, ideas, or flat AI slides into 100% native, editable PowerPoint decks — without blurry OCR hacks.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB.svg)](https://www.python.org/)
[![Agent Skill](https://img.shields.io/badge/Agent_Skill-Claude_Code_%7C_Cursor_%7C_Antigravity_%7C_Codex-111827.svg)](SKILL.md)

**SlideForge** 是一款面向 AI Coding Agent 的 **100% 原生可编辑 PPTX 混合生产技能（Agent Skill）**。  
它彻底抛弃传统「图片转 PPT」脆弱的 OCR 擦字硬抠路线，首创 **「视觉解构 / 创意设计 → 纯白底多批次组件图集 → 代码分层正向重建 → 真实 PowerPoint 渲染闭环自检」** 四层生产引擎：

- **文字与结构 100% 原生可编辑**：标题、正文、数字、渐变卡片、虚线 Bento 容器、精准圆环图、原生表格、箭头连线全部由 OOXML 矢量构建，双击即改、无限缩放锐利。
- **视觉灵魂交由 AI 图集批产**：只把代码画不出来的定制图标、3D 立体图示、高光 Hero 底板、丝缎渐隐箭头交由生图模型在纯白 `#FFFFFF` 底上分批生成，再经外围 BFS 泛洪 / 反混色算法无损抠成透明 PNG。
- **天然适配任意企业母版**：直接垫入你自己的企业背景底板（自带 Logo、页眉、页脚或纯白画布），零裁切、零变形、自动吸取品牌配色体系。
- **双工作模式自由切换**：支持「**极速直出**（跳过整页蓝图，省时省费）」与「**大师蓝图复刻**（先定顶级设计总监蓝图，再 1:1 复刻为真 PPT）」两种模式。
- **真实 PowerPoint / WPS / LibreOffice 视觉自检**：每页装配完毕自动导出 1080P PNG 回看，排查文字换行溢出、元素重叠与留白失衡，迭代至交付级水准。

---

## 为什么传统「图片转 PPT」总是一团糟？

| 维度 | 传统 OCR / 逆向抠图工具 | **SlideForge（正向分层重建）** |
| :--- | :--- | :--- |
| **底层路线** | 对着成品图硬擦文字 + 暴力抠块 | **4 层语义解耦正向施工**（底板 / 原生矢量 / 透明 AI 组件 / 真文字） |
| **背景与渐变** | 擦字必糊，背景色块发脏 | 原始背景底板 100% 原封不动锁死在底层，矢量渐变丝滑锐利 |
| **图标与插画** | 边缘白边锯齿，白底内图标被误抠穿 | 锚定参考图分批重绘纯白底无文字图集 → BFS 泛洪 / 反混色无损抠透明 |
| **中文字体与阴影** | 经常退化回宋体，自带厚重默认黑影 | 强制写入 `<a:ea>` 东亚字体 + 清空默认主题重阴影 `<a:effectLst/>` |
| **复杂页面组件多** | 塞进一张图糊成一团 | 超过 6–8 个自动拆分多批次生成，跨批绑定 `style_ref` 锁死统一画风 |
| **排版质检** | 盲出文件，开盖即惊喜 | 自动调本机 PowerPoint / WPS / LibreOffice 渲染回看，闭环精修 |

---

## 核心架构：4 层解耦 × 双工作模式

```text
用户输入：业务文档 / 创意需求 / 已有成品图  +  任意背景母版（可选）
   │
   ├── 模式 A · 极速直出（Direct-Design）
   │     跳过整页蓝图 → 脑中直出非对称版式 → 拆组件图集 → 代码分层直出 → 实渲自检
   │
   └── 模式 B · 大师蓝图复刻（Blueprint-Replica）
         已有成品图（或先生成整页视觉蓝图）→ 锚定蓝图拆批出透明组件 → 4 层 1:1 复刻 → 实渲精修
   ↓
┌─────────────────────────────────────────────────────────────┐
│ Layer 4 · 真文字与原生表格层  （<a:ea> 东亚字体 / 富文本 / 矩阵表） │
│ Layer 3 · 透明 AI 组件层     （多批次纯白底图集 → RGBA 透明 PNG）  │
│ Layer 2 · 原生矢量结构层     （渐变图形 / 精准圆环 / 虚线 Bento）   │
│ Layer 1 · 用户原始背景层     （企业母版原图或纯净画布，零失真）      │
└─────────────────────────────────────────────────────────────┘
   ↓
本机 PowerPoint / WPS COM（支持多 Agent 并发互斥锁）或 LibreOffice 渲染回看 → 交付终版 .pptx
```

---

## 快速开始（3 步上手）

### 1. 安装依赖与配置 API Key

```bash
# 克隆仓库
git clone https://github.com/123zhaozheng/SlideForge.git
cd SlideForge

# 复制环境变量模板并填入你的 OpenAI 兼容生图 API Key
cp .env.example .env
```

Windows 用户可一键运行安装脚本（依赖 [`uv`](https://docs.astral.sh/uv/) 或标准 Python 3.10+）：
```powershell
.\install.ps1
```

macOS / Linux 用户：
```bash
chmod +x install.sh && ./install.sh
```

### 2. 作为 Agent Skill 接入你喜欢的 AI 编程工具

将整个 `SlideForge` 文件夹复制到对应 Agent 的技能目录即可自动识别 `SKILL.md`：
- **Claude Code / Cursor / Codex / Antigravity**：放入全局或项目级 `skills/slide-forge/` 目录
- 或直接在对话中告诉 Agent：
  > 「请阅读 `SlideForge/SKILL.md`，按照模式 B 把 `target.png` 复刻成 100% 原生可编辑 PPTX」
  > 或：「请阅读 `SlideForge/SKILL.md`，使用模式 A 基于这份文档和 `template_bg.png` 生成 8 页可编辑高管汇报 PPT」

### 3. 核心脚本一览

| 脚本 | 职责 |
| :--- | :--- |
| `scripts/prepare_baseboard.py` | （可选）按坐标无痕擦除企业底板上的占位示例标题，自动适配纯色/渐变底纹 |
| `scripts/generate_components.py` | 多批次调用生图 API 生成纯白 `#FFFFFF` 无文字组件图集（自动拆批 ≤8 图标 / ≤4 大卡、限流重试、跨批画风锁死） |
| `scripts/extract_components.py` | 自适应投影谷底切分 + 外围 BFS 泛洪抠图（`solid`）+ 白底反混色透明提取（`soft_fade`）+ 杂点清理 + 棋盘格验图 |
| `scripts/pptx_kit.py` | 大师级 `python-pptx` 原生组件库（多主题色板、东亚字体锁死、去重影、原生线性渐变、精确 `add_donut_ring` 圆环、虚线 Bento、高定表格） |
| `scripts/render_pptx.py` | 零 `pywin32` 依赖的跨平台 PPTX → PNG 渲染器（Windows PowerPoint/WPS COM 带跨进程互斥锁 + LibreOffice 兜底） |

---

## 目录结构

```text
SlideForge/
├── SKILL.md                             # Agent 技能核心操作手册（双模式 SOP + 红线规范）
├── README.md                            # 项目介绍与使用文档
├── LICENSE                              # MIT 开源协议
├── .env.example                         # 通用 API 环境变量模板（不含密钥）
├── requirements.txt                     # Python 依赖清单
├── install.ps1 / install.sh             # 一键环境检测脚本
├── examples/
│   └── components_manifest.sample.json  # 多批次组件清单示例
├── references/
│   ├── component-splitting.md           # 组件职责划分与多批次拆图规范
│   ├── design-systems.md                # 品牌色自动推导法则 + 6 大设计体系
│   ├── layout-archetypes.md             # 10 大反平庸版式原型与防雷同节奏律
│   └── pptx-pitfalls.md                 # python-pptx 底层 7 大视觉排雷手册
└── scripts/
    ├── generate_components.py
    ├── extract_components.py
    ├── prepare_baseboard.py
    ├── pptx_kit.py
    └── render_pptx.py
```

---

## Star History & 贡献

如果 **SlideForge** 帮你做出了真正拿得出手、每一笔都能自由编辑的高级 PPT，欢迎点一颗 **Star ⭐** 支持！  
也欢迎提交 Issue / PR 贡献更多行业设计体系、原生图表组件或跨平台渲染优化。

Licensed under the [MIT License](LICENSE).

# Teaching PPT Engine

> 把粗糙讲课大纲 + 本地参考资料 → 一键变成专业教学 PPT + AI 图像生成 prompt

## 这是什么？

一个面向高校教师的 **PPT 半自动生产流水线**。你只需要：

1. 填写一份 5 分钟的 brief（主题、页数、参考文件路径）
2. 引擎自动读取你本地的参考 PPTX 文件，提取全部内容
3. 结合你的需求，生成结构化的幻灯片配置文件
4. 一键输出：**可编辑 PPTX** + **关键页 AI 图像生成 prompt**

```
你的 brief.md + 参考 PPTX  →  read_refs.py  →  context_package.json
                                                    ↓
                                           Claude / 你的 AI 工具
                                                    ↓
                                          slides_config.json
                                                    ↓
                                            make_ppt.py
                                               ↙     ↘
                                     📊 PPTX        🤖 DALL-E Prompts
```

## 效果展示

以下是使用本引擎生成的 20 页「数据交易与合规核心」教学 PPT（13 页 python-pptx 排版 + 7 页 DALL-E prompt）：

| 页面类型 | 数量 | 示例 |
|---------|------|------|
| 封面 & 总结 | 2 页 | 深色封面 + 三问总结 |
| 文字排版页 | 8 页 | 定义、案例、要点 |
| 对比表格 | 1 页 | 三类来源对比 |
| 左右对比 | 1 页 | 三个不要 × 三个一定 |
| AI 图像生成页 | 7 页 | 流程架构图、原理简图、阶梯图 |

## 与同类项目的对比

2024-2025 年 GitHub 上 AI PPT 生成项目可分为三类：

### 第一类：全自动黑盒型

| 代表项目 | 做法 | 缺点 |
|------|------|------|
| [ai-ppt-slide-generator](https://github.com/ysskrishna/ai-ppt-slide-generator) | 主题 → Gemini 生成全稿 → PPTX | 无法逐页控制内容 |
| [LandPPT](https://github.com/sligter/LandPPT) | 主题 → GPT 联网研究 → 自动排版 | AI 猜你需求，不可控 |
| [Presenton](https://github.com/presenton/presenton) | 主题 → 多模型生成 → PPTX+PDF | 不能读本地参考文件 |

### 第二类：实时语音转 PPT

| 代表项目 | 做法 | 缺点 |
|------|------|------|
| [pptGEN-dev](https://github.com/NerdyVisky/pptGEN-dev) | 讲课语音 → 实时生成 slide | 适合即兴，不适合教案 |

### 第三类：本引擎

| 对比维度 | 其他项目 | **Teaching PPT Engine** |
|------|------|------|
| 内容控制 | AI 猜你讲什么 | **你精确指定每页标题、要点、配图** |
| 本地参考文件 | ❌ 都不支持 | ✅ **`read_refs.py` 读取本地 625 页教案** |
| 图文分工 | 全 python-pptx 或全 DALL-E | ✅ **文字页 python-pptx + 关键页 DALL-E 解耦** |
| API 依赖 | 必须填 Gemini/GPT/DALL-E Key | ✅ **引擎不调用任何外部 API** |
| 设计方法论 | 模板卡片堆叠 | ✅ **咨询级 Proof Object 方法论** |
| 教学特有功能 | 无 | ✅ **课程头标/互动提问/可删减提示/进度条** |
| 可编辑性 | 部分只能出图片 | ✅ **全文字可编辑 PPTX** |

### 我们的劣势（已知，正在解决）

- 7 个关键流程图需要**手动复制 prompt 到即梦/DALL-E 生图**再插入 PPT（因为即梦无公开 API，手动操作等于最后一道质量控制）
- 需要一个 AI 工具（Claude/ChatGPT）辅助生成 `slides_config.json`（但只需一分钟）

## 快速开始

### 环境要求

```bash
pip3 install python-pptx
```

### 第一步：准备 brief

复制 `examples/brief_template.md`，填写你的需求：

```markdown
主题: 数据交易与合规核心
期望页数: 20
课程名称: 《数据要素×：理论实务》第五周 · 20分钟
第一章：合法数据来源（公开/授权/自采）—— 约5页
第二章：数据合规技术武器（隐私计算/脱敏加密）—— 约5页
第三章：数据交易三不原则 —— 约5页
第四章：综合案例与总结 —— 约5页
参考文件目录: /path/to/your/reference/pptx/files/
```

### 第二步：读取参考文件

```bash
python3 read_refs.py "你的参考PPTX文件夹路径" brief.md
# 输出: context_package.json（包含所有参考文件的全文提取）
```

### 第三步：生成幻灯片配置

将 `context_package.json` 和你的 `brief.md` 发给 AI 工具（Claude / ChatGPT 等），
请它生成 `slides_config.json`。推荐 prompt：

```
请根据以下上下文包和 brief，生成一份 slides_config.json。
配置格式参考 examples/slides_compliance.json。
每页需指定 type（cover/content/cards/table/layers/comparison/summary/dalle），
需要 AI 图像生成的页面 type 设为 "dalle"。
```

### 第四步：一键生成 PPT

```bash
python3 make_ppt.py slides_config.json
# 输出:
#   ~/Desktop/<output_name>.pptx           — 可编辑教学PPT
#   ~/Desktop/<output_name>-DALLE-prompts.md — AI图像生成prompt
```

## 配置格式说明

`slides_config.json` 中的每页支持以下类型：

| type | 用途 | 必需字段 |
|------|------|---------|
| `cover` | 封面 | title, subtitle |
| `content` | 通用内容页 | title, lines/cards/info_boxes |
| `cards` | 多卡片并列 | title, cards[{title, lines, icon, color}] |
| `table` | 对比表格 | title, table{headers, rows, widths} |
| `layers` | 多层架构图 | title, layers[{label, desc, color}] |
| `comparison` | 左右对比 | title, left{header, items}, right{header, items} |
| `summary` | 编号问题总结 | title, items[{q, detail, color}] |
| `dalle` | AI图像生成页 | title, dalle_desc, image_hint |

完整示例见 `examples/slides_compliance.json`。

## 关于 AI 图像生成

本引擎 **不内置 AI 图像生成能力**。`make_ppt.py` 会：

- 为 `type: "dalle"` 的页面在 PPTX 中生成文字排版占位
- 同时在 `-DALLE-prompts.md` 文件中输出该页的 AI 图像生成 prompt

你需要**使用自己的 API Key**，将 prompt 复制到以下任一工具中生成图片：

- [DALL-E](https://openai.com/dall-e)（需 OpenAI API Key）
- [即梦 AI](https://jimeng.jianying.com/)（字节跳动，中文 prompt 友好）
- [Midjourney](https://www.midjourney.com/)
- 任何支持文本生成图像的工具

生成 PNG 后，拖入 PPTX 替换对应页的占位内容即可。

> ⚠️ 引擎本身不调用任何外部 API，也不会消耗你的 token。

## 项目结构

```
teaching-ppt-engine/
├── README.md                          ← 本文件
├── LICENSE                            ← 非商业使用许可
├── make_ppt.py                        ← 通用 PPT 渲染引擎（核心）
├── read_refs.py                       ← 参考 PPTX 文件读取器
├── examples/
│   ├── brief_template.md              ← brief 模板
│   └── slides_compliance.json         ← 20页示例配置
└── .gitignore
```

## 常见问题

### 为什么不直接生成 AI 图片？

因为每个人的 API Key 不同，且图片生成消耗较大。引擎把"文字排版"和"AI 图像"解耦：
13 页纯文字排版免费即时生成，7 页关键流程图按需用你自己的 API 生成。

### 生成的 PPTX 能改吗？

可以。`make_ppt.py` 输出的是标准 `.pptx` 文件，所有文字、颜色、位置都可以在 PowerPoint / WPS 中编辑。

### 支持英文 PPT 吗？

引擎内置字体为 PingFang SC（中文）。如需要英文 PPT，修改 `make_ppt.py` 中 `font_name` 即可。

### 为什么不能用作商业用途？

本工具基于开源学习目的开发，引用的咨询 PPT 方法论来自 [rw-consulting-ppt](https://github.com/Pikapika260214/rw-consulting-ppt)（MIT License）。若需商业使用，请自行替换视觉设计系统。

## 致谢

- PPT 设计方法论参考：[rw-consulting-ppt](https://github.com/Pikapika260214/rw-consulting-ppt)
- PPTX 生成基于：[python-pptx](https://github.com/scanny/python-pptx)

## License

本项目的所有代码和文档采用 **CC BY-NC 4.0**（知识共享 署名-非商业性使用 4.0 国际）许可。

- ✅ 个人学习、教学、研究 — 自由使用
- ✅ 修改、分享 — 需署名
- ❌ 商业用途 — 需另行授权

完整条款见 [LICENSE](./LICENSE) 文件。

# Doc_Fix

WPS/Word `.doc` 申报书刚性规则检查工具。

Doc_Fix 面向实际申报材料中常见的 `.doc` 文件。v1 会先将模板和目标文档转换为 `.docx` 工作副本，再检查最关键的刚性限制：各板块字数限制、表格格式和图片格式。

## v1 重点

- **`.doc` 优先**：支持输入 WPS/Word `.doc` 文件，原文件只读处理。
- **保留副本**：自动生成并保留 `.docx` 工作副本，便于复核。
- **字数检查**：从模板中识别“限xxxx字以内”“不超过xxxx字”“不包括表格”等规则。
- **表格检查**：检查表格数量、列数、宽度、边框、字体、对齐、行高和是否超页宽。
- **图片检查**：检查图片数量、所在板块、宽高、对齐、环绕方式和是否超页宽。
- **报告输出**：输出终端摘要、JSON 详情、按目录章节分组的 Markdown/HTML 人工核对报告。
- **AI 辅助**：可选调用 DeepSeek 生成摘要和人工处理建议。

## 技术栈

Python 3.10+ / LibreOffice (Linux) or Microsoft Word (Windows) / python-docx / lxml / click / rich / openai / python-dotenv / pytest

## 运行逻辑

一次检查从模板文档和目标文档开始，最终输出机器可读报告和人工核对报告：

1. **输入标准化**
   - 用户传入 `--template` 和 `--input`，可为 `.doc` 或 `.docx`。
   - `.doc` 文件先通过本机 Microsoft Word / Word Converter 转换为 `.docx` 工作副本。
   - 原始 `.doc` 只读处理，不会被覆盖或修改。

2. **模板规则抽取**
   - 程序读取模板工作副本中的段落、标题层级、表格、图片等结构。
   - 自动识别“限1500字以内”“限2000字以内（不包括表格）”“每个课题限3000字以内”等字数规则。
   - 可通过 `--config rules.json` 覆盖、补充或禁用自动识别的字数规则。

3. **目标文档结构抽取**
   - 程序读取目标工作副本，按 Word body 顺序识别段落、目录层级、表格和图片。
   - 中文章节标题会被解析成路径，例如 `第二部分 > 一、项目目标 > （二）考核指标`。
   - 表格和图片会记录所在章节、章节内序号、全文序号、附近标题、表题/图题和内容摘录。

4. **刚性规则检查**
   - 字数检查按“中文字符 + 英文/数字词”统计。
   - 表格检查数量、行列数、列宽、边框、字体、字号、对齐、行高和是否超页宽。
   - 图片检查数量、宽高、环绕方式、段落对齐和是否超页宽。
   - 检查结论由规则引擎决定，DeepSeek 只做可选解释，不改变通过/失败结果。

5. **报告输出**
   - 默认在 `--out-dir` 下生成 `report.json`、`report.md`、`report.html`。
   - JSON 面向后续自动修正、批处理和程序读取。
   - Markdown/HTML 面向人工核对，按目标文档章节分组，方便在 WPS 目录中定位。

## 技术实现

项目按“转换、抽取、检查、报告、CLI”分层，避免把 Word COM、文档解析和业务规则混在一起。

| 模块 | 职责 |
|------|------|
| `doc_fix.converter` | 使用 LibreOffice (Linux) / Microsoft Word COM (Windows) 将 `.doc` 标准化为 `.docx` 工作副本 |
| `doc_fix.extractor` | 使用 `python-docx` 和底层 OOXML 提取段落、标题、表格、图片和字数规则 |
| `doc_fix.checker` | 执行字数、表格、图片刚性检查，生成结构化 issue |
| `doc_fix.reporter` | 生成 JSON、Markdown、HTML 和终端摘要 |
| `doc_fix.ai` | 可选调用 DeepSeek 生成报告摘要和人工处理建议 |
| `doc_fix.cli` | 使用 `click` 编排命令行流程 |
| `doc_fix.model` | 定义不可变数据模型，作为模块之间的数据契约 |

核心技术选择：

- **Microsoft Word / Word Converter**：负责 `.doc -> .docx`，因为 `.doc` 是老二进制格式，`python-docx` 不能直接读取。
- **python-docx**：读取 `.docx` 的段落、表格、图片、样式和页面设置。
- **lxml / OOXML**：补充读取 `python-docx` 高层 API 不直接暴露的表格边框、宽度、标题大纲级别等信息。
- **click**：提供 `doc-fix check` 命令。
- **rich**：输出终端摘要。
- **openai + python-dotenv**：通过 OpenAI-compatible API 调用 DeepSeek，并从本地 `.env` 读取 API key。
- **pytest**：覆盖转换、规则抽取、字数统计、检查器、报告和 AI 辅助解析。

## 报告定位逻辑

人工报告的核心目标是“能在 WPS 里找到问题”。因此每条问题都会尽量附带：

- `章节路径`：从目标文档目录/标题层级推导。
- `怎么找`：例如“本章节第 2 个表格 / 全文第 5 个表格”或“目标文档段落 78”。
- `附近标题/表题`：例如“表4 项目目标、成果与考核指标表”。
- `期望/实际`：展示规则要求和当前检测结果。
- `内容摘录`：截取附近文本，帮助人工确认定位是否正确。

如果模板和目标不是同一类文档，程序会尽量给出章节定位，但表格匹配可能退回到“按全局序号兜底匹配”。这类提示表示结果需要人工复核。

## 开发环境

当前 Linux 版本在 `linux_ver` 分支开发，使用 `doc-fix` 虚拟环境：

```bash
python3 -m venv doc-fix
source doc-fix/bin/activate
pip install -r requirements.txt
pip install -e .
```

## 快速开始

安装依赖：

```bash
pip install -r requirements.txt
```

运行检查：

```bash
python -m doc_fix.cli.main check --template template.doc --input target.doc --out-dir output
```

命令会自动在 `--out-dir` 下生成：

- `report.json`
- `report.md`
- `report.html`
- `template.converted.docx`
- `input.converted.docx`

如果需要覆盖某个报告路径，可以单独指定：

```bash
python -m doc_fix.cli.main check --template template.doc --input target.doc --out-dir output --html review.html
```

使用 JSON 配置覆盖模板自动识别的字数规则：

```bash
python -m doc_fix.cli.main check --template template.doc --input target.doc --config rules.json --out-dir output
```

`rules.json` 示例：

```json
{
  "word_count_rules": [
    {
      "section_title": "项目简介",
      "limit": 1500,
      "exclude_tables": true
    },
    {
      "section_title": "研究基础",
      "limit": 1000,
      "enabled": false
    }
  ]
}
```

启用 DeepSeek 辅助摘要和处理建议：

```bash
python -m doc_fix.cli.main check --template template.doc --input target.doc --out-dir output --ai
```

本地 `.env` 示例见 [.env.example](./.env.example)。真实 API key 只应写入本地 `.env`，不要提交到仓库。

如果希望使用 `doc-fix` 命令，可在虚拟环境中以可编辑模式安装：

```bash
pip install -e .
doc-fix check --template template.doc --input target.doc --out-dir output
```

## 重要说明

`.doc` 是老二进制格式，`python-docx` 不能直接读取。Doc_Fix v1 采用“`.doc` 转 `.docx` 工作副本后检查”的方案。后续如果需要将 `.docx` 再导出为 `.doc`，可以通过 Microsoft Word 实现，但不能保证 100% 无损，必须导出后再次检查刚性规则。

## 项目状态

当前处于 Phase 1 — Linux 版本适配开发阶段，在 `linux_ver` 分支进行。详见 [progress.md](./progress.md)。

## 协作规则

参与本项目前请阅读 [AGENTS.md](./AGENTS.md)。

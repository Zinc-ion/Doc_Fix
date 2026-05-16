# Doc_Fix

WPS/Word `.doc` 申报书刚性规则检查工具。

Doc_Fix 面向实际申报材料中常见的 `.doc` 文件。v1 会先将模板和目标文档转换为 `.docx` 工作副本，再检查最关键的刚性限制：各板块字数限制、表格格式和图片格式。

## v1 重点

- **`.doc` 优先**：支持输入 WPS/Word `.doc` 文件，原文件只读处理。
- **保留副本**：自动生成并保留 `.docx` 工作副本，便于复核。
- **字数检查**：从模板中识别“限xxxx字以内”“不超过xxxx字”“不包括表格”等规则。
- **表格检查**：检查表格数量、列数、宽度、边框、字体、对齐、行高和是否超页宽。
- **图片检查**：检查图片数量、所在板块、宽高、对齐、环绕方式和是否超页宽。
- **报告输出**：输出终端摘要和 JSON 详情。

## 技术栈

Python 3.10+ / Microsoft Word or Word Converter / python-docx / lxml / click / rich / pytest

## 重要说明

`.doc` 是老二进制格式，`python-docx` 不能直接读取。Doc_Fix v1 采用“`.doc` 转 `.docx` 工作副本后检查”的方案。后续如果需要将 `.docx` 再导出为 `.doc`，可以通过 Microsoft Word 实现，但不能保证 100% 无损，必须导出后再次检查刚性规则。

## 项目状态

当前处于 Phase 0 — 文档与方向校准阶段。详见 [progress.md](./progress.md)。

## 协作规则

参与本项目前请阅读 [AGENTS.md](./AGENTS.md)。

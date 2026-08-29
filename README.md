# 大白核查 / DOCX Content Guard

本地比较两个 DOCX 版本，过滤常见格式归一化差异，并生成可人工复核的 HTML 内容变化报告。

这个工具处理一个具体风险：使用自动化工具调整论文格式时，正文可能同时发生非预期修改。它不会把文档上传到网络，也不调用外部模型。

## 主要能力

- 比较修改前后的正文段落
- 过滤全角与半角、空白、引号等格式差异
- 识别参考文献编号调整和 DOI 删除
- 标出新增、删除和改写内容
- 生成字符级差异 HTML 报告
- 使用 `--strict` 作为自动化流程的内容变更检查

## 安装

需要 Python 3.10 或更高版本。

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

Windows PowerShell 激活命令：

```powershell
.venv\Scripts\Activate.ps1
pip install -e .
```

## 使用

指定两个文件：

```bash
docx-content-guard 修改前.docx 修改后.docx
```

指定报告路径，并禁止自动打开浏览器：

```bash
docx-content-guard 修改前.docx 修改后.docx \
  --output report.html \
  --no-open
```

如果不传文件名，工具会在当前目录查找名称中分别含有“修改前”和“修改后”的 DOCX 文件。

## 检测边界

当前版本主要比较 Word 正文段落，不覆盖表格、文本框、批注、修订记录、脚注、尾注、页眉和页脚。报告中的“未发现”只表示当前覆盖范围内没有检测到变化，不能替代人工终稿检查。

相似段落通过启发式规则配对。段落大规模重排、重复段落或结构变化较大的文档，可能需要使用者进一步判断。

## 测试

```bash
python -m unittest discover -s tests -v
```

测试使用临时生成的 DOCX 文件，不包含真实论文内容。

## 项目角色

我定义了“格式调整不得改变正文”的检查目标、允许差异规则和人工复核报告形式，并使用 AI 编程工具协助实现和测试。最终判断仍由使用者根据原文完成。

## 许可证

[MIT License](LICENSE)

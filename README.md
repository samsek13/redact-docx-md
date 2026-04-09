# 文件脱敏工具

本地运行的文件脱敏工具，支持将文档中的敏感词替换为占位符，并可恢复原文。

## 功能

- **脱敏 (Redact)**：将敏感词替换为 `[REDACTED_N]` 占位符（支持大小写不敏感匹配）
- **恢复 (De-redact)**：根据映射文件还原原文
- 支持 `.txt`, `.md`, `.docx` 格式
- 批量处理（最多 10 个文件）
- Web UI 界面，本地运行

### DOCX 脱敏覆盖范围

脱敏功能覆盖 Word 文档中的以下所有内容：

- 正文段落及嵌套表格
- 页眉页脚
- 超链接 URL
- 批注（Comments）
- 脚注（Footnotes）
- 尾注（Endnotes）
- 文本框（Textbox）
- 修订追踪中的删除文本
- 结构化文档标签（SDT）

敏感词即使被拆分到不同格式元素（如超链接、修订标记等），也能被正确识别和脱敏。

## 安装

```bash
pip install python-docx
```

## 使用

```bash
python app.py
```

浏览器访问 http://localhost:8080

### 脱敏

1. 上传文件
2. 输入敏感词（一行一个）或勾选使用默认词库（两者可叠加使用）
3. 点击执行，下载脱敏文件和映射文件

### 恢复

1. 上传脱敏文件和对应的映射文件
2. 点击执行，下载恢复文件

## 默认词库

首次使用请从示例文件创建本地词库：

```bash
cp words.txt.sample words.txt
```

编辑 `words.txt` 文件，每行一个敏感词。

> 注意：`words.txt` 不会提交到仓库，请勿在其中存储真实敏感数据。

## 依赖

- `python-docx` - 处理 Word 文档
- Python 内置库：`http.server`, `json`, `pathlib`, `re`, `logging`
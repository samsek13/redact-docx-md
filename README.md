# 文件脱敏工具

本地运行的文件脱敏工具，支持将文档中的敏感词替换为占位符，并可恢复原文。

## 功能

- **脱敏 (Redact)**：将敏感词替换为 `[REDACTED_N]` 占位符
- **恢复 (De-redact)**：根据映射文件还原原文
- 支持 `.txt`, `.md`, `.docx` 格式
- 批量处理（最多 10 个文件）
- Web UI 界面，本地运行

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
2. 输入敏感词（一行一个）或勾选使用默认词库
3. 点击执行，下载脱敏文件和映射文件

### 恢复

1. 上传脱敏文件和对应的映射文件
2. 点击执行，下载恢复文件

## 默认词库

编辑 `words.txt` 文件，每行一个敏感词。

## 依赖

- `python-docx` - 处理 Word 文档
- Python 内置库：`http.server`, `json`, `pathlib`, `re`
"""
文件脱敏与恢复核心模块
支持格式：.txt, .docx
"""

import re
import json
from pathlib import Path
from datetime import datetime
from typing import Tuple, Dict, List

# 尝试导入 python-docx
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class Redactor:
    """脱敏器：负责文本脱敏和恢复"""

    def __init__(self):
        self.counter = 0
        self.mapping: Dict[str, str] = {}

    def reset(self):
        """重置计数器和映射"""
        self.counter = 0
        self.mapping = {}

    def redact_text(self, text: str, words: List[str]) -> str:
        """
        对文本进行脱敏处理

        Args:
            text: 原始文本
            words: 需要脱敏的词列表

        Returns:
            脱敏后的文本
        """
        if not text or not words:
            return text

        # 过滤空词并去重
        words = list(set(w for w in words if w and w.strip()))
        if not words:
            return text

        # 按词长降序排序，避免短词误匹配
        words_sorted = sorted(words, key=len, reverse=True)

        def replace_match(match):
            """替换匹配到的词"""
            matched_word = match.group(0)
            # 检查是否已有映射（同一词可能多次出现）
            for placeholder, original in self.mapping.items():
                if original == matched_word:
                    return placeholder

            # 创建新映射
            self.counter += 1
            placeholder = f"[REDACTED_{self.counter}]"
            self.mapping[placeholder] = matched_word
            return placeholder

        # 为每个词创建正则模式并替换
        result = text
        for word in words_sorted:
            # 使用 re.escape 处理特殊字符，确保精确匹配
            pattern = re.escape(word)
            result = re.sub(pattern, replace_match, result)

        return result

    def deredact_text(self, text: str, mapping: Dict[str, str]) -> str:
        """
        根据映射字典恢复脱敏文本

        Args:
            text: 脱敏后的文本
            mapping: 占位符到原文的映射字典

        Returns:
            恢复后的文本
        """
        if not text or not mapping:
            return text

        result = text
        # 按占位符长度降序排序，避免部分匹配问题
        sorted_mappings = sorted(mapping.items(), key=lambda x: len(x[0]), reverse=True)

        for placeholder, original in sorted_mappings:
            result = result.replace(placeholder, original)

        return result


def redact_txt(file_path: Path, words: List[str], output_dir: Path = None) -> Tuple[Path, Path]:
    """
    处理 txt 文件脱敏

    Args:
        file_path: 原始 txt 文件路径
        words: 需要脱敏的词列表
        output_dir: 输出目录，默认为原文件所在目录

    Returns:
        (脱敏文件路径, 映射文件路径)
    """
    if output_dir is None:
        output_dir = file_path.parent

    # 读取文件
    text = file_path.read_text(encoding='utf-8')

    # 脱敏处理
    redactor = Redactor()
    redacted_text = redactor.redact_text(text, words)

    # 生成输出文件名
    stem = file_path.stem
    output_file = output_dir / f"{stem}_redacted.txt"
    mapping_file = output_dir / f"{stem}_redacted.mapping.json"

    # 写入脱敏文件
    output_file.write_text(redacted_text, encoding='utf-8')

    # 写入映射文件
    mapping_data = {
        "original_filename": file_path.name,
        "created_at": datetime.now().isoformat(),
        "mappings": redactor.mapping
    }
    mapping_file.write_text(json.dumps(mapping_data, ensure_ascii=False, indent=2), encoding='utf-8')

    return output_file, mapping_file


def deredact_txt(file_path: Path, mapping_file: Path, output_dir: Path = None) -> Path:
    """
    恢复 txt 脱敏文件

    Args:
        file_path: 脱敏 txt 文件路径
        mapping_file: 映射文件路径
        output_dir: 输出目录，默认为原文件所在目录

    Returns:
        恢复后的文件路径
    """
    if output_dir is None:
        output_dir = file_path.parent

    # 读取文件
    text = file_path.read_text(encoding='utf-8')
    mapping_data = json.loads(mapping_file.read_text(encoding='utf-8'))

    # 恢复处理
    redactor = Redactor()
    restored_text = redactor.deredact_text(text, mapping_data['mappings'])

    # 生成输出文件名
    stem = file_path.stem.replace('_redacted', '')
    output_file = output_dir / f"{stem}_restored.txt"

    # 写入恢复文件
    output_file.write_text(restored_text, encoding='utf-8')

    return output_file


def redact_docx(file_path: Path, words: List[str], output_dir: Path = None) -> Tuple[Path, Path]:
    """
    处理 docx 文件脱敏

    Args:
        file_path: 原始 docx 文件路径
        words: 需要脱敏的词列表
        output_dir: 输出目录，默认为原文件所在目录

    Returns:
        (脱敏文件路径, 映射文件路径)
    """
    if not HAS_DOCX:
        raise ImportError("python-docx 未安装，无法处理 docx 文件。请运行: pip install python-docx")

    if output_dir is None:
        output_dir = file_path.parent

    # 打开文档
    doc = Document(str(file_path))
    redactor = Redactor()

    # 处理所有段落
    for paragraph in doc.paragraphs:
        if paragraph.text:
            redacted_text = redactor.redact_text(paragraph.text, words)
            # 保持原有格式，替换文本
            # 清除原有 runs
            for run in paragraph.runs:
                run.text = ""
            # 如果有 runs，在第一个 run 中设置新文本
            if paragraph.runs:
                paragraph.runs[0].text = redacted_text
            else:
                # 如果没有 runs，添加一个新 run
                paragraph.add_run(redacted_text)

    # 处理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text:
                        redacted_text = redactor.redact_text(paragraph.text, words)
                        if paragraph.runs:
                            paragraph.runs[0].text = redacted_text
                            for run in paragraph.runs[1:]:
                                run.text = ""
                        else:
                            paragraph.add_run(redacted_text)

    # 处理页眉页脚
    for section in doc.sections:
        # 页眉
        for paragraph in section.header.paragraphs:
            if paragraph.text:
                redacted_text = redactor.redact_text(paragraph.text, words)
                if paragraph.runs:
                    paragraph.runs[0].text = redacted_text
                    for run in paragraph.runs[1:]:
                        run.text = ""
                else:
                    paragraph.add_run(redacted_text)

        # 页脚
        for paragraph in section.footer.paragraphs:
            if paragraph.text:
                redacted_text = redactor.redact_text(paragraph.text, words)
                if paragraph.runs:
                    paragraph.runs[0].text = redacted_text
                    for run in paragraph.runs[1:]:
                        run.text = ""
                else:
                    paragraph.add_run(redacted_text)

    # 生成输出文件名
    stem = file_path.stem
    output_file = output_dir / f"{stem}_redacted.docx"
    mapping_file = output_dir / f"{stem}_redacted.mapping.json"

    # 保存文档
    doc.save(str(output_file))

    # 写入映射文件
    mapping_data = {
        "original_filename": file_path.name,
        "created_at": datetime.now().isoformat(),
        "mappings": redactor.mapping
    }
    mapping_file.write_text(json.dumps(mapping_data, ensure_ascii=False, indent=2), encoding='utf-8')

    return output_file, mapping_file


def deredact_docx(file_path: Path, mapping_file: Path, output_dir: Path = None) -> Path:
    """
    恢复 docx 脱敏文件

    Args:
        file_path: 脱敏 docx 文件路径
        mapping_file: 映射文件路径
        output_dir: 输出目录，默认为原文件所在目录

    Returns:
        恢复后的文件路径
    """
    if not HAS_DOCX:
        raise ImportError("python-docx 未安装，无法处理 docx 文件。请运行: pip install python-docx")

    if output_dir is None:
        output_dir = file_path.parent

    # 打开文档
    doc = Document(str(file_path))
    mapping_data = json.loads(mapping_file.read_text(encoding='utf-8'))
    redactor = Redactor()

    # 处理所有段落
    for paragraph in doc.paragraphs:
        if paragraph.text:
            restored_text = redactor.deredact_text(paragraph.text, mapping_data['mappings'])
            if paragraph.runs:
                paragraph.runs[0].text = restored_text
                for run in paragraph.runs[1:]:
                    run.text = ""
            else:
                paragraph.add_run(restored_text)

    # 处理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text:
                        restored_text = redactor.deredact_text(paragraph.text, mapping_data['mappings'])
                        if paragraph.runs:
                            paragraph.runs[0].text = restored_text
                            for run in paragraph.runs[1:]:
                                run.text = ""
                        else:
                            paragraph.add_run(restored_text)

    # 处理页眉页脚
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            if paragraph.text:
                restored_text = redactor.deredact_text(paragraph.text, mapping_data['mappings'])
                if paragraph.runs:
                    paragraph.runs[0].text = restored_text
                    for run in paragraph.runs[1:]:
                        run.text = ""
                else:
                    paragraph.add_run(restored_text)

        for paragraph in section.footer.paragraphs:
            if paragraph.text:
                restored_text = redactor.deredact_text(paragraph.text, mapping_data['mappings'])
                if paragraph.runs:
                    paragraph.runs[0].text = restored_text
                    for run in paragraph.runs[1:]:
                        run.text = ""
                else:
                    paragraph.add_run(restored_text)

    # 生成输出文件名
    stem = file_path.stem.replace('_redacted', '')
    output_file = output_dir / f"{stem}_restored.docx"

    # 保存文档
    doc.save(str(output_file))

    return output_file


def get_supported_extensions() -> List[str]:
    """获取支持的文件扩展名"""
    extensions = ['.txt', '.md']
    if HAS_DOCX:
        extensions.append('.docx')
    return extensions


def is_supported(file_path: Path) -> bool:
    """检查文件是否支持"""
    return file_path.suffix.lower() in get_supported_extensions()
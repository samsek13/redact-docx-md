"""
本地文件脱敏工具 - Web 服务器
使用 Python 内置 http.server 模块
支持 Python 3.13+（不依赖已移除的 cgi 模块）
"""

import os
import json
import re
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse, unquote, quote
from email.parser import BytesParser

# 导入核心模块
from redactor import (
    redact_txt, redact_docx, deredact_txt, deredact_docx,
    is_supported, get_supported_extensions, HAS_DOCX
)

# 配置
PORT = 8080
HOST = 'localhost'
TEMPLATE_DIR = Path(__file__).parent / 'templates'
UPLOAD_DIR = Path(__file__).parent / 'uploads'
OUTPUT_DIR = Path(__file__).parent / 'outputs'
WORDS_FILE = Path(__file__).parent / 'words.txt'

# 确保目录存在
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def parse_multipart(content_type: str, body: bytes) -> dict:
    """
    解析 multipart/form-data 表单数据
    返回: {'字段名': {'filename': str|None, 'data': bytes|str, 'type': 'file'|'text'}}
    """
    # 获取 boundary
    boundary_match = re.search(r'boundary=(.+)', content_type)
    if not boundary_match:
        raise ValueError("无法找到 boundary")

    boundary = boundary_match.group(1).strip()
    boundary_bytes = boundary.encode('utf-8')

    # 分割 parts
    parts = body.split(b'--' + boundary_bytes)
    result = {}

    for part in parts:
        if not part or part == b'--\r\n' or part == b'--':
            continue

        # 移除前导 \r\n
        if part.startswith(b'\r\n'):
            part = part[2:]

        # 分离 headers 和内容
        header_end = part.find(b'\r\n\r\n')
        if header_end == -1:
            continue

        headers_raw = part[:header_end]
        content = part[header_end + 4:]

        # 移除尾部 \r\n
        if content.endswith(b'\r\n'):
            content = content[:-2]

        # 解析 headers
        headers = {}
        for line in headers_raw.split(b'\r\n'):
            if ':' in line.decode('utf-8', errors='ignore'):
                key, value = line.decode('utf-8', errors='ignore').split(':', 1)
                headers[key.strip().lower()] = value.strip()

        # 获取字段名
        disposition = headers.get('content-disposition', '')
        name_match = re.search(r'name="([^"]+)"', disposition)
        if not name_match:
            continue

        field_name = name_match.group(1)

        # 检查是否是文件
        filename_match = re.search(r'filename="([^"]+)"', disposition)

        if filename_match:
            # 文件字段
            filename = filename_match.group(1)
            result[field_name] = {
                'filename': filename,
                'data': content,
                'type': 'file'
            }
        else:
            # 文本字段
            result[field_name] = {
                'filename': None,
                'data': content.decode('utf-8', errors='ignore'),
                'type': 'text'
            }

    return result


class RedactHandler(BaseHTTPRequestHandler):
    """HTTP 请求处理器"""

    def log_message(self, format, *args):
        """自定义日志格式"""
        print(f"[{self.address_string()}] {format % args}")

    def send_html(self, html: str, status=200):
        """发送 HTML 响应"""
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(html.encode('utf-8'))

    def send_json(self, data: dict, status=200):
        """发送 JSON 响应"""
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def send_file(self, file_path: Path):
        """发送文件响应"""
        filename = file_path.name
        # 对中文文件名进行 RFC 5987 编码
        encoded_filename = quote(filename, safe='')

        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        # 使用 filename* 参数支持非 ASCII 文件名
        self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{encoded_filename}")
        self.end_headers()
        with open(file_path, 'rb') as f:
            self.wfile.write(f.read())

    def do_GET(self):
        """处理 GET 请求"""
        parsed = urlparse(self.path)

        if parsed.path == '/' or parsed.path == '/index.html':
            # 返回主页
            html_path = TEMPLATE_DIR / 'index.html'
            if html_path.exists():
                html = html_path.read_text(encoding='utf-8')
                # 注入支持的格式信息
                extensions = get_supported_extensions()
                supported_str = ', '.join(extensions)
                html = html.replace('{{SUPPORTED_FORMATS}}', supported_str)
                self.send_html(html)
            else:
                self.send_html('<h1>模板文件不存在</h1>', 404)

        elif parsed.path == '/words':
            # 返回默认词库
            if WORDS_FILE.exists():
                words = WORDS_FILE.read_text(encoding='utf-8')
                self.send_json({'words': words, 'exists': True})
            else:
                self.send_json({'words': '', 'exists': False})

        elif parsed.path.startswith('/download/'):
            # 下载文件
            filename = unquote(parsed.path.replace('/download/', ''))
            file_path = OUTPUT_DIR / filename
            if file_path.exists():
                self.send_file(file_path)
            else:
                self.send_html('<h1>文件不存在</h1>', 404)

        elif parsed.path == '/clear':
            # 清理上传和输出目录
            cleanup_dirs()
            self.send_json({'status': 'ok', 'message': '临时文件已清理'})

        else:
            self.send_html('<h1>404 - 页面不存在</h1>', 404)

    def do_POST(self):
        """处理 POST 请求"""
        parsed = urlparse(self.path)

        if parsed.path == '/redact':
            self.handle_redact()
        elif parsed.path == '/deredact':
            self.handle_deredact()
        else:
            self.send_json({'error': '未知的操作'}, 404)

    def handle_redact(self):
        """处理脱敏请求"""
        try:
            # 获取 Content-Type 和 Content-Length
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', 0))

            # 读取请求体
            body = self.rfile.read(content_length)

            # 解析 multipart 表单
            form = parse_multipart(content_type, body)

            # 获取上传的文件
            files_data = form.get('files')
            if not files_data:
                self.send_json({'error': '请选择文件'}, 400)
                return

            # 处理多文件（可能是单个或列表）
            if isinstance(files_data, list):
                files = files_data
            else:
                files = [files_data]

            # 获取敏感词
            words_text = form.get('words', {}).get('data', '') if 'words' in form else ''
            use_default = form.get('use_default', {}).get('data', 'false') == 'true' if 'use_default' in form else False

            # 合并敏感词来源
            words = []
            if words_text:
                words.extend([w.strip() for w in words_text.split('\n') if w.strip()])

            if use_default and WORDS_FILE.exists():
                default_words = [w.strip() for w in WORDS_FILE.read_text(encoding='utf-8').split('\n') if w.strip()]
                words.extend(default_words)

            # 去重
            words = list(set(words))

            if not words:
                self.send_json({'error': '请输入敏感词或选择使用默认词库'}, 400)
                return

            # 检查文件数量
            if len(files) > 10:
                self.send_json({'error': '最多支持10个文件同时处理'}, 400)
                return

            # 处理每个文件
            results = []
            for file_item in files:
                filename = file_item['filename']
                file_path = UPLOAD_DIR / filename

                # 保存上传文件
                file_path.write_bytes(file_item['data'])

                # 检查格式支持
                if not is_supported(file_path):
                    results.append({
                        'filename': filename,
                        'error': f'不支持的文件格式，支持: {get_supported_extensions()}'
                    })
                    continue

                # 执行脱敏
                try:
                    if file_path.suffix.lower() in ('.txt', '.md'):
                        output_file, mapping_file = redact_txt(file_path, words, OUTPUT_DIR)
                    elif file_path.suffix.lower() == '.docx':
                        output_file, mapping_file = redact_docx(file_path, words, OUTPUT_DIR)

                    results.append({
                        'filename': filename,
                        'status': 'success',
                        'output': output_file.name,
                        'mapping': mapping_file.name
                    })
                except Exception as e:
                    results.append({
                        'filename': filename,
                        'error': str(e)
                    })

            self.send_json({'results': results})

        except Exception as e:
            self.send_json({'error': f'处理失败: {str(e)}'}, 500)

    def handle_deredact(self):
        """处理恢复请求"""
        try:
            # 获取 Content-Type 和 Content-Length
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', 0))

            # 读取请求体
            body = self.rfile.read(content_length)

            # 解析 multipart 表单
            form = parse_multipart(content_type, body)

            # 获取上传的脱敏文件
            redacted_data = form.get('redacted_file')
            if not redacted_data:
                self.send_json({'error': '请选择脱敏文件'}, 400)
                return

            redacted_path = UPLOAD_DIR / redacted_data['filename']
            redacted_path.write_bytes(redacted_data['data'])

            # 获取上传的映射文件
            mapping_data = form.get('mapping_file')
            if not mapping_data:
                self.send_json({'error': '请选择映射文件'}, 400)
                return

            mapping_path = UPLOAD_DIR / mapping_data['filename']
            mapping_path.write_bytes(mapping_data['data'])

            # 检查格式支持
            if not is_supported(redacted_path):
                self.send_json({
                    'error': f'不支持的文件格式，支持: {get_supported_extensions()}'
                }, 400)
                return

            # 执行恢复
            try:
                if redacted_path.suffix.lower() in ('.txt', '.md'):
                    output_file = deredact_txt(redacted_path, mapping_path, OUTPUT_DIR)
                elif redacted_path.suffix.lower() == '.docx':
                    output_file = deredact_docx(redacted_path, mapping_path, OUTPUT_DIR)

                self.send_json({
                    'status': 'success',
                    'output': output_file.name
                })
            except Exception as e:
                self.send_json({'error': str(e)}, 500)

        except Exception as e:
            self.send_json({'error': f'处理失败: {str(e)}'}, 500)


def cleanup_dirs():
    """清理临时目录"""
    for dir_path in [UPLOAD_DIR, OUTPUT_DIR]:
        if dir_path.exists():
            for file in dir_path.iterdir():
                if file.is_file():
                    file.unlink()


def run_server():
    """启动服务器"""
    print(f"\n{'='*50}")
    print("本地文件脱敏工具")
    print(f"{'='*50}")
    print(f"支持的格式: {', '.join(get_supported_extensions())}")
    print(f"服务器地址: http://{HOST}:{PORT}")
    print(f"默认词库: {WORDS_FILE.name} ({'存在' if WORDS_FILE.exists() else '不存在'})")
    print(f"{'='*50}")
    print("\n提示:")
    print("  - 在浏览器中打开上述地址即可使用")
    print("  - 按 Ctrl+C 停止服务器")
    print("\n")

    try:
        server = HTTPServer((HOST, PORT), RedactHandler)
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        cleanup_dirs()
    except Exception as e:
        print(f"\n服务器错误: {e}")
        cleanup_dirs()


if __name__ == '__main__':
    run_server()
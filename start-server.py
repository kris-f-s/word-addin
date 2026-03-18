#!/usr/bin/env python3
"""
Word Add-in 本地服务器启动脚本
适用于 macOS（系统自带 Python 3）
"""

import http.server
import socketserver
import ssl
import os
import sys
import webbrowser
import threading
from pathlib import Path

PORT = 3000
CERT_FILE = "localhost.pem"
KEY_FILE = "localhost-key.pem"

def generate_self_signed_cert():
    """生成自签名 SSL 证书（如果不存在）"""
    if os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE):
        print(f"✓ 找到现有证书文件")
        return True
    
    print("正在生成 SSL 证书...")
    try:
        import subprocess
        
        # 生成私钥
        subprocess.run([
            "openssl", "genrsa", "-out", KEY_FILE, "2048"
        ], check=True, capture_output=True)
        
        # 生成证书
        subprocess.run([
            "openssl", "req", "-new", "-x509", "-key", KEY_FILE,
            "-out", CERT_FILE, "-days", "365",
            "-subj", "/C=CN/ST=State/L=City/O=Organization/CN=localhost"
        ], check=True, capture_output=True)
        
        print("✓ SSL 证书生成成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 证书生成失败: {e}")
        return False
    except FileNotFoundError:
        print("✗ 未找到 openssl，请先安装 OpenSSL")
        print("  或使用已生成的证书文件")
        return False

def open_browser():
    """延迟打开浏览器（可选）"""
    import time
    time.sleep(1)
    url = f"https://localhost:{PORT}/taskpane.html"
    print(f"正在打开浏览器: {url}")
    webbrowser.open(url)

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    """自定义 HTTP 请求处理器"""
    
    # 忽略这些无害的404请求，不记录日志
    IGNORE_PATHS = [
        '/favicon.ico',
        '/.well-known/',
        '/apple-touch-icon',
        '/robots.txt'
    ]
    
    def log_message(self, format, *args):
        """自定义日志输出，忽略无害的404"""
        # log_message 的参数格式: format, *args
        # 对于 HTTP 请求日志，args 的格式：
        # - log_request: ('"GET /path HTTP/1.1"', '200', '1234')
        #   其中 args[0] 是请求行（字符串），args[1] 是状态码（字符串），args[2] 是大小（字符串）
        # - log_error: ('code 404, message File not found',)
        
        path = ''
        status_code = None
        
        try:
            if len(args) >= 2:
                # 标准 log_request 格式
                request_line = str(args[0])
                status_str = str(args[1])  # 状态码是字符串，需要转换
                
                # 提取路径
                if '"' in request_line:
                    # 格式: "GET /path HTTP/1.1"
                    parts = request_line.strip('"').split()
                    if len(parts) >= 2:
                        path = parts[1]
                
                # 转换状态码为整数
                try:
                    status_code = int(status_str)
                except (ValueError, TypeError):
                    pass
            
            elif len(args) == 1:
                # log_error 格式: ('code 404, message File not found',)
                first_arg = str(args[0])
                if 'code' in first_arg.lower():
                    import re
                    match = re.search(r'code (\d+)', first_arg)
                    if match:
                        status_code = int(match.group(1))
                        # 尝试提取路径
                        match_path = re.search(r'"GET ([^"]+)"', first_arg)
                        if match_path:
                            path = match_path.group(1)
        except Exception:
            # 如果解析失败，继续正常记录
            pass
        
        # 如果是404且是应该忽略的路径，不记录
        if status_code == 404 and path:
            for ignore_path in self.IGNORE_PATHS:
                if ignore_path in path:
                    return  # 不记录这个404
        
        # 其他情况正常记录
        super().log_message(format, *args)
    
    def end_headers(self):
        # 添加 CORS 头
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()
    
    def do_OPTIONS(self):
        """处理 OPTIONS 请求（CORS）"""
        self.send_response(200)
        self.end_headers()
    
    def do_GET(self):
        """处理 GET 请求，对404返回更友好的响应"""
        # 检查文件是否存在
        if self.path == '/favicon.ico':
            # 返回空的favicon，避免404
            self.send_response(204)  # No Content
            self.end_headers()
            return
        
        # 调用父类方法处理其他请求
        super().do_GET()

def main():
    """主函数"""
    # 切换到脚本所在目录
    script_dir = Path(__file__).parent
    os.chdir(script_dir)
    
    print("=" * 50)
    print("Word Add-in 本地服务器")
    print("=" * 50)
    print()
    
    # 检查证书
    if not generate_self_signed_cert():
        print("\n警告: 无法生成证书，将使用 HTTP（不安全）")
        use_https = False
    else:
        use_https = True
    
    # 创建服务器
    try:
        with socketserver.TCPServer(("", PORT), MyHTTPRequestHandler) as httpd:
            if use_https:
                # 配置 SSL
                context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                context.load_cert_chain(CERT_FILE, KEY_FILE)
                httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
                protocol = "HTTPS"
            else:
                protocol = "HTTP"
            
            print(f"✓ 服务器已启动")
            print(f"  协议: {protocol}")
            print(f"  地址: https://localhost:{PORT}")
            print(f"  目录: {script_dir}")
            print()
            print("提示:")
            print("1. 在 Word 中加载 manifest.xml")
            print("2. 如果出现证书警告，请选择'继续'或'信任'")
            print("3. 按 Ctrl+C 停止服务器")
            print()
            print("-" * 50)
            
            # 可选：自动打开浏览器
            # threading.Thread(target=open_browser, daemon=True).start()
            
            # 启动服务器
            httpd.serve_forever()
            
    except OSError as e:
        if e.errno == 48:  # Address already in use
            print(f"✗ 错误: 端口 {PORT} 已被占用")
            print(f"  请关闭占用该端口的程序，或修改脚本中的 PORT 变量")
        else:
            print(f"✗ 错误: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n服务器已停止")
        sys.exit(0)

if __name__ == "__main__":
    main()


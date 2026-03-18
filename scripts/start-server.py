#!/usr/bin/env python3
"""
Word Add-in 本地服务器启动脚本
适用于 macOS / Windows（Python 3）

仓库内：从项目根目录的 public/ 提供静态文件，证书在 certs/
离线包：脚本与 HTML 同目录时为扁平布局（无 public/ 子目录）
"""

import http.server
import socketserver
import ssl
import os
import sys
from pathlib import Path

PORT = 3000
CERT_BASENAME = "localhost.pem"
KEY_BASENAME = "localhost-key.pem"


def resolve_paths():
    """返回 (document_root, cert_dir)。"""
    script_dir = Path(__file__).resolve().parent
    public = script_dir.parent / "public"
    if public.is_dir():
        # 开发仓库：scripts/start-server.py
        project_root = script_dir.parent
        return project_root / "public", project_root / "certs"
    # 离线包等扁平目录：start-server.py 与 HTML 同级
    return script_dir, script_dir


def generate_self_signed_cert(cert_dir: Path) -> bool:
    """在 cert_dir 下生成自签名证书（若不存在）。"""
    cert_file = cert_dir / CERT_BASENAME
    key_file = cert_dir / KEY_BASENAME
    if cert_file.exists() and key_file.exists():
        print("✓ 找到现有证书文件")
        return True

    print("正在生成 SSL 证书...")
    try:
        import subprocess

        subprocess.run(
            ["openssl", "genrsa", "-out", str(key_file), "2048"],
            check=True,
            capture_output=True,
        )
        subprocess.run(
            [
                "openssl",
                "req",
                "-new",
                "-x509",
                "-key",
                str(key_file),
                "-out",
                str(cert_file),
                "-days",
                "365",
                "-subj",
                "/C=CN/ST=State/L=City/O=Organization/CN=localhost",
            ],
            check=True,
            capture_output=True,
        )
        print("✓ SSL 证书生成成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 证书生成失败: {e}")
        return False
    except FileNotFoundError:
        print("✗ 未找到 openssl，请先安装 OpenSSL")
        print("  或使用已生成的证书文件")
        return False


class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    IGNORE_PATHS = [
        "/favicon.ico",
        "/.well-known/",
        "/apple-touch-icon",
        "/robots.txt",
    ]

    def log_message(self, format, *args):
        path = ""
        status_code = None
        try:
            if len(args) >= 2:
                request_line = str(args[0])
                status_str = str(args[1])
                if '"' in request_line:
                    parts = request_line.strip('"').split()
                    if len(parts) >= 2:
                        path = parts[1]
                try:
                    status_code = int(status_str)
                except (ValueError, TypeError):
                    pass
            elif len(args) == 1:
                first_arg = str(args[0])
                if "code" in first_arg.lower():
                    import re

                    match = re.search(r"code (\d+)", first_arg)
                    if match:
                        status_code = int(match.group(1))
        except Exception:
            pass

        if status_code == 404 and path:
            for ignore_path in self.IGNORE_PATHS:
                if ignore_path in path:
                    return
        super().log_message(format, *args)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        if self.path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return
        super().do_GET()


def main():
    doc_root, cert_dir = resolve_paths()
    cert_dir.mkdir(parents=True, exist_ok=True)
    os.chdir(doc_root)

    print("=" * 50)
    print("Word Add-in 本地服务器")
    print("=" * 50)
    print()
    print(f"  站点目录: {doc_root}")

    cert_file = cert_dir / CERT_BASENAME
    key_file = cert_dir / KEY_BASENAME

    if not generate_self_signed_cert(cert_dir):
        print("\n警告: 无法生成证书，将使用 HTTP（不安全）")
        use_https = False
    else:
        use_https = True

    try:
        with socketserver.TCPServer(("", PORT), MyHTTPRequestHandler) as httpd:
            if use_https and cert_file.exists() and key_file.exists():
                context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                context.load_cert_chain(str(cert_file), str(key_file))
                httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
                protocol = "HTTPS"
            else:
                protocol = "HTTP"

            print(f"✓ 服务器已启动")
            print(f"  协议: {protocol}")
            print(f"  地址: https://localhost:{PORT}")
            print()
            print("提示:")
            print("1. 在 Word 中加载 manifest.xml / manifest-local.xml")
            print("2. 如出现证书警告，请选择「继续」或「信任」")
            print("3. 按 Ctrl+C 停止服务器")
            print()
            print("-" * 50)
            httpd.serve_forever()

    except OSError as e:
        if e.errno == 48:
            print(f"✗ 错误: 端口 {PORT} 已被占用")
        else:
            print(f"✗ 错误: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n服务器已停止")
        sys.exit(0)


if __name__ == "__main__":
    main()

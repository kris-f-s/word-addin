#!/usr/bin/env python3
"""离线包专用：脚本与 HTML 同目录，证书与脚本同目录。"""
import http.server
import socketserver
import ssl
import os
import sys
from pathlib import Path

PORT = 3000
CERT_FILE = "localhost.pem"
KEY_FILE = "localhost-key.pem"


def generate_self_signed_cert():
    d = Path(__file__).resolve().parent
    cf, kf = d / CERT_FILE, d / KEY_FILE
    if cf.exists() and kf.exists():
        print("✓ 找到现有证书文件")
        return True
    print("正在生成 SSL 证书...")
    try:
        import subprocess
        subprocess.run(["openssl", "genrsa", "-out", str(kf), "2048"], check=True, capture_output=True)
        subprocess.run(
            [
                "openssl", "req", "-new", "-x509", "-key", str(kf),
                "-out", str(cf), "-days", "365",
                "-subj", "/C=CN/ST=State/L=City/O=Organization/CN=localhost",
            ],
            check=True,
            capture_output=True,
        )
        print("✓ SSL 证书生成成功")
        return True
    except Exception as e:
        print(f"✗ 证书生成失败: {e}")
        return False


class H(http.server.SimpleHTTPRequestHandler):
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
    d = Path(__file__).resolve().parent
    os.chdir(d)
    print("=" * 50)
    print("Word Add-in 本地服务器（离线包）")
    print("=" * 50)
    use_https = generate_self_signed_cert()
    try:
        with socketserver.TCPServer(("", PORT), H) as httpd:
            if use_https and (d / CERT_FILE).exists():
                ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                ctx.load_cert_chain(str(d / CERT_FILE), str(d / KEY_FILE))
                httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
                print(f"✓ https://localhost:{PORT}")
            else:
                print(f"✓ http://localhost:{PORT}")
            print("在 Word 中加载 manifest-local.xml")
            httpd.serve_forever()
    except OSError as e:
        if e.errno == 48:
            print(f"端口 {PORT} 已被占用")
        else:
            print(e)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n已停止")
        sys.exit(0)


if __name__ == "__main__":
    main()

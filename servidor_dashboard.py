import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import http.server
import socketserver
import webbrowser
import socket
import os
import threading

PASTA = r"C:\Users\robson.noberto\Desktop\Controle Kips Motoristas"
PORTA = 8081
ARQUIVO_PRINCIPAL = "Dashboard_KPI_JT.html"


def obter_ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PASTA, **kwargs)

    def do_GET(self):
        if self.path == "/":
            self.path = f"/{ARQUIVO_PRINCIPAL}"
        super().do_GET()

    def log_message(self, _format, *args):
        cliente = self.client_address[0]
        recurso = args[0] if args else ""
        print(f"  Acesso de {cliente}  ->  {recurso}")


def abrir_navegador(url):
    import time
    time.sleep(1)
    webbrowser.open(url)


def porta_em_uso(porta):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("localhost", porta)) == 0


if __name__ == "__main__":
    os.chdir(PASTA)

    ip_local = obter_ip_local()
    url_local = f"http://localhost:{PORTA}"
    url_rede  = f"http://{ip_local}:{PORTA}"

    print("=" * 55)
    print("   SERVIDOR DASHBOARD - JET SP")
    print("=" * 55)
    print()

    if porta_em_uso(PORTA):
        print(f"   Servidor ja esta rodando na porta {PORTA}.")
        print()
        print("   Dashboard disponivel em:")
        print()
        print(f"   Voce (este PC) : {url_local}")
        print(f"   Rede local     : {url_rede}")
        print()
        print("   (nenhuma acao necessaria — ja esta no ar)")
        webbrowser.open(url_local)
    else:
        print("   Dashboard disponivel em:")
        print()
        print(f"   Voce (este PC) : {url_local}")
        print(f"   Rede local     : {url_rede}")
        print()
        print("   Compartilhe o link 'Rede local' com os colegas.")
        print("   Eles precisam estar na mesma rede Wi-Fi/corporativa.")
        print()
        print("   Pressione Ctrl+C para encerrar o servidor.")
        print("-" * 55)

        threading.Thread(target=abrir_navegador, args=(url_local,), daemon=True).start()

        with socketserver.TCPServer(("", PORTA), Handler) as httpd:
            httpd.allow_reuse_address = True
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\n[OK] Servidor encerrado.")

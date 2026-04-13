#!/usr/bin/env python3
"""
Launcher para o Social Media Analytics Scraper.

Localiza o Edge, abre com porta de debug, aguarda login do usuario
e chama scraper.py com os argumentos fornecidos.

Uso (pelo bat ou diretamente):
    python executar.py planilha.xlsx
    python executar.py "https://docs.google.com/spreadsheets/d/..."
"""

import sys
import os
import subprocess
import time
import socket
import shutil
from pathlib import Path

# Garante que o diretorio do script e o diretorio de trabalho
os.chdir(Path(__file__).resolve().parent)


# ---------------------------------------------------------------------------
def check_and_install_deps() -> bool:
    """Verifica pacotes necessarios e instala se faltarem."""
    try:
        import selenium       # noqa: F401
        import openpyxl      # noqa: F401
        import gspread        # noqa: F401
        import google_auth_oauthlib  # noqa: F401
        return True
    except ImportError:
        print("Instalando dependencias (aguarde)...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"]
        )
        if result.returncode != 0:
            print("[ERRO] Falha ao instalar dependencias. Verifique sua conexao.")
            return False
        print()
        return True


# ---------------------------------------------------------------------------
def find_edge() -> str | None:
    """Retorna o caminho do msedge.exe ou None se nao encontrado."""
    # 1) Via PATH
    edge = shutil.which("msedge")
    if edge:
        return edge

    # 2) Caminhos padrao do Windows
    prog_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    prog     = os.environ.get("ProgramFiles",       r"C:\Program Files")
    local    = os.environ.get("LOCALAPPDATA",       "")

    candidates = [
        Path(prog_x86) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(prog)     / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(local)    / "Microsoft" / "Edge" / "Application" / "msedge.exe",
    ]
    for c in candidates:
        if c.exists():
            return str(c)

    return None


# ---------------------------------------------------------------------------
def kill_edge() -> None:
    """Encerra processos Edge existentes."""
    subprocess.run(["taskkill", "/f", "/im", "msedge.exe"], capture_output=True)
    time.sleep(2)


# ---------------------------------------------------------------------------
def wait_for_port(port: int, timeout: int = 30) -> bool:
    """Aguarda ate a porta de debug responder."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except OSError:
            time.sleep(0.5)
    return False


# ---------------------------------------------------------------------------
def launch_edge(edge_exe: str, port: int) -> None:
    """Abre o Edge com a porta de debug remota."""
    local = os.environ.get("LOCALAPPDATA", "")
    user_data = str(Path(local) / "Microsoft" / "Edge" / "User Data")

    cmd = [
        edge_exe,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data}",
    ]
    subprocess.Popen(cmd)
    print("Edge aberto. Aguardando inicializar...")

    if wait_for_port(port, timeout=30):
        print(f"[OK] Edge respondendo na porta {port}.")
    else:
        print("[AVISO] Edge pode nao ter inicializado completamente. Continue mesmo assim.")


# ---------------------------------------------------------------------------
def main() -> None:
    port = 9222

    print("=" * 60)
    print("  Social Media Analytics Scraper  [browser: Edge]")
    print("=" * 60)
    print()

    if len(sys.argv) < 2:
        print("Arraste sua planilha ou URL do Google Sheets para cima deste .bat")
        print("ou execute pelo terminal:")
        print()
        print("  python executar.py planilha.xlsx")
        print('  python executar.py "https://docs.google.com/spreadsheets/d/..."')
        print()
        input("Pressione Enter para sair...")
        sys.exit(1)

    planilha = sys.argv[1]

    # ---------- Dependencias ----------
    if not check_and_install_deps():
        input("Pressione Enter para sair...")
        sys.exit(1)

    # ---------- Verifica se ja existe uma sessao Edge com debug ativo ----------
    if wait_for_port(port, timeout=2):
        print(f"[OK] Edge ja esta aberto com debug na porta {port}.")
        print("     Usando a sessao existente (logins preservados).")
    else:
        # Nao ha sessao ativa — precisa abrir o Edge com a porta de debug
        edge_exe = find_edge()

        if not edge_exe:
            print("[AVISO] Edge nao encontrado nos caminhos padrao.")
            print("O scraper tentara conectar ou abrir automaticamente.")
        else:
            # Encerra Edge comum (sem debug) se estiver rodando, para abrir com debug
            print("Edge nao esta em modo debug. Reabrindo com porta de debug...")
            kill_edge()

            print(f"Abrindo Edge: {edge_exe}")
            launch_edge(edge_exe, port)

            print()
            print("IMPORTANTE: Faca login no LinkedIn / Instagram / TikTok no Edge aberto.")
            print("Quando estiver logado, pressione Enter para iniciar a extracao.")
            print()
            input(">>> Pressione Enter para continuar...")

    # ---------- Executa o scraper ----------
    print()
    print("Iniciando extracao...")
    print(f"  Fonte: {planilha}")
    print()

    result = subprocess.run(
        [sys.executable, "scraper.py", planilha, "--porta-debug", str(port)]
    )

    print()
    if result.returncode != 0:
        print("[ERRO] O scraper terminou com erro. Veja scraper.log para detalhes.")
    else:
        print("Concluido! Verifique a planilha para os resultados.")

    print()
    input("Pressione Enter para sair...")


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()

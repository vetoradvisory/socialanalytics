#!/usr/bin/env python3
"""
Configura o acesso ao Google Sheets API (execute uma vez).

O que este script faz:
  1. Verifica se as dependências estão instaladas
  2. Guia o usuário para obter o arquivo credentials.json
  3. Abre o browser para autorização OAuth
  4. Salva o token para uso futuro (não precisa autorizar de novo)
  5. Testa o acesso à planilha informada

Uso:
  python configurar_google.py
  python configurar_google.py --planilha "https://docs.google.com/spreadsheets/d/..."
"""

import sys
import argparse
from pathlib import Path


# ---------------------------------------------------------------------------
def _check_deps() -> bool:
    missing = []
    for pkg in ("gspread", "google_auth_oauthlib"):
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg.replace("_", "-"))
    if missing:
        print(f"[!] Pacotes não instalados: {', '.join(missing)}")
        print("    Execute: pip install -r requirements.txt")
        return False
    return True


# ---------------------------------------------------------------------------
def _print_setup_instructions() -> None:
    print("""
Para usar o scraper com Google Sheets você precisa de um arquivo
'credentials.json' com as credenciais OAuth do Google Cloud.

Siga os passos abaixo (leva ~5 minutos):

╔══════════════════════════════════════════════════════════════╗
║  PASSO 1 — Abra o Google Cloud Console                       ║
║  https://console.cloud.google.com/                           ║
╠══════════════════════════════════════════════════════════════╣
║  PASSO 2 — Crie um projeto                                   ║
║  • Clique em "Selecionar projeto" (topo da tela)             ║
║  • Clique em "Novo Projeto"                                   ║
║  • Nome: Social Analytics (ou qualquer nome)                 ║
║  • Clique em "Criar"                                         ║
╠══════════════════════════════════════════════════════════════╣
║  PASSO 3 — Ative a Google Sheets API                         ║
║  • Menu lateral: "APIs e Serviços" > "Biblioteca"            ║
║  • Pesquise: "Google Sheets API"                             ║
║  • Clique em "Ativar"                                        ║
╠══════════════════════════════════════════════════════════════╣
║  PASSO 4 — Configure a Tela de Consentimento OAuth           ║
║  • Menu lateral: "APIs e Serviços" >                         ║
║    "Tela de consentimento OAuth"                             ║
║  • Tipo de usuário: "Externo" > "Criar"                      ║
║  • Preencha apenas o nome do app e seu e-mail                ║
║  • Salve e continue (clique "Salvar e continuar"             ║
║    em todas as telas)                                        ║
║  • Em "Usuários de teste", adicione SEU e-mail Google        ║
╠══════════════════════════════════════════════════════════════╣
║  PASSO 5 — Crie as credenciais OAuth                         ║
║  • Menu lateral: "APIs e Serviços" > "Credenciais"           ║
║  • Clique em "+ Criar Credenciais"                           ║
║  • Escolha "ID do cliente OAuth"                             ║
║  • Tipo de aplicativo: "App para computador" (Desktop app)   ║
║  • Nome: Scraper (qualquer nome)                             ║
║  • Clique em "Criar"                                         ║
╠══════════════════════════════════════════════════════════════╣
║  PASSO 6 — Baixe o arquivo JSON                              ║
║  • Na lista de credenciais OAuth 2.0, clique no ícone ↓      ║
║    (download) ao lado da credencial criada                   ║
║  • Renomeie o arquivo baixado para: credentials.json         ║
║  • Mova para a pasta do scraper:                             ║
║    C:\\Vetor\\SocialAnalytics\\credentials.json               ║
╚══════════════════════════════════════════════════════════════╝

Depois execute este script novamente:
  python configurar_google.py
""")


# ---------------------------------------------------------------------------
def authorize(credentials_file: str = "credentials.json",
              token_file: str = "token.json") -> object:
    """Executa o fluxo OAuth e retorna o cliente gspread autenticado."""
    import gspread

    print("[>] Abrindo o browser para autorização do Google...")
    print("    Faça login com a conta que tem acesso à planilha.")
    print()

    gc = gspread.oauth(
        credentials_filename=credentials_file,
        authorized_user_filename=token_file,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    print("[✓] Autorização concluída!")
    print(f"[✓] Token salvo em '{token_file}' — próximas execuções não precisarão de login.")
    return gc


# ---------------------------------------------------------------------------
def test_sheet_access(gc, sheet_url: str) -> None:
    """Testa o acesso a uma planilha específica."""
    import re
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", sheet_url)
    if not m:
        print(f"[!] URL inválida: {sheet_url}")
        return

    sheet_id = m.group(1)
    print(f"\n[>] Testando acesso à planilha...")
    try:
        spreadsheet = gc.open_by_key(sheet_id)
        tabs = [ws.title for ws in spreadsheet.worksheets()]
        print(f"[✓] Planilha acessada: '{spreadsheet.title}'")
        print(f"[✓] Abas disponíveis: {', '.join(tabs)}")

        if "posts" in tabs:
            ws = spreadsheet.worksheet("posts")
            # Verifica se coluna AE tem conteúdo
            ae_values = ws.col_values(31)  # coluna AE = 31
            urls = [v for v in ae_values[1:] if v.startswith("http")]
            print(f"[✓] Aba 'posts' encontrada — {len(urls)} URL(s) na coluna AE")
        else:
            print(f"[!] Aba 'posts' não encontrada. Abas existentes: {', '.join(tabs)}")
    except Exception as exc:
        print(f"[✗] Não foi possível acessar a planilha: {exc}")


# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Configura o acesso ao Google Sheets API."
    )
    parser.add_argument(
        "--planilha",
        default="https://docs.google.com/spreadsheets/d/1P8foX40o4yX1Ew2VobuBJKwIoGI9vWZ0KYmB2zMsVrU/edit",
        help="URL da planilha Google Sheets para testar o acesso",
    )
    args = parser.parse_args()

    print("=" * 60)
    print("  Configuração Google Sheets API")
    print("=" * 60)

    if not _check_deps():
        sys.exit(1)

    credentials_file = "credentials.json"
    token_file = "token.json"

    if not Path(credentials_file).exists():
        _print_setup_instructions()
        print(f"[!] Arquivo '{credentials_file}' não encontrado.")
        print("    Siga os passos acima e execute novamente.")
        sys.exit(1)

    print(f"[✓] credentials.json encontrado.")

    if Path(token_file).exists():
        print(f"[✓] token.json encontrado — já autorizado anteriormente.")
        import gspread
        try:
            gc = gspread.oauth(
                credentials_filename=credentials_file,
                authorized_user_filename=token_file,
            )
            print("[✓] Token válido.")
        except Exception:
            print("[!] Token expirado ou inválido — reautorizando...")
            Path(token_file).unlink(missing_ok=True)
            gc = authorize(credentials_file, token_file)
    else:
        gc = authorize(credentials_file, token_file)

    test_sheet_access(gc, args.planilha)

    print()
    print("=" * 60)
    print("  Configuração concluída!")
    print("  Agora execute o scraper:")
    print()
    print('  python scraper.py "' + args.planilha + '" --porta-debug 9222')
    print("=" * 60)


if __name__ == "__main__":
    main()

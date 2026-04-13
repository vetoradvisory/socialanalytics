#!/usr/bin/env python3
"""
Social Media Analytics Scraper
================================
Lê uma planilha Excel com URLs de analytics do LinkedIn, Instagram ou TikTok,
abre cada página no Chrome (usando o perfil existente para manter o login),
rola a tela para carregar todo o conteúdo e extrai as métricas desejadas.

Uso básico:
    python scraper.py planilha.xlsx

Uso avançado:
    python scraper.py planilha.xlsx --saida resultados.xlsx
    python scraper.py planilha.xlsx --perfil /home/user/.config/google-chrome
    python scraper.py --url https://www.linkedin.com/analytics/post-summary/urn:li:activity:...
"""

import re
import sys
import time
import logging
import argparse
import platform
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import config

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("scraper.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ===========================================================================
# UTILITÁRIOS DE TEXTO / NÚMERO
# ===========================================================================

def _clean_number(raw: str) -> str:
    """Normaliza um valor numérico bruto extraído da tela.

    Suporta formatos como: '1.234', '1,2K', '1.2M', '10 mil', '10K', '1 234'.
    Retorna o texto limpo mantendo o sufixo (K/M/B) se existir.
    """
    if not raw:
        return ""
    raw = raw.strip().replace("\u00a0", " ").replace("\u202f", " ")
    # Normalizar "mil" → K
    raw = re.sub(r"(?i)\bmil\b", "K", raw)
    m = re.search(r"([\d][\d\s.,]*[KkMmBb]?)", raw)
    return m.group(1).strip() if m else raw.strip()


def _extract_by_label(
    page_text: str,
    label_variants: list[str],
) -> str:
    """Procura um rótulo no texto da página e retorna o número adjacente.

    Aceita o número antes ou depois do rótulo (separado por nova linha ou espaço).
    """
    for label in label_variants:
        escaped = re.escape(label)
        patterns = [
            # número ANTES do rótulo (ex: "1.234\nImpressões")
            rf"([\d][\d\s.,]*[KkMmBb]?)\s*\n\s*{escaped}",
            # rótulo ANTES do número (ex: "Impressões\n1.234")
            rf"{escaped}\s*\n\s*([\d][\d\s.,]*[KkMmBb]?)",
            # mesmo na linha  (ex: "Impressões 1.234")
            rf"{escaped}\s+([\d][\d\s.,]*[KkMmBb]?)",
        ]
        for pat in patterns:
            m = re.search(pat, page_text, re.IGNORECASE)
            if m:
                return _clean_number(m.group(1))
    return ""


def _detect_sponsored(page_text: str) -> str:
    lower = page_text.lower()
    return "Sim" if any(t in lower for t in config.SPONSORED_TERMS) else "Não"


# ===========================================================================
# SCROLL
# ===========================================================================

def _scroll_full_page(driver: webdriver.Chrome) -> None:
    """Rola a página inteira de cima a baixo em passos, depois volta ao topo.

    Isso força o carregamento de conteúdo dinâmico (lazy-load, infinite scroll).
    Também tenta rolar dentro de containers internos que possam ter scroll próprio.
    """
    step = config.SCROLL_STEP_PX
    pause = config.SCROLL_PAUSE

    # Rola a janela principal
    current = 0
    while True:
        page_h = driver.execute_script("return document.body.scrollHeight")
        if current >= page_h:
            break
        current += step
        driver.execute_script(f"window.scrollTo(0, {current});")
        time.sleep(pause)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h > page_h:
            # Conteúdo novo carregou — continua
            page_h = new_h

    # Tenta rolar containers internos com overflow (modais, painéis de insights)
    driver.execute_script("""
        const candidates = document.querySelectorAll(
            '[class*="modal"], [class*="panel"], [class*="drawer"],\
             [class*="sheet"], [class*="overlay"], [class*="scroll"],\
             [class*="insight"], [role="dialog"], main, article'
        );
        candidates.forEach(el => {
            if (el.scrollHeight > el.clientHeight + 50) {
                el.scrollTo(0, el.scrollHeight);
            }
        });
    """)
    time.sleep(pause)

    # Volta ao topo
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.4)


# ===========================================================================
# EXTRATORES POR PLATAFORMA
# ===========================================================================

class _BaseExtractor:
    """Interface comum para todos os extratores."""

    label_map: dict[str, list[str]] = {}
    fields: list[str] = []

    def __init__(self, driver: webdriver.Chrome) -> None:
        self.driver = driver

    # ------------------------------------------------------------------
    def extract(self, url: str) -> dict:
        """Navega até a URL, espera, rola e extrai as métricas."""
        log.info("Navegando para: %s", url)
        self.driver.get(url)
        self._wait_for_content()
        time.sleep(config.POST_SCROLL_WAIT)
        self._expand_sections()
        _scroll_full_page(self.driver)
        time.sleep(config.POST_SCROLL_WAIT)

        page_text = self._get_page_text()
        data = self._extract_metrics(page_text)
        data["Patrocinado"] = _detect_sponsored(page_text)
        return data

    # ------------------------------------------------------------------
    def _wait_for_content(self) -> None:
        """Aguarda o carregamento inicial da página."""
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        except TimeoutException:
            log.warning("Timeout aguardando conteúdo da página")

    # ------------------------------------------------------------------
    def _expand_sections(self) -> None:
        """Clica em botões "Ver mais" ou "Mostrar mais" para expandir seções."""
        xpaths = [
            "//button[contains(., 'Ver mais')]",
            "//button[contains(., 'Mostrar mais')]",
            "//button[contains(., 'See more')]",
            "//button[contains(., 'Show more')]",
            "//a[contains(., 'Ver mais')]",
        ]
        for xp in xpaths:
            for btn in self.driver.find_elements(By.XPATH, xp)[:3]:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.6)
                except Exception:
                    pass

    # ------------------------------------------------------------------
    def _get_page_text(self) -> str:
        """Obtém o texto visível da página."""
        try:
            return self.driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            return ""

    # ------------------------------------------------------------------
    def _extract_metrics(self, page_text: str) -> dict:
        """Extrai todas as métricas usando text matching."""
        data: dict[str, str] = {}
        for field in self.fields:
            variants = self.label_map.get(field, [field.lower()])
            data[field] = _extract_by_label(page_text, variants)

        # Segunda passagem: tenta extração via DOM para campos ainda vazios
        missing = [f for f in self.fields if not data.get(f)]
        if missing:
            dom_data = self._extract_from_dom(missing)
            for f in missing:
                if dom_data.get(f):
                    data[f] = dom_data[f]

        return data

    # ------------------------------------------------------------------
    def _extract_from_dom(self, fields: list[str]) -> dict:
        """Varredura DOM: busca elementos que contenham apenas um número
        e cujo elemento pai (ou vizinho) contenha um rótulo conhecido."""
        data: dict[str, str] = {}
        try:
            # JS: retorna lista de {label, value} de todos os pares encontrados
            pairs: list[dict] = self.driver.execute_script("""
                const results = [];
                document.querySelectorAll('*').forEach(el => {
                    if (el.children.length > 0) return;
                    const txt = (el.textContent || '').trim();
                    if (/^[\d][\d\s.,]*[KkMmBb]?$/.test(txt) && txt.length < 20) {
                        const parent = el.parentElement;
                        if (parent) {
                            results.push({
                                value: txt,
                                context: parent.textContent.trim().slice(0, 200)
                            });
                        }
                    }
                });
                return results;
            """) or []
        except Exception:
            return data

        for field in fields:
            variants = self.label_map.get(field, [field.lower()])
            for pair in pairs:
                ctx = pair.get("context", "").lower()
                if any(v.lower() in ctx for v in variants):
                    data[field] = _clean_number(pair.get("value", ""))
                    break

        return data


# ---------------------------------------------------------------------------
class LinkedInExtractor(_BaseExtractor):
    label_map = config.LINKEDIN_LABEL_MAP
    fields = config.LINKEDIN_FIELDS

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "[data-test-analytics-summary-metric]")
                    ),
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "[class*='analytics']")
                    ),
                    EC.presence_of_element_located((By.TAG_NAME, "main")),
                )
            )
        except TimeoutException:
            log.warning("LinkedIn: timeout aguardando painel de analytics")


# ---------------------------------------------------------------------------
class InstagramExtractor(_BaseExtractor):
    label_map = config.INSTAGRAM_LABEL_MAP
    fields = config.INSTAGRAM_FIELDS

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(., 'Visualizações') or contains(., 'Views')]")
                    ),
                    EC.presence_of_element_located((By.TAG_NAME, "article")),
                    EC.presence_of_element_located((By.TAG_NAME, "main")),
                )
            )
        except TimeoutException:
            log.warning("Instagram: timeout aguardando painel de insights")

    def _expand_sections(self) -> None:
        super()._expand_sections()
        # Fecha popups de cookies / notificações que bloqueiam o conteúdo
        close_xpaths = [
            "//button[@aria-label='Close']",
            "//button[@aria-label='Fechar']",
            "//button[contains(., 'Não agora')]",
            "//button[contains(., 'Not Now')]",
            "//button[contains(., 'Aceitar')]",
            "//button[contains(., 'Accept')]",
        ]
        for xp in close_xpaths:
            for btn in self.driver.find_elements(By.XPATH, xp)[:1]:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
                except Exception:
                    pass


# ---------------------------------------------------------------------------
class TikTokExtractor(_BaseExtractor):
    label_map = config.TIKTOK_LABEL_MAP
    fields = config.TIKTOK_FIELDS

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "[class*='analytics']")
                    ),
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "[class*='metric']")
                    ),
                    EC.presence_of_element_located((By.TAG_NAME, "main")),
                )
            )
        except TimeoutException:
            log.warning("TikTok: timeout aguardando painel de analytics")


# ===========================================================================
# DETECÇÃO DE PLATAFORMA
# ===========================================================================

_PLATFORM_MAP = {
    "linkedin.com": "linkedin",
    "instagram.com": "instagram",
    "tiktok.com": "tiktok",
}

_EXTRACTOR_MAP = {
    "linkedin": LinkedInExtractor,
    "instagram": InstagramExtractor,
    "tiktok": TikTokExtractor,
}

_ALL_FIELDS_MAP = {
    "linkedin": config.LINKEDIN_FIELDS,
    "instagram": config.INSTAGRAM_FIELDS,
    "tiktok": config.TIKTOK_FIELDS,
}


def _detect_platform(url: str) -> Optional[str]:
    lower = url.lower()
    for domain, name in _PLATFORM_MAP.items():
        if domain in lower:
            return name
    return None


# ===========================================================================
# CHROME DRIVER
# ===========================================================================

def _default_chrome_profile() -> str:
    system = platform.system()
    home = Path.home()
    if system == "Windows":
        return str(home / "AppData" / "Local" / "Google" / "Chrome" / "User Data")
    if system == "Darwin":
        return str(home / "Library" / "Application Support" / "Google" / "Chrome")
    return str(home / ".config" / "google-chrome")


def create_driver(
    profile_path: Optional[str] = None,
    debug_port: Optional[int] = None,
    headless: bool = False,
) -> webdriver.Chrome:
    """Cria o WebDriver do Chrome.

    Ordem de preferência:
    1. Conecta a uma instância Chrome já aberta (debug_port).
    2. Usa o perfil Chrome do usuário (mantém sessões de login).
    3. Abre Chrome "limpo" (sem perfil, sem login).
    """
    opts = Options()

    if debug_port:
        opts.add_experimental_option(
            "debuggerAddress", f"127.0.0.1:{debug_port}"
        )
        log.info("Conectando ao Chrome existente na porta %d", debug_port)
    else:
        path = profile_path or _default_chrome_profile()
        opts.add_argument(f"--user-data-dir={path}")
        opts.add_argument("--profile-directory=Default")
        log.info("Usando perfil Chrome: %s", path)

    # Anti-detecção de automação
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    # Estabilidade
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1440,900")

    if headless:
        opts.add_argument("--headless=new")

    try:
        from webdriver_manager.chrome import ChromeDriverManager

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
    except Exception:
        # Fallback: chromedriver no PATH
        driver = webdriver.Chrome(options=opts)

    # Remove a flag navigator.webdriver
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.implicitly_wait(2)
    return driver


# ===========================================================================
# LEITURA DO EXCEL
# ===========================================================================

_URL_HEADER_ALIASES = {"url", "link", "endereço", "address", "urls", "links"}


def read_urls_from_excel(path: str) -> list[str]:
    """Lê todas as URLs da planilha Excel.

    Procura a coluna cujo cabeçalho seja "URL", "Link", etc.
    Se não houver cabeçalho, assume que a coluna A inteira contém URLs.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    url_col = None
    start_row = 1

    # Procura cabeçalho na linha 1
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and str(val).strip().lower() in _URL_HEADER_ALIASES:
            url_col = col
            start_row = 2
            break

    if url_col is None:
        url_col = 1

    urls = []
    for row in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row, column=url_col).value
        if val:
            stripped = str(val).strip()
            if stripped.startswith("http"):
                urls.append(stripped)

    log.info("Planilha '%s': %d URLs encontradas", path, len(urls))
    return urls


# ===========================================================================
# ESCRITA DO EXCEL
# ===========================================================================

_BASE_COLS = ["Plataforma", "URL", "Patrocinado", "Data Extração"]


def _header_fill(hex_color: str) -> PatternFill:
    return PatternFill(
        start_color=hex_color, end_color=hex_color, fill_type="solid"
    )


def _write_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict],
    color: str,
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    fill = _header_fill(color)
    white_font = Font(bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Cabeçalho
    for ci, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = white_font
        cell.fill = fill
        cell.alignment = center
    ws.row_dimensions[1].height = 32

    # Dados
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    for ri, row_data in enumerate(rows, 2):
        for ci, col in enumerate(columns, 1):
            cell = ws.cell(row=ri, column=ci, value=row_data.get(col, ""))
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0:
                cell.fill = alt_fill

    # Largura das colunas
    for ci, col in enumerate(columns, 1):
        letter = ws.cell(row=1, column=ci).column_letter
        max_w = max(
            len(col) + 4,
            *(
                len(str(ws.cell(row=ri, column=ci).value or "")) + 2
                for ri in range(2, len(rows) + 2)
            ),
        )
        ws.column_dimensions[letter].width = min(max_w, 45)


def save_results_to_excel(results: list[dict], output_path: str) -> None:
    """Grava os resultados em um Excel com abas separadas por plataforma
    e uma aba 'Resumo' com todos os dados."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove aba padrão

    # Agrupa por plataforma
    by_platform: dict[str, list[dict]] = {}
    for r in results:
        p = r.get("Plataforma", "Outros").lower()
        by_platform.setdefault(p, []).append(r)

    # Aba por plataforma
    for platform, rows in by_platform.items():
        fields = _ALL_FIELDS_MAP.get(platform, [])
        columns = _BASE_COLS + [f for f in fields if f not in _BASE_COLS]
        # Inclui campo "Erro" se presente em alguma linha
        if any("Erro" in r for r in rows):
            columns = columns + ["Erro"]
        color = config.PLATFORM_COLORS.get(platform, config.PLATFORM_COLORS["outros"])
        _write_sheet(wb, platform.capitalize(), columns, rows, color)

    # Aba Resumo (todos juntos)
    all_keys: list[str] = list(
        dict.fromkeys(
            _BASE_COLS
            + config.LINKEDIN_FIELDS
            + config.INSTAGRAM_FIELDS
            + config.TIKTOK_FIELDS
            + ["Erro"]
        )
    )
    _write_sheet(wb, "Resumo", all_keys, results, "444444")
    # Move Resumo para primeiro lugar
    wb.move_sheet("Resumo", offset=-len(wb.sheetnames) + 1)

    wb.save(output_path)
    log.info("Resultados gravados em: %s", output_path)


# ===========================================================================
# ORQUESTRADOR PRINCIPAL
# ===========================================================================

class SocialMediaScraper:
    """Orquestra a extração de analytics para múltiplas URLs."""

    def __init__(
        self,
        profile_path: Optional[str] = None,
        debug_port: Optional[int] = None,
        headless: bool = False,
    ) -> None:
        self.profile_path = profile_path
        self.debug_port = debug_port
        self.headless = headless
        self.driver: Optional[webdriver.Chrome] = None

    def __enter__(self) -> "SocialMediaScraper":
        self.driver = create_driver(
            profile_path=self.profile_path,
            debug_port=self.debug_port,
            headless=self.headless,
        )
        return self

    def __exit__(self, *_) -> None:
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass

    # ------------------------------------------------------------------
    def scrape_url(self, url: str) -> dict:
        """Extrai métricas de uma única URL."""
        assert self.driver, "Use como context manager: with SocialMediaScraper() as s:"

        platform = _detect_platform(url)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

        if platform is None:
            log.warning("Plataforma não reconhecida para: %s", url)
            return {
                "Plataforma": "Desconhecida",
                "URL": url,
                "Patrocinado": "Não",
                "Data Extração": timestamp,
                "Erro": "Plataforma não reconhecida. URLs suportadas: LinkedIn, Instagram, TikTok.",
            }

        extractor_cls = _EXTRACTOR_MAP[platform]
        try:
            extractor = extractor_cls(self.driver)
            data = extractor.extract(url)
            data["Plataforma"] = platform.capitalize()
            data["URL"] = url
            data["Data Extração"] = timestamp
            return data
        except WebDriverException as exc:
            log.error("Erro ao processar %s: %s", url, exc)
            return {
                "Plataforma": platform.capitalize(),
                "URL": url,
                "Patrocinado": "Não",
                "Data Extração": timestamp,
                "Erro": str(exc)[:200],
            }

    # ------------------------------------------------------------------
    def scrape_excel(
        self,
        input_path: str,
        output_path: Optional[str] = None,
    ) -> str:
        """Processa todas as URLs da planilha e salva os resultados."""
        urls = read_urls_from_excel(input_path)
        if not urls:
            print("Nenhuma URL encontrada na planilha.")
            return ""

        results = []
        total = len(urls)
        for i, url in enumerate(urls, 1):
            print(f"  [{i}/{total}] {url}")
            result = self.scrape_url(url)
            results.append(result)
            if i < total:
                time.sleep(config.BETWEEN_URLS_PAUSE)

        if not output_path:
            stem = Path(input_path).stem
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = str(Path(input_path).parent / f"{stem}_resultados_{ts}.xlsx")

        save_results_to_excel(results, output_path)
        return output_path


# ===========================================================================
# CLI
# ===========================================================================

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="scraper.py",
        description="Extrai métricas de analytics do LinkedIn, Instagram e TikTok.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  # Processar planilha com perfil Chrome existente (recomendado)
  python scraper.py planilha.xlsx

  # Especificar perfil Chrome manualmente
  python scraper.py planilha.xlsx --perfil "/home/user/.config/google-chrome"

  # Conectar ao Chrome já aberto (inicie o Chrome com --remote-debugging-port=9222)
  python scraper.py planilha.xlsx --porta-debug 9222

  # Arquivo de saída personalizado
  python scraper.py planilha.xlsx --saida relatorio.xlsx

  # Processar uma única URL (sem planilha)
  python scraper.py --url "https://www.linkedin.com/analytics/post-summary/urn:li:activity:..."

Dica: se o Chrome estiver aberto, feche-o antes de executar (ou use --porta-debug).
        """,
    )
    p.add_argument(
        "planilha",
        nargs="?",
        help="Caminho para a planilha Excel com URLs (coluna 'URL' ou coluna A)",
    )
    p.add_argument(
        "--url",
        metavar="URL",
        help="Processar uma única URL diretamente (sem planilha)",
    )
    p.add_argument(
        "--saida", "-o",
        metavar="ARQUIVO",
        help="Arquivo Excel de saída (padrão: <planilha>_resultados_<timestamp>.xlsx)",
    )
    p.add_argument(
        "--perfil",
        metavar="CAMINHO",
        help="Caminho para o diretório User Data do Chrome",
    )
    p.add_argument(
        "--porta-debug",
        type=int,
        metavar="PORTA",
        help="Porta de depuração remota do Chrome (ex: 9222)",
    )
    p.add_argument(
        "--headless",
        action="store_true",
        help="Executar sem janela de browser (não funciona para páginas que exigem login interativo)",
    )
    return p


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if not args.planilha and not args.url:
        parser.print_help()
        sys.exit(1)

    print("=" * 60)
    print("  Social Media Analytics Scraper")
    print("=" * 60)

    with SocialMediaScraper(
        profile_path=args.perfil,
        debug_port=args.porta_debug,
        headless=args.headless,
    ) as scraper:

        if args.url:
            # Modo URL única
            print(f"\nProcessando: {args.url}\n")
            result = scraper.scrape_url(args.url)
            print("\nResultados:")
            print("-" * 40)
            for k, v in result.items():
                if v:
                    print(f"  {k:<50} {v}")
            print("-" * 40)

            if args.saida:
                save_results_to_excel([result], args.saida)
                print(f"\nSalvo em: {args.saida}")

        else:
            # Modo planilha
            if not Path(args.planilha).exists():
                print(f"\nErro: arquivo não encontrado — {args.planilha}")
                sys.exit(1)

            print(f"\nPlanilha: {args.planilha}")
            print("Iniciando extração...\n")
            output = scraper.scrape_excel(args.planilha, args.saida)
            if output:
                print(f"\n✓ Concluído. Resultados em: {output}")


if __name__ == "__main__":
    main()

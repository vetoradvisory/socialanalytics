#!/usr/bin/env python3
"""
Social Media Analytics Scraper  —  versão anti-bot
====================================================
Lê uma planilha Excel com URLs de analytics, abre cada página no Chrome
simulando comportamento humano (rolagem variável, pausas aleatórias,
movimentos de mouse) para evitar bloqueio por bot-detection.

Resultado escrito de volta na mesma planilha, nas colunas designadas:
  • LinkedIn   → AF : AN  (9 campos)
  • Instagram  → AO : AW  (9 campos)
  • TikTok     → AX : BF  (9 campos)

Uso:
    python scraper.py planilha.xlsx
    python scraper.py planilha.xlsx --porta-debug 9222
    python scraper.py --url "https://www.linkedin.com/analytics/..."
"""

import random
import re
import sys
import time
import logging
import argparse
import platform as _platform_module
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import config

# ---------------------------------------------------------------------------
# undetected-chromedriver (opcional — fallback para selenium puro se ausente)
# ---------------------------------------------------------------------------
try:
    import undetected_chromedriver as uc
    _UC_AVAILABLE = True
except ImportError:
    _UC_AVAILABLE = False

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

if not _UC_AVAILABLE:
    log.warning(
        "undetected-chromedriver não encontrado — install: pip install undetected-chromedriver. "
        "Usando selenium com patches manuais."
    )


# ===========================================================================
# COMPORTAMENTO HUMANO
# ===========================================================================

class HumanBehavior:
    """
    Simula padrões de interação humana no browser para evitar detecção
    como bot pelos sistemas anti-automação do LinkedIn, Instagram e TikTok.

    Técnicas aplicadas:
      • Delays aleatórios com distribuição gaussiana
      • Scroll de velocidade variável com recuos e pausas de "leitura"
      • Movimentos de mouse para posições aleatórias
      • Viewport com resolução comum ligeiramente variada
      • User-Agent realista de versão recente do Chrome
    """

    # User-Agents reais de Chrome (Windows / Mac) — versões recentes
    _USER_AGENTS: list[str] = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    ]

    # Resoluções de monitor comuns
    _RESOLUTIONS: list[tuple[int, int]] = [
        (1366, 768), (1440, 900), (1536, 864), (1920, 1080), (1280, 800),
    ]

    # ------------------------------------------------------------------
    @classmethod
    def user_agent(cls) -> str:
        return random.choice(cls._USER_AGENTS)

    @classmethod
    def resolution(cls) -> tuple[int, int]:
        """Retorna uma resolução comum com pequena variação (+/-8 px)."""
        w, h = random.choice(cls._RESOLUTIONS)
        return w + random.randint(-8, 8), h + random.randint(-8, 8)

    # ------------------------------------------------------------------
    @staticmethod
    def sleep(min_s: float, max_s: float) -> None:
        """Pausa com distribuição gaussiana — mais realista que uniform."""
        mu = (min_s + max_s) / 2
        sigma = (max_s - min_s) / 4
        t = max(min_s, min(max_s, random.gauss(mu, sigma)))
        time.sleep(t)

    @staticmethod
    def micro() -> None:
        """Micro-pausa (60–350 ms) simulando latência de resposta."""
        time.sleep(random.uniform(0.06, 0.35))

    @classmethod
    def between_pages(cls) -> None:
        """Pausa entre URLs — simula o usuário decidindo o próximo link."""
        cls.sleep(config.BETWEEN_URLS_MIN, config.BETWEEN_URLS_MAX)

    # ------------------------------------------------------------------
    @staticmethod
    def move_mouse(driver: webdriver.Chrome, n: int = 4) -> None:
        """Move o mouse para N posições aleatórias na viewport."""
        try:
            size = driver.get_window_size()
            w, h = size["width"], size["height"]
            body = driver.find_element(By.TAG_NAME, "body")
            actions = ActionChains(driver)
            for _ in range(n):
                ox = random.randint(int(w * 0.05), int(w * 0.95)) - w // 2
                oy = random.randint(int(h * 0.05), int(h * 0.85)) - h // 2
                actions.move_to_element_with_offset(body, ox, oy)
                actions.pause(random.uniform(0.06, 0.3))
            actions.perform()
        except Exception:
            pass  # movimento de mouse não é crítico

    # ------------------------------------------------------------------
    @classmethod
    def scroll_full_page(cls, driver: webdriver.Chrome) -> None:
        """
        Rola a página inteira de forma humana:
          - velocidade variável (rápida ↔ lenta)
          - pausas de "leitura" aleatórias
          - recuos ocasionais (releitura)
          - rola containers internos (modais de insights)
          - sobe ao topo ao final
        """
        # Pausa inicial: o usuário "vê" a página antes de rolar
        cls.sleep(1.2, 3.0)
        cls.move_mouse(driver, n=random.randint(2, 4))

        page_h = driver.execute_script("return document.body.scrollHeight")
        pos = 0
        iterations = 0
        max_iter = 25  # segurança contra scroll infinito

        while pos < page_h and iterations < max_iter:
            iterations += 1

            # Recuo ocasional — simula releitura
            if random.random() < config.SCROLL_BACK_PROB and pos > 400:
                back = random.randint(*config.SCROLL_BACK_PX)
                pos = max(0, pos - back)
                driver.execute_script(f"window.scrollTo(0, {pos});")
                time.sleep(random.uniform(0.3, 0.8))

            # Scroll
            step = random.randint(config.SCROLL_STEP_MIN, config.SCROLL_STEP_MAX)
            pos += step
            driver.execute_script(f"window.scrollTo(0, {pos});")

            # Pausa de leitura ou pausa rápida
            if random.random() < config.SCROLL_READ_PROB:
                cls.sleep(*config.SCROLL_READ_PAUSE)
            else:
                cls.sleep(*config.SCROLL_FAST_PAUSE)

            # Conteúdo novo pode ter carregado (lazy-load)
            new_h = driver.execute_script("return document.body.scrollHeight")
            if new_h > page_h:
                page_h = new_h

        # Rola containers internos (painéis, modais de insights)
        cls._scroll_internals(driver)

        # Retorno ao topo suave
        cls.sleep(0.4, 1.0)
        steps_up = random.randint(3, 6)
        for i in range(steps_up, 0, -1):
            driver.execute_script(f"window.scrollTo(0, {pos * i // steps_up});")
            time.sleep(random.uniform(0.05, 0.15))
        driver.execute_script("window.scrollTo(0, 0);")
        cls.sleep(0.3, 0.7)

    @staticmethod
    def _scroll_internals(driver: webdriver.Chrome) -> None:
        """Rola elementos com scroll próprio (modais, side-panels)."""
        try:
            driver.execute_script("""
                const sel = [
                    '[class*="modal"]','[class*="panel"]','[class*="drawer"]',
                    '[class*="sheet"]','[class*="overlay"]','[class*="insight"]',
                    '[class*="scroll"]','[role="dialog"]','main','article'
                ].join(',');
                document.querySelectorAll(sel).forEach(el => {
                    if (el.scrollHeight > el.clientHeight + 60) {
                        el.scrollTop = el.scrollHeight;
                    }
                });
            """)
            time.sleep(random.uniform(0.5, 1.2))
        except Exception:
            pass


# ===========================================================================
# DRIVER CHROME
# ===========================================================================

def _default_profile() -> str:
    system = _platform_module.system()
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
    """
    Cria o WebDriver do Chrome com máxima resistência à detecção.

    Prioridade:
      1. Porta de debug → conecta ao Chrome já aberto pelo usuário (mais seguro).
      2. undetected-chromedriver + perfil → usa sessão existente sem markers de bot.
      3. Selenium puro + patches manuais → fallback se uc não estiver instalado.
    """
    if debug_port:
        # Modo "attach": conecta ao Chrome aberto pelo usuário, sem nenhuma
        # alteração de perfil — máxima naturalidade.
        opts = Options()
        opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
        log.info("Conectando ao Chrome existente — porta %d", debug_port)
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            return webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=opts
            )
        except Exception:
            return webdriver.Chrome(options=opts)

    profile = profile_path or _default_profile()
    w, h = HumanBehavior.resolution()
    ua = HumanBehavior.user_agent()

    if _UC_AVAILABLE:
        return _create_uc_driver(profile, w, h, ua, headless)
    return _create_selenium_driver(profile, w, h, ua, headless)


def _create_uc_driver(
    profile: str, w: int, h: int, ua: str, headless: bool
) -> webdriver.Chrome:
    """Driver via undetected-chromedriver — remove automaticamente os markers
    de automação que o Selenium deixa no browser."""
    opts = uc.ChromeOptions()
    opts.add_argument(f"--user-data-dir={profile}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument(f"--window-size={w},{h}")
    opts.add_argument(f"--user-agent={ua}")
    opts.add_argument("--lang=pt-BR,pt;q=0.9,en;q=0.8")
    # Desativa flags que sinalizam automação
    opts.add_argument("--disable-blink-features=AutomationControlled")
    if headless:
        opts.add_argument("--headless=new")

    log.info("Chrome (undetected) — perfil: %s  res: %dx%d", profile, w, h)
    driver = uc.Chrome(options=opts, headless=headless)
    # Garante que navigator.webdriver está indefinido
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
    })
    return driver


def _create_selenium_driver(
    profile: str, w: int, h: int, ua: str, headless: bool
) -> webdriver.Chrome:
    """Fallback com selenium puro + patches manuais de anti-detecção."""
    opts = Options()
    opts.add_argument(f"--user-data-dir={profile}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument(f"--window-size={w},{h}")
    opts.add_argument(f"--user-agent={ua}")
    opts.add_argument("--lang=pt-BR,pt;q=0.9,en;q=0.8")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    # Desativa o aviso "Chrome está sendo controlado por software automatizado"
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    })
    if headless:
        opts.add_argument("--headless=new")

    log.info("Chrome (selenium+patches) — perfil: %s  res: %dx%d", profile, w, h)

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=opts
        )
    except Exception:
        driver = webdriver.Chrome(options=opts)

    # Patches via CDP / JS
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "\n".join([
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});",
            "window.chrome={runtime:{}};",
            "Object.defineProperty(navigator,'languages',{get:()=>['pt-BR','pt','en-US','en']});",
            "Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]});",
        ])
    })
    return driver


# ===========================================================================
# UTILITÁRIOS DE EXTRAÇÃO
# ===========================================================================

def _detect_platform(url: str) -> Optional[str]:
    lower = url.lower()
    if "linkedin.com"  in lower: return "linkedin"
    if "instagram.com" in lower: return "instagram"
    if "tiktok.com"    in lower: return "tiktok"
    return None


def _clean_number(raw: str) -> str:
    """Normaliza texto numérico da tela: '1.234', '1,2K', '10 mil', etc."""
    if not raw:
        return ""
    raw = raw.strip().replace("\u00a0", " ").replace("\u202f", " ")
    raw = re.sub(r"(?i)\bmil\b", "K", raw)
    m = re.search(r"([\d][\d\s.,]*[KkMmBb]?)", raw)
    return m.group(1).strip() if m else raw.strip()


def _extract_by_label(page_text: str, label_variants: list[str]) -> str:
    """Procura rótulo no texto da página e retorna o número adjacente.
    Aceita número antes ou depois do rótulo (mesma linha ou próxima)."""
    for label in label_variants:
        esc = re.escape(label)
        for pat in (
            rf"([\d][\d\s.,]*[KkMmBb]?)\s*\n\s*{esc}",  # número → rótulo
            rf"{esc}\s*\n\s*([\d][\d\s.,]*[KkMmBb]?)",   # rótulo → número
            rf"{esc}\s+([\d][\d\s.,]*[KkMmBb]?)",         # mesma linha
        ):
            m = re.search(pat, page_text, re.IGNORECASE)
            if m:
                return _clean_number(m.group(1))
    return ""


def _detect_sponsored(page_text: str) -> str:
    lower = page_text.lower()
    return "Sim" if any(t in lower for t in config.SPONSORED_TERMS) else "Não"


# ===========================================================================
# EXTRATORES POR PLATAFORMA
# ===========================================================================

class _BaseExtractor:
    """Interface comum para os extratores de cada plataforma."""

    label_map: dict[str, list[str]] = {}
    fields: list[str] = []

    def __init__(self, driver: webdriver.Chrome) -> None:
        self.driver = driver

    # ------------------------------------------------------------------
    def extract(self, url: str) -> dict:
        log.info("Abrindo: %s", url)
        self.driver.get(url)
        self._wait_for_content()
        HumanBehavior.sleep(1.5, 3.0)         # "lê" o topo da página
        HumanBehavior.move_mouse(self.driver)  # posiciona o mouse naturalmente
        self._expand_sections()
        HumanBehavior.scroll_full_page(self.driver)
        HumanBehavior.sleep(config.POST_SCROLL_WAIT * 0.8,
                            config.POST_SCROLL_WAIT * 1.4)

        page_text = self._page_text()
        data = self._extract_metrics(page_text)
        data["Patrocinado"] = _detect_sponsored(page_text)
        return data

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        except TimeoutException:
            log.warning("Timeout aguardando conteúdo")

    def _expand_sections(self) -> None:
        xpaths = [
            "//button[contains(.,'Ver mais')]",
            "//button[contains(.,'Mostrar mais')]",
            "//button[contains(.,'See more')]",
            "//button[contains(.,'Show more')]",
        ]
        for xp in xpaths:
            for btn in self.driver.find_elements(By.XPATH, xp)[:3]:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                    HumanBehavior.micro()
                except Exception:
                    pass

    def _page_text(self) -> str:
        try:
            return self.driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            return ""

    def _extract_metrics(self, page_text: str) -> dict:
        data: dict[str, str] = {}
        for field in self.fields:
            variants = self.label_map.get(field, [field.lower()])
            data[field] = _extract_by_label(page_text, variants)

        # Segunda passagem via DOM para campos ainda vazios
        missing = [f for f in self.fields if not data.get(f)]
        if missing:
            for f, v in self._dom_fallback(missing).items():
                if v:
                    data[f] = v
        return data

    def _dom_fallback(self, fields: list[str]) -> dict:
        """JS varre o DOM procurando pares (valor numérico → rótulo vizinho)."""
        data: dict[str, str] = {}
        try:
            pairs: list[dict] = self.driver.execute_script("""
                const res = [];
                document.querySelectorAll('*').forEach(el => {
                    if (el.children.length > 0) return;
                    const t = (el.textContent || '').trim();
                    if (/^[\\d][\\d\\s.,]*[KkMmBb]?$/.test(t) && t.length < 20) {
                        const p = el.parentElement;
                        if (p) res.push({value: t, ctx: p.textContent.trim().slice(0,250)});
                    }
                });
                return res;
            """) or []
        except Exception:
            return data

        for field in fields:
            variants = self.label_map.get(field, [field.lower()])
            for pair in pairs:
                ctx = pair.get("ctx", "").lower()
                if any(v.lower() in ctx for v in variants):
                    data[field] = _clean_number(pair.get("value", ""))
                    break
        return data


# ---------------------------------------------------------------------------
class LinkedInExtractor(_BaseExtractor):
    label_map = config.LINKEDIN_LABEL_MAP
    fields    = config.LINKEDIN_FIELDS

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
            log.warning("LinkedIn: timeout aguardando analytics")


# ---------------------------------------------------------------------------
class InstagramExtractor(_BaseExtractor):
    label_map = config.INSTAGRAM_LABEL_MAP
    fields    = config.INSTAGRAM_FIELDS

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         "//*[contains(.,'Visualizações') or contains(.,'Views')]")
                    ),
                    EC.presence_of_element_located((By.TAG_NAME, "article")),
                    EC.presence_of_element_located((By.TAG_NAME, "main")),
                )
            )
        except TimeoutException:
            log.warning("Instagram: timeout aguardando insights")

    def _expand_sections(self) -> None:
        # Fecha popups (cookies, notificações) que bloqueiam o conteúdo
        for xp in [
            "//button[@aria-label='Close']",
            "//button[@aria-label='Fechar']",
            "//button[contains(.,'Não agora')]",
            "//button[contains(.,'Not Now')]",
            "//button[contains(.,'Aceitar')]",
            "//button[contains(.,'Accept')]",
        ]:
            for btn in self.driver.find_elements(By.XPATH, xp)[:1]:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                    HumanBehavior.sleep(0.4, 0.9)
                except Exception:
                    pass
        super()._expand_sections()


# ---------------------------------------------------------------------------
class TikTokExtractor(_BaseExtractor):
    label_map = config.TIKTOK_LABEL_MAP
    fields    = config.TIKTOK_FIELDS

    def _wait_for_content(self) -> None:
        try:
            WebDriverWait(self.driver, config.PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "[class*='analytics'],[class*='metric']")
                    ),
                    EC.presence_of_element_located((By.TAG_NAME, "main")),
                )
            )
        except TimeoutException:
            log.warning("TikTok: timeout aguardando analytics")


# ---------------------------------------------------------------------------
_EXTRACTOR_MAP: dict[str, type[_BaseExtractor]] = {
    "linkedin":  LinkedInExtractor,
    "instagram": InstagramExtractor,
    "tiktok":    TikTokExtractor,
}
_FIELDS_MAP: dict[str, list[str]] = {
    "linkedin":  config.LINKEDIN_FIELDS,
    "instagram": config.INSTAGRAM_FIELDS,
    "tiktok":    config.TIKTOK_FIELDS,
}


# ===========================================================================
# GERENCIADOR DE EXCEL  (leitura + escrita in-place)
# ===========================================================================

# Colunas iniciais de saída por plataforma (vem do config)
_OUTPUT_COL: dict[str, int] = {
    "linkedin":  column_index_from_string(config.LINKEDIN_OUTPUT_COL),   # 32 = AF
    "instagram": column_index_from_string(config.INSTAGRAM_OUTPUT_COL),  # 41 = AO
    "tiktok":    column_index_from_string(config.TIKTOK_OUTPUT_COL),     # 50 = AX
}

_URL_ALIASES = {"url", "link", "endereço", "address", "urls", "links"}


class ExcelManager:
    """
    Lê as URLs da planilha e escreve os resultados de volta
    nas colunas designadas por plataforma.

    Escrita incremental: salva após cada linha processada para que
    uma eventual interrupção não perca o progresso já feito.
    """

    def __init__(self, path: str) -> None:
        self.path = path
        self.wb   = openpyxl.load_workbook(path)
        self.ws   = self.wb.active
        self._url_col   = self._find_url_col()
        self._has_hdr   = self._detect_header()
        self._data_start = 2 if self._has_hdr else 1

    # ------------------------------------------------------------------
    def _find_url_col(self) -> int:
        """Detecta coluna de URLs pelo cabeçalho ou pelo conteúdo."""
        # 1) Cabeçalho na linha 1
        for col in range(1, self.ws.max_column + 1):
            val = self.ws.cell(row=1, column=col).value
            if val and str(val).strip().lower() in _URL_ALIASES:
                return col
        # 2) Primeira célula que pareça uma URL
        for col in range(1, 11):
            for row in range(1, 6):
                val = self.ws.cell(row=row, column=col).value
                if val and str(val).strip().startswith("http"):
                    return col
        return 1  # fallback: coluna A

    def _detect_header(self) -> bool:
        val = self.ws.cell(row=1, column=self._url_col).value
        return bool(val) and not str(val).strip().startswith("http")

    # ------------------------------------------------------------------
    def url_rows(self) -> dict[int, str]:
        """Retorna {linha: url} para todas as linhas com URL válida."""
        result = {}
        for row in range(self._data_start, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=self._url_col).value
            if val and str(val).strip().startswith("http"):
                result[row] = str(val).strip()
        return result

    # ------------------------------------------------------------------
    def ensure_headers(self) -> None:
        """Escreve os cabeçalhos dos campos nas colunas designadas (linha 1)."""
        for platform, start_col in _OUTPUT_COL.items():
            fields = _FIELDS_MAP[platform]
            color  = config.PLATFORM_COLORS.get(platform, "444444")
            fill   = PatternFill(start_color=color, end_color=color,
                                 fill_type="solid")
            bold_w = Font(bold=True, color="FFFFFF", size=10)
            center = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
            for i, field in enumerate(fields):
                cell = self.ws.cell(row=1, column=start_col + i)
                cell.value     = field
                cell.font      = bold_w
                cell.fill      = fill
                cell.alignment = center
            self.ws.row_dimensions[1].height = 34

        self._save()

    # ------------------------------------------------------------------
    def write_row(self, row_num: int, platform: str, data: dict) -> None:
        """Escreve os campos de uma plataforma na linha indicada e salva."""
        p = platform.lower()
        if p not in _OUTPUT_COL:
            log.warning("Plataforma '%s' sem mapeamento de colunas.", p)
            return

        start_col = _OUTPUT_COL[p]
        fields    = _FIELDS_MAP[p]
        alt_fill  = (
            PatternFill(start_color="F2F2F2", end_color="F2F2F2",
                        fill_type="solid")
            if row_num % 2 == 0 else None
        )
        for i, field in enumerate(fields):
            cell = self.ws.cell(row=row_num, column=start_col + i,
                                value=data.get(field, ""))
            cell.alignment = Alignment(vertical="center")
            if alt_fill:
                cell.fill = alt_fill

        self._save()  # salva incrementalmente

    # ------------------------------------------------------------------
    def _save(self) -> None:
        self.wb.save(self.path)


# ===========================================================================
# ORQUESTRADOR PRINCIPAL
# ===========================================================================

class SocialMediaScraper:
    """Coordena a extração de analytics para múltiplas URLs."""

    def __init__(
        self,
        profile_path: Optional[str] = None,
        debug_port: Optional[int] = None,
        headless: bool = False,
    ) -> None:
        self.profile_path = profile_path
        self.debug_port   = debug_port
        self.headless     = headless
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
        assert self.driver, "Use como context manager."
        platform  = _detect_platform(url)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

        if platform is None:
            return {
                "Plataforma": "Desconhecida",
                "URL": url,
                "Patrocinado": "Não",
                "Data Extração": timestamp,
                "Erro": "URL não reconhecida. Suportadas: LinkedIn, Instagram, TikTok.",
            }

        try:
            extractor = _EXTRACTOR_MAP[platform](self.driver)
            data = extractor.extract(url)
            data.update({
                "Plataforma":    platform.capitalize(),
                "URL":           url,
                "Data Extração": timestamp,
            })
            return data
        except WebDriverException as exc:
            log.error("Erro em %s: %s", url, exc)
            return {
                "Plataforma":    platform.capitalize(),
                "URL":           url,
                "Patrocinado":   "Não",
                "Data Extração": timestamp,
                "Erro":          str(exc)[:200],
            }

    # ------------------------------------------------------------------
    def scrape_excel(self, input_path: str) -> None:
        """
        Processa todas as URLs da planilha e escreve os resultados
        de volta nas colunas designadas por plataforma.
        """
        manager   = ExcelManager(input_path)
        url_rows  = manager.url_rows()

        if not url_rows:
            print("Nenhuma URL encontrada na planilha.")
            return

        manager.ensure_headers()
        total = len(url_rows)

        for idx, (row_num, url) in enumerate(url_rows.items(), 1):
            print(f"  [{idx:>{len(str(total))}}/{total}] linha {row_num} — {url}")
            result   = self.scrape_url(url)
            platform = result.get("Plataforma", "").lower()

            if platform in _OUTPUT_COL:
                manager.write_row(row_num, platform, result)
                log.info("Linha %d salva (%s)", row_num, platform)
            else:
                log.warning("Linha %d ignorada — plataforma '%s' sem colunas mapeadas.",
                            row_num, platform)

            if idx < total:
                HumanBehavior.between_pages()

        print(f"\nConcluído. Planilha atualizada: {input_path}")


# ===========================================================================
# CLI
# ===========================================================================

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="scraper.py",
        description="Extrai analytics do LinkedIn, Instagram e TikTok simulando comportamento humano.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  # Processar planilha (Chrome deve estar FECHADO para compartilhar o perfil)
  python scraper.py planilha.xlsx

  # Usar perfil Chrome em caminho personalizado
  python scraper.py planilha.xlsx --perfil "/home/user/.config/google-chrome"

  # Modo recomendado: abrir Chrome manualmente, logar, depois conectar via debug
  #   1. Abra o Chrome com:  google-chrome --remote-debugging-port=9222
  #   2. Faça login nas redes sociais normalmente
  #   3. Execute:
  python scraper.py planilha.xlsx --porta-debug 9222

  # URL única (sem planilha)
  python scraper.py --url "https://www.linkedin.com/analytics/post-summary/urn:li:activity:..."

Colunas de saída na planilha:
  LinkedIn   → AF : AN
  Instagram  → AO : AW
  TikTok     → AX : BF
        """,
    )
    p.add_argument("planilha", nargs="?",
                   help="Planilha Excel com coluna 'URL' (ou coluna A)")
    p.add_argument("--url", metavar="URL",
                   help="Processar uma única URL diretamente")
    p.add_argument("--perfil", metavar="CAMINHO",
                   help="Diretório User Data do Chrome")
    p.add_argument("--porta-debug", type=int, metavar="PORTA",
                   help="Porta de depuração remota do Chrome (ex: 9222)")
    p.add_argument("--headless", action="store_true",
                   help="Executar sem janela de browser (requer login prévio via --porta-debug)")
    return p


def main() -> None:
    parser = _build_parser()
    args   = parser.parse_args()

    if not args.planilha and not args.url:
        parser.print_help()
        sys.exit(1)

    print("=" * 60)
    print("  Social Media Analytics Scraper")
    if _UC_AVAILABLE:
        print("  [anti-bot: undetected-chromedriver ativo]")
    else:
        print("  [anti-bot: patches manuais — instale undetected-chromedriver]")
    print("=" * 60)

    with SocialMediaScraper(
        profile_path=args.perfil,
        debug_port=args.porta_debug,
        headless=args.headless,
    ) as scraper:

        if args.url:
            print(f"\nProcessando: {args.url}\n")
            result = scraper.scrape_url(args.url)
            print("\nResultados:")
            print("-" * 45)
            for k, v in result.items():
                if v:
                    print(f"  {k:<52} {v}")
            print("-" * 45)

        else:
            if not Path(args.planilha).exists():
                print(f"\nErro: arquivo não encontrado — {args.planilha}")
                sys.exit(1)
            print(f"\nPlanilha: {args.planilha}\n")
            scraper.scrape_excel(args.planilha)


if __name__ == "__main__":
    main()

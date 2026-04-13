"""
Configurações do Social Media Analytics Scraper.

Campos esperados por plataforma e variações de rótulo aceitas
(português e inglês) para extração resiliente a mudanças de UI.
"""

# ---------------------------------------------------------------------------
# CAMPOS LINKEDIN
# ---------------------------------------------------------------------------
LINKEDIN_FIELDS = [
    "Impressões",
    "Usuários alcançados",
    "Visualizações do perfil a partir desta publicação",
    "Seguidores obtidos com esta publicação",
    "Reações",
    "Comentários",
    "Compartilhamentos",
    "Salvamentos",
    "Envios no LinkedIn",
]

# Variações de rótulo que podem aparecer na tela (PT-BR e EN)
LINKEDIN_LABEL_MAP: dict[str, list[str]] = {
    "Impressões": [
        "impressões",
        "impressions",
    ],
    "Usuários alcançados": [
        "usuários alcançados",
        "unique impressions",
        "reach",
        "alcançados",
    ],
    "Visualizações do perfil a partir desta publicação": [
        "visualizações do perfil a partir desta publicação",
        "profile views from this post",
        "visualizações do perfil",
        "profile views",
    ],
    "Seguidores obtidos com esta publicação": [
        "seguidores obtidos com esta publicação",
        "followers gained from this post",
        "seguidores obtidos",
        "new followers",
    ],
    "Reações": [
        "reações",
        "reactions",
        "likes",
    ],
    "Comentários": [
        "comentários",
        "comments",
    ],
    "Compartilhamentos": [
        "compartilhamentos",
        "shares",
        "reposts",
    ],
    "Salvamentos": [
        "salvamentos",
        "saves",
        "bookmarks",
    ],
    "Envios no LinkedIn": [
        "envios no linkedin",
        "direct sends",
        "sends",
    ],
}

# ---------------------------------------------------------------------------
# CAMPOS INSTAGRAM
# ---------------------------------------------------------------------------
INSTAGRAM_FIELDS = [
    "Visualizações",
    "Contas alcançadas",
    "Interações",
    "Curtidas",
    "Comentários",
    "Salvamentos",
    "Compartilhamentos",
    "Atividade do perfil",
    "Seguidores",
]

INSTAGRAM_LABEL_MAP: dict[str, list[str]] = {
    "Visualizações": [
        "visualizações",
        "views",
        "plays",
        "reproduções",
        "impressions",
    ],
    "Contas alcançadas": [
        "contas alcançadas",
        "accounts reached",
        "reach",
        "alcance",
    ],
    "Interações": [
        "interações",
        "interactions",
        "engajamento",
        "engagement",
        "total de interações",
    ],
    "Curtidas": [
        "curtidas",
        "likes",
        "gostos",
    ],
    "Comentários": [
        "comentários",
        "comments",
    ],
    "Salvamentos": [
        "salvamentos",
        "saves",
        "bookmarks",
        "guardados",
    ],
    "Compartilhamentos": [
        "compartilhamentos",
        "shares",
        "retweets",
    ],
    "Atividade do perfil": [
        "atividade do perfil",
        "profile activity",
        "ações no perfil",
        "profile actions",
        "visitas ao perfil",
        "profile visits",
    ],
    "Seguidores": [
        "seguidores",
        "followers",
        "novos seguidores",
        "new followers",
    ],
}

# ---------------------------------------------------------------------------
# CAMPOS TIKTOK
# ---------------------------------------------------------------------------
TIKTOK_FIELDS = [
    "Visualizações do vídeo",
    "Curtidas",
    "Comentários",
    "Compartilhamentos",
    "Salvamentos",
    "Alcance",
    "Impressões",
    "Tempo médio de exibição",
    "Reproduções completas",
]

TIKTOK_LABEL_MAP: dict[str, list[str]] = {
    "Visualizações do vídeo": [
        "visualizações do vídeo",
        "video views",
        "views",
        "visualizações",
    ],
    "Curtidas": [
        "curtidas",
        "likes",
    ],
    "Comentários": [
        "comentários",
        "comments",
    ],
    "Compartilhamentos": [
        "compartilhamentos",
        "shares",
    ],
    "Salvamentos": [
        "salvamentos",
        "saves",
        "favoritos",
        "favorites",
        "bookmarks",
    ],
    "Alcance": [
        "alcance",
        "reach",
        "usuários únicos",
        "unique viewers",
    ],
    "Impressões": [
        "impressões",
        "impressions",
    ],
    "Tempo médio de exibição": [
        "tempo médio de exibição",
        "average watch time",
        "avg watch time",
        "tempo médio assistido",
    ],
    "Reproduções completas": [
        "reproduções completas",
        "full video watches",
        "watched full video",
        "vídeo completo",
        "finished watching",
    ],
}

# ---------------------------------------------------------------------------
# COLUNAS DE SAÍDA NA PLANILHA (escrita in-place)
# LinkedIn → AF:AN  |  Instagram → AO:AW  |  TikTok → AX:BF
# ---------------------------------------------------------------------------
LINKEDIN_OUTPUT_COL  = "AF"   # AF(32) a AN(40) — 9 campos LinkedIn
INSTAGRAM_OUTPUT_COL = "AO"   # AO(41) a AW(49) — 9 campos Instagram
TIKTOK_OUTPUT_COL    = "AX"   # AX(50) a BF(58) — 9 campos TikTok

# ---------------------------------------------------------------------------
# TERMOS DE PATROCÍNIO
# ---------------------------------------------------------------------------
SPONSORED_TERMS = [
    "patrocinado",
    "sponsored",
    "impulsionado",
    "promoted",
    "publicidade",
    "anúncio",
    "paid",
    "campanha paga",
]

# ---------------------------------------------------------------------------
# CORES DAS ABAS NO EXCEL (ARGB hex)
# ---------------------------------------------------------------------------
PLATFORM_COLORS = {
    "linkedin":  "0A66C2",   # azul LinkedIn
    "instagram": "E1306C",   # rosa Instagram
    "tiktok":    "010101",   # preto TikTok
    "outros":    "555555",   # cinza padrão
}

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES DE ESPERA / SCROLL  (todos em segundos)
# ---------------------------------------------------------------------------
PAGE_LOAD_TIMEOUT    = 20     # timeout do carregamento inicial
POST_SCROLL_WAIT     = 2.0    # pausa antes de extrair texto após scroll
BETWEEN_URLS_MIN     = 4      # pausa mínima entre URLs (anti-rate-limit)
BETWEEN_URLS_MAX     = 9      # pausa máxima entre URLs

# --- Parâmetros de scroll humano ---
SCROLL_STEP_MIN      = 220    # pixels mínimos por "gesto" de scroll
SCROLL_STEP_MAX      = 520    # pixels máximos por "gesto" de scroll
SCROLL_FAST_PAUSE    = (0.35, 1.1)   # (min, max) pausa scroll rápido
SCROLL_READ_PAUSE    = (2.0,  5.0)   # (min, max) pausa de "leitura"
SCROLL_READ_PROB     = 0.18          # probabilidade de pausar para "ler"
SCROLL_BACK_PROB     = 0.10          # probabilidade de rolar levemente para cima
SCROLL_BACK_PX       = (70, 220)     # quantidade de pixels no recuo

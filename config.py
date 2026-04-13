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
# CONFIGURAÇÕES DE ESPERA / SCROLL
# ---------------------------------------------------------------------------
PAGE_LOAD_TIMEOUT = 20        # segundos para esperar o carregamento inicial
SCROLL_PAUSE = 1.5            # segundos entre cada passo de scroll
SCROLL_STEP_PX = 600          # pixels por passo de scroll
POST_SCROLL_WAIT = 2.0        # segundos após scroll completo antes de extrair
BETWEEN_URLS_PAUSE = 3        # segundos entre URLs para evitar rate-limit

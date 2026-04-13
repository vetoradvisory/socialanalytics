#!/usr/bin/env python3
"""
Cria uma planilha Excel modelo com URLs de exemplo e cabeçalhos explicativos.

Uso:
    python criar_planilha_modelo.py
    python criar_planilha_modelo.py --saida minha_planilha.xlsx
"""

import argparse
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import config

# ---------------------------------------------------------------------------
# CORES
# ---------------------------------------------------------------------------
_HEADER_BLUE = "0A66C2"       # LinkedIn blue
_HEADER_GRAY = "444444"
_EXAMPLE_FILL = "FFF9E6"      # fundo amarelado para linhas de exemplo


def _fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


# ---------------------------------------------------------------------------
# LINHAS DE EXEMPLO
# ---------------------------------------------------------------------------
_EXAMPLE_URLS = [
    # LinkedIn sem patrocínio
    "https://www.linkedin.com/analytics/post-summary/urn:li:activity:7427034257539129344",
    # LinkedIn com patrocínio (sponsored campaign manager — formato alternativo)
    "https://www.linkedin.com/analytics/post-summary/urn:li:activity:7427034257539129345",
    # Instagram Reels / Post insights
    "https://www.instagram.com/insights/media/3873963123355803288/",
    # Instagram Story insights
    "https://www.instagram.com/insights/media/3873963123355803289/",
    # TikTok — substitua pelo link real da sua conta
    "https://www.tiktok.com/@suaconta/video/7427034257539129344",
]

_EXAMPLE_NOTES = [
    "LinkedIn — Post orgânico",
    "LinkedIn — Post patrocinado (substitua o ID)",
    "Instagram — Reels/Post orgânico",
    "Instagram — Story (substitua o ID)",
    "TikTok — substitua pela URL real",
]


# ---------------------------------------------------------------------------
# CRIAÇÃO
# ---------------------------------------------------------------------------

def create_template(output_path: str = "planilha_modelo.xlsx") -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "URLs"

    # --- Linha de cabeçalho ---
    headers = ["URL", "Observação (opcional)"]
    header_fill = _fill(_HEADER_BLUE)
    white_bold = Font(bold=True, color="FFFFFF", size=12)
    center = Alignment(horizontal="center", vertical="center")

    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = white_bold
        cell.fill = header_fill
        cell.alignment = center
    ws.row_dimensions[1].height = 28

    # --- Linhas de exemplo ---
    example_fill = _fill(_EXAMPLE_FILL)
    for ri, (url, note) in enumerate(zip(_EXAMPLE_URLS, _EXAMPLE_NOTES), 2):
        url_cell = ws.cell(row=ri, column=1, value=url)
        url_cell.fill = example_fill
        url_cell.alignment = Alignment(vertical="center")

        note_cell = ws.cell(row=ri, column=2, value=note)
        note_cell.fill = example_fill
        note_cell.alignment = Alignment(vertical="center")

    # Linhas em branco para o usuário preencher
    for ri in range(len(_EXAMPLE_URLS) + 2, len(_EXAMPLE_URLS) + 12):
        ws.cell(row=ri, column=1, value="").alignment = Alignment(vertical="center")

    # --- Largura das colunas ---
    ws.column_dimensions["A"].width = 80
    ws.column_dimensions["B"].width = 40

    # --- Aba de referência dos campos ---
    _add_fields_reference_sheet(wb)

    wb.save(output_path)
    print(f"Planilha modelo criada: {output_path}")
    print()
    print("Como usar:")
    print("  1. Abra a planilha e substitua as URLs de exemplo pelas suas URLs reais.")
    print("  2. Certifique-se de estar logado nas redes sociais no Chrome.")
    print("  3. Execute: python scraper.py planilha_modelo.xlsx")
    print()
    print("Dica: a coluna 'Observação' é opcional e não afeta o processamento.")


def _add_fields_reference_sheet(wb: openpyxl.Workbook) -> None:
    """Cria uma aba 'Campos Extraídos' com a referência dos campos por plataforma."""
    ws = wb.create_sheet(title="Campos Extraídos")

    header_style = {
        "linkedin":  (_fill(config.PLATFORM_COLORS["linkedin"]),  Font(bold=True, color="FFFFFF")),
        "instagram": (_fill(config.PLATFORM_COLORS["instagram"]), Font(bold=True, color="FFFFFF")),
        "tiktok":    (_fill(config.PLATFORM_COLORS["tiktok"]),    Font(bold=True, color="FFFFFF")),
    }

    platform_data = [
        ("LinkedIn", config.LINKEDIN_FIELDS),
        ("Instagram", config.INSTAGRAM_FIELDS),
        ("TikTok", config.TIKTOK_FIELDS),
    ]

    row = 1
    for platform_name, fields in platform_data:
        key = platform_name.lower()
        fill, font = header_style[key]

        # Título da plataforma
        title_cell = ws.cell(row=row, column=1, value=f"  {platform_name}")
        title_cell.font = font
        title_cell.fill = fill
        title_cell.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.row_dimensions[row].height = 24
        row += 1

        # Sub-cabeçalho
        ws.cell(row=row, column=1, value="Campo").font = Font(bold=True)
        ws.cell(row=row, column=2, value="Descrição").font = Font(bold=True)
        row += 1

        # Campos
        for field in fields:
            ws.cell(row=row, column=1, value=field)
            ws.cell(row=row, column=2, value="Extraído automaticamente")
            row += 1

        row += 1  # espaço entre plataformas

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 30


# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Cria planilha Excel modelo para o Social Media Analytics Scraper."
    )
    parser.add_argument(
        "--saida", "-o",
        default="planilha_modelo.xlsx",
        help="Nome do arquivo de saída (padrão: planilha_modelo.xlsx)",
    )
    args = parser.parse_args()
    create_template(args.saida)


if __name__ == "__main__":
    main()

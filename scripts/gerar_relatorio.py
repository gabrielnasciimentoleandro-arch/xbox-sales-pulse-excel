#!/usr/bin/env python3
"""Gera um relatório PDF curto para apresentação do Xbox Sales Pulse."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch, Rectangle
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
FILE = ROOT / "xbox_sales_pulse.xlsx"
OUTPUT = DOCS / "xbox-sales-pulse-apresentacao.pdf"
AUTHOR = "0Barone (@gabrielnasciimentoleandro-arch)"

INK = "#0B0F14"
DARK = "#111827"
SLATE = "#1F2937"
GREEN = "#22C55E"
LIME = "#9BC848"
MINT = "#2AE6B1"
BLUE = "#38BDF8"
PURPLE = "#8B5CF6"
AMBER = "#F59E0B"
RED = "#EF4444"
WHITE = "#FFFFFF"
MUTED = "#94A3B8"
LIGHT = "#E7EDF3"


def money(value: float, decimals: bool = False) -> str:
    places = 2 if decimals else 0
    raw = f"{value:,.{places}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {raw}"


def card(ax, x, y, w, h, face=DARK, edge="#334155", radius=.02):
    patch = FancyBboxPatch(
        (x, y), w, h,
        boxstyle=f"round,pad=.008,rounding_size={radius}",
        transform=ax.transAxes, facecolor=face, edgecolor=edge, linewidth=1.1,
    )
    ax.add_patch(patch)
    return patch


def footer(ax, page: int) -> None:
    ax.text(.06, .035, "XBOX SALES PULSE", transform=ax.transAxes,
            color=MUTED, fontsize=7, fontweight="bold")
    ax.text(.5, .035, AUTHOR, transform=ax.transAxes,
            color=MUTED, fontsize=7, ha="center")
    ax.text(.94, .035, f"{page} / 4", transform=ax.transAxes,
            color=MUTED, fontsize=7, ha="right")


def base_page() -> tuple[plt.Figure, plt.Axes]:
    fig = plt.figure(figsize=(13.333, 7.5), facecolor=INK)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    return fig, ax


def cover(pdf: PdfPages, data: dict[str, float]) -> None:
    fig, ax = base_page()
    ax.add_patch(Rectangle((0, 0), .022, 1, transform=ax.transAxes,
                           facecolor=GREEN, edgecolor=GREEN))
    ax.text(.075, .82, "DIO • PROJETO AUTORAL", transform=ax.transAxes,
            color=GREEN, fontsize=11, fontweight="bold")
    ax.text(.075, .66, "XBOX\nSALES PULSE", transform=ax.transAxes,
            color=WHITE, fontsize=38, fontweight="bold", linespacing=.9, va="center")
    ax.text(.078, .48,
            "Dashboard de vendas em Excel para transformar\ndados de assinaturas em decisões claras.",
            transform=ax.transAxes, color="#CBD5E1", fontsize=14, linespacing=1.5)

    metrics = [
        ("Receita", money(data["revenue"]), GREEN),
        ("Assinaturas", f"{int(data['subscribers'])}", BLUE),
        ("Ticket médio", money(data["ticket"], True), MINT),
        ("Base validada", "100%", PURPLE),
    ]
    x = .075
    for label, value, color in metrics:
        card(ax, x, .205, .185, .14)
        ax.text(x + .018, .31, label.upper(), transform=ax.transAxes,
                color=MUTED, fontsize=7.5, fontweight="bold")
        ax.text(x + .018, .245, value, transform=ax.transAxes,
                color=color, fontsize=17, fontweight="bold")
        x += .205

    ax.text(.075, .115, "EXCEL • OPENPYXL • ANÁLISE REPRODUTÍVEL • SEM MACROS",
            transform=ax.transAxes, color=WHITE, fontsize=9, fontweight="bold")
    ax.text(.94, .82, "2024", transform=ax.transAxes, color="#172033",
            fontsize=70, fontweight="bold", ha="right")
    footer(ax, 1)
    pdf.savefig(fig, facecolor=fig.get_facecolor())
    plt.close(fig)


def dashboard_page(pdf: PdfPages) -> None:
    fig, ax = base_page()
    ax.text(.05, .94, "01  VISÃO GERAL", transform=ax.transAxes,
            color=GREEN, fontsize=10, fontweight="bold")
    ax.text(.05, .885, "Painel executivo", transform=ax.transAxes,
            color=WHITE, fontsize=22, fontweight="bold")
    image = plt.imread(DOCS / "dashboard-geral.png")
    image_ax = fig.add_axes([.055, .095, .89, .75])
    image_ax.imshow(image)
    image_ax.set_axis_off()
    footer(ax, 2)
    pdf.savefig(fig, facecolor=fig.get_facecolor())
    plt.close(fig)


def insights_page(pdf: PdfPages, data: dict[str, float]) -> None:
    fig, ax = base_page()
    ax.text(.05, .94, "02  LEITURAS DE NEGÓCIO", transform=ax.transAxes,
            color=GREEN, fontsize=10, fontweight="bold")
    ax.text(.05, .885, "O que os dados indicam", transform=ax.transAxes,
            color=WHITE, fontsize=22, fontweight="bold")
    ax.text(.05, .84, "Indicadores calculados sobre os 295 registros da base didática.",
            transform=ax.transAxes, color=MUTED, fontsize=10)

    insights = [
        ("01", "Ultimate concentra a receita",
         f"{data['ultimate']/data['revenue']:.1%} do total — {money(data['ultimate'])}.", GREEN),
        ("02", "Renovação anual é relevante",
         f"{data['annual_yes']/data['annual']:.1%} da receita anual tem auto renovação.", LIME),
        ("03", "Minecraft é o maior adicional",
         f"{money(data['minecraft'])} contra {money(data['ea'])} de EA Play.", MINT),
        ("04", "Renovação está equilibrada",
         f"148 clientes com renovação e 147 sem; receitas de {money(data['renew_yes'])} e {money(data['renew_no'])}.", BLUE),
        ("05", "A série exige contexto",
         "Janeiro e fevereiro têm 2 registros cada; dezembro termina em 16/12.", AMBER),
    ]
    positions = [(.05, .62, .42, .16), (.53, .62, .42, .16),
                 (.05, .40, .42, .16), (.53, .40, .42, .16),
                 (.05, .18, .90, .16)]
    for (num, title, body, color), (x, y, w, h) in zip(insights, positions):
        card(ax, x, y, w, h)
        ax.text(x + .02, y + h - .035, num, transform=ax.transAxes,
                color=color, fontsize=8.5, fontweight="bold")
        ax.text(x + .02, y + h - .078, title, transform=ax.transAxes,
                color=WHITE, fontsize=12, fontweight="bold")
        ax.text(x + .02, y + .035, body, transform=ax.transAxes,
                color="#CBD5E1", fontsize=9)
    footer(ax, 3)
    pdf.savefig(fig, facecolor=fig.get_facecolor())
    plt.close(fig)


def quality_page(pdf: PdfPages) -> None:
    fig, ax = base_page()
    ax.text(.05, .94, "03  QUALIDADE E REPRODUÇÃO", transform=ax.transAxes,
            color=GREEN, fontsize=10, fontweight="bold")
    ax.text(.05, .885, "Do dado bruto ao painel", transform=ax.transAxes,
            color=WHITE, fontsize=22, fontweight="bold")

    stages = [
        ("01", "Base original", "Preservada sem edição", BLUE),
        ("02", "Normalização", "295 linhas + CSV", PURPLE),
        ("03", "Cálculos", "Fórmulas e nomes", MINT),
        ("04", "Dashboard", "3 filtros + 4 gráficos", GREEN),
        ("05", "Validação", "OOXML + recálculo", LIME),
    ]
    x = .05
    for i, (num, title, subtitle, color) in enumerate(stages):
        card(ax, x, .62, .16, .16)
        ax.text(x + .018, .745, num, transform=ax.transAxes,
                color=color, fontsize=8, fontweight="bold")
        ax.text(x + .018, .69, title, transform=ax.transAxes,
                color=WHITE, fontsize=10.5, fontweight="bold")
        ax.text(x + .018, .645, subtitle, transform=ax.transAxes,
                color=MUTED, fontsize=7.5)
        if i < len(stages) - 1:
            ax.text(x + .176, .695, "→", transform=ax.transAxes,
                    color="#475569", fontsize=18, fontweight="bold")
        x += .185

    checks = [
        "295 IDs únicos e 295 totais conferidos",
        "Receita: planos + EA Play + Minecraft − cupons",
        "4 gráficos e 3 filtros sem macros",
        "Base original preservada e rastreada por SHA-256",
        "Sem vínculos externos de pasta de trabalho",
        "Gerador e verificador disponíveis em scripts/",
    ]
    card(ax, .05, .18, .55, .32)
    ax.text(.075, .455, "CHECKLIST DE VALIDAÇÃO", transform=ax.transAxes,
            color=WHITE, fontsize=10, fontweight="bold")
    y = .405
    for text in checks:
        ax.text(.078, y, "✓", transform=ax.transAxes, color=GREEN,
                fontsize=10, fontweight="bold")
        ax.text(.105, y, text, transform=ax.transAxes, color="#CBD5E1",
                fontsize=8.3)
        y -= .044

    card(ax, .64, .18, .31, .32, face=GREEN, edge=GREEN)
    ax.text(.67, .445, "REPRODUZIR", transform=ax.transAxes,
            color=INK, fontsize=9, fontweight="bold")
    commands = [
        "python scripts/gerar_dashboard.py",
        "python scripts/gerar_previas.py",
        "python scripts/verificar_dashboard.py",
    ]
    y = .375
    for cmd in commands:
        ax.text(.67, y, cmd, transform=ax.transAxes, color=INK,
                fontsize=8.2, family="monospace", fontweight="bold")
        y -= .065
    ax.text(.67, .225, "Requer Python 3.10+ e as dependências\ndo requirements.txt.",
            transform=ax.transAxes, color="#153326", fontsize=8)

    ax.text(.05, .105,
            "SHA-256 da base: 0b60de55f2e54ef2759b6363fe1ef1436ea7ce3de5493ddd5f0e987ac8052a13",
            transform=ax.transAxes, color=MUTED, fontsize=7.5, family="monospace")
    footer(ax, 4)
    pdf.savefig(fig, facecolor=fig.get_facecolor())
    plt.close(fig)


def main() -> None:
    wb = load_workbook(FILE, data_only=True)
    dash = wb["Dashboard"]
    analysis = wb["Análise"]
    data = {
        "revenue": dash["A11"].value,
        "subscribers": dash["D11"].value,
        "ticket": dash["G11"].value,
        "ultimate": analysis["F28"].value,
        "annual": analysis["B17"].value,
        "annual_yes": analysis["B18"].value,
        "minecraft": analysis["B21"].value,
        "ea": analysis["B20"].value,
        "renew_yes": analysis["J26"].value,
        "renew_no": analysis["J27"].value,
    }
    DOCS.mkdir(parents=True, exist_ok=True)
    with PdfPages(OUTPUT, metadata={
        "Title": "Xbox Sales Pulse — Apresentação",
        "Author": AUTHOR,
        "Subject": "Dashboard autoral de vendas no Excel",
        "Keywords": "DIO, Excel, Xbox, dashboard, vendas, openpyxl",
    }) as pdf:
        cover(pdf, data)
        dashboard_page(pdf)
        insights_page(pdf, data)
        quality_page(pdf)
    print(f"Apresentação criada: {OUTPUT}")


if __name__ == "__main__":
    main()

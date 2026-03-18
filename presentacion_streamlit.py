from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List
from html import escape

import streamlit as st
from pptx import Presentation


PPTX_DEFAULT = Path("Presentacion-Ejecutiva-Corporativa-2025-2026.pptx")

ECIPSA_AZUL = "#0B3E53"
ECIPSA_NARANJA = "#F18019"
ECIPSA_CELESTE = "#009AC4"
ECIPSA_GRAFITO = "#2E2F2F"

IMAGENES_ECIPSA = [
    "https://www.ecipsa.com/wp-content/uploads/2021/06/El-Bosque.jpg",
    "https://www.ecipsa.com/wp-content/uploads/2021/06/Valle-Escondido.jpg",
    "https://www.ecipsa.com/wp-content/uploads/2021/06/Valle-Cercano.jpg",
    "https://www.ecipsa.com/wp-content/uploads/2021/06/Tower.jpg",
    "https://www.ecipsa.com/wp-content/uploads/2021/09/foto-milaires.jpg",
]

FONDO_TOWER = "https://www.ecipsa.com/wp-content/uploads/2021/06/Tower.jpg"


@dataclass
class DeckSlide:
    number: int
    title: str
    bullets: List[str]


def extraer_deck(ruta: Path) -> List[DeckSlide]:
    prs = Presentation(str(ruta))
    result: List[DeckSlide] = []

    for idx, slide in enumerate(prs.slides, start=1):
        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = slide.shapes.title.text.strip()

        bullets: List[str] = []
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape == slide.shapes.title:
                continue
            tf = shape.text_frame
            if tf is None:
                continue
            for paragraph in tf.paragraphs:
                text = (paragraph.text or "").strip()
                if text:
                    bullets.append(text)

        if not title:
            title = f"Slide {idx}"

        result.append(DeckSlide(number=idx, title=title, bullets=bullets))

    return result


def estilos() -> None:
    st.markdown(
        f"""
        <style>
            .stApp {{
                background: radial-gradient(circle at 10% 10%, #e9f6fb 0%, #f8fafc 45%, #eef4f8 100%);
                color: {ECIPSA_GRAFITO};
            }}
            .stApp [data-testid="stAppViewContainer"] {{
                position: relative;
                background: transparent;
            }}
            .app-watermark {{
                position: fixed;
                inset: 0;
                background-size: cover;
                background-position: center center;
                opacity: 0.24;
                filter: saturate(0.95) contrast(0.92);
                transform: scale(1.03);
                pointer-events: none;
                z-index: 0;
            }}
            .app-watermark-overlay {{
                position: fixed;
                inset: 0;
                background: linear-gradient(90deg, rgba(248,250,252,0.76) 0%, rgba(248,250,252,0.72) 40%, rgba(248,250,252,0.6) 100%);
                pointer-events: none;
                z-index: 0;
            }}
            .stApp [data-testid="stAppViewContainer"] > .main,
            .stApp header,
            .stApp [data-testid="stToolbar"] {{
                position: relative;
                z-index: 1;
            }}
            .stApp h1,
            .stApp h2,
            .stApp h3,
            .stApp label,
            .stCaption,
            .stMarkdown p,
            .stSlider label {{
                color: {ECIPSA_AZUL} !important;
            }}
            .hero {{
                border: 1px solid {ECIPSA_AZUL};
                border-radius: 20px;
                padding: 20px 24px;
                background: linear-gradient(120deg, {ECIPSA_AZUL} 0%, #134f69 55%, {ECIPSA_CELESTE} 100%);
                color: #f8fafc;
                margin-bottom: 1rem;
                box-shadow: 0 12px 32px rgba(15, 23, 42, 0.22);
            }}
            .hero-title {{
                font-size: 1.7rem;
                font-weight: 700;
                margin: 0.1rem 0 0.5rem 0;
                color: #FFFFFF !important;
            }}
            .hero-sub {{
                font-size: 1rem;
                margin: 0;
                color: #EAF7FC;
            }}
            .deck-wrap {{
                border: 1px solid #cfe4eb;
                border-radius: 20px;
                background: #ffffff;
                padding: 28px 34px;
                min-height: 64vh;
                box-shadow: 0 14px 35px rgba(15, 23, 42, 0.08);
            }}
            .deck-kicker {{
                font-size: 0.9rem;
                color: {ECIPSA_CELESTE};
                letter-spacing: 0.02em;
                margin-bottom: 0.2rem;
                font-weight: 600;
            }}
            .deck-title {{
                font-size: 2.1rem;
                line-height: 1.15;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin: 0.3rem 0 1.2rem 0;
            }}
            .deck-bullet {{
                font-size: 1.2rem;
                color: {ECIPSA_GRAFITO};
                margin: 0.45rem 0;
            }}
            .side-card {{
                border: 1px solid #cfe4eb;
                border-radius: 16px;
                background: #ffffff;
                padding: 14px 16px;
                margin-bottom: 0.7rem;
            }}
            .side-title {{
                font-size: 0.92rem;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin: 0;
            }}
            .side-sub {{
                font-size: 0.84rem;
                color: {ECIPSA_GRAFITO};
                margin: 0.2rem 0 0 0;
            }}
            .deck-footer {{
                margin-top: 1.2rem;
                font-size: 0.9rem;
                color: {ECIPSA_GRAFITO};
            }}
            .badge {{
                display: inline-block;
                border: 1px solid #ffd8b5;
                border-radius: 9999px;
                padding: 0.15rem 0.7rem;
                margin-right: 0.45rem;
                font-size: 0.8rem;
                color: #8a4900;
                background: #fff3e8;
            }}
            .stButton > button {{
                border: 1px solid #f4b67f;
                background: #fff7f0;
                color: {ECIPSA_AZUL};
                font-weight: 600;
                min-height: 2.6rem;
                padding: 0.2rem 0.35rem;
                font-size: 1.05rem;
            }}
            .stButton > button:hover {{
                border-color: {ECIPSA_NARANJA};
                color: {ECIPSA_NARANJA};
            }}
            .stProgress > div > div > div > div {{
                background-color: {ECIPSA_NARANJA};
            }}
            .obra-wrap {{
                border: 1px solid #cfe4eb;
                border-radius: 14px;
                background: #ffffff;
                padding: 12px 14px;
                margin: 0.3rem 0 0.8rem 0;
            }}
            .obra-head {{
                color: {ECIPSA_AZUL};
                font-weight: 700;
                font-size: 0.9rem;
                margin-bottom: 0.45rem;
            }}
            .nodes-wrap {{
                display: flex;
                align-items: center;
                gap: 0;
                margin: 0.35rem 0 0.55rem 0;
            }}
            .node {{
                width: 18px;
                height: 18px;
                border-radius: 9999px;
                border: 2px solid #d1d5db;
                background: #f8fafc;
                flex: 0 0 auto;
            }}
            .node.done {{
                border-color: {ECIPSA_NARANJA};
                background: {ECIPSA_NARANJA};
            }}
            .node.current {{
                border-color: {ECIPSA_AZUL};
                background: {ECIPSA_AZUL};
                box-shadow: 0 0 0 3px rgba(11, 62, 83, 0.15);
                animation: pulseNode 1.6s ease-in-out infinite;
            }}
            .node-line {{
                height: 3px;
                background: #dbe3ea;
                flex: 1 1 auto;
                margin: 0 3px;
                border-radius: 99px;
            }}
            .node-line.done {{
                background: {ECIPSA_CELESTE};
            }}
            .tower-svg {{
                margin-top: 6px;
                display: block;
                width: 100%;
                max-width: 320px;
                animation: floatTower 3.8s ease-in-out infinite;
            }}
            .obra-legend {{
                color: {ECIPSA_GRAFITO};
                font-size: 0.82rem;
                margin-top: 0.45rem;
            }}
            @keyframes floatTower {{
                0% {{ transform: translateY(0px); }}
                50% {{ transform: translateY(-2px); }}
                100% {{ transform: translateY(0px); }}
            }}
            @keyframes pulseNode {{
                0% {{ box-shadow: 0 0 0 0 rgba(11, 62, 83, 0.25); }}
                70% {{ box-shadow: 0 0 0 6px rgba(11, 62, 83, 0); }}
                100% {{ box-shadow: 0 0 0 0 rgba(11, 62, 83, 0); }}
            }}
            .metrics-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 1rem;
                margin: 1.2rem 0 1.4rem 0;
            }}
            .metric-card {{
                border: 1px solid #cfe4eb;
                border-radius: 16px;
                background: #ffffff;
                padding: 1.1rem 1.2rem 1rem 1.2rem;
                box-shadow: 0 4px 14px rgba(15,23,42,0.06);
                display: flex;
                flex-direction: column;
                align-items: center;
                text-align: center;
                gap: 0.25rem;
            }}
            .metric-icon {{
                font-size: 1.6rem;
                line-height: 1;
            }}
            .metric-value {{
                font-size: 2.4rem;
                font-weight: 800;
                color: {ECIPSA_AZUL};
                line-height: 1.1;
            }}
            .metric-label {{
                font-size: 0.9rem;
                color: {ECIPSA_GRAFITO};
                line-height: 1.35;
            }}
            .area-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 0.75rem;
                margin: 0.9rem 0 1rem 0;
            }}
            .area-card {{
                border: 1px solid #cfe4eb;
                border-radius: 14px;
                background: #ffffff;
                padding: 0.85rem 1rem;
                box-shadow: 0 3px 10px rgba(15,23,42,0.05);
            }}
            .area-card-title {{
                font-size: 0.85rem;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin-bottom: 0.35rem;
            }}
            .area-card-item {{
                font-size: 0.88rem;
                color: {ECIPSA_GRAFITO};
                line-height: 1.45;
                margin: 0.18rem 0;
            }}
            .infra-banner {{
                border-left: 4px solid {ECIPSA_CELESTE};
                background: #eaf7fc;
                border-radius: 0 12px 12px 0;
                padding: 0.75rem 1.1rem;
                margin-top: 0.6rem;
            }}
            .infra-banner-title {{
                font-size: 0.88rem;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin-bottom: 0.3rem;
            }}
            .infra-banner-item {{
                font-size: 0.88rem;
                color: {ECIPSA_GRAFITO};
                margin: 0.15rem 0;
            }}
            .area-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 0.75rem;
                margin: 0.9rem 0 1rem 0;
            }}
            .area-card {{
                border: 1px solid #cfe4eb;
                border-radius: 14px;
                background: #ffffff;
                padding: 0.85rem 1rem;
                box-shadow: 0 3px 10px rgba(15,23,42,0.05);
            }}
            .area-card-title {{
                font-size: 0.85rem;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin-bottom: 0.35rem;
            }}
            .area-card-item {{
                font-size: 0.88rem;
                color: {ECIPSA_GRAFITO};
                line-height: 1.45;
                margin: 0.18rem 0;
            }}
            .infra-banner {{
                border-left: 4px solid {ECIPSA_CELESTE};
                background: #eaf7fc;
                border-radius: 0 12px 12px 0;
                padding: 0.75rem 1.1rem;
                margin-top: 0.6rem;
            }}
            .infra-banner-title {{
                font-size: 0.88rem;
                font-weight: 700;
                color: {ECIPSA_AZUL};
                margin-bottom: 0.3rem;
            }}
            .infra-banner-item {{
                font-size: 0.88rem;
                color: {ECIPSA_GRAFITO};
                margin: 0.15rem 0;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def imagen_para_slide(numero: int) -> str:
    return IMAGENES_ECIPSA[(numero - 1) % len(IMAGENES_ECIPSA)]


def render_fondo_app(image_url: str) -> None:
    st.markdown(
        f"""
        <div class="app-watermark" style="background-image: url('{image_url}');"></div>
        <div class="app-watermark-overlay"></div>
        """,
        unsafe_allow_html=True,
    )


def icono_para_slide(titulo: str, numero: int) -> str:
    t = titulo.lower()
    if "agenda" in t:
        return "🧭"
    if "resultado" in t or "desempeño" in t:
        return "📈"
    if "impacto" in t:
        return "🏢"
    if "infraestructura" in t:
        return "🖥️"
    if "progreso" in t or "cartera" in t:
        return "🚀"
    if "inteligencia" in t or "ia" in t:
        return "🤖"
    if "riesgo" in t:
        return "⚠️"
    if "roadmap" in t:
        return "🛣️"
    if "decisiones" in t:
        return "✅"
    if numero == 1:
        return "🎯"
    return "📌"


def render_hero(total: int, current: int) -> None:
    st.markdown(
        f"""
        <div class="hero">
            <p class="hero-title">Presentación Ejecutiva Corporativa</p>
            <p class="hero-sub">Formato de exposición moderna para Dirección · Asistida por IA · {current}/{total} slides</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_panel_lateral(slide: DeckSlide) -> None:
    icon = icono_para_slide(slide.title, slide.number)
    titulo = escape(slide.title)

    if slide.number == 1:
        st.markdown(
            """
            <div class="side-card">
                <p class="side-title">🎯 Propósito de esta presentación</p>
                <p class="side-sub">Comunicar los logros de 2025, el estado actual de las iniciativas estratégicas y la hoja de ruta para 2026.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    if slide.number == 2:
        st.markdown(
            """
            <div class="side-card" style="padding: 1.2rem 1.4rem;">
                <p class="side-title" style="font-size:1rem;margin-bottom:0.6rem;">🎯 Objetivo a lograr</p>
            </div>
            <div class="metric-banner">🏆 Consolidación del área como habilitador estratégico de la organización</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_cover(slide: DeckSlide) -> None:
    st.markdown(
        f"""
        <div class="deck-wrap" style="display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;min-height:64vh;">
            <p class="deck-kicker" style="font-size:1rem;letter-spacing:0.08em;">ECIPSA</p>
            <h1 class="deck-title" style="font-size:3rem;margin:0.6rem 0 1.2rem 0;">{escape(slide.title)}</h1>
            <p style="font-size:1.15rem;color:#2E2F2F;max-width:480px;line-height:1.55;">Marzo 2026</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_slide(slide: DeckSlide, total: int) -> None:
    if slide.number == 1:
        render_cover(slide)
        return

    if slide.number == 2:
        render_resultados(slide)
        return

    tam_bullet = "1.35rem"
    tam_titulo = "2.5rem"
    icono = icono_para_slide(slide.title, slide.number)

    bullets_html = "".join(
        f"<li class='deck-bullet' style='font-size:{tam_bullet};'>{escape(b)}</li>" for b in slide.bullets[:8]
    )

    st.markdown(
        f"""
        <div class="deck-wrap">
            <h1 class="deck-title" style="font-size:{tam_titulo};">{icono} {escape(slide.title)}</h1>
            <ul style="padding-left: 1.4rem; margin-top: 0.2rem;">
                {bullets_html}
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_navegacion_obra(actual: int, total: int) -> None:
    """Edificio que crece piso a piso: cada slide = un piso nuevo."""
    FLOOR_H = max(8, min(18, 140 // total))
    BLD_X = 50
    BLD_W = 80
    GROUND_Y = 168
    NUM_WINS = 4
    WIN_W = 9
    win_h = max(4, FLOOR_H - 6)
    spacing = (BLD_W - NUM_WINS * WIN_W) / (NUM_WINS + 1)

    parts: list = []

    # ── floors ────────────────────────────────────────────────────────────
    for i in range(actual):
        floor_y = GROUND_Y - (i + 1) * FLOOR_H
        is_current = (i == actual - 1)
        fill = "#0f5272" if is_current else "#0B3E53"

        # floor slab
        if is_current and actual > 1:
            parts.append(
                f"<rect x='{BLD_X}' y='{floor_y}' width='{BLD_W}' height='{FLOOR_H - 1}' fill='{fill}' rx='2'>"
                "<animate attributeName='opacity' values='0;1' dur='0.5s' fill='freeze'/></rect>"
            )
        else:
            parts.append(
                f"<rect x='{BLD_X}' y='{floor_y}' width='{BLD_W}' height='{FLOOR_H - 1}' fill='{fill}' rx='2'/>"
            )

        # windows
        win_y = floor_y + 2
        for j in range(NUM_WINS):
            wx = round(BLD_X + spacing + j * (WIN_W + spacing))
            if is_current:
                parts.append(
                    f"<rect x='{wx}' y='{win_y}' width='{WIN_W}' height='{win_h}' rx='1.5' fill='#FFFDE7' stroke='#F18019' stroke-width='0.8'>"
                    f"<animate attributeName='opacity' values='0.5;1;0.7;1' dur='1.8s' begin='{j * 0.15:.2f}s' repeatCount='indefinite'/></rect>"
                )
            else:
                parts.append(
                    f"<rect x='{wx}' y='{win_y}' width='{WIN_W}' height='{win_h}' rx='1.5' fill='#BFD9E8' stroke='#9BC9DF' stroke-width='0.6'/>"
                )

    # ── crane when building is in progress ────────────────────────────────
    if actual < total:
        crane_top = GROUND_Y - actual * FLOOR_H
        arm_h = max(16, int(FLOOR_H * 1.5))
        mast_top = crane_top - arm_h
        arm_len = 26
        wire_end_y = mast_top + arm_h // 2
        ball_cx = BLD_X + BLD_W - 2  # right side of building
        parts += [
            f"<rect x='88' y='{mast_top}' width='4' height='{arm_h}' fill='{ECIPSA_NARANJA}' rx='1'/>",
            f"<rect x='88' y='{mast_top}' width='{arm_len}' height='3' fill='{ECIPSA_NARANJA}' rx='1'/>",
            f"<line x1='{88 + arm_len - 1}' y1='{mast_top + 3}' x2='{88 + arm_len - 1}' y2='{wire_end_y}' stroke='{ECIPSA_NARANJA}' stroke-width='1' stroke-dasharray='2,2'/>",
            f"<circle cx='{88 + arm_len - 1}' cy='{wire_end_y + 4}' r='3.5' fill='{ECIPSA_NARANJA}'>"
            "<animate attributeName='opacity' values='0.3;1;0.3' dur='0.9s' repeatCount='indefinite'/>"
            "</circle>",
        ]

    # ── flag when complete ────────────────────────────────────────────────
    if actual == total:
        top_y = GROUND_Y - total * FLOOR_H
        parts += [
            f"<rect x='88' y='{top_y - 18}' width='2.5' height='18' fill='{ECIPSA_GRAFITO}'/>",
            f"<polygon points='90.5,{top_y - 18} 101,{top_y - 13} 90.5,{top_y - 8}' fill='{ECIPSA_NARANJA}'/>",
        ]

    # ── ground line ───────────────────────────────────────────────────────
    parts.append(f"<rect x='0' y='{GROUND_Y}' width='180' height='5' fill='#cfe4eb' rx='2'/>")

    svg = (
        "<svg viewBox='0 0 180 200' xmlns='http://www.w3.org/2000/svg' "
        "style='width:100%;max-width:200px;display:block;margin:0 auto;'>"
        + "".join(parts)
        + "</svg>"
    )

    icono = "🏗️" if actual < total else "🏢"
    piso_txt = "Planta baja" if actual == 1 else f"Piso {actual - 1}"

    st.markdown(
        f"""
        <div class="obra-wrap">
            <div class="obra-head">{icono} {piso_txt} &mdash; slide {actual} de {total}</div>
            {svg}
        </div>
        """,
        unsafe_allow_html=True,
    )


def app() -> None:
    st.set_page_config(page_title="Presentación Ejecutiva", page_icon="📊", layout="wide")
    estilos()

    pptx_path = PPTX_DEFAULT

    if not pptx_path.exists():
        st.error("No se encontró la presentación corporativa base en la carpeta del proyecto.")
        st.stop()

    slides = extraer_deck(pptx_path)
    total = len(slides)

    if "slide_idx" not in st.session_state:
        st.session_state.slide_idx = 0

    # Procesar navegacion ANTES de renderizar para que la slide correcta aparezca al primer click
    if st.session_state.get("_nav") == "prev":
        st.session_state.slide_idx = max(0, st.session_state.slide_idx - 1)
        del st.session_state["_nav"]
    elif st.session_state.get("_nav") == "next":
        st.session_state.slide_idx = min(total - 1, st.session_state.slide_idx + 1)
        del st.session_state["_nav"]

    render_fondo_app(FONDO_TOWER)

    content_l, content_r = st.columns([3.2, 1.2])
    with content_l:
        render_slide(slides[st.session_state.slide_idx], total)
    with content_r:
        render_panel_lateral(slides[st.session_state.slide_idx])

        render_navegacion_obra(st.session_state.slide_idx + 1, total)

        nav_l, _, nav_r = st.columns([1, 1.2, 1])
        with nav_l:
            if st.button("◀", use_container_width=True):
                st.session_state["_nav"] = "prev"
                st.rerun()
        with nav_r:
            if st.button("▶", use_container_width=True):
                st.session_state["_nav"] = "next"
                st.rerun()

    st.progress((st.session_state.slide_idx + 1) / total)


if __name__ == "__main__":
    app()

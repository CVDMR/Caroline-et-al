"""Exibe figuras Plotly e permite salvá-las na pasta `figuras/` do projeto."""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import streamlit as st

PASTA = Path(__file__).parent / "figuras"

_registro: list[tuple[str, object]] = []  # figuras da execução atual
_contador = 0


def reinicia() -> None:
    """Chamar no início de cada execução do app."""
    global _contador
    _registro.clear()
    _contador = 0


def _slug(texto: str) -> str:
    s = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").lower()
    return s or "figura"


def opcoes_sidebar() -> tuple[list[str], int]:
    st.subheader("Salvar figuras")
    formatos = st.multiselect("Formatos", ["PNG", "HTML (interativo)"], default=["PNG", "HTML (interativo)"])
    escala = st.slider("Resolução do PNG (escala)", 1, 4, 3, help="3 ≈ 300 dpi para figuras de artigo.")
    st.caption(f"Pasta: `{PASTA}`")
    return formatos, escala


def salva(fig, nome: str, formatos: list[str], escala: int) -> list[Path]:
    PASTA.mkdir(exist_ok=True)
    gravados: list[Path] = []
    if "PNG" in formatos:
        p = PASTA / f"{nome}.png"
        fig.write_image(p, scale=escala)
        gravados.append(p)
    if "HTML (interativo)" in formatos:
        p = PASTA / f"{nome}.html"
        fig.write_html(p, include_plotlyjs="cdn")
        gravados.append(p)
    return gravados


@st.fragment
def _botao(fig, nome: str, n: int, formatos: list[str], escala: int) -> None:
    if st.button("💾 Salvar na pasta", key=f"salvar_{n}", disabled=not formatos):
        try:
            arqs = salva(fig, nome, formatos, escala)
            st.caption("Salvo: " + ", ".join(f"`{a.name}`" for a in arqs))
        except Exception as e:  # noqa: BLE001
            st.error(f"Não consegui salvar: {e}")


def mostra(fig, nome: str | None = None) -> None:
    """Mostra a figura, com botão de salvar e câmera (download pelo navegador)."""
    global _contador
    _contador += 1
    titulo = fig.layout.title.text if fig.layout.title and fig.layout.title.text else ""
    nome = _slug(nome or titulo or f"figura_{_contador}")
    # evita sobrescrever figuras com o mesmo título dentro da mesma execução
    usados = {n for n, _ in _registro}
    base, i = nome, 2
    while nome in usados:
        nome, i = f"{base}_{i}", i + 1
    _registro.append((nome, fig))

    st.plotly_chart(
        fig, width="stretch", key=f"fig_{_contador}",
        config={"toImageButtonOptions": {"format": "png", "filename": nome, "scale": 3}},
    )
    formatos, escala = st.session_state.get("_fmt", (["PNG"], 3))
    _botao(fig, nome, _contador, formatos, escala)


def salvar_todas(formatos: list[str], escala: int) -> list[Path]:
    out: list[Path] = []
    for nome, fig in _registro:
        out += salva(fig, nome, formatos, escala)
    return out

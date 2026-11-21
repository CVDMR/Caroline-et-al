"""Tabelas descritivas: exibição no app e exportação em .docx (padrão de artigo).

Formato: título acima ("Tabela N. ..."), três linhas horizontais (sem linhas
verticais), Times New Roman, vírgula decimal e nota abaixo.
"""
from __future__ import annotations

import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

PASTA = Path(__file__).parent / "tabelas"
ARQUIVO_UNICO = "todas_as_tabelas.docx"
FONTE = "Times New Roman"

_registro: list[dict] = []  # tabelas da execução atual


def reinicia() -> None:
    """Chamar no início de cada execução do app."""
    _registro.clear()


# ============================================================ construção
def _num(x, casas: int = 2) -> str:
    """Número com vírgula decimal, arredondando 'meio para cima' (como Excel/SPSS); '–' se ausente."""
    if pd.isna(x):
        return "–"
    q = Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP)
    return f"{q:.{casas}f}".replace(".", ",")


def _cat(s: pd.Series, binaria: bool = False) -> pd.Series:
    """Converte a série em texto legível (0/1 -> Não/Sim; 2.0 -> '2'; NA -> '(em branco)')."""
    def f(v):
        if pd.isna(v):
            return "(em branco)"
        if binaria and v in (0, 1):
            return "Sim" if v == 1 else "Não"
        if pd.api.types.is_number(v) and float(v).is_integer():
            return str(int(v))
        return str(v)
    return s.map(f)


def resumo_artigo(base: pd.DataFrame, var: str, por: str | None = None, casas: int = 2,
                  rotulo_grupo: str = "Grupo", rotulo_total: str = "Total") -> pd.DataFrame:
    """n, média (DP), mediana (Q1–Q3) e mín–máx; com linha Total (sem grupo: uma só linha)."""
    def linha(s: pd.Series) -> list[str]:
        s = s.dropna()
        if s.empty:
            return ["0", "–", "–", "–"]
        dp = s.std() if len(s) > 1 else float("nan")
        return [
            str(len(s)),
            f"{_num(s.mean(), casas)} ({_num(dp, casas)})",
            f"{_num(s.median(), casas)} ({_num(s.quantile(.25), casas)}–{_num(s.quantile(.75), casas)})",
            f"{_num(s.min(), casas)}–{_num(s.max(), casas)}",
        ]

    linhas = []
    if por:
        for g, sub in base.groupby(por, observed=True, dropna=False):
            linhas.append(["(em branco)" if pd.isna(g) else str(g)] + linha(sub[var]))
    linhas.append([rotulo_total] + linha(base[var]))
    return pd.DataFrame(linhas, columns=[rotulo_grupo, "n", "Média (DP)", "Mediana (Q1–Q3)", "Mín–Máx"])


def freq_tabela(s: pd.Series, rotulo_col: str, ordenar: str = "n", binaria: bool = False) -> pd.DataFrame:
    """Frequência absoluta e relativa, com linha Total."""
    cont = _cat(s, binaria).value_counts()
    if ordenar == "indice":  # numéricos em ordem crescente; textos depois
        def chave(x: str):
            num = x.lstrip("-").isdigit()
            return (not num, int(x) if num else 0, x)
        cont = cont.reindex(sorted(cont.index, key=chave))
    total = int(cont.sum())
    linhas = [[k, str(int(v)), _num(100 * v / total, 1)] for k, v in cont.items()]
    linhas.append(["Total", str(total), "100,0"])
    return pd.DataFrame(linhas, columns=[rotulo_col, "n", "%"])


def cruzada(base: pd.DataFrame, var: str, por: str, rotulo_var: str,
            bin_var: bool = False, bin_por: bool = False) -> pd.DataFrame:
    """Categorias de `var` nas linhas, grupos de `por` nas colunas; células 'n (%)' (% da coluna)."""
    v, g = _cat(base[var], bin_var), _cat(base[por], bin_por)
    tab = pd.crosstab(v, g)
    tab["Total"] = tab.sum(axis=1)
    tot = tab.sum(axis=0)
    linhas = [[cat] + [f"{int(tab.loc[cat, c])} ({_num(100 * tab.loc[cat, c] / tot[c], 1)})" for c in tab.columns]
              for cat in tab.index]
    linhas.append(["Total"] + [f"{int(tot[c])} (100,0)" for c in tab.columns])
    return pd.DataFrame(linhas, columns=[rotulo_var] + list(tab.columns))


def proporcao_tabela(itens: list[tuple[str, int, int]], rotulo_col: str, col_n: str = "Sim (n)",
                     ordenar: bool = True) -> pd.DataFrame:
    """Linhas (rótulo, n, N) -> rótulo | n | N | %."""
    linhas = [(r, n, N, 100 * n / N if N else float("nan")) for r, n, N in itens]
    if ordenar:
        linhas.sort(key=lambda t: (-(t[3] if t[3] == t[3] else -1), t[0]))
    return pd.DataFrame(
        [[r, str(n), str(N), _num(p, 1)] for r, n, N, p in linhas],
        columns=[rotulo_col, col_n, "Válidos (N)", "%"],
    )


# ============================================================ .docx
def _slug(texto: str) -> str:
    s = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return (re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").lower() or "tabela")[:70]


def _fonte(run, tamanho: float, negrito=False, italico=False) -> None:
    run.font.name = FONTE
    run.font.size = Pt(tamanho)
    run.bold, run.italic = negrito, italico
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONTE)


def _documento(paisagem: bool) -> Document:
    doc = Document()
    sec = doc.sections[0]
    _orienta(sec, paisagem)
    estilo = doc.styles["Normal"]
    estilo.font.name, estilo.font.size = FONTE, Pt(10)
    rpr = estilo.element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), FONTE)
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "pt-BR")
    rpr.append(lang)
    return doc


def _orienta(sec, paisagem: bool) -> None:
    larg, alt = Cm(21), Cm(29.7)
    sec.orientation = WD_ORIENT.LANDSCAPE if paisagem else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = (alt, larg) if paisagem else (larg, alt)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2.5)


def _borda(celula, **bordas) -> None:
    """bordas: top=12, bottom=6 ... (espessura em 1/8 pt)."""
    tcpr = celula._tc.get_or_add_tcPr()
    el = tcpr.find(qn("w:tcBorders"))
    if el is None:
        el = OxmlElement("w:tcBorders")
        tcpr.append(el)
    for lado, sz in bordas.items():
        b = OxmlElement(f"w:{lado}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        el.append(b)


def _paisagem(df: pd.DataFrame) -> bool:
    return df.shape[1] > 7


def _adiciona(doc: Document, df: pd.DataFrame, titulo: str, nota: str, numero: int) -> None:
    larg_util = doc.sections[-1].page_width - doc.sections[-1].left_margin - doc.sections[-1].right_margin

    cap = doc.add_paragraph()
    cap.paragraph_format.keep_with_next = True
    cap.paragraph_format.space_after = Pt(4)
    _fonte(cap.add_run(f"Tabela {numero}. "), 10, negrito=True)
    _fonte(cap.add_run(titulo), 10)

    n_col, n_lin = df.shape[1], len(df)
    tab = doc.add_table(rows=n_lin + 1, cols=n_col)
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab.autofit = False

    # larguras proporcionais ao conteúdo (1ª coluna pode ser mais larga)
    pesos = []
    for j, c in enumerate(df.columns):
        m = max([len(str(c))] + [len(str(x)) for x in df.iloc[:, j]])
        pesos.append(min(max(m, 6), 46 if j == 0 else 24))
    larguras = [int(larg_util * p / sum(pesos)) for p in pesos]

    def celula(i: int, j: int, texto: str, negrito=False) -> None:
        c = tab.cell(i, j)
        c.width = larguras[j]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        _fonte(p.add_run(str(texto)), 10, negrito=negrito)

    for j, c in enumerate(df.columns):
        celula(0, j, str(c), negrito=True)
    for i in range(n_lin):
        total = str(df.iloc[i, 0]) == "Total"
        for j in range(n_col):
            celula(i + 1, j, df.iloc[i, j])
            if total:
                _borda(tab.cell(i + 1, j), top=4)

    # três linhas: topo e base do cabeçalho, e base da tabela
    for j in range(n_col):
        _borda(tab.cell(0, j), top=12, bottom=6)
        _borda(tab.cell(n_lin, j), bottom=12)
    trpr = tab.rows[0]._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:tblHeader"))  # repete o cabeçalho em nova página

    if nota:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        _fonte(p.add_run("Nota. "), 9, italico=True)
        _fonte(p.add_run(nota), 9)
    doc.add_paragraph()


def salva_docx(item: dict, numero: int) -> Path:
    PASTA.mkdir(exist_ok=True)
    doc = _documento(_paisagem(item["df"]))
    _adiciona(doc, item["df"], item["titulo"], item["nota"], numero)
    caminho = PASTA / f"{item['nome']}.docx"
    doc.save(caminho)
    return caminho


def salvar_todas(arquivo_unico: bool = True) -> tuple[list[Path], list[str]]:
    """Grava todas as tabelas; devolve (gravados, falhas). Uma falha não interrompe as demais."""
    out: list[Path] = []
    falhas: list[str] = []
    for n, it in enumerate(_registro, 1):
        try:
            out.append(salva_docx(it, n))
        except PermissionError:
            falhas.append(f"{it['nome']}.docx (aberto em outro programa?)")
    if arquivo_unico and _registro:
        try:
            out.append(_salva_unico())
        except PermissionError:
            falhas.append(f"{ARQUIVO_UNICO} (aberto em outro programa?)")
    return out, falhas


def _salva_unico() -> Path:
    """Um único .docx com todas as tabelas (nova seção quando muda a orientação da página)."""
    PASTA.mkdir(exist_ok=True)
    atual = _paisagem(_registro[0]["df"])
    doc = _documento(atual)
    for n, it in enumerate(_registro, 1):
        p = _paisagem(it["df"])
        if p != atual:
            _orienta(doc.add_section(WD_SECTION.NEW_PAGE), p)
            atual = p
        _adiciona(doc, it["df"], it["titulo"], it["nota"], n)
    caminho = PASTA / ARQUIVO_UNICO
    doc.save(caminho)
    return caminho


# ============================================================ interface
def opcoes_sidebar() -> bool:
    st.subheader("Salvar tabelas (.docx)")
    unico = st.checkbox("Gerar também um arquivo único com todas", value=True)
    st.caption(f"Pasta: `{PASTA}`")
    return unico


@st.fragment
def _botao(n: int) -> None:
    item = _registro[n - 1] if n - 1 < len(_registro) else None
    if item is None:
        return
    if st.button("💾 Salvar tabela (.docx)", key=f"salvar_tab_{n}"):
        try:
            st.caption(f"Salvo: `{salva_docx(item, n).name}`")
        except Exception as e:  # noqa: BLE001
            st.error(f"Não consegui salvar: {e}")


def mostra_tabela(df: pd.DataFrame, titulo: str, nome: str | None = None, nota: str = "") -> None:
    """Exibe a tabela (já formatada como texto) e permite salvá-la em .docx."""
    nome = _slug(nome or titulo)
    usados = {it["nome"] for it in _registro}
    base, i = nome, 2
    while nome in usados:
        nome, i = f"{base}_{i}", i + 1
    _registro.append({"nome": nome, "titulo": titulo, "df": df, "nota": nota})
    n = len(_registro)

    st.markdown(f"**Tabela {n}.** {titulo}")
    st.dataframe(df, hide_index=True, width="stretch")
    if nota:
        st.caption(f"Nota. {nota}")
    _botao(n)

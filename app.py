"""Protótipo 1 — Análise descritiva (blocos 1 e 2 do plano de análises).

Rodar:  streamlit run app.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import dados as d
import figuras as fg
from figuras import mostra

ARQUIVO_PADRAO = Path(__file__).parent / "Caroline_Camila.xlsx"
COR = "#3b6ea5"
COR_OK, COR_NOK = "#4c9f70", "#d95f5f"

st.set_page_config(page_title="Rotulagem de cárneos — Descritiva", layout="wide")


# ------------------------------------------------------------------ dados
@st.cache_data(show_spinner="Lendo planilha…")
def le(origem) -> pd.DataFrame:
    return d.carrega(origem)


with st.sidebar:
    st.header("Dados")
    envio = st.file_uploader("Planilha (.xlsx)", type="xlsx")
    origem = envio if envio is not None else ARQUIVO_PADRAO
    st.caption("Usando: " + (envio.name if envio else ARQUIVO_PADRAO.name))

if envio is None and not ARQUIVO_PADRAO.exists():
    st.warning("Envie a planilha .xlsx na barra lateral.")
    st.stop()

try:
    df_all = le(origem)
except Exception as e:  # noqa: BLE001
    st.error(f"Não consegui ler a aba '{d.ABA_DADOS}': {e}")
    st.stop()

dic = d.dicionario(df_all)

with st.sidebar:
    st.header("Filtros")
    subs = sorted(df_all["SUBGRUPO"].dropna().unique()) if "SUBGRUPO" in df_all else []
    sel_sub = st.multiselect("Subgrupo", subs, default=subs)
    marcas = sorted(df_all["MARCA_EMPRESA"].dropna().unique()) if "MARCA_EMPRESA" in df_all else []
    sel_marca = st.multiselect("Marca/empresa", marcas, default=marcas)
    st.divider()
    fg.reinicia()
    st.session_state["_fmt"] = fg.opcoes_sidebar()
    espaco_salvar_todas = st.empty()  # preenchido no fim do script

df = df_all[df_all["SUBGRUPO"].isin(sel_sub) & df_all["MARCA_EMPRESA"].isin(sel_marca)]
if df.empty:
    st.warning("Nenhum produto com os filtros atuais.")
    st.stop()

st.title("Rotulagem e composição de produtos cárneos")
st.caption(f"{len(df)} de {len(df_all)} produtos selecionados")

aba_dados, aba_car, aba_rot = st.tabs(["📋 Dados", "1 · Caracterização", "2 · Conformidade da rotulagem"])


def tabela_resumo(base: pd.DataFrame, var: str, por: str | None = None) -> pd.DataFrame:
    """Estatísticas descritivas de uma variável numérica (opcionalmente por grupo)."""
    def f(s: pd.Series) -> pd.Series:
        s = s.dropna()
        return pd.Series(
            {"n": len(s), "Média": s.mean(), "DP": s.std(), "Mediana": s.median(),
             "Q1": s.quantile(.25), "Q3": s.quantile(.75), "Mín": s.min(), "Máx": s.max()}
        )
    if por:
        out = base.groupby(por, observed=True)[var].apply(f).unstack()
    else:
        out = f(base[var]).to_frame(var).T
    out["n"] = out["n"].astype(int)
    return out.round(2)


def frequencias(s: pd.Series) -> pd.DataFrame:
    cont = s.fillna("(em branco)").astype(str).value_counts()
    return pd.DataFrame({"n": cont, "%": (100 * cont / cont.sum()).round(1)})


# ------------------------------------------------------------------ Dados
with aba_dados:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Produtos", len(df_all))
    c2.metric("Variáveis", df_all.shape[1])
    c3.metric("Marcas/empresas", df_all["MARCA_EMPRESA"].nunique())
    c4.metric("Subgrupos", df_all["SUBGRUPO"].nunique())

    st.subheader("Dicionário de variáveis")
    st.caption(
        "Tipos definidos a partir da aba Legenda. `ALTERNATIVA_SODIO` e `TIPO_ALTERNATIVA_SODIO` "
        "foram deixadas de fora por enquanto."
    )
    tipos_sel = st.multiselect("Filtrar por tipo", d.TIPO_ORDEM, default=d.TIPO_ORDEM)
    st.dataframe(dic[dic["Tipo"].isin(tipos_sel)], hide_index=True, width="stretch")

    st.subheader("Tabela de dados")
    st.dataframe(df, hide_index=True, width="stretch")
    st.download_button("Baixar dados filtrados (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                       "dados_filtrados.csv", "text/csv")

# ------------------------------------------------------------ Caracterização
with aba_car:
    st.subheader("Perfil geral dos produtos")
    g1, g2 = st.columns(2)

    with g1:
        por_sub = frequencias(df["SUBGRUPO"]).reset_index(names="Subgrupo")
        fig = px.bar(por_sub, x="n", y="Subgrupo", orientation="h", text="n",
                     title="Produtos por subgrupo", color_discrete_sequence=[COR])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Nº de produtos")
        mostra(fig)

    with g2:
        por_marca = frequencias(df["MARCA_EMPRESA"]).reset_index(names="Marca")
        fig = px.bar(por_marca, x="n", y="Marca", orientation="h", text="n",
                     title="Produtos por marca/empresa", color_discrete_sequence=[COR])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Nº de produtos")
        mostra(fig)

    st.subheader("Preço e peso por subgrupo")
    p1, p2 = st.columns(2)
    for col_ui, var, titulo in [(p1, "PRECO_100G", "Preço por 100 g"), (p2, "PESO_GRAMAS", "Peso adquirido (g)")]:
        with col_ui:
            fig = px.box(df, x="SUBGRUPO", y=var, points="all", hover_data=["ID", "PRODUTO"],
                         title=titulo, color_discrete_sequence=[COR])
            fig.update_layout(xaxis_title="", yaxis_title=d.rotulo(var))
            mostra(fig)
            st.dataframe(tabela_resumo(df, var, "SUBGRUPO"), width="stretch")

    st.subheader("Ingredientes e aditivos")
    i1, i2 = st.columns(2)
    for col_ui, var, titulo in [(i1, "NUMERO_INGREDIENTES", "Nº de ingredientes"), (i2, "N_ADITIVOS", "Nº de aditivos")]:
        with col_ui:
            fig = px.histogram(df, x=var, nbins=15, title=titulo, color_discrete_sequence=[COR])
            fig.update_layout(xaxis_title=d.rotulo(var), yaxis_title="Produtos", bargap=.05)
            mostra(fig)

    adit = ["GLUTAMATO_PRESENCA", "NITRATO_PRESENCA", "NITRITO_PRESENCA"]
    adit = [a for a in adit if a in df.columns]
    pct = pd.DataFrame({d.rotulo(a): [100 * (df[a] == 1).sum() / df[a].notna().sum()] for a in adit}).T
    pct.columns = ["% de produtos com presença"]
    fig = px.bar(pct.reset_index(names="Aditivo"), x="Aditivo", y="% de produtos com presença",
                 text_auto=".0f", title="Presença de glutamato, nitrato e nitrito",
                 color_discrete_sequence=[COR])
    fig.update_layout(yaxis_range=[0, 100], xaxis_title="")
    mostra(fig)

    st.divider()
    st.subheader("Explorar qualquer variável")
    cand = d.colunas_por_tipo(dic, "Contínua", "Contagem", "Binária (0/1)", "Categórica")
    e1, e2 = st.columns(2)
    var = e1.selectbox("Variável", cand, index=cand.index("PRECO_100G") if "PRECO_100G" in cand else 0,
                       format_func=d.rotulo)
    cat_cols = ["(nenhum)"] + d.colunas_por_tipo(dic, "Categórica", "Binária (0/1)")
    por = e2.selectbox("Comparar por", [c for c in cat_cols if c != var], format_func=lambda x: x if x == "(nenhum)" else d.rotulo(x))
    por = None if por == "(nenhum)" else por
    tipo = dic.set_index("Variável").loc[var, "Tipo"]

    if tipo in ("Contínua", "Contagem"):
        a, b = st.columns(2)
        with a:
            fig = px.histogram(df, x=var, color=por, nbins=20, barmode="overlay", opacity=.75,
                               title="Histograma")
            fig.update_layout(xaxis_title=d.rotulo(var), yaxis_title="Produtos")
            mostra(fig, f"histograma_{var}" + (f"_por_{por}" if por else ""))
        with b:
            fig = px.box(df, y=var, x=por, color=por, points="all", hover_data=["ID", "PRODUTO"], title="Boxplot")
            fig.update_layout(showlegend=False, yaxis_title=d.rotulo(var), xaxis_title="")
            mostra(fig, f"boxplot_{var}" + (f"_por_{por}" if por else ""))
        st.dataframe(tabela_resumo(df, var, por), width="stretch")
    else:
        base = df.assign(**{var: df[var].astype("string").fillna("(em branco)")})
        if por:
            base = base.assign(**{por: base[por].astype("string").fillna("(em branco)")})
            tab = base.groupby([por, var]).size().reset_index(name="n")
            fig = px.bar(tab, x=var, y="n", color=por, barmode="group", text="n", title="Frequência")
        else:
            tab = frequencias(df[var]).reset_index(names=var)
            fig = px.bar(tab, x=var, y="n", text="n", title="Frequência", color_discrete_sequence=[COR])
        fig.update_layout(xaxis_title=d.rotulo(var), yaxis_title="Produtos")
        mostra(fig, f"frequencia_{var}" + (f"_por_{por}" if por else ""))
        st.dataframe(frequencias(df[var]), width="stretch")

# ------------------------------------------------------------ Rotulagem
with aba_rot:
    st.subheader("Itens obrigatórios de rotulagem")
    st.caption("Percentual de produtos que apresentam / estão em conformidade com cada item "
               "(1 = sim/conforme). Células em branco ficam fora do denominador.")

    itens = [c for c in d.ITENS_ROTULAGEM if c in df.columns]
    res = pd.DataFrame(
        {
            "Item": [d.rotulo(c) for c in itens],
            "Conformes": [int((df[c] == 1).sum()) for c in itens],
            "Válidos": [int(df[c].notna().sum()) for c in itens],
        }
    )
    res["% conforme"] = (100 * res["Conformes"] / res["Válidos"]).round(1)
    res = res.sort_values("% conforme")
    fig = go.Figure(
        go.Bar(
            x=res["% conforme"], y=res["Item"], orientation="h",
            marker_color=[COR_OK if v == 100 else COR for v in res["% conforme"]],
            text=[f"{p:.0f}%  ({c}/{n})" for p, c, n in zip(res["% conforme"], res["Conformes"], res["Válidos"])],
            textposition="outside", cliponaxis=False,
        )
    )
    fig.update_layout(xaxis={"range": [0, 115], "title": "% de produtos"}, height=max(380, 34 * len(res)),
                      margin={"l": 10, "r": 40})
    mostra(fig, "conformidade_itens_obrigatorios")

    st.subheader("Rotulagem nutricional frontal (RNF)")
    r1, r2, r3, r4 = st.columns(4)
    exig = df["RNF_EXIGIDA"] == 1
    pres = df["RNF_PRESENTE"] == 1
    r1.metric("RNF exigida", f"{int(exig.sum())} / {int(df['RNF_EXIGIDA'].notna().sum())}")
    r2.metric("Presente (entre as exigidas)", f"{int((exig & pres).sum())} / {int(exig.sum())}")
    conf = (pres & (df["RNF_CONFORME"] == 1))
    r3.metric("Conforme (entre as presentes)", f"{int(conf.sum())} / {int(pres.sum())}")
    r4.metric("Média de alertas", f"{df['N_ALERTAS_RNF'].mean():.2f}")

    a, b = st.columns(2)
    with a:
        tab = frequencias(df["N_ALERTAS_RNF"]).reset_index(names="Nº de alertas")
        fig = px.bar(tab, x="Nº de alertas", y="n", text="n", title="Nº de alertas na RNF",
                     color_discrete_sequence=[COR])
        fig.update_layout(yaxis_title="Produtos")
        mostra(fig)
    with b:
        tab = frequencias(df["ESPEC_RNF"]).reset_index(names="Nutriente(s) no alerta")
        fig = px.bar(tab, x="n", y="Nutriente(s) no alerta", orientation="h", text="n",
                     title="Nutriente(s) indicado(s) na RNF", color_discrete_sequence=[COR])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Produtos")
        mostra(fig)

    st.subheader("Conformidade por subgrupo")
    long = df.melt(id_vars="SUBGRUPO", value_vars=itens, var_name="Item", value_name="v").dropna(subset=["v"])
    long["Item"] = long["Item"].map(d.rotulo)
    heat = long.groupby(["SUBGRUPO", "Item"])["v"].mean().mul(100).round(1).unstack()
    fig = px.imshow(heat, text_auto=".0f", aspect="auto", color_continuous_scale="RdYlGn",
                    zmin=0, zmax=100, labels={"color": "% conforme"})
    fig.update_layout(xaxis_title="", yaxis_title="", height=max(260, 70 * len(heat)))
    mostra(fig, "conformidade_por_subgrupo")

    st.subheader("Mapa produto × item")
    if st.checkbox("Mostrar mapa de calor por produto", value=False):
        m = df.set_index("ID")[itens].apply(pd.to_numeric, errors="coerce")
        m = m.loc[(m == 0).sum(axis=1).sort_values(ascending=False).index]
        fig = go.Figure(go.Heatmap(z=m.values, x=[d.rotulo(c) for c in m.columns], y=m.index,
                                   colorscale=[[0, COR_NOK], [1, COR_OK]], zmin=0, zmax=1,
                                   hoverongaps=False, showscale=False))
        fig.update_layout(height=max(400, 16 * len(m)), yaxis={"autorange": "reversed"})
        mostra(fig, "mapa_produto_x_item")
        st.caption("Verde = conforme/presente · Vermelho = não · Branco = sem informação.")

    st.subheader("Produtos com algum item não conforme")
    falhas = df[["ID", "MARCA_EMPRESA", "PRODUTO", "SUBGRUPO"]].copy()
    falhas["Itens não conformes"] = [
        ", ".join(d.rotulo(c) for c in itens if row[c] == 0) for _, row in df.iterrows()
    ]
    falhas["Nº"] = falhas["Itens não conformes"].map(lambda s: len([x for x in s.split(", ") if x]))
    falhas = falhas[falhas["Nº"] > 0].sort_values("Nº", ascending=False)
    st.caption(f"{len(falhas)} de {len(df)} produtos.")
    st.dataframe(falhas, hide_index=True, width="stretch")


# ------------------------------------------------------- salvar todas as figuras
with espaco_salvar_todas.container():
    _formatos, _escala = st.session_state["_fmt"]
    if st.button("💾 Salvar TODAS as figuras (todas as abas)", disabled=not _formatos):
        with st.spinner("Gravando figuras…"):
            try:
                _arqs = fg.salvar_todas(_formatos, _escala)
                st.success(f"{len(_arqs)} arquivos salvos em `figuras/`.")
            except Exception as e:  # noqa: BLE001
                st.error(f"Não consegui salvar: {e}")

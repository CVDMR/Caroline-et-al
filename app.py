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
import tabelas as tb
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
    st.divider()
    tb.reinicia()
    st.session_state["_tab_unico"] = tb.opcoes_sidebar()
    espaco_salvar_tabelas = st.empty()  # preenchido no fim do script

df = df_all[df_all["SUBGRUPO"].isin(sel_sub) & df_all["MARCA_EMPRESA"].isin(sel_marca)]
if df.empty:
    st.warning("Nenhum produto com os filtros atuais.")
    st.stop()

st.title("Rotulagem e composição de produtos cárneos")
st.caption(f"{len(df)} de {len(df_all)} produtos selecionados")

aba_dados, aba_car, aba_rot = st.tabs(["📋 Dados", "1 · Caracterização", "2 · Conformidade da rotulagem"])


NOTA_FILTRO = (f"Amostra filtrada: {len(df)} de {len(df_all)} produtos."
               if len(df) != len(df_all) else "")
NOTA_FREQ = "n = número de produtos; % = percentual sobre o total de produtos da linha Total."
NOTA_RESUMO = ("n = número de produtos com informação; DP = desvio padrão; "
               "Q1–Q3 = primeiro e terceiro quartis; Mín–Máx = menor e maior valor.")
NOTA_PROP = "Percentual calculado sobre os produtos com informação (células em branco excluídas)."


def tabela(df_: pd.DataFrame, titulo: str, nota: str = "", nome: str | None = None) -> None:
    """Mostra a tabela com botão para salvar em .docx (acrescenta a nota do filtro, se houver)."""
    tb.mostra_tabela(df_, titulo, nome, " ".join(x for x in (nota, NOTA_FILTRO) if x))


def frequencias(s: pd.Series) -> pd.DataFrame:
    cont = tb._cat(s).value_counts()
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
        tabela(tb.freq_tabela(df["SUBGRUPO"], "Subgrupo"), "Distribuição dos produtos por subgrupo",
               NOTA_FREQ, "produtos_por_subgrupo")

    with g2:
        por_marca = frequencias(df["MARCA_EMPRESA"]).reset_index(names="Marca")
        fig = px.bar(por_marca, x="n", y="Marca", orientation="h", text="n",
                     title="Produtos por marca/empresa", color_discrete_sequence=[COR])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Nº de produtos")
        mostra(fig)
        tabela(tb.freq_tabela(df["MARCA_EMPRESA"], "Marca/empresa"), "Distribuição dos produtos por marca/empresa",
               NOTA_FREQ, "produtos_por_marca_empresa")

    st.subheader("Preço e peso por subgrupo")
    p1, p2 = st.columns(2)
    for col_ui, var, titulo in [(p1, "PRECO_100G", "Preço por 100 g"), (p2, "PESO_GRAMAS", "Peso adquirido (g)")]:
        with col_ui:
            fig = px.box(df, x="SUBGRUPO", y=var, points="all", hover_data=["ID", "PRODUTO"],
                         title=titulo, color_discrete_sequence=[COR])
            fig.update_layout(xaxis_title="", yaxis_title=d.rotulo(var))
            mostra(fig)
            tabela(tb.resumo_artigo(df, var, "SUBGRUPO", casas=2 if var == "PRECO_100G" else 0,
                                    rotulo_grupo="Subgrupo"),
                   f"{titulo} segundo o subgrupo", NOTA_RESUMO, f"{var.lower()}_por_subgrupo")

    st.subheader("Ingredientes e aditivos")
    i1, i2 = st.columns(2)
    for col_ui, var, titulo in [(i1, "NUMERO_INGREDIENTES", "Nº de ingredientes"), (i2, "N_ADITIVOS", "Nº de aditivos")]:
        with col_ui:
            fig = px.histogram(df, x=var, nbins=15, title=titulo, color_discrete_sequence=[COR])
            fig.update_layout(xaxis_title=d.rotulo(var), yaxis_title="Produtos", bargap=.05)
            mostra(fig)
            tabela(tb.resumo_artigo(df, var, "SUBGRUPO", casas=1, rotulo_grupo="Subgrupo"),
                   f"{titulo} segundo o subgrupo", NOTA_RESUMO, f"{var.lower()}_por_subgrupo")

    adit = ["GLUTAMATO_PRESENCA", "NITRATO_PRESENCA", "NITRITO_PRESENCA"]
    adit = [a for a in adit if a in df.columns]
    pct = pd.DataFrame({d.rotulo(a): [100 * (df[a] == 1).sum() / df[a].notna().sum()] for a in adit}).T
    pct.columns = ["% de produtos com presença"]
    fig = px.bar(pct.reset_index(names="Aditivo"), x="Aditivo", y="% de produtos com presença",
                 text_auto=".0f", title="Presença de glutamato, nitrato e nitrito",
                 color_discrete_sequence=[COR])
    fig.update_layout(yaxis_range=[0, 100], xaxis_title="")
    mostra(fig)
    tabela(tb.proporcao_tabela([(d.rotulo(a), int((df[a] == 1).sum()), int(df[a].notna().sum())) for a in adit],
                               "Aditivo", "Presença (n)", ordenar=False),
           "Presença de glutamato, nitrato e nitrito nos produtos", NOTA_PROP, "presenca_aditivos")

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
        tabela(tb.resumo_artigo(df, var, por, casas=2,
                                rotulo_grupo=d.rotulo(por) if por else "Variável",
                                rotulo_total="Total" if por else d.rotulo(var)),
               d.rotulo(var) + (f" segundo {d.rotulo(por).lower()}" if por else ""), NOTA_RESUMO,
               f"explorador_{var.lower()}" + (f"_por_{por.lower()}" if por else ""))
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
        bin_var = tipo == "Binária (0/1)"
        if por:
            bin_por = dic.set_index("Variável").loc[por, "Tipo"] == "Binária (0/1)"
            tab_exp = tb.cruzada(df, var, por, d.rotulo(var), bin_var, bin_por)
            titulo_exp = f"{d.rotulo(var)} segundo {d.rotulo(por).lower()}"
            nota_exp = "Valores: n (%); percentuais calculados sobre o total de cada coluna."
        else:
            tab_exp = tb.freq_tabela(df[var], d.rotulo(var), binaria=bin_var)
            titulo_exp, nota_exp = d.rotulo(var), NOTA_FREQ
        tabela(tab_exp, titulo_exp, nota_exp,
               f"explorador_{var.lower()}" + (f"_por_{por.lower()}" if por else ""))

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
                      margin={"l": 10, "r": 120})
    mostra(fig, "conformidade_itens_obrigatorios")
    tabela(tb.proporcao_tabela([(r.Item, int(r.Conformes), int(r.Válidos)) for r in res.itertuples()],
                               "Item", "Conformes (n)"),
           "Conformidade dos produtos quanto aos itens obrigatórios de rotulagem", NOTA_PROP,
           "conformidade_itens_obrigatorios")

    st.subheader("Rotulagem nutricional frontal (RNF)")
    r1, r2, r3, r4 = st.columns(4)
    exig = df["RNF_EXIGIDA"] == 1
    pres = df["RNF_PRESENTE"] == 1
    r1.metric("RNF exigida", f"{int(exig.sum())} / {int(df['RNF_EXIGIDA'].notna().sum())}")
    r2.metric("Presente (entre as exigidas)", f"{int((exig & pres).sum())} / {int(exig.sum())}")
    conf = (pres & (df["RNF_CONFORME"] == 1))
    r3.metric("Conforme (entre as presentes)", f"{int(conf.sum())} / {int(pres.sum())}")
    r4.metric("Média de alertas", f"{df['N_ALERTAS_RNF'].mean():.2f}")
    tabela(tb.proporcao_tabela(
        [("RNF exigida", int(exig.sum()), int(df["RNF_EXIGIDA"].notna().sum())),
         ("RNF presente (entre as exigidas)", int((exig & pres).sum()), int(exig.sum())),
         ("RNF conforme (entre as presentes)", int(conf.sum()), int(pres.sum()))],
        "Indicador", "n", ordenar=False),
        "Rotulagem nutricional frontal (RNF) nos produtos",
        "RNF = rotulagem nutricional frontal. A base (N) de cada linha é o grupo indicado no rótulo da linha.",
        "rnf_resumo")

    a, b = st.columns(2)
    with a:
        tab = frequencias(df["N_ALERTAS_RNF"]).reset_index(names="Nº de alertas")
        fig = px.bar(tab, x="Nº de alertas", y="n", text="n", title="Nº de alertas na RNF",
                     color_discrete_sequence=[COR])
        fig.update_layout(yaxis_title="Produtos")
        mostra(fig)
        tabela(tb.freq_tabela(df["N_ALERTAS_RNF"], "Nº de alertas", ordenar="indice"),
               "Distribuição dos produtos segundo o número de alertas na RNF", NOTA_FREQ, "rnf_n_alertas")
    with b:
        tab = frequencias(df["ESPEC_RNF"]).reset_index(names="Nutriente(s) no alerta")
        fig = px.bar(tab, x="n", y="Nutriente(s) no alerta", orientation="h", text="n",
                     title="Nutriente(s) indicado(s) na RNF", color_discrete_sequence=[COR])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Produtos")
        mostra(fig)
        tabela(tb.freq_tabela(df["ESPEC_RNF"], "Nutriente(s) indicado(s)"),
               "Distribuição dos produtos segundo o(s) nutriente(s) indicado(s) na RNF", NOTA_FREQ,
               "rnf_nutrientes")

    st.subheader("Conformidade por subgrupo")
    long = df.melt(id_vars="SUBGRUPO", value_vars=itens, var_name="Item", value_name="v").dropna(subset=["v"])
    long["Item"] = long["Item"].map(d.rotulo)
    heat = long.groupby(["SUBGRUPO", "Item"])["v"].mean().mul(100).round(1).unstack()
    fig = px.imshow(heat, text_auto=".0f", aspect="auto", color_continuous_scale="RdYlGn",
                    zmin=0, zmax=100, labels={"color": "% conforme"})
    fig.update_layout(xaxis_title="", yaxis_title="", height=max(260, 70 * len(heat)))
    mostra(fig, "conformidade_por_subgrupo")
    tab_sub = heat.T.reset_index(names="Item").rename_axis(columns=None)
    for c in tab_sub.columns[1:]:
        tab_sub[c] = tab_sub[c].map(lambda v: tb._num(v, 1))
    tabela(tab_sub, "Percentual de conformidade dos itens obrigatórios segundo o subgrupo",
           "Valores em % de produtos conformes dentro de cada subgrupo (células em branco excluídas).",
           "conformidade_por_subgrupo")

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

# ------------------------------------------------------- salvar todas as tabelas
with espaco_salvar_tabelas.container():
    if st.button("💾 Salvar TODAS as tabelas (.docx)"):
        with st.spinner("Gravando tabelas…"):
            try:
                _arqs, _falhas = tb.salvar_todas(st.session_state["_tab_unico"])
                st.success(f"{len(_arqs)} arquivos salvos em `tabelas/`.")
                if _falhas:
                    st.warning("Não consegui gravar (feche o arquivo no Word e tente de novo):\n- "
                               + "\n- ".join(_falhas))
            except Exception as e:  # noqa: BLE001
                st.error(f"Não consegui salvar: {e}")

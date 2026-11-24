"""Gera o Word de curadoria/controle das análises estatísticas feitas em R (Quarto).

Uso (venv do projeto):
    C:\\Users\\Claudio\\venvs\\caroline-stats\\Scripts\\python.exe relatorio\\controle_analises.py

Saída: controle_analises_R.docx na raiz do projeto. Atualize a lista ANALISES abaixo sempre que
uma análise for criada ou alterada nos .qmd, e gere o arquivo de novo.
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SAIDA = Path(__file__).resolve().parent.parent / "controle_analises_R.docx"
FONTE = "Calibri"

# ---------------------------------------------------------------------------------------------
# Inventário das análises: (arquivo, saída no relatório, tipo de análise, variáveis, pacote/função)
# ---------------------------------------------------------------------------------------------
CAR = "caracterizacao.qmd"
CON = "conformidade.qmd"
B4 = "bloco4.qmd"

ITENS = ("DENOMINACAO_OBRIGATORIA, INF_PRODUTOR_OBRIGATORIO, PESO_LIQUIDO_OBRIGATORIO, LOTE_OBRIGATORIO, "
         "VALIDADE_OBRIGATORIO, PORCAO_OBRIGATORIA, TABELA_NUTRI_OBRIGATORIA, TABELA_NUTRI_CONFORMIDADE, "
         "LISTA_INGREDIENTES_OBRIGATORIO, INF_GLUTEN_OBRIGATORIO, SELO_INSPECAO_OBRIGATORIO, "
         "ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO")

ANALISES = [
    # ---- 1 · Caracterização --------------------------------------------------------------
    (CAR, "fig-subgrupo, tbl-subgrupo",
     "Estatística descritiva: distribuição de frequência (n e %) de variável categórica; gráfico de barras.",
     "SUBGRUPO", "freq_tabela() e barras_h() (R/funcoes.R); ggplot2, gt"),
    (CAR, "fig-marca, tbl-marca",
     "Estatística descritiva: distribuição de frequência (n e %); gráfico de barras.",
     "MARCA_EMPRESA", "freq_tabela() e barras_h(); ggplot2, gt"),
    (CAR, "fig-box-preco-peso, tbl-preco, tbl-peso",
     "Estatística descritiva de variável contínua por grupo: n, média (DP), mediana (Q1–Q3), mínimo–máximo; "
     "boxplot com pontos individuais. Sem teste de hipótese.",
     "PRECO_100G, PESO_GRAMAS (por SUBGRUPO)", "resumo_artigo() (R/funcoes.R); ggplot2, gt"),
    (CAR, "fig-ingredientes-aditivos, tbl-ingredientes, tbl-aditivos",
     "Estatística descritiva de contagens (nº de ingredientes e de aditivos): n, média (DP), mediana (Q1–Q3), "
     "mín–máx por grupo; histograma. Sem teste de hipótese (não foi ajustado modelo de Poisson).",
     "NUMERO_INGREDIENTES, N_ADITIVOS (por SUBGRUPO)", "resumo_artigo(); ggplot2, gt, tidyr"),
    (CAR, "fig-aditivos, tbl-presenca-aditivos",
     "Estatística descritiva: proporção de produtos com presença do aditivo (n/N e %), N = produtos com informação. "
     "Sem intervalo de confiança binomial.",
     "GLUTAMATO_PRESENCA, NITRATO_PRESENCA, NITRITO_PRESENCA (0/1)", "proporcao_tabela(); ggplot2, gt"),
    # ---- 2 · Conformidade ----------------------------------------------------------------
    (CON, "fig-itens, tbl-itens",
     "Estatística descritiva: proporção de produtos conformes em cada item obrigatório (n/N e %); "
     "células em branco fora do denominador. Sem IC binomial.",
     ITENS + " (0/1)", "proporcao_tabela(); ggplot2, gt"),
    (CON, "tbl-rnf-resumo",
     "Estatística descritiva: proporções encadeadas da rotulagem nutricional frontal (exigida; presente entre as "
     "exigidas; conforme entre as presentes).",
     "RNF_EXIGIDA, RNF_PRESENTE, RNF_CONFORME (0/1)", "proporcao_tabela(); gt"),
    (CON, "fig-alertas, tbl-alertas (e média no texto)",
     "Estatística descritiva de contagem: distribuição de frequência do nº de alertas e média.",
     "N_ALERTAS_RNF", "freq_tabela(), contagem(); ggplot2, gt"),
    (CON, "fig-nutrientes, tbl-nutrientes",
     "Estatística descritiva: distribuição de frequência do(s) nutriente(s) indicado(s) na RNF.",
     "ESPEC_RNF", "freq_tabela(), barras_h(); ggplot2, gt"),
    (CON, "fig-heat-subgrupo, tbl-heat-subgrupo",
     "Estatística descritiva estratificada: % de conformidade de cada item dentro de cada subgrupo (mapa de calor). "
     "Sem teste de associação (qui-quadrado/Fisher não aplicados).",
     "Itens obrigatórios (lista acima) × SUBGRUPO", "dplyr, tidyr; ggplot2, gt"),
    (CON, "fig-mapa, tbl-falhas",
     "Estatística descritiva / conferência: mapa produto × item e lista dos produtos com ao menos um item "
     "não conforme.",
     "Itens obrigatórios (lista acima); ID, MARCA_EMPRESA, PRODUTO, SUBGRUPO", "dplyr, tidyr; ggplot2, gt"),
    # ---- 4 · Rótulo × laboratório --------------------------------------------------------
    (B4, "tbl-descritiva",
     "Estatística descritiva dos valores pareados: mediana [Q1–Q3] do rótulo e do laboratório, n = 32 produtos "
     "com FoodScan.",
     "PROTEINA_ROT_100G / PROTEINA_LAB_G_100G; GORD_TOTAL_ROT_100G / GORDURA_LAB_G_100G; "
     "CHO_ROT_100G / CARBOIDRATOS_LAB_G_100G; KCAL_ROT_100G / ENERGIA_LAB_KCAL_CONV_100G; "
     "SODIO_ROT_100G / SODIO_LAB_MG_100G", "gtsummary::tbl_summary()"),
    (B4, "tbl-wilcoxon",
     "Teste de Wilcoxon pareado (postos sinalizados), bicaudal, α = 5%; diferença mediana (pseudomediana) com IC 95%; "
     "valores de p ajustados por Holm entre os 5 nutrientes. Diferença = rótulo − laboratório.",
     "Os 5 pares rótulo/laboratório acima", "stats::wilcox.test(paired = TRUE, conf.int = TRUE, exact = FALSE); "
     "stats::p.adjust(method = \"holm\")"),
    (B4, "tbl-concordancia (coluna Spearman)",
     "Correlação de postos de Spearman (associação) com IC 95%.",
     "Os 5 pares rótulo/laboratório", "DescTools::SpearmanRho()"),
    (B4, "tbl-concordancia (coluna CCC de Lin)",
     "Coeficiente de correlação de concordância de Lin (concordância absoluta entre dois métodos) com IC 95%.",
     "Os 5 pares rótulo/laboratório", "DescTools::CCC()"),
    (B4, "tbl-concordancia (coluna CCI)",
     "Coeficiente de correlação intraclasse, modelo two-way, concordância absoluta, medida única, com IC 95%.",
     "Os 5 pares rótulo/laboratório", "irr::icc(model = \"twoway\", type = \"agreement\", unit = \"single\")"),
    (B4, "tbl-concordancia (colunas Viés e Limites) e gráficos Bland-Altman",
     "Análise de Bland-Altman: viés (média das diferenças, IC 95%) e limites de concordância de 95% "
     "(viés ± 1,96 DP); gráficos de dispersão com linha de identidade e de Bland-Altman, com produtos fora dos "
     "limites identificados.",
     "Os 5 pares rótulo/laboratório", "blandr::blandr.statistics(); ggplot2, ggrepel, patchwork"),
    (B4, "tbl-extremos-sodio",
     "Conferência de dados (não é teste): cinco maiores valores de sódio de laboratório com linha e célula na "
     "planilha. Nenhum valor excluído; análise de sódio provisória.",
     "SODIO_LAB_MG_100G (coluna BB), SODIO_ROT_100G (coluna AD), LINHA_EXCEL, ID", "dplyr; gt"),
]

NAO_USADAS = [
    ("Teste binomial / IC binomial (Wilson, Clopper-Pearson)", "Proporções são apenas descritivas (n/N e %)."),
    ("Qui-quadrado / exato de Fisher", "Sem teste de associação entre categóricas."),
    ("ANOVA / Kruskal-Wallis (com post hoc)", "Sem comparação inferencial entre subgrupos."),
    ("Regressão de Poisson / binomial negativa", "Contagens (ingredientes, aditivos, alertas) só descritas."),
    ("Regressão linear / logística", "Nenhum modelo ajustado."),
    ("Correlação entre variáveis do mesmo método", "Correlação usada apenas na concordância rótulo × laboratório (Spearman)."),
]

APOIO = [
    ("relatorio/_quarto.yml", "Configuração do site Quarto (idioma pt, tema claro/escuro, menu, execução sem código)."),
    ("relatorio/R/dados.R", "Leitura e limpeza da planilha em memória; ITENS_ROTULAGEM; LINHA_EXCEL e coluna do Excel; "
                            "conversão de energia do laboratório (kJ ÷ 4,184 → kcal)."),
    ("relatorio/R/funcoes.R", "Formatação (vírgula decimal, arredondamento meio para cima) e tabelas/gráficos padrão: fmt, "
                              "freq_tabela, resumo_artigo, proporcao_tabela, tabela_gt, barras_h."),
    ("relatorio/R/concordancia.R", "Pares rótulo × laboratório (PARES), pareia(), resume_concordancia() e "
                                   "graficos_concordancia() do bloco 4."),
    ("relatorio/index.qmd", "Página inicial (contagens gerais e índice das análises)."),
    ("relatorio/copiar_para_docs.py", "Pós-render: copia o site para docs/r/ (publicação no GitHub Pages)."),
    ("relatorio/controle_analises.py", "Gera este documento Word."),
]

DECISOES = [
    "A planilha Caroline_Camila.xlsx nunca é alterada; toda limpeza ocorre em memória (aba Planilha_análise).",
    "Variáveis ALTERNATIVA_SODIO e TIPO_ALTERNATIVA_SODIO não são usadas (decisão do pesquisador).",
    "A tolerância de ±20% da ANVISA não é aplicada (decisão do pesquisador); o bloco 5 (conformidade nutricional) ainda não foi feito.",
    "Células em branco ficam fora do denominador das proporções.",
    "Energia do laboratório: valor em kJ convertido para kcal (÷ 4,184) em ENERGIA_LAB_KCAL_CONV_100G; difere em até "
    "2,1% da coluna ENERGIA_LAB_KCAL_100G existente na planilha (a conferir qual é a oficial).",
    "Valores extremos de sódio de laboratório (BB9, BB10, BB12, BB21, BB44) mantidos; checagem pendente. "
    "O P043 (BB44: 11.876,7 no laboratório contra 3.735 no rótulo) foi o único ponto fora dos limites do Bland-Altman.",
    "Testes de hipótese: bicaudais, α = 5%, correção de Holm para os cinco nutrientes; não paramétricos por causa do n pequeno (n = 32).",
    "Tabelas no estilo de artigo (três linhas, Times New Roman, vírgula decimal, arredondamento meio para cima).",
]


# ---------------------------------------------------------------------------------------------
def fonte(run, tam=10, negrito=False, cor=None, italico=False):
    run.font.name = FONTE
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONTE)
    run.font.size = Pt(tam)
    run.bold = negrito
    run.italic = italico
    if cor:
        run.font.color.rgb = RGBColor.from_string(cor)


def paragrafo(doc, texto, tam=10, negrito=False, italico=False, depois=4, cor=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(depois)
    fonte(p.add_run(texto), tam, negrito, cor, italico)
    return p


def titulo(doc, texto, nivel=1):
    p = doc.add_heading(level=nivel)
    p.paragraph_format.space_before = Pt(12 if nivel == 1 else 8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    fonte(p.add_run(texto), 14 if nivel == 1 else 12, True, "1F3864")
    return p


def marcador(doc, texto):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    fonte(p.add_run(texto), 10)


def sombra(celula, cor):
    tc_pr = celula._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tc_pr.append(shd)


def repetir_cabecalho(linha):
    tr_pr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def nao_quebrar_linha(linha):
    tr_pr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def tabela(doc, cabecalho, linhas, larguras_cm, tam=8.5, zebra=True):
    t = doc.add_table(rows=1, cols=len(cabecalho))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, (c, w) in enumerate(zip(t.rows[0].cells, larguras_cm)):
        c.width = Cm(w)
        c.text = ""
        fonte(c.paragraphs[0].add_run(cabecalho[i]), tam, True, "FFFFFF")
        sombra(c, "1F3864")
    repetir_cabecalho(t.rows[0])
    for n, linha in enumerate(linhas):
        r = t.add_row()
        nao_quebrar_linha(r)
        for i, (c, w) in enumerate(zip(r.cells, larguras_cm)):
            c.width = Cm(w)
            c.text = ""
            fonte(c.paragraphs[0].add_run(str(linha[i])), tam)
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
            if zebra and n % 2 == 1:
                sombra(c, "EEF2F8")
    return t


def rodape_paginas(secao):
    p = secao.footer.paragraphs[0]
    fonte(p.add_run("Controle das análises estatísticas em R · página "), 8, cor="666666")
    for tipo, texto in (("begin", None), (None, "PAGE"), ("end", None)):
        r = p.add_run()
        fonte(r, 8, cor="666666")
        if tipo:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), tipo)
            r._r.append(f)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = texto
            r._r.append(it)


def main():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONTE
    st.font.size = Pt(10)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    rodape_paginas(sec)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    fonte(p.add_run("Controle das análises estatísticas em R"), 20, True, "1F3864")
    paragrafo(doc, "Rotulagem e composição de produtos cárneos da Bahia · projeto caroline-stats", 11, depois=2)
    paragrafo(doc, f"Documento para curadoria e controle · gerado em {date.today():%d/%m/%Y}", 10, italico=True,
              cor="666666", depois=8)

    titulo(doc, "1. Resumo")
    paragrafo(doc,
              "Este documento lista cada análise feita nos arquivos Quarto (.qmd) do site de resultados "
              "(pasta relatorio/), com o tipo de análise estatística, as variáveis usadas (nomes padronizados "
              "da planilha Caroline_Camila.xlsx) e o pacote/função de R responsável. A coluna \"Revisão\" é "
              "para o controle de quem fizer a curadoria (visto e data).")
    paragrafo(doc, "Arquivos de análise: caracterizacao.qmd (1), conformidade.qmd (2) e bloco4.qmd (4, rótulo × "
                   "laboratório). O bloco 3 (composição) e o bloco 5 (conformidade nutricional) ainda não foram feitos.")
    paragrafo(doc, f"Total de análises registradas: {len(ANALISES)}. Testes de hipótese e medidas de concordância "
                   "aparecem apenas em bloco4.qmd; os demais arquivos são descritivos.", depois=2)

    titulo(doc, "2. Análises realizadas")
    cab = ["Nº", "Arquivo", "Saída no relatório (rótulo Quarto)", "Tipo de análise estatística", "Variáveis (planilha)",
           "Pacote / função (R)", "Revisão"]
    larg = [0.9, 2.8, 3.6, 7.0, 6.2, 4.6, 1.6]
    linhas = [(f"{i}", a, s, t, v, f, "") for i, (a, s, t, v, f) in enumerate(ANALISES, 1)]
    tabela(doc, cab, linhas, larg)

    titulo(doc, "3. Análises que ainda não foram usadas")
    paragrafo(doc, "Para evitar dúvida na curadoria: os métodos abaixo NÃO foram aplicados até agora.", depois=4)
    tabela(doc, ["Método", "Situação atual"], NAO_USADAS, [9.0, 17.7], tam=9)

    titulo(doc, "4. Arquivos de apoio")
    tabela(doc, ["Arquivo", "Função"], APOIO, [6.5, 20.2], tam=9)

    titulo(doc, "5. Tratamento dos dados e decisões")
    for d in DECISOES:
        marcador(doc, d)

    titulo(doc, "6. Ambiente")
    for t in ("R 4.6.0; Quarto 1.9.37.",
              "Pacotes: readxl 1.5.0, dplyr 1.2.1, tidyr 1.3.2, ggplot2 4.0.3, ggrepel 0.9.8, patchwork 1.3.2, "
              "scales 1.4.0, gt 1.3.0, gtsummary 2.5.1, DescTools 0.99.60, irr 0.85, blandr 0.6.0.",
              "Para renderizar: colocar o Rscript no PATH (C:\\Program Files\\R\\R-4.6.0\\bin) e executar "
              "quarto render dentro de relatorio/. O resultado é copiado para docs/r/ (GitHub Pages)."):
        marcador(doc, t)

    doc.save(SAIDA)
    print("Gerado:", SAIDA)


if __name__ == "__main__":
    main()

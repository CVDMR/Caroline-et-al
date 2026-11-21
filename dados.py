"""Leitura, limpeza e dicionário de variáveis do Caroline_Camila.xlsx.

O arquivo original NUNCA é alterado: toda a limpeza acontece em memória.
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

ABA_DADOS = "Planilha_análise"

# Variáveis que ainda não serão usadas (decisão do pesquisador).
EXCLUIDAS = {"ALTERNATIVA_SODIO", "TIPO_ALTERNATIVA_SODIO"}

# ---- Tipos explícitos -------------------------------------------------------
IDENTIFICACAO = ["ID", "MARCA_EMPRESA", "PRODUTO"]
CATEGORICAS = [
    "SUBGRUPO",
    "ESPEC_RNF",
    "TIPO_SELO",
    "TIPO_ALTERNATIVA_GLUTAMATO",
    "TIPO_ALTERNATIVA_NITRATO_NITRITO",
]
CONTAGENS = ["N_ALERTAS_RNF", "NUMERO_INGREDIENTES", "N_ADITIVOS"]
CONTINUAS_FIXAS = ["PRECO_100G", "PESO_GRAMAS", "PORCAO_TAMANHO"]

# ---- Grupos (seguem as seções da aba "Legenda") ------------------------------
GRUPOS = {
    "Identificação": ["ID", "MARCA_EMPRESA", "PRODUTO", "SUBGRUPO", "PRECO_100G", "PESO_GRAMAS"],
    "Informações obrigatórias": [
        "DENOMINACAO_OBRIGATORIA", "INF_PRODUTOR_OBRIGATORIO", "PESO_LIQUIDO_OBRIGATORIO",
        "LOTE_OBRIGATORIO", "VALIDADE_OBRIGATORIO", "PORCAO_OBRIGATORIA", "PORCAO_TAMANHO",
        "TABELA_NUTRI_OBRIGATORIA", "TABELA_NUTRI_CONFORMIDADE", "LISTA_INGREDIENTES_OBRIGATORIO",
        "INF_GLUTEN_OBRIGATORIO", "SELO_INSPECAO_OBRIGATORIO", "ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO",
    ],
    "Rotulagem frontal (RNF)": ["RNF_EXIGIDA", "RNF_PRESENTE", "RNF_CONFORME", "N_ALERTAS_RNF", "ESPEC_RNF"],
    "Ingredientes e aditivos": [
        "NUMERO_INGREDIENTES", "N_ADITIVOS", "GLUTAMATO_PRESENCA", "ALTERNATIVA_GLUTAMATO",
        "TIPO_ALTERNATIVA_GLUTAMATO", "NITRATO_PRESENCA", "NITRITO_PRESENCA",
        "ALTERNATIVA_NITRATO_NITRITO", "TIPO_ALTERNATIVA_NITRATO_NITRITO",
    ],
    "Outras informações": ["INC_PRESENTE", "ALERGENICOS_PRESENTE", "TIPO_SELO"],
}

# Itens usados no bloco 2 (conformidade da rotulagem): presença/conformidade 0/1.
ITENS_ROTULAGEM = [
    "DENOMINACAO_OBRIGATORIA", "INF_PRODUTOR_OBRIGATORIO", "PESO_LIQUIDO_OBRIGATORIO",
    "LOTE_OBRIGATORIO", "VALIDADE_OBRIGATORIO", "PORCAO_OBRIGATORIA",
    "TABELA_NUTRI_OBRIGATORIA", "TABELA_NUTRI_CONFORMIDADE", "LISTA_INGREDIENTES_OBRIGATORIO",
    "INF_GLUTEN_OBRIGATORIO", "SELO_INSPECAO_OBRIGATORIO", "ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO",
]

TIPO_ORDEM = ["Identificação", "Categórica", "Binária (0/1)", "Contagem", "Contínua"]


def normaliza_nome(nome: str) -> str:
    """'MARCA/ EMPRESA' -> 'MARCA_EMPRESA'; 'KCAL_ROT__100G' -> 'KCAL_ROT_100G'."""
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "_", s.strip()).strip("_")
    return s.upper()


# ---- Nomes amigáveis (com acento) -------------------------------------------
# Chave = nome padronizado da coluna. Usados em eixos, tabelas e listas do app.
ROTULOS = {
    "ID": "ID",
    "MARCA_EMPRESA": "Marca/empresa",
    "PRODUTO": "Produto",
    "SUBGRUPO": "Subgrupo",
    "PRECO_100G": "Preço por 100 g",
    "PESO_GRAMAS": "Peso adquirido (g)",
    # informações obrigatórias
    "DENOMINACAO_OBRIGATORIA": "Denominação de venda",
    "INF_PRODUTOR_OBRIGATORIO": "Identificação do produtor",
    "PESO_LIQUIDO_OBRIGATORIO": "Conteúdo líquido",
    "LOTE_OBRIGATORIO": "Lote",
    "VALIDADE_OBRIGATORIO": "Validade",
    "PORCAO_OBRIGATORIA": "Porção declarada",
    "PORCAO_TAMANHO": "Tamanho da porção",
    "TABELA_NUTRI_OBRIGATORIA": "Tabela nutricional",
    "TABELA_NUTRI_CONFORMIDADE": "Conformidade da tabela nutricional",
    "LISTA_INGREDIENTES_OBRIGATORIO": "Lista de ingredientes",
    "INF_GLUTEN_OBRIGATORIO": "Declaração sobre glúten",
    "SELO_INSPECAO_OBRIGATORIO": "Selo de inspeção",
    "ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO": "Orientação de armazenamento",
    # rotulagem frontal
    "RNF_EXIGIDA": "RNF exigida",
    "RNF_PRESENTE": "RNF presente",
    "RNF_CONFORME": "RNF conforme",
    "N_ALERTAS_RNF": "Nº de alertas na RNF",
    "ESPEC_RNF": "Nutriente(s) indicado(s) na RNF",
    # informação nutricional do rótulo (por 100 g)
    "KCAL_ROT_100G": "Energia – rótulo (kcal/100 g)",
    "CHO_ROT_100G": "Carboidratos – rótulo (g/100 g)",
    "ACUCARES_TOTAIS_ROT_100G": "Açúcares totais – rótulo (g/100 g)",
    "ACUCARES_ADIC_ROT_100G": "Açúcares adicionados – rótulo (g/100 g)",
    "PROTEINA_ROT_100G": "Proteína – rótulo (g/100 g)",
    "GORD_TOTAL_ROT_100G": "Gordura total – rótulo (g/100 g)",
    "GORD_SAT_ROT_100G": "Gordura saturada – rótulo (g/100 g)",
    "GORD_TRANS_ROT_100G": "Gordura trans – rótulo (g/100 g)",
    "FIBRA_ROT_100G": "Fibra alimentar – rótulo (g/100 g)",
    "SODIO_ROT_100G": "Sódio – rótulo (mg/100 g)",
    # ingredientes e aditivos
    "NUMERO_INGREDIENTES": "Nº de ingredientes",
    "N_ADITIVOS": "Nº de aditivos",
    "GLUTAMATO_PRESENCA": "Presença de glutamato",
    "ALTERNATIVA_GLUTAMATO": "Alternativa ao glutamato",
    "TIPO_ALTERNATIVA_GLUTAMATO": "Tipo de alternativa ao glutamato",
    "NITRATO_PRESENCA": "Presença de nitrato",
    "NITRITO_PRESENCA": "Presença de nitrito",
    "ALTERNATIVA_NITRATO_NITRITO": "Alternativa a nitrato/nitrito",
    "TIPO_ALTERNATIVA_NITRATO_NITRITO": "Tipo de alternativa a nitrato/nitrito",
    # outras informações
    "INC_PRESENTE": "Informação nutricional complementar",
    "ALERGENICOS_PRESENTE": "Declaração de alergênicos",
    "TIPO_SELO": "Tipo de selo",
    # análise laboratorial (por 100 g)
    "PROTEINA_LAB_G_100G": "Proteína – laboratório (g/100 g)",
    "GORDURA_LAB_G_100G": "Gordura – laboratório (g/100 g)",
    "CARBOIDRATOS_LAB_G_100G": "Carboidratos – laboratório (g/100 g)",
    "ENERGIA_LAB_KCAL_100G": "Energia – laboratório (kcal/100 g)",
    "ENERGIA_LAB_KJ_100G": "Energia – laboratório (kJ/100 g)",
    "SODIO_LAB_MG_100G": "Sódio – laboratório (mg/100 g)",
}

# Nomes amigáveis para VALORES de variáveis categóricas (aplicados ao carregar).
# "NAO" na planilha = ausência da informação (ver aba Legenda).
ROTULOS_VALORES = {
    "NAO": "Ausente (NAO)",
    "SODIO": "Sódio",
    "GORDURA": "Gordura",
    "GORDURA_SODIO": "Gordura e sódio",
    "SIE_SISBI": "SIE/SISBI",
    "Inosinato dissodico": "Inosinato dissódico",
    "Inosinato dissodico + Guanilato dissodico": "Inosinato dissódico + Guanilato dissódico",
}


# Colunas em que o valor "NAO" deve aparecer exatamente como na planilha (decisão do pesquisador).
MANTER_NAO = {"ESPEC_RNF"}


def rotulo(nome: str) -> str:
    """Nome amigável (com acento) de uma variável, para eixos e tabelas."""
    if nome in ROTULOS:
        return ROTULOS[nome]
    # variável nova, sem entrada na tabela: melhor esforço
    return nome.replace("_100G", " /100g").replace("_", " ").capitalize()


def _grupo_de(col: str) -> str:
    for g, cols in GRUPOS.items():
        if col in cols:
            return g
    if col.endswith("_ROT_100G"):
        return "Informação nutricional do rótulo"
    if "_LAB_" in col:
        return "Análise laboratorial"
    return "Não classificada"


def _tipo_de(col: str, serie: pd.Series) -> str:
    if col in IDENTIFICACAO:
        return "Identificação"
    if col in CATEGORICAS:
        return "Categórica"
    if col in CONTAGENS:
        return "Contagem"
    if col in CONTINUAS_FIXAS or col.endswith("_ROT_100G") or "_LAB_" in col:
        return "Contínua"
    # demais: binária se só tiver 0/1
    if pd.api.types.is_numeric_dtype(serie) and set(serie.dropna().unique()) <= {0, 1}:
        return "Binária (0/1)"
    if pd.api.types.is_numeric_dtype(serie):
        return "Contínua"
    return "Categórica"


def carrega(origem) -> pd.DataFrame:
    """Lê a aba de dados (caminho ou arquivo enviado) e devolve o DataFrame limpo."""
    bruto = pd.read_excel(origem, sheet_name=ABA_DADOS, engine="openpyxl")
    bruto.columns = [normaliza_nome(c) for c in bruto.columns]
    bruto = bruto.drop(columns=[c for c in EXCLUIDAS if c in bruto.columns])
    bruto = bruto.dropna(how="all")

    df = bruto.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype("string").str.strip()
            df[col] = df[col].replace("", pd.NA)
            # colunas que são números gravados como texto
            num = pd.to_numeric(df[col], errors="coerce")
            if col not in IDENTIFICACAO + CATEGORICAS and num.notna().sum() == df[col].notna().sum():
                df[col] = num
    for col in df.columns:
        if col in IDENTIFICACAO + CATEGORICAS:
            df[col] = df[col].astype("string")
    # valores categóricos com nomes legíveis (só em memória; a planilha não muda)
    for col in CATEGORICAS:
        if col in df.columns:
            mapa = {k: v for k, v in ROTULOS_VALORES.items() if not (k == "NAO" and col in MANTER_NAO)}
            df[col] = df[col].replace(mapa)
    return df.reset_index(drop=True)


def dicionario(df: pd.DataFrame) -> pd.DataFrame:
    """Dicionário de variáveis: tipo, grupo, completude e resumo."""
    linhas = []
    for col in df.columns:
        s = df[col]
        tipo = _tipo_de(col, s)
        n_valid = int(s.notna().sum())
        if tipo in ("Contínua", "Contagem"):
            resumo = f"{s.min():.4g} a {s.max():.4g}" if n_valid else ""
        elif tipo == "Binária (0/1)":
            resumo = f"1 = {int((s == 1).sum())} | 0 = {int((s == 0).sum())}" if n_valid else ""
        else:
            resumo = f"{s.nunique()} categorias"
        linhas.append(
            {
                "Variável": col,
                "Nome": rotulo(col),
                "Grupo": _grupo_de(col),
                "Tipo": tipo,
                "N válidos": n_valid,
                "% faltante": round(100 * (1 - n_valid / len(df)), 1),
                "Resumo": resumo,
            }
        )
    return pd.DataFrame(linhas)


def colunas_por_tipo(dic: pd.DataFrame, *tipos: str) -> list[str]:
    return dic.loc[dic["Tipo"].isin(tipos), "Variável"].tolist()

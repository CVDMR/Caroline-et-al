# Leitura, limpeza e dicionário de variáveis em R (espelha dados.py).
# A planilha original NUNCA é alterada: toda a limpeza acontece em memória.
library(readxl)
library(dplyr)

ARQUIVO <- file.path("..", "Caroline_Camila.xlsx")
ABA <- "Planilha_análise"
EXCLUIDAS <- c("ALTERNATIVA_SODIO", "TIPO_ALTERNATIVA_SODIO")  # decisão do pesquisador

IDENTIFICACAO <- c("ID", "MARCA_EMPRESA", "PRODUTO")
CATEGORICAS <- c("SUBGRUPO", "ESPEC_RNF", "TIPO_SELO",
                 "TIPO_ALTERNATIVA_GLUTAMATO", "TIPO_ALTERNATIVA_NITRATO_NITRITO")

# Itens do bloco 2 (conformidade da rotulagem): presença/conformidade 0/1.
ITENS_ROTULAGEM <- c(
  "DENOMINACAO_OBRIGATORIA", "INF_PRODUTOR_OBRIGATORIO", "PESO_LIQUIDO_OBRIGATORIO",
  "LOTE_OBRIGATORIO", "VALIDADE_OBRIGATORIO", "PORCAO_OBRIGATORIA",
  "TABELA_NUTRI_OBRIGATORIA", "TABELA_NUTRI_CONFORMIDADE", "LISTA_INGREDIENTES_OBRIGATORIO",
  "INF_GLUTEN_OBRIGATORIO", "SELO_INSPECAO_OBRIGATORIO", "ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO"
)

# Nomes amigáveis (com acento) das variáveis, usados em eixos e tabelas.
ROTULOS <- c(
  ID = "ID", MARCA_EMPRESA = "Marca/empresa", PRODUTO = "Produto", SUBGRUPO = "Subgrupo",
  PRECO_100G = "Preço por 100 g", PESO_GRAMAS = "Peso adquirido (g)",
  DENOMINACAO_OBRIGATORIA = "Denominação de venda",
  INF_PRODUTOR_OBRIGATORIO = "Identificação do produtor",
  PESO_LIQUIDO_OBRIGATORIO = "Conteúdo líquido", LOTE_OBRIGATORIO = "Lote",
  VALIDADE_OBRIGATORIO = "Validade", PORCAO_OBRIGATORIA = "Porção declarada",
  PORCAO_TAMANHO = "Tamanho da porção", TABELA_NUTRI_OBRIGATORIA = "Tabela nutricional",
  TABELA_NUTRI_CONFORMIDADE = "Conformidade da tabela nutricional",
  LISTA_INGREDIENTES_OBRIGATORIO = "Lista de ingredientes",
  INF_GLUTEN_OBRIGATORIO = "Declaração sobre glúten",
  SELO_INSPECAO_OBRIGATORIO = "Selo de inspeção",
  ORIENTACAO_ARMAZENAMENTO_OBRIGATORIO = "Orientação de armazenamento",
  RNF_EXIGIDA = "RNF exigida", RNF_PRESENTE = "RNF presente", RNF_CONFORME = "RNF conforme",
  N_ALERTAS_RNF = "Nº de alertas na RNF", ESPEC_RNF = "Nutriente(s) indicado(s) na RNF",
  NUMERO_INGREDIENTES = "Nº de ingredientes", N_ADITIVOS = "Nº de aditivos",
  GLUTAMATO_PRESENCA = "Presença de glutamato", NITRATO_PRESENCA = "Presença de nitrato",
  NITRITO_PRESENCA = "Presença de nitrito", TIPO_SELO = "Tipo de selo"
)

# Nomes legíveis para VALORES categóricos ("NAO" = ausência da informação, ver aba Legenda).
ROTULOS_VALORES <- c(
  NAO = "Ausente (NAO)", SODIO = "Sódio", GORDURA = "Gordura", GORDURA_SODIO = "Gordura e sódio",
  SIE_SISBI = "SIE/SISBI", "Inosinato dissodico" = "Inosinato dissódico",
  "Inosinato dissodico + Guanilato dissodico" = "Inosinato dissódico + Guanilato dissódico"
)
MANTER_NAO <- c("ESPEC_RNF")  # nestas colunas "NAO" aparece como na planilha

rotulo <- function(nome) {
  ifelse(nome %in% names(ROTULOS), ROTULOS[nome],
         paste0(toupper(substring(nome, 1, 1)), tolower(substring(gsub("_", " ", nome), 2))))
}

normaliza_nome <- function(x) {
  x <- iconv(x, from = "UTF-8", to = "ASCII//TRANSLIT")
  x <- gsub("[^A-Za-z0-9]+", "_", trimws(x))
  toupper(gsub("^_|_$", "", x))
}

carrega_dados <- function(arquivo = ARQUIVO) {
  df <- read_excel(arquivo, sheet = ABA, col_types = "text")
  names(df) <- normaliza_nome(names(df))
  # posição original na planilha (cabeçalho na linha 1; letra da coluna do Excel)
  col_excel <- setNames(col_letra(seq_along(names(df))), names(df))
  df$LINHA_EXCEL <- seq_len(nrow(df)) + 1
  df <- df |>
    select(-any_of(EXCLUIDAS)) |>
    filter(if_any(-LINHA_EXCEL, ~ !is.na(.x)))
  df <- df |> mutate(across(everything(), ~ ifelse(trimws(.x) == "", NA, trimws(.x))))
  # números gravados como texto -> numérico (menos identificação e categóricas)
  livres <- setdiff(names(df), c(IDENTIFICACAO, CATEGORICAS))
  for (col in livres) {
    num <- suppressWarnings(as.numeric(gsub(",", ".", df[[col]])))
    if (sum(!is.na(num)) == sum(!is.na(df[[col]]))) df[[col]] <- num
  }
  # valores categóricos com nomes legíveis (só em memória)
  for (col in intersect(CATEGORICAS, names(df))) {
    mapa <- ROTULOS_VALORES
    if (col %in% MANTER_NAO) mapa <- mapa[names(mapa) != "NAO"]
    i <- df[[col]] %in% names(mapa)
    df[[col]][i] <- unname(mapa[df[[col]][i]])
  }
  # energia do laboratório: kJ convertido em kcal (1 kcal = 4,184 kJ)
  df$ENERGIA_LAB_KCAL_CONV_100G <- df$ENERGIA_LAB_KJ_100G / 4.184
  attr(df, "col_excel") <- col_excel
  df
}

# 1 -> "A", 27 -> "AA", 54 -> "BB"
col_letra <- function(i) {
  vapply(i, function(n) {
    s <- ""
    while (n > 0) { r <- (n - 1) %% 26; s <- paste0(LETTERS[r + 1], s); n <- (n - 1) %/% 26 }
    s
  }, "")
}

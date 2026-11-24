# Funções de tabelas e gráficos (equivalem a tabelas.py / estilo das figuras do app).
library(dplyr)
library(ggplot2)
library(gt)

COR <- "#3b6ea5"; COR_OK <- "#4c9f70"; COR_NOK <- "#d95f5f"

# Número com vírgula decimal, arredondando "meio para cima" (como Excel/SPSS); "–" se ausente.
fmt <- function(x, d = 1) {
  r <- floor(abs(x) * 10^d + 0.5 + 1e-9) / 10^d * sign(x)
  out <- formatC(r, format = "f", digits = d, decimal.mark = ",")
  out[is.na(x)] <- "–"
  out
}

# Converte em texto legível (0/1 -> Não/Sim; 2 -> "2"; NA -> "(em branco)").
cat_txt <- function(x, binaria = FALSE) {
  if (is.numeric(x)) {
    s <- ifelse(!is.na(x) & x == floor(x), as.character(round(x)), as.character(x))
    if (binaria) s <- ifelse(!is.na(x) & x == 1, "Sim", ifelse(!is.na(x) & x == 0, "Não", s))
  } else s <- as.character(x)
  s[is.na(x)] <- "(em branco)"
  s
}

freq_tabela <- function(x, rotulo_col, ordenar = "n", binaria = FALSE) {
  s <- cat_txt(x, binaria)
  cont <- table(factor(s, levels = unique(s)))
  nomes <- names(cont); cont <- as.integer(cont)
  if (ordenar == "indice") {
    num <- grepl("^-?[0-9]+$", nomes)
    o <- order(!num, ifelse(num, suppressWarnings(as.integer(nomes)), 0L), nomes)
  } else o <- order(-cont)
  nomes <- nomes[o]; cont <- cont[o]
  tab <- data.frame(nomes, as.character(cont), fmt(100 * cont / sum(cont), 1))
  tab <- rbind(tab, c("Total", as.character(sum(cont)), "100,0"))
  setNames(tab, c(rotulo_col, "n", "%"))
}

# n, média (DP), mediana (Q1–Q3) e mín–máx, por grupo e no total.
resumo_artigo <- function(base, var, por = NULL, casas = 2, rotulo_grupo = "Grupo",
                          rotulo_total = "Total") {
  linha <- function(s) {
    s <- s[!is.na(s)]
    if (!length(s)) return(c("0", "–", "–", "–"))
    dp <- if (length(s) > 1) sd(s) else NA_real_
    c(as.character(length(s)),
      sprintf("%s (%s)", fmt(mean(s), casas), fmt(dp, casas)),
      sprintf("%s (%s–%s)", fmt(median(s), casas), fmt(quantile(s, .25, names = FALSE), casas),
              fmt(quantile(s, .75, names = FALSE), casas)),
      sprintf("%s–%s", fmt(min(s), casas), fmt(max(s), casas)))
  }
  linhas <- list()
  if (!is.null(por)) {
    g <- base[[por]]
    for (nivel in sort(unique(g[!is.na(g)]))) {
      linhas[[length(linhas) + 1]] <- c(nivel, linha(base[[var]][!is.na(g) & g == nivel]))
    }
    if (anyNA(g)) linhas[[length(linhas) + 1]] <- c("(em branco)", linha(base[[var]][is.na(g)]))
  }
  linhas[[length(linhas) + 1]] <- c(rotulo_total, linha(base[[var]]))
  tab <- as.data.frame(do.call(rbind, linhas))
  setNames(tab, c(rotulo_grupo, "n", "Média (DP)", "Mediana (Q1–Q3)", "Mín–Máx"))
}

# itens: data.frame(item, n, N) -> item | n | Válidos (N) | %
proporcao_tabela <- function(itens, rotulo_col, col_n = "Sim (n)", ordenar = TRUE) {
  itens$p <- ifelse(itens$N > 0, 100 * itens$n / itens$N, NA_real_)
  if (ordenar) itens <- itens[order(-ifelse(is.na(itens$p), -1, itens$p), itens$item), ]
  tab <- data.frame(itens$item, as.character(itens$n), as.character(itens$N), fmt(itens$p, 1))
  setNames(tab, c(rotulo_col, col_n, "Válidos (N)", "%"))
}

# Tabela no estilo de artigo: três linhas horizontais, Times, primeira coluna à esquerda.
tabela_gt <- function(tab, nota = NULL) {
  g <- gt(tab) |>
    cols_align("center", columns = -1) |>
    cols_align("left", columns = 1) |>
    opt_table_font(font = list("Times New Roman", "Times", "serif")) |>
    tab_options(
      table.border.top.style = "solid", table.border.top.width = px(2), table.border.top.color = "black",
      table.border.bottom.style = "solid", table.border.bottom.width = px(2), table.border.bottom.color = "black",
      column_labels.border.top.style = "none",
      column_labels.border.bottom.style = "solid", column_labels.border.bottom.width = px(1),
      column_labels.border.bottom.color = "black",
      table_body.hlines.style = "none", table_body.border.bottom.style = "none",
      table.font.size = px(14), data_row.padding = px(3),
      table.background.color = "white", table.font.color = "black",
      column_labels.font.weight = "bold"
    )
  total <- which(tab[[1]] == "Total")
  if (length(total)) {
    g <- g |> tab_style(style = cell_borders(sides = "top", color = "black", weight = px(1)),
                        locations = cells_body(rows = total))
  }
  if (!is.null(nota) && nzchar(nota)) {
    g <- g |> tab_source_note(html(paste0("<i>Nota.</i> ", nota)))
  }
  g
}

tema_fig <- function() theme_minimal(base_size = 13) + theme(panel.grid.minor = element_blank())

# Contagens (sem a linha Total) para gráficos.
contagem <- function(x, nome) {
  out <- as.data.frame(table(cat_txt(x)), stringsAsFactors = FALSE)
  setNames(out, c(nome, "n"))
}

# Barras horizontais de contagem, ordenadas.
barras_h <- function(tab, rotulo_x = "Nº de produtos") {
  ggplot(tab, aes(n, reorder(.data[[names(tab)[1]]], n))) +
    geom_col(fill = COR) +
    geom_text(aes(label = n), hjust = -0.3, size = 3.8) +
    scale_x_continuous(expand = expansion(mult = c(0, .12))) +
    labs(x = rotulo_x, y = NULL) + tema_fig()
}

NOTA_FREQ <- "n = número de produtos; % = percentual sobre o total de produtos da linha Total."
NOTA_RESUMO <- paste("n = número de produtos com informação; DP = desvio padrão;",
                     "Q1–Q3 = primeiro e terceiro quartis; Mín–Máx = menor e maior valor.")
NOTA_PROP <- "Percentual calculado sobre os produtos com informação (células em branco excluídas)."

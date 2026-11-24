# Concordância rótulo × laboratório (FoodScan): funções do bloco 4.
# Formatação numérica (fmt) e tabelas (tabela_gt) vêm de R/funcoes.R.
suppressMessages({
  library(dplyr); library(ggplot2); library(ggrepel); library(patchwork)
})

# nutriente, coluna do rótulo, coluna do laboratório, unidade
PARES <- tibble::tribble(
  ~nutriente,        ~rotulo,               ~lab,                          ~unidade,
  "Proteína",        "PROTEINA_ROT_100G",   "PROTEINA_LAB_G_100G",         "g/100 g",
  "Gordura total",   "GORD_TOTAL_ROT_100G", "GORDURA_LAB_G_100G",          "g/100 g",
  "Carboidratos",    "CHO_ROT_100G",        "CARBOIDRATOS_LAB_G_100G",     "g/100 g",
  "Energia",         "KCAL_ROT_100G",       "ENERGIA_LAB_KCAL_CONV_100G",  "kcal/100 g",
  "Sódio",           "SODIO_ROT_100G",      "SODIO_LAB_MG_100G",           "mg/100 g"
)

# Produtos com os dois valores (rótulo e laboratório). Diferença = rótulo − laboratório.
pareia <- function(df, nutriente) {
  p <- filter(PARES, nutriente == !!nutriente)
  df |>
    transmute(ID, LINHA_EXCEL, rot = .data[[p$rotulo]], lab = .data[[p$lab]]) |>
    filter(!is.na(rot), !is.na(lab)) |>
    mutate(media = (rot + lab) / 2, dif = rot - lab)
}

# Todas as estatísticas de concordância de um nutriente (uma linha)
resume_concordancia <- function(d, nutriente) {
  wil <- suppressWarnings(wilcox.test(d$rot, d$lab, paired = TRUE, conf.int = TRUE, exact = FALSE))
  r_rho <- DescTools::SpearmanRho(d$rot, d$lab, conf.level = 0.95)
  r_ccc <- DescTools::CCC(d$rot, d$lab)$rho.c
  r_icc <- irr::icc(cbind(d$rot, d$lab), model = "twoway", type = "agreement", unit = "single")
  r_ba <- blandr::blandr.statistics(d$rot, d$lab, sig.level = 0.95)
  tibble::tibble(
    Nutriente = nutriente, n = nrow(d),
    med_rot = median(d$rot), q1_rot = quantile(d$rot, .25), q3_rot = quantile(d$rot, .75),
    med_lab = median(d$lab), q1_lab = quantile(d$lab, .25), q3_lab = quantile(d$lab, .75),
    dif_med = unname(wil$estimate), dif_lo = wil$conf.int[1], dif_hi = wil$conf.int[2],
    p_wilcoxon = wil$p.value,
    rho = unname(r_rho["rho"]), rho_lo = unname(r_rho["lwr.ci"]), rho_hi = unname(r_rho["upr.ci"]),
    ccc = r_ccc$est, ccc_lo = r_ccc$lwr.ci, ccc_hi = r_ccc$upr.ci,
    icc = r_icc$value, icc_lo = r_icc$lbound, icc_hi = r_icc$ubound,
    vies = r_ba$bias, vies_lo = r_ba$biasLowerCI, vies_hi = r_ba$biasUpperCI,
    loa_inf = r_ba$lowerLOA, loa_sup = r_ba$upperLOA
  )
}

num_br <- scales::label_number(decimal.mark = ",", big.mark = ".")

# Dois gráficos lado a lado: dispersão com a linha de identidade e Bland-Altman
graficos_concordancia <- function(d, r, nutriente, unidade) {
  lim <- range(c(d$rot, d$lab))
  disp <- ggplot(d, aes(rot, lab)) +
    geom_abline(slope = 1, intercept = 0, linetype = "dashed", colour = "grey40") +
    geom_point(size = 2.2, alpha = .85, colour = "#2c7fb8") +
    coord_equal(xlim = lim, ylim = lim) +
    scale_x_continuous(labels = num_br) + scale_y_continuous(labels = num_br) +
    labs(x = paste0("Rótulo (", unidade, ")"), y = paste0("Laboratório (", unidade, ")"),
         subtitle = "Linha tracejada: concordância perfeita") +
    theme_minimal(base_size = 12)

  fora <- filter(d, dif > r$loa_sup | dif < r$loa_inf)
  ba <- ggplot(d, aes(media, dif)) +
    geom_hline(yintercept = 0, colour = "grey70") +
    geom_hline(yintercept = r$vies, colour = "#d95f0e", linewidth = .8) +
    geom_hline(yintercept = c(r$loa_inf, r$loa_sup), colour = "#d95f0e", linetype = "dashed") +
    geom_point(size = 2.2, alpha = .85, colour = "#2c7fb8") +
    geom_text_repel(data = fora, aes(label = ID), size = 3, min.segment.length = 0) +
    scale_x_continuous(labels = num_br) + scale_y_continuous(labels = num_br) +
    labs(x = paste0("Média rótulo e laboratório (", unidade, ")"),
         y = paste0("Rótulo − laboratório (", unidade, ")"),
         subtitle = paste0("Viés ", num_br(r$vies), "; limites ", num_br(r$loa_inf), " a ", num_br(r$loa_sup))) +
    theme_minimal(base_size = 12)
  disp + ba
}

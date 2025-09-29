#!/usr/bin/env Rscript
# -*- coding: utf-8 -*-

# =============================================================================
# VISUALIZAÇÕES EM R - CONTROLE DE DEPÓSITOS
# =============================================================================
# Autor: Sistema de Análise Automatizada
# Data: 19/09/2025
# Descrição: Criação de gráficos e visualizações dos dados consolidados
# =============================================================================

# Carregar bibliotecas necessárias
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(ggplot2)
  library(tidyr)
})

# Configurações globais para gráficos
theme_set(theme_minimal())
options(stringsAsFactors = FALSE)

# =============================================================================
# FUNÇÃO: ENCONTRAR E CARREGAR DADOS
# =============================================================================
carregar_dados_para_visualizacao <- function(arquivo_excel = NULL) {
  cat("=== CARREGANDO DADOS PARA VISUALIZAÇÃO ===\n")
  
  # Se não foi especificado arquivo, tentar encontrar
  if (is.null(arquivo_excel)) {
    arquivos_excel <- list.files(pattern = "\\.xlsx$", recursive = TRUE)
    
    if (length(arquivos_excel) > 0) {
      arquivo_excel <- arquivos_excel[1]
      cat("Usando arquivo encontrado:", arquivo_excel, "\n")
    } else {
      cat("ERRO: Nenhum arquivo Excel encontrado!\n")
      return(NULL)
    }
  }
  
  # Verificar se o arquivo existe
  if (!file.exists(arquivo_excel)) {
    cat("ERRO: Arquivo não encontrado:", arquivo_excel, "\n")
    return(NULL)
  }
  
  # Carregar dados
  dados <- read_excel(arquivo_excel)
  dados <- dados[!apply(is.na(dados), 1, all), ]
  
  # Definir colunas de depósitos
  colunas_depositos <- c(
    "Nº de Depósito Inspecionado_B",
    "Nº de Depósito Inspecionado_C", 
    "Nº de Depósito Inspecionado_D1",
    "Nº de Depósito Inspecionado_D2"
  )
  
  # Verificar se as colunas existem
  colunas_existentes <- colunas_depositos[colunas_depositos %in% names(dados)]
  
  if (length(colunas_existentes) == 0) {
    cat("ERRO: Colunas de depósitos não encontradas\n")
    return(NULL)
  }
  
  # Converter colunas numéricas
  for (col in colunas_existentes) {
    dados[[col]] <- as.numeric(dados[[col]])
    dados[[col]][is.na(dados[[col]])] <- 0
  }
  
  # Converter outras colunas numéricas
  colunas_numericas <- c("Qtde. Dep. Tratado", "Depósito Eliminado")
  for (col in colunas_numericas) {
    if (col %in% names(dados)) {
      dados[[col]] <- as.numeric(dados[[col]])
      dados[[col]][is.na(dados[[col]])] <- 0
    }
  }
  
  # Verificar se existe coluna de localidade
  if (!"Nome da Localidade" %in% names(dados)) {
    cat("ERRO: Coluna 'Nome da Localidade' não encontrada\n")
    return(NULL)
  }
  
  # Consolidar por localidade
  consolidado <- dados %>%
    group_by(`Nome da Localidade`) %>%
    summarise(
      Deposito_B = sum(.data[["Nº de Depósito Inspecionado_B"]], na.rm = TRUE),
      Deposito_C = sum(.data[["Nº de Depósito Inspecionado_C"]], na.rm = TRUE),
      Deposito_D1 = sum(.data[["Nº de Depósito Inspecionado_D1"]], na.rm = TRUE),
      Deposito_D2 = sum(.data[["Nº de Depósito Inspecionado_D2"]], na.rm = TRUE),
      Dep_Tratados = sum(.data[["Qtde. Dep. Tratado"]], na.rm = TRUE),
      Dep_Eliminados = sum(.data[["Depósito Eliminado"]], na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    mutate(
      Total_Inspecionados = Deposito_B + Deposito_C + Deposito_D1 + Deposito_D2,
      Eficiencia_Eliminacao = ifelse(Total_Inspecionados > 0, 
                                   (Dep_Eliminados / Total_Inspecionados) * 100, 
                                   0)
    )
  
  cat("Dados processados para", nrow(consolidado), "localidades\n")
  return(consolidado)
}

# =============================================================================
# FUNÇÃO: GRÁFICO 1 - TOTAL DE DEPÓSITOS INSPECIONADOS POR LOCALIDADE
# =============================================================================
grafico_total_inspecionados <- function(consolidado) {
  cat("Criando gráfico: Total de Depósitos Inspecionados\n")
  
  p1 <- ggplot(consolidado, aes(x = reorder(`Nome da Localidade`, Total_Inspecionados), 
                               y = Total_Inspecionados)) +
    geom_col(fill = "skyblue", color = "black", alpha = 0.8) +
    geom_text(aes(label = Total_Inspecionados), 
              hjust = -0.1, size = 3.5, fontface = "bold") +
    coord_flip() +
    labs(
      title = "Total de Depósitos Inspecionados por Localidade",
      x = "Localidade",
      y = "Quantidade de Depósitos",
      caption = "Fonte: Dados de Controle de Depósitos"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.text = element_text(size = 10),
      axis.title = element_text(size = 12, face = "bold"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank()
    )
  
  return(p1)
}

# =============================================================================
# FUNÇÃO: GRÁFICO 2 - DEPÓSITOS TRATADOS VS ELIMINADOS
# =============================================================================
grafico_tratados_eliminados <- function(consolidado) {
  cat("Criando gráfico: Depósitos Tratados vs Eliminados\n")
  
  # Preparar dados para o gráfico
  dados_tratamento <- consolidado %>%
    select(`Nome da Localidade`, Dep_Tratados, Dep_Eliminados) %>%
    pivot_longer(cols = c(Dep_Tratados, Dep_Eliminados),
                names_to = "Tipo", values_to = "Quantidade") %>%
    mutate(Tipo = case_when(
      Tipo == "Dep_Tratados" ~ "Tratados",
      Tipo == "Dep_Eliminados" ~ "Eliminados"
    ))
  
  p2 <- ggplot(dados_tratamento, aes(x = `Nome da Localidade`, y = Quantidade, fill = Tipo)) +
    geom_col(position = "dodge", alpha = 0.8, color = "black") +
    geom_text(aes(label = Quantidade), 
              position = position_dodge(width = 0.9), 
              vjust = -0.3, size = 3) +
    scale_fill_manual(values = c("Tratados" = "lightgreen", "Eliminados" = "coral")) +
    labs(
      title = "Depósitos Tratados vs Eliminados por Localidade",
      x = "Localidade",
      y = "Quantidade de Depósitos",
      fill = "Tipo de Ação",
      caption = "Fonte: Dados de Controle de Depósitos"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1, size = 10),
      axis.text.y = element_text(size = 10),
      axis.title = element_text(size = 12, face = "bold"),
      legend.position = "top",
      panel.grid.minor = element_blank()
    )
  
  return(p2)
}

# =============================================================================
# FUNÇÃO: GRÁFICO 3 - DISTRIBUIÇÃO POR TIPO DE DEPÓSITO (PIZZA)
# =============================================================================
grafico_distribuicao_tipos <- function(consolidado) {
  cat("Criando gráfico: Distribuição por Tipo de Depósito\n")
  
  # Preparar dados para gráfico de pizza
  tipos_deposito <- data.frame(
    Tipo = c("Tipo B", "Tipo C", "Tipo D1", "Tipo D2"),
    Quantidade = c(
      sum(consolidado$Deposito_B),
      sum(consolidado$Deposito_C),
      sum(consolidado$Deposito_D1),
      sum(consolidado$Deposito_D2)
    )
  ) %>%
    filter(Quantidade > 0) %>%  # Remover tipos com quantidade zero
    mutate(
      Percentual = round((Quantidade / sum(Quantidade)) * 100, 1),
      Label = paste0(Tipo, "\n", Percentual, "%")
    )
  
  p3 <- ggplot(tipos_deposito, aes(x = "", y = Quantidade, fill = Tipo)) +
    geom_col(color = "white", size = 1) +
    coord_polar(theta = "y") +
    geom_text(aes(label = Label), 
              position = position_stack(vjust = 0.5),
              size = 3.5, fontface = "bold") +
    scale_fill_brewer(type = "qual", palette = "Set3") +
    labs(
      title = "Distribuição por Tipo de Depósito",
      caption = "Fonte: Dados de Controle de Depósitos"
    ) +
    theme_void() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      legend.position = "bottom"
    )
  
  return(p3)
}

# =============================================================================
# FUNÇÃO: GRÁFICO 4 - EFICIÊNCIA DE ELIMINAÇÃO
# =============================================================================
grafico_eficiencia <- function(consolidado) {
  cat("Criando gráfico: Eficiência de Eliminação\n")
  
  p4 <- ggplot(consolidado, aes(x = reorder(`Nome da Localidade`, Eficiencia_Eliminacao), 
                               y = Eficiencia_Eliminacao)) +
    geom_col(fill = "gold", color = "black", alpha = 0.8) +
    geom_text(aes(label = paste0(round(Eficiencia_Eliminacao, 1), "%")), 
              hjust = -0.1, size = 3.5, fontface = "bold") +
    coord_flip() +
    labs(
      title = "Eficiência de Eliminação por Localidade",
      x = "Localidade",
      y = "Eficiência (%)",
      caption = "Fonte: Dados de Controle de Depósitos"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.text = element_text(size = 10),
      axis.title = element_text(size = 12, face = "bold"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank()
    )
  
  return(p4)
}

# =============================================================================
# FUNÇÃO: GRÁFICO 5 - COMPARATIVO DETALHADO POR TIPO
# =============================================================================
grafico_comparativo_detalhado <- function(consolidado) {
  cat("Criando gráfico: Comparativo Detalhado por Tipo\n")
  
  # Preparar dados para gráfico de barras agrupadas
  dados_detalhados <- consolidado %>%
    select(`Nome da Localidade`, Deposito_B, Deposito_C, Deposito_D1, Deposito_D2, Dep_Eliminados) %>%
    pivot_longer(cols = -`Nome da Localidade`,
                names_to = "Tipo", values_to = "Quantidade") %>%
    mutate(Tipo = case_when(
      Tipo == "Deposito_B" ~ "Tipo B",
      Tipo == "Deposito_C" ~ "Tipo C", 
      Tipo == "Deposito_D1" ~ "Tipo D1",
      Tipo == "Deposito_D2" ~ "Tipo D2",
      Tipo == "Dep_Eliminados" ~ "Eliminados"
    ))
  
  p5 <- ggplot(dados_detalhados, aes(x = `Nome da Localidade`, y = Quantidade, fill = Tipo)) +
    geom_col(position = "dodge", alpha = 0.8, color = "black") +
    scale_fill_brewer(type = "qual", palette = "Set2") +
    labs(
      title = "Comparativo Detalhado por Localidade e Tipo de Depósito",
      x = "Localidade",
      y = "Quantidade de Depósitos",
      fill = "Tipo de Depósito",
      caption = "Fonte: Dados de Controle de Depósitos"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1, size = 10),
      axis.text.y = element_text(size = 10),
      axis.title = element_text(size = 12, face = "bold"),
      legend.position = "top",
      panel.grid.minor = element_blank()
    )
  
  return(p5)
}

# =============================================================================
# FUNÇÃO: SALVAR GRÁFICOS
# =============================================================================
salvar_graficos <- function(consolidado) {
  cat("\n=== CRIANDO E SALVANDO VISUALIZAÇÕES ===\n")
  
  if (is.null(consolidado) || nrow(consolidado) == 0) {
    cat("ERRO: Dados consolidados não disponíveis\n")
    return(NULL)
  }
  
  # Criar gráficos
  g1 <- grafico_total_inspecionados(consolidado)
  g2 <- grafico_tratados_eliminados(consolidado)
  g3 <- grafico_distribuicao_tipos(consolidado)
  g4 <- grafico_eficiencia(consolidado)
  g5 <- grafico_comparativo_detalhado(consolidado)
  
  # Salvar gráficos individuais
  ggsave("grafico_total_inspecionados.png", g1, 
         width = 10, height = 6, dpi = 300, bg = "white")
  
  ggsave("grafico_tratados_eliminados.png", g2, 
         width = 12, height = 8, dpi = 300, bg = "white")
  
  ggsave("grafico_distribuicao_tipos.png", g3, 
         width = 8, height = 8, dpi = 300, bg = "white")
  
  ggsave("grafico_eficiencia.png", g4, 
         width = 10, height = 6, dpi = 300, bg = "white")
  
  ggsave("grafico_comparativo_detalhado.png", g5, 
         width = 14, height = 8, dpi = 300, bg = "white")
  
  cat("Gráficos salvos:\n")
  cat("- grafico_total_inspecionados.png\n")
  cat("- grafico_tratados_eliminados.png\n")
  cat("- grafico_distribuicao_tipos.png\n")
  cat("- grafico_eficiencia.png\n")
  cat("- grafico_comparativo_detalhado.png\n")
  
  return(list(g1 = g1, g2 = g2, g3 = g3, g4 = g4, g5 = g5))
}

# =============================================================================
# FUNÇÃO PRINCIPAL PARA VISUALIZAÇÕES
# =============================================================================
main_visualizacoes <- function(arquivo_excel = NULL) {
  cat("INICIANDO CRIAÇÃO DE VISUALIZAÇÕES EM R\n")
  cat(paste(rep("=", 45), collapse = ""), "\n")
  
  tryCatch({
    # Carregar dados processados
    consolidado <- carregar_dados_para_visualizacao(arquivo_excel)
    
    if (is.null(consolidado)) {
      cat("ERRO: Não foi possível carregar dados para visualização\n")
      return(NULL)
    }
    
    # Criar e salvar gráficos
    graficos <- salvar_graficos(consolidado)
    
    cat("\n", paste(rep("=", 45), collapse = ""), "\n")
    cat("VISUALIZAÇÕES CRIADAS COM SUCESSO!\n")
    cat(paste(rep("=", 45), collapse = ""), "\n")
    
    return(graficos)
    
  }, error = function(e) {
    cat("Erro durante criação das visualizações:", e$message, "\n")
    return(NULL)
  })
}

# =============================================================================
# EXECUTAR SE SCRIPT FOR CHAMADO DIRETAMENTE
# =============================================================================
if (!interactive()) {
  graficos <- main_visualizacoes()
}

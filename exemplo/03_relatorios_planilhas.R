#!/usr/bin/env Rscript
# -*- coding: utf-8 -*-

# =============================================================================
# GERAÇÃO DE RELATÓRIOS E PLANILHAS EM R - CONTROLE DE DEPÓSITOS
# =============================================================================
# Autor: Sistema de Análise Automatizada
# Data: 19/09/2025
# Descrição: Criação de planilhas Excel consolidadas e relatórios detalhados
# =============================================================================

# Carregar bibliotecas necessárias
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(openxlsx)
})

# Configurações globais
options(stringsAsFactors = FALSE)

# =============================================================================
# FUNÇÃO: ENCONTRAR E PROCESSAR DADOS COMPLETOS
# =============================================================================
processar_dados_completos <- function(arquivo_excel = NULL) {
  cat("=== PROCESSANDO DADOS COMPLETOS ===\n")
  
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
  
  # Carregar dados originais
  dados_originais <- read_excel(arquivo_excel)
  dados_originais <- dados_originais[!apply(is.na(dados_originais), 1, all), ]
  
  # Definir colunas de depósitos
  colunas_depositos <- c(
    "Nº de Depósito Inspecionado_B",
    "Nº de Depósito Inspecionado_C", 
    "Nº de Depósito Inspecionado_D1",
    "Nº de Depósito Inspecionado_D2"
  )
  
  # Verificar se as colunas existem
  colunas_existentes <- colunas_depositos[colunas_depositos %in% names(dados_originais)]
  
  if (length(colunas_existentes) == 0) {
    cat("ERRO: Colunas de depósitos não encontradas\n")
    return(NULL)
  }
  
  # Converter colunas numéricas
  for (col in colunas_existentes) {
    dados_originais[[col]] <- as.numeric(dados_originais[[col]])
    dados_originais[[col]][is.na(dados_originais[[col]])] <- 0
  }
  
  # Converter outras colunas numéricas
  colunas_numericas <- c("Qtde. Dep. Tratado", "Depósito Eliminado")
  for (col in colunas_numericas) {
    if (col %in% names(dados_originais)) {
      dados_originais[[col]] <- as.numeric(dados_originais[[col]])
      dados_originais[[col]][is.na(dados_originais[[col]])] <- 0
    }
  }
  
  # Verificar se existe coluna de localidade
  if (!"Nome da Localidade" %in% names(dados_originais)) {
    cat("ERRO: Coluna 'Nome da Localidade' não encontrada\n")
    return(NULL)
  }
  
  # Consolidar por localidade
  consolidado_localidades <- dados_originais %>%
    group_by(`Nome da Localidade`) %>%
    summarise(
      Num_Registros = n(),
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
      Eficiencia_Eliminacao_Pct = ifelse(Total_Inspecionados > 0, 
                                        round((Dep_Eliminados / Total_Inspecionados) * 100, 2), 
                                        0)
    )
  
  # Calcular resumo geral
  resumo_geral <- data.frame(
    Metrica = c(
      "Total de Localidades",
      "Total de Registros",
      "Total Depósitos Tipo B",
      "Total Depósitos Tipo C",
      "Total Depósitos Tipo D1",
      "Total Depósitos Tipo D2",
      "Total Depósitos Inspecionados",
      "Total Depósitos Tratados",
      "Total Depósitos Eliminados",
      "Eficiência Geral (%)"
    ),
    Valor = c(
      nrow(consolidado_localidades),
      sum(consolidado_localidades$Num_Registros),
      sum(consolidado_localidades$Deposito_B),
      sum(consolidado_localidades$Deposito_C),
      sum(consolidado_localidades$Deposito_D1),
      sum(consolidado_localidades$Deposito_D2),
      sum(consolidado_localidades$Total_Inspecionados),
      sum(consolidado_localidades$Dep_Tratados),
      sum(consolidado_localidades$Dep_Eliminados),
      ifelse(sum(consolidado_localidades$Total_Inspecionados) > 0,
             round((sum(consolidado_localidades$Dep_Eliminados) / 
                   sum(consolidado_localidades$Total_Inspecionados)) * 100, 2),
             0)
    )
  )
  
  # Análise por tipos de depósito
  analise_tipos <- data.frame(
    Tipo_Deposito = c("Tipo B", "Tipo C", "Tipo D1", "Tipo D2"),
    Total_Inspecionados = c(
      sum(consolidado_localidades$Deposito_B),
      sum(consolidado_localidades$Deposito_C),
      sum(consolidado_localidades$Deposito_D1),
      sum(consolidado_localidades$Deposito_D2)
    )
  )
  
  # Calcular percentuais
  total_geral <- sum(analise_tipos$Total_Inspecionados)
  analise_tipos$Percentual_Pct <- ifelse(total_geral > 0,
                                        round((analise_tipos$Total_Inspecionados / total_geral) * 100, 2),
                                        0)
  
  cat("Processamento completo finalizado\n")
  cat("- Dados originais:", nrow(dados_originais), "registros\n")
  cat("- Consolidado:", nrow(consolidado_localidades), "localidades\n")
  
  return(list(
    dados_originais = dados_originais,
    consolidado_localidades = consolidado_localidades,
    resumo_geral = resumo_geral,
    analise_tipos = analise_tipos
  ))
}

# =============================================================================
# FUNÇÃO: CRIAR PLANILHA EXCEL CONSOLIDADA
# =============================================================================
criar_planilha_excel <- function(dados_processados, arquivo_saida = "planilha_consolidada.xlsx") {
  cat("\n=== CRIANDO PLANILHA EXCEL CONSOLIDADA ===\n")
  
  if (is.null(dados_processados)) {
    cat("ERRO: Dados processados não disponíveis\n")
    return(FALSE)
  }
  
  # Criar workbook
  wb <- createWorkbook()
  
  # Estilo para cabeçalhos
  header_style <- createStyle(
    fontSize = 12,
    fontColour = "white",
    halign = "center",
    fgFill = "#4F81BD",
    border = "TopBottomLeftRight",
    borderColour = "black",
    textDecoration = "bold"
  )
  
  # Aba 1: Dados Originais
  addWorksheet(wb, "Dados_Originais")
  writeData(wb, "Dados_Originais", dados_processados$dados_originais, startRow = 1)
  addStyle(wb, "Dados_Originais", header_style, rows = 1, cols = 1:ncol(dados_processados$dados_originais))
  
  # Aba 2: Consolidado por Localidades
  addWorksheet(wb, "Consolidado_Localidades")
  writeData(wb, "Consolidado_Localidades", dados_processados$consolidado_localidades, startRow = 1)
  addStyle(wb, "Consolidado_Localidades", header_style, rows = 1, cols = 1:ncol(dados_processados$consolidado_localidades))
  
  # Aba 3: Resumo Geral
  addWorksheet(wb, "Resumo_Geral")
  writeData(wb, "Resumo_Geral", dados_processados$resumo_geral, startRow = 1)
  addStyle(wb, "Resumo_Geral", header_style, rows = 1, cols = 1:ncol(dados_processados$resumo_geral))
  
  # Aba 4: Análise por Tipos
  addWorksheet(wb, "Analise_Tipos")
  writeData(wb, "Analise_Tipos", dados_processados$analise_tipos, startRow = 1)
  addStyle(wb, "Analise_Tipos", header_style, rows = 1, cols = 1:ncol(dados_processados$analise_tipos))
  
  # Aba 5: Informações do Relatório
  info_relatorio <- data.frame(
    Campo = c(
      "Data de Geração",
      "Sistema",
      "Linguagem",
      "Total de Abas",
      "Descrição"
    ),
    Valor = c(
      as.character(Sys.time()),
      "Sistema de Análise Automatizada",
      paste("R versão", R.version.string),
      "5 abas",
      "Análise completa de controle de depósitos"
    )
  )
  
  addWorksheet(wb, "Info_Relatorio")
  writeData(wb, "Info_Relatorio", info_relatorio, startRow = 1)
  addStyle(wb, "Info_Relatorio", header_style, rows = 1, cols = 1:ncol(info_relatorio))
  
  # Salvar planilha
  tryCatch({
    saveWorkbook(wb, arquivo_saida, overwrite = TRUE)
    cat("Planilha Excel criada:", arquivo_saida, "\n")
    cat("Abas criadas:\n")
    cat("- Dados_Originais: Dados brutos da planilha original\n")
    cat("- Consolidado_Localidades: Dados agrupados por localidade\n")
    cat("- Resumo_Geral: Métricas gerais do projeto\n")
    cat("- Analise_Tipos: Distribuição por tipos de depósito\n")
    cat("- Info_Relatorio: Informações sobre o relatório\n")
    return(TRUE)
  }, error = function(e) {
    cat("ERRO ao salvar planilha:", e$message, "\n")
    return(FALSE)
  })
}

# =============================================================================
# FUNÇÃO: GERAR RELATÓRIO DETALHADO EM MARKDOWN
# =============================================================================
gerar_relatorio_markdown <- function(dados_processados, arquivo_saida = "relatorio_detalhado.md") {
  cat("\n=== GERANDO RELATÓRIO MARKDOWN ===\n")
  
  if (is.null(dados_processados)) {
    cat("ERRO: Dados processados não disponíveis\n")
    return(FALSE)
  }
  
  # Criar conteúdo do relatório
  relatorio_md <- c(
    "# Relatório de Análise - Controle de Depósitos",
    "",
    paste("**Data de geração:** ", Sys.time()),
    "**Sistema:** Sistema de Análise Automatizada em R",
    paste("**Linguagem:** R versão", R.version.string),
    "",
    "---",
    "",
    "## Resumo Executivo",
    "",
    paste("Este relatório apresenta a análise completa dos dados de controle de depósitos,"),
    paste("processando", nrow(dados_processados$dados_originais), "registros distribuídos em"),
    paste(nrow(dados_processados$consolidado_localidades), "localidades distintas."),
    "",
    "### Principais Indicadores",
    ""
  )
  
  # Adicionar tabela de resumo geral
  relatorio_md <- c(relatorio_md, "| Métrica | Valor |")
  relatorio_md <- c(relatorio_md, "|---------|-------|")
  
  for (i in 1:nrow(dados_processados$resumo_geral)) {
    linha <- paste("|", dados_processados$resumo_geral$Metrica[i], "|", 
                  dados_processados$resumo_geral$Valor[i], "|")
    relatorio_md <- c(relatorio_md, linha)
  }
  
  relatorio_md <- c(relatorio_md, "", "---", "", "## Análise por Localidade", "")
  
  # Adicionar tabela consolidada por localidade
  relatorio_md <- c(relatorio_md, "| Localidade | Inspecionados | Tratados | Eliminados | Eficiência (%) |")
  relatorio_md <- c(relatorio_md, "|------------|---------------|----------|------------|----------------|")
  
  for (i in 1:nrow(dados_processados$consolidado_localidades)) {
    linha <- paste("|", dados_processados$consolidado_localidades$`Nome da Localidade`[i], "|",
                  dados_processados$consolidado_localidades$Total_Inspecionados[i], "|",
                  dados_processados$consolidado_localidades$Dep_Tratados[i], "|",
                  dados_processados$consolidado_localidades$Dep_Eliminados[i], "|",
                  dados_processados$consolidado_localidades$Eficiencia_Eliminacao_Pct[i], "|")
    relatorio_md <- c(relatorio_md, linha)
  }
  
  relatorio_md <- c(relatorio_md, "", "---", "", "## Distribuição por Tipos de Depósito", "")
  
  # Adicionar tabela de análise por tipos
  relatorio_md <- c(relatorio_md, "| Tipo de Depósito | Total Inspecionados | Percentual (%) |")
  relatorio_md <- c(relatorio_md, "|------------------|---------------------|----------------|")
  
  for (i in 1:nrow(dados_processados$analise_tipos)) {
    linha <- paste("|", dados_processados$analise_tipos$Tipo_Deposito[i], "|",
                  dados_processados$analise_tipos$Total_Inspecionados[i], "|",
                  dados_processados$analise_tipos$Percentual_Pct[i], "|")
    relatorio_md <- c(relatorio_md, linha)
  }
  
  relatorio_md <- c(relatorio_md, "", "---", "", "## Conclusões", "")
  
  # Adicionar conclusões automáticas
  if (nrow(dados_processados$consolidado_localidades) > 0) {
    melhor_localidade <- dados_processados$consolidado_localidades[
      which.max(dados_processados$consolidado_localidades$Eficiencia_Eliminacao_Pct), 
      "Nome da Localidade"
    ]
    
    maior_volume <- dados_processados$consolidado_localidades[
      which.max(dados_processados$consolidado_localidades$Total_Inspecionados), 
      "Nome da Localidade"
    ]
    
    eficiencia_geral <- dados_processados$resumo_geral[
      dados_processados$resumo_geral$Metrica == "Eficiência Geral (%)", "Valor"
    ]
    
    relatorio_md <- c(relatorio_md,
      paste("- **Melhor eficiência:** A localidade", melhor_localidade, "apresentou a maior eficiência de eliminação."),
      paste("- **Maior volume:** A localidade", maior_volume, "registrou o maior número de depósitos inspecionados."),
      paste("- **Eficiência geral:** O projeto alcançou uma eficiência geral de eliminação de", eficiencia_geral, "%."),
      "",
      "---",
      "",
      "*Relatório gerado automaticamente pelo Sistema de Análise em R*"
    )
  }
  
  # Salvar relatório
  tryCatch({
    writeLines(relatorio_md, arquivo_saida, useBytes = TRUE)
    cat("Relatório Markdown salvo:", arquivo_saida, "\n")
    return(TRUE)
  }, error = function(e) {
    cat("ERRO ao salvar relatório:", e$message, "\n")
    return(FALSE)
  })
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================
main_relatorios <- function(arquivo_excel = NULL) {
  cat("INICIANDO GERAÇÃO DE RELATÓRIOS E PLANILHAS EM R\n")
  cat(paste(rep("=", 50), collapse = ""), "\n")
  
  tryCatch({
    # 1. Processar dados completos
    dados_processados <- processar_dados_completos(arquivo_excel)
    
    if (is.null(dados_processados)) {
      cat("ERRO: Não foi possível processar os dados\n")
      return(NULL)
    }
    
    # 2. Criar planilha Excel consolidada
    sucesso_excel <- criar_planilha_excel(dados_processados)
    
    # 3. Gerar relatório em Markdown
    sucesso_md <- gerar_relatorio_markdown(dados_processados)
    
    if (sucesso_excel && sucesso_md) {
      cat("\n", paste(rep("=", 50), collapse = ""), "\n")
      cat("RELATÓRIOS E PLANILHAS GERADOS COM SUCESSO!\n")
      cat(paste(rep("=", 50), collapse = ""), "\n")
    } else {
      cat("AVISO: Alguns arquivos podem não ter sido gerados corretamente\n")
    }
    
    return(dados_processados)
    
  }, error = function(e) {
    cat("Erro durante geração de relatórios:", e$message, "\n")
    return(NULL)
  })
}

# =============================================================================
# EXECUTAR SE SCRIPT FOR CHAMADO DIRETAMENTE
# =============================================================================
if (!interactive()) {
  resultado <- main_relatorios()
}

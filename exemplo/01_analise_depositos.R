#!/usr/bin/env Rscript
# -*- coding: utf-8 -*-

# =============================================================================
# ANÁLISE DE CONTROLE DE DEPÓSITOS - SCRIPT PRINCIPAL EM R
# =============================================================================
# Autor: Sistema de Análise Automatizada
# Data: 19/09/2025
# Descrição: Análise completa dos dados de controle de depósitos
# =============================================================================

# Carregar bibliotecas necessárias
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(ggplot2)
})

# Configurações globais
options(stringsAsFactors = FALSE)

# =============================================================================
# FUNÇÃO: ENCONTRAR ARQUIVO DE DADOS
# =============================================================================
encontrar_arquivo_dados <- function() {
  # Possíveis localizações do arquivo
  caminhos_possiveis <- c(
    "dados.xlsx",
    "teste.xlsx", 
    "planilha.xlsx",
    "dados/dados.xlsx",
    "dados/teste.xlsx",
    "data/dados.xlsx"
  )
  
  # Procurar arquivos Excel no diretório atual
  arquivos_excel <- list.files(pattern = "\\.xlsx$", recursive = TRUE)
  
  if (length(arquivos_excel) > 0) {
    cat("Arquivos Excel encontrados:\n")
    for (i in seq_along(arquivos_excel)) {
      cat(paste0(i, ". ", arquivos_excel[i], "\n"))
    }
    return(arquivos_excel[1])  # Usar o primeiro encontrado
  }
  
  # Verificar caminhos específicos
  for (caminho in caminhos_possiveis) {
    if (file.exists(caminho)) {
      return(caminho)
    }
  }
  
  return(NULL)
}

# =============================================================================
# FUNÇÃO: CARREGAR E PROCESSAR DADOS
# =============================================================================
carregar_dados <- function(arquivo_excel = NULL) {
  cat("=== CARREGANDO DADOS DA PLANILHA ===\n")
  
  # Se não foi especificado arquivo, tentar encontrar
  if (is.null(arquivo_excel)) {
    arquivo_excel <- encontrar_arquivo_dados()
  }
  
  # Verificar se o arquivo existe
  if (is.null(arquivo_excel) || !file.exists(arquivo_excel)) {
    cat("ERRO: Nenhum arquivo Excel encontrado!\n")
    cat("Coloque sua planilha (.xlsx) no diretório do projeto\n")
    cat("Ou especifique o caminho correto na função carregar_dados()\n")
    return(NULL)
  }
  
  cat("Usando arquivo:", arquivo_excel, "\n")
  
  # Carregar dados
  dados <- read_excel(arquivo_excel)
  
  # Remover linhas completamente vazias
  dados <- dados[!apply(is.na(dados), 1, all), ]
  
  cat("Total de registros carregados:", nrow(dados), "\n")
  cat("Total de colunas:", ncol(dados), "\n")
  
  # Exibir estrutura dos dados
  cat("\nPrimeiras colunas encontradas:\n")
  cat(paste(names(dados)[1:min(5, ncol(dados))], collapse = ", "), "\n")
  
  return(dados)
}

# =============================================================================
# FUNÇÃO: PROCESSAR E CALCULAR ESTATÍSTICAS
# =============================================================================
processar_dados <- function(dados) {
  cat("\n=== PROCESSANDO DADOS ===\n")
  
  if (is.null(dados)) {
    cat("Dados não disponíveis para processamento\n")
    return(NULL)
  }
  
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
    cat("AVISO: Colunas de depósitos não encontradas com nomes esperados\n")
    cat("Colunas disponíveis:\n")
    cat(paste(names(dados), collapse = ", "), "\n")
    return(NULL)
  }
  
  # Converter colunas numéricas e substituir NA por 0
  for (col in colunas_existentes) {
    dados[[col]] <- as.numeric(dados[[col]])
    dados[[col]][is.na(dados[[col]])] <- 0
  }
  
  # Converter outras colunas numéricas importantes
  colunas_numericas <- c("Qtde. Dep. Tratado", "Depósito Eliminado")
  for (col in colunas_numericas) {
    if (col %in% names(dados)) {
      dados[[col]] <- as.numeric(dados[[col]])
      dados[[col]][is.na(dados[[col]])] <- 0
    }
  }
  
  # Consolidar por localidade
  if ("Nome da Localidade" %in% names(dados)) {
    consolidado <- dados %>%
      group_by(`Nome da Localidade`) %>%
      summarise(
        Num_Registros = n(),
        Deposito_B = sum(get("Nº de Depósito Inspecionado_B", pos = cur_data()), na.rm = TRUE),
        Deposito_C = sum(get("Nº de Depósito Inspecionado_C", pos = cur_data()), na.rm = TRUE),
        Deposito_D1 = sum(get("Nº de Depósito Inspecionado_D1", pos = cur_data()), na.rm = TRUE),
        Deposito_D2 = sum(get("Nº de Depósito Inspecionado_D2", pos = cur_data()), na.rm = TRUE),
        Dep_Tratados = sum(get("Qtde. Dep. Tratado", pos = cur_data()), na.rm = TRUE),
        Dep_Eliminados = sum(get("Depósito Eliminado", pos = cur_data()), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      mutate(
        Total_Inspecionados = Deposito_B + Deposito_C + Deposito_D1 + Deposito_D2,
        Eficiencia_Eliminacao = ifelse(Total_Inspecionados > 0, 
                                     round((Dep_Eliminados / Total_Inspecionados) * 100, 2), 
                                     0)
      )
    
    cat("Consolidado por localidade criado com", nrow(consolidado), "localidades\n")
    print(consolidado)
    
    return(list(dados_originais = dados, consolidado = consolidado))
  } else {
    cat("Coluna 'Nome da Localidade' não encontrada\n")
    cat("Colunas disponíveis:", paste(names(dados), collapse = ", "), "\n")
    return(list(dados_originais = dados, consolidado = NULL))
  }
}

# =============================================================================
# FUNÇÃO: CALCULAR RESUMO GERAL
# =============================================================================
calcular_resumo_geral <- function(consolidado) {
  cat("\n=== CALCULANDO RESUMO GERAL ===\n")
  
  if (is.null(consolidado)) {
    cat("Consolidado não disponível\n")
    return(NULL)
  }
  
  resumo <- data.frame(
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
      nrow(consolidado),
      sum(consolidado$Num_Registros),
      sum(consolidado$Deposito_B),
      sum(consolidado$Deposito_C), 
      sum(consolidado$Deposito_D1),
      sum(consolidado$Deposito_D2),
      sum(consolidado$Total_Inspecionados),
      sum(consolidado$Dep_Tratados),
      sum(consolidado$Dep_Eliminados),
      ifelse(sum(consolidado$Total_Inspecionados) > 0,
             round((sum(consolidado$Dep_Eliminados) / sum(consolidado$Total_Inspecionados)) * 100, 2),
             0)
    )
  )
  
  cat("Resumo geral calculado:\n")
  print(resumo)
  
  return(resumo)
}

# =============================================================================
# FUNÇÃO: ANÁLISE POR TIPOS DE DEPÓSITO
# =============================================================================
analisar_tipos_deposito <- function(consolidado) {
  cat("\n=== ANALISANDO TIPOS DE DEPÓSITO ===\n")
  
  if (is.null(consolidado)) {
    cat("Consolidado não disponível\n")
    return(NULL)
  }
  
  analise_tipos <- data.frame(
    Tipo_Deposito = c("Tipo B", "Tipo C", "Tipo D1", "Tipo D2"),
    Total_Inspecionados = c(
      sum(consolidado$Deposito_B),
      sum(consolidado$Deposito_C),
      sum(consolidado$Deposito_D1), 
      sum(consolidado$Deposito_D2)
    )
  )
  
  # Calcular percentuais
  total_geral <- sum(analise_tipos$Total_Inspecionados)
  analise_tipos$Percentual <- ifelse(total_geral > 0,
                                   round((analise_tipos$Total_Inspecionados / total_geral) * 100, 2),
                                   0)
  
  cat("Análise por tipos de depósito:\n")
  print(analise_tipos)
  
  return(analise_tipos)
}

# =============================================================================
# FUNÇÃO: GERAR RELATÓRIO EM TEXTO
# =============================================================================
gerar_relatorio <- function(dados_originais, consolidado, resumo, analise_tipos, arquivo_saida = "relatorio_analise_R.txt") {
  cat("\n=== GERANDO RELATÓRIO ===\n")
  
  if (is.null(dados_originais)) {
    cat("Dados não disponíveis para gerar relatório\n")
    return(NULL)
  }
  
  # Criar conteúdo do relatório
  relatorio <- c(
    "RELATÓRIO DE ANÁLISE - CONTROLE DE DEPÓSITOS",
    paste(rep("=", 50), collapse = ""),
    paste("Data de geração:", Sys.time()),
    "",
    "RESUMO GERAL:",
    paste("- Total de registros processados:", nrow(dados_originais)),
    ""
  )
  
  # Adicionar resumo geral se disponível
  if (!is.null(resumo)) {
    for (i in 1:nrow(resumo)) {
      relatorio <- c(relatorio, paste("-", resumo$Metrica[i], ":", resumo$Valor[i]))
    }
  }
  
  relatorio <- c(relatorio, "", "CONSOLIDADO POR LOCALIDADE:", paste(rep("-", 30), collapse = ""))
  
  # Adicionar dados por localidade
  if (!is.null(consolidado)) {
    for (i in 1:nrow(consolidado)) {
      localidade <- consolidado$`Nome da Localidade`[i]
      relatorio <- c(
        relatorio,
        "",
        paste(toupper(as.character(localidade)), ":"),
        paste("  • Depósitos Tipo B:", consolidado$Deposito_B[i]),
        paste("  • Depósitos Tipo C:", consolidado$Deposito_C[i]),
        paste("  • Depósitos Tipo D1:", consolidado$Deposito_D1[i]),
        paste("  • Depósitos Tipo D2:", consolidado$Deposito_D2[i]),
        paste("  • Total Inspecionados:", consolidado$Total_Inspecionados[i]),
        paste("  • Depósitos Tratados:", consolidado$Dep_Tratados[i]),
        paste("  • Depósitos Eliminados:", consolidado$Dep_Eliminados[i]),
        paste("  • Eficiência de Eliminação:", paste0(consolidado$Eficiencia_Eliminacao[i], "%"))
      )
    }
  }
  
  # Salvar relatório
  writeLines(relatorio, arquivo_saida, useBytes = TRUE)
  cat("Relatório salvo em:", arquivo_saida, "\n")
  
  return(relatorio)
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================
main <- function(arquivo_dados = NULL) {
  cat("INICIANDO ANÁLISE DA PLANILHA EM R\n")
  cat(paste(rep("=", 40), collapse = ""), "\n")
  
  tryCatch({
    # 1. Carregar dados
    dados <- carregar_dados(arquivo_dados)
    
    if (is.null(dados)) {
      cat("ERRO: Não foi possível carregar os dados\n")
      return(NULL)
    }
    
    # 2. Processar dados
    resultado <- processar_dados(dados)
    
    if (is.null(resultado)) {
      cat("ERRO: Não foi possível processar os dados\n")
      return(NULL)
    }
    
    dados_originais <- resultado$dados_originais
    consolidado <- resultado$consolidado
    
    # 3. Calcular resumo geral
    resumo <- calcular_resumo_geral(consolidado)
    
    # 4. Analisar tipos de depósito
    analise_tipos <- analisar_tipos_deposito(consolidado)
    
    # 5. Gerar relatório
    relatorio <- gerar_relatorio(dados_originais, consolidado, resumo, analise_tipos)
    
    cat("\n", paste(rep("=", 40), collapse = ""), "\n")
    cat("ANÁLISE CONCLUÍDA COM SUCESSO!\n")
    cat(paste(rep("=", 40), collapse = ""), "\n")
    
    # Retornar resultados
    return(list(
      dados_originais = dados_originais,
      consolidado = consolidado,
      resumo = resumo,
      analise_tipos = analise_tipos,
      relatorio = relatorio
    ))
    
  }, error = function(e) {
    cat("Erro durante a análise:", e$message, "\n")
    return(NULL)
  })
}

# =============================================================================
# EXECUTAR ANÁLISE SE SCRIPT FOR CHAMADO DIRETAMENTE
# =============================================================================
if (!interactive()) {
  resultado <- main()
}

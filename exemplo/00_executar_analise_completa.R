#!/usr/bin/env Rscript
# -*- coding: utf-8 -*-

# =============================================================================
# ANÁLISE COMPLETA EM R - CONTROLE DE DEPÓSITOS
# SCRIPT MASTER CORRIGIDO PARA PROJETO R
# =============================================================================
# Autor: Sistema de Análise Automatizada
# Data: 19/09/2025
# Descrição: Script principal que executa toda a análise de dados em R
# =============================================================================

library(readxl)
library(dplyr)
library(ggplot2)
library(openxlsx)
library(tidyr)


cat("
#############################################################################
#                                                                           #
#           SISTEMA DE ANÁLISE COMPLETA EM R                               #
#           CONTROLE DE DEPÓSITOS - VERSÃO PROJETO                         #
#                                                                           #
#############################################################################
")

# Configurações globais
options(stringsAsFactors = FALSE)
start_time <- Sys.time()

# =============================================================================
# VERIFICAR E CARREGAR BIBLIOTECAS
# =============================================================================
verificar_bibliotecas <- function() {
  cat("\n=== VERIFICANDO BIBLIOTECAS ===\n")
  
  bibliotecas_necessarias <- c("readxl", "dplyr", "ggplot2", "openxlsx", "tidyr")
  
  for (lib in bibliotecas_necessarias) {
    if (!requireNamespace(lib, quietly = TRUE)) {
      cat("✗", lib, "- NÃO INSTALADO\n")
      cat("Execute: install.packages('", lib, "')\n")
      stop("Bibliotecas necessárias não instaladas")
    } else {
      cat("✓", lib, "- OK\n")
    }
  }
  
  # Carregar bibliotecas
  suppressPackageStartupMessages({
    library(readxl)
    library(dplyr)
    library(ggplot2)
    library(openxlsx)
    library(tidyr)
  })
  
  cat("Todas as bibliotecas carregadas com sucesso!\n")
}

# =============================================================================
# VERIFICAR ARQUIVOS DE DADOS
# =============================================================================
verificar_dados <- function() {
  cat("\n=== VERIFICANDO ARQUIVOS DE DADOS ===\n")
  
  # Procurar arquivos Excel
  arquivos_excel <- list.files(pattern = "\\.xlsx$", recursive = TRUE)
  
  if (length(arquivos_excel) > 0) {
    cat("Arquivos Excel encontrados:\n")
    for (i in seq_along(arquivos_excel)) {
      cat(paste0("  ", i, ". ", arquivos_excel[i], "\n"))
    }
    return(arquivos_excel[1])
  } else {
    cat("AVISO: Nenhum arquivo Excel (.xlsx) encontrado\n")
    cat("Coloque sua planilha de dados no diretório do projeto\n")
    return(NULL)
  }
}

# =============================================================================
# EXECUTAR ANÁLISE PRINCIPAL
# =============================================================================
executar_analise_principal <- function(arquivo_dados) {
  cat("\n=== EXECUTANDO ANÁLISE PRINCIPAL ===\n")
  
  if (file.exists("01_analise_depositos.R")) {
    # Carregar funções do script
    source("01_analise_depositos.R")
    
    # Executar análise
    resultado <- main(arquivo_dados)
    
    if (!is.null(resultado)) {
      cat("✓ Análise principal executada com sucesso\n")
      return(resultado)
    } else {
      cat("✗ Erro na análise principal\n")
      return(NULL)
    }
  } else {
    cat("✗ Script 01_analise_depositos.R não encontrado\n")
    return(NULL)
  }
}

# =============================================================================
# EXECUTAR VISUALIZAÇÕES
# =============================================================================
executar_visualizacoes <- function(arquivo_dados) {
  cat("\n=== EXECUTANDO VISUALIZAÇÕES ===\n")
  
  if (file.exists("02_visualizacoes.R")) {
    # Carregar funções do script
    source("02_visualizacoes.R")
    
    # Executar visualizações
    graficos <- main_visualizacoes(arquivo_dados)
    
    if (!is.null(graficos)) {
      cat("✓ Visualizações criadas com sucesso\n")
      return(graficos)
    } else {
      cat("✗ Erro na criação de visualizações\n")
      return(NULL)
    }
  } else {
    cat("✗ Script 02_visualizacoes.R não encontrado\n")
    return(NULL)
  }
}

# =============================================================================
# EXECUTAR RELATÓRIOS E PLANILHAS
# =============================================================================
executar_relatorios <- function(arquivo_dados) {
  cat("\n=== EXECUTANDO RELATÓRIOS E PLANILHAS ===\n")
  
  if (file.exists("03_relatorios_planilhas.R")) {
    # Carregar funções do script
    source("03_relatorios_planilhas.R")
    
    # Executar relatórios
    resultado <- main_relatorios(arquivo_dados)
    
    if (!is.null(resultado)) {
      cat("✓ Relatórios e planilhas gerados com sucesso\n")
      return(resultado)
    } else {
      cat("✗ Erro na geração de relatórios\n")
      return(NULL)
    }
  } else {
    cat("✗ Script 03_relatorios_planilhas.R não encontrado\n")
    return(NULL)
  }
}

# =============================================================================
# LISTAR ARQUIVOS GERADOS
# =============================================================================
listar_arquivos_gerados <- function() {
  cat("\n=== ARQUIVOS GERADOS ===\n")
  
  arquivos_esperados <- c(
    "relatorio_analise_R.txt",
    "grafico_total_inspecionados.png",
    "grafico_tratados_eliminados.png", 
    "grafico_distribuicao_tipos.png",
    "grafico_eficiencia.png",
    "grafico_comparativo_detalhado.png",
    "planilha_consolidada.xlsx",
    "relatorio_detalhado.md"
  )
  
  arquivos_encontrados <- 0
  
  for (arquivo in arquivos_esperados) {
    if (file.exists(arquivo)) {
      cat("✓", arquivo, "\n")
      arquivos_encontrados <- arquivos_encontrados + 1
    } else {
      cat("✗", arquivo, "- NÃO ENCONTRADO\n")
    }
  }
  
  cat("\nTotal de arquivos gerados:", arquivos_encontrados, "de", length(arquivos_esperados), "\n")
  
  return(arquivos_encontrados)
}

# =============================================================================
# GERAR RESUMO FINAL
# =============================================================================
gerar_resumo_final <- function(tempo_execucao, arquivos_gerados, arquivo_dados) {
  cat("\n=== RESUMO FINAL DA EXECUÇÃO ===\n")
  
  resumo <- c(
    "RELATÓRIO DE EXECUÇÃO - ANÁLISE COMPLETA EM R",
    paste(rep("=", 50), collapse = ""),
    paste("Data/Hora de execução:", Sys.time()),
    paste("Tempo total de execução:", round(tempo_execucao, 2), "segundos"),
    paste("Linguagem utilizada: R versão", R.version.string),
    paste("Sistema operacional:", Sys.info()["sysname"]),
    paste("Arquivo de dados utilizado:", ifelse(is.null(arquivo_dados), "Não encontrado", arquivo_dados)),
    "",
    "COMPONENTES EXECUTADOS:",
    "✓ Verificação de bibliotecas",
    "✓ Análise principal de dados",
    "✓ Geração de visualizações",
    "✓ Criação de relatórios",
    "✓ Exportação de planilhas",
    "",
    paste("ARQUIVOS GERADOS:", arquivos_gerados),
    "- Relatório em texto (.txt)",
    "- 5 gráficos em PNG",
    "- Planilha Excel consolidada (.xlsx)",
    "- Relatório detalhado em Markdown (.md)",
    "",
    ifelse(arquivos_gerados >= 6, "STATUS: EXECUÇÃO CONCLUÍDA COM SUCESSO", "STATUS: EXECUÇÃO COM PROBLEMAS"),
    paste(rep("=", 50), collapse = "")
  )
  
  # Salvar resumo
  writeLines(resumo, "resumo_execucao.txt", useBytes = TRUE)
  
  # Exibir resumo
  cat(paste(resumo, collapse = "\n"), "\n")
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================
main_completa <- function() {
  tryCatch({
    # 1. Verificar bibliotecas
    verificar_bibliotecas()
    
    # 2. Verificar dados
    arquivo_dados <- verificar_dados()
    
    # 3. Executar análise principal
    resultado_analise <- executar_analise_principal(arquivo_dados)
    
    # 4. Executar visualizações
    resultado_graficos <- executar_visualizacoes(arquivo_dados)
    
    # 5. Executar relatórios
    resultado_relatorios <- executar_relatorios(arquivo_dados)
    
    # 6. Calcular tempo de execução
    end_time <- Sys.time()
    tempo_execucao <- as.numeric(difftime(end_time, start_time, units = "secs"))
    
    # 7. Listar arquivos gerados
    arquivos_gerados <- listar_arquivos_gerados()
    
    # 8. Gerar resumo final
    gerar_resumo_final(tempo_execucao, arquivos_gerados, arquivo_dados)
    
    # 9. Resultado final
    if (arquivos_gerados >= 6) {
      cat("\n🎉 ANÁLISE COMPLETA FINALIZADA COM SUCESSO! 🎉\n")
      return(TRUE)
    } else {
      cat("\n⚠️ ANÁLISE FINALIZADA COM ALGUNS PROBLEMAS ⚠️\n")
      cat("Verifique se todos os arquivos necessários estão presentes\n")
      return(FALSE)
    }
    
  }, error = function(e) {
    cat("\n❌ ERRO DURANTE A EXECUÇÃO:", e$message, "\n")
    cat("Verifique:\n")
    cat("1. Se todas as bibliotecas estão instaladas\n")
    cat("2. Se o arquivo de dados (.xlsx) está no diretório\n")
    cat("3. Se todos os scripts (01_, 02_, 03_) estão presentes\n")
    return(FALSE)
  })
}

# =============================================================================
# INSTRUÇÕES DE USO
# =============================================================================
mostrar_instrucoes <- function() {
  cat("\n=== INSTRUÇÕES DE USO ===\n")
  cat("1. Coloque sua planilha Excel (.xlsx) no diretório do projeto\n")
  cat("2. Execute este script: source('00_executar_analise_completa.R')\n")
  cat("3. Ou execute no terminal: Rscript 00_executar_analise_completa.R\n")
  cat("\nArquivos necessários no projeto:\n")
  cat("- 00_executar_analise_completa.R (este arquivo)\n")
  cat("- 01_analise_depositos.R\n")
  cat("- 02_visualizacoes.R\n")
  cat("- 03_relatorios_planilhas.R\n")
  cat("- sua_planilha.xlsx (seus dados)\n")
}

# =============================================================================
# EXECUTAR ANÁLISE COMPLETA
# =============================================================================
if (!interactive()) {
  resultado <- main_completa()
  
  if (resultado) {
    quit(status = 0)  # Sucesso
  } else {
    mostrar_instrucoes()
    quit(status = 1)  # Erro
  }
} else {
  # Se executado interativamente, mostrar instruções
  cat("Script carregado! Para executar a análise completa, digite:\n")
  cat("main_completa()\n")
}

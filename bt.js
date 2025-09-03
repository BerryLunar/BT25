/**
 * ========================================================================
 * SISTEMA BANCO DE TALENTOS - PROGRAMA GOVERNO EFICAZ
 * Santana de Parnaíba - SP
 * ========================================================================
 * 
 * Sistema automatizado para consolidação de dados do Banco de Talentos
 * das 23 secretarias municipais em uma planilha central.
 * 
 * Funcionalidades:
 * - Consolidação automática de dados das secretarias
 * - Prevenção de duplicatas por nome e prontuário
 * - Interface com botões Atualizar e Sobre
 * - Logs detalhados de execução
 * - Tratamento robusto de erros
 * 
 * @author Sistema desenvolvido para SMA - Secretaria Municipal de Administração lUANA 41331
 * @version 2.0 - Otimizada para Google Apps Script
 * ========================================================================
 */

// ========================================================================
// CONFIGURAÇÕES GLOBAIS
// ========================================================================

/**
 * IDs das 23 planilhas das secretarias municipais
 * Extraídos automaticamente das URLs fornecidas
 */
const PLANILHAS_SECRETARIAS = [
    "1Z6rfDo09m2nUQIjjAS0S3iuweToHAtIXK9idRb-RzuA",
    "14W2ewZBf-MgoKUYewj7--FLjDC1lVL6-5EUpHTHDoqU", 
    "15ztEc0wlK1mkrRd-DYkfXUn-Vq-k2JT0yLV-JaMeYV4",
    "1u7L6Qh57zFqQRCPNTJLGGReuEBlykCnD65SwwNFeRNA",
    "1Nc9O1Ha038gKY5LcfxUVClhTq6rsR0zghdSJqfScI6k",
    "1QrRoRoOsyrKgFIitAYWl1g53zYyxymWnoQCVBsRyHsM",
    "1rlyWJDE3srgUMyJy2eDlNrA4Sp57Njca9oypUkRSpP8",
    "1_4d9POGUbjKHHGPCcbQGx3-wLQuSmThKlLsFG98vDpw",
    "1n5JfTcpy8EtSlBY-bT3JLqV5kECrN_QeSM79TXbEZjk",
    "1giGaLo8jtOJ2VCSFcwhULFR5aYZcgPdifSfIKHXAm1A",
    "1WBOuLGg7hwFY7ehP1qKmuzN8-ZJR12GhMdYcNDCYHz0",
    "1cWASt4ldQbIEFm0XW1xPd7zJJZ9kNjGCylgc5XQIbPg",
    "1n2UuXYvKzz1Dau32aFwVhHLWzt2brClbJ_CFcpVa3Ks",
    "16SHJbhAb7XVmEhREFX_cVFeR2xtczizXjvpaUAwpFAk",
    "1xIWtAH9P7HZzjzroFG01KhtvRmy-sKV1YxCCBmOBxJo",
    "1qWSF0f7wVmJPM7Ht2_dtBPDBztt8PyyFwdL9rwwGTBM",
    "1eSp1C9K-AO2ZJr3ApFuDDlhGhsJXcvbxZZUn04VildM",
    "1AhWPtTgLqF_VEBNM6HxcxxJl5bkaRG5MLHPIX3Pa-3Y",
    "1EObh9xVjRqrPY1Fdz_ji7Nzq-2lcHXZKJysF6SJf66w",
    "17dEVkFJNNGanitiJYnT7Lwqz3DWsz8qODZA0K7Quggs",
    "1klqCbpMJVyCXTdpBkNrsrVfZozEOFWi1VT0xZjjSf8g",
    "1SjSlad1XQTPA0PqPUrpB7WMfPO0WxWV1f7SpHvhubrQ",
    "1LPaScCjVYVK6OVA5ZTeUXj3PXiM3nq_1GUdyiTru0jk"
  ];
  
  /**
   * Configurações do sistema
   */
  const CONFIG = {
    ABA_CENTRAL: "BT 2025",
    ABA_ORIGEM: "Banco de Talentos (Externo)",
    LINHA_INICIO_DADOS: 4, // Ignorar linhas 1-4 (cabeçalhos)
    TIMEOUT_POR_PLANILHA: 10000, // 10 segundos por planilha
    MAX_TENTATIVAS: 3
  };
  
  /**
   * Cabeçalhos da planilha central (baseados na análise do CSV)
   */
  const CABECALHOS_CENTRAL = [
    "Secretaria",
    "Nome", 
    "Prontuário",
    "Formação Acadêmica",
    "Área de Formação", 
    "Cargo Concurso",
    "CC / FE",
    "Função Gratificada",
    "Readaptado",
    "Justificativa",
    "Ação (o que)",
    "Condicionalidade", 
    "Data da Liberação"
  ];
  
  // ========================================================================
  // FUNÇÃO PRINCIPAL - EXECUTADA NA ABERTURA
  // ========================================================================
  
  /**
   * Função executada automaticamente ao abrir a planilha
   * Configura o menu personalizado e executa importação inicial
   */
  function onOpen() {
    try {
      criarMenuPersonalizado();
      Logger.log("✅ Menu personalizado criado com sucesso");
      
      // Executar importação automática
      SpreadsheetApp.getUi().alert(
        "🔄 Sistema Banco de Talentos",
        "Iniciando importação automática dos dados...\nEste processo pode demorar alguns minutos.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      importarBancoDeTalentos();
      
    } catch (erro) {
      Logger.log("❌ Erro na inicialização: " + erro.toString());
      SpreadsheetApp.getUi().alert(
        "⚠️ Erro na Inicialização", 
        "Ocorreu um erro ao inicializar o sistema. Verifique os logs para mais detalhes.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  /**
   * Cria menu personalizado na planilha
   */
  function criarMenuPersonalizado() {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu("🏛️ Banco de Talentos")
      .addItem("🔄 Atualizar Dados", "importarBancoDeTalentos")
      .addSeparator()
      .addItem("ℹ️ Sobre o Sistema", "exibirSobre")
      .addItem("📊 Relatório de Execução", "gerarRelatorioExecucao")
      .addItem("🧹 Limpar Dados", "limparDadosCentral")
      .addToUi();
  }
  
  // ========================================================================
  // FUNÇÃO PRINCIPAL DE IMPORTAÇÃO
  // ========================================================================
  
  /**
   * Função principal para importar dados do Banco de Talentos
   * Consolida dados de todas as 23 secretarias
   */
  function importarBancoDeTalentos() {
    const inicioExecucao = new Date();
    let relatorioExecucao = {
      inicio: inicioExecucao,
      secretariasProcessadas: 0,
      registrosImportados: 0,
      erros: [],
      duplicatasEvitadas: 0
    };
    
    try {
      Logger.log("🚀 === INICIANDO IMPORTAÇÃO DO BANCO DE TALENTOS ===");
      Logger.log(`📅 Data/Hora: ${inicioExecucao.toLocaleString('pt-BR')}`);
      
      // Obter planilha central
      const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
      const abaCentral = obterOuCriarAbaCentral(planilhaCentral);
      
      // Limpar dados existentes e configurar cabeçalhos
      configurarAbaCentral(abaCentral);
      
      // Controle de duplicatas
      const registrosExistentes = new Set();
      let ultimaLinha = 2; // Começar após o cabeçalho
      
      // Processar cada planilha
      PLANILHAS_SECRETARIAS.forEach((planilhaId, indice) => {
        try {
          Logger.log(`📂 Processando planilha ${indice + 1}/23: ${planilhaId}`);
          
          const resultado = processarPlanilhaSecretaria(planilhaId, registrosExistentes);
          
          if (resultado.sucesso && resultado.dados.length > 0) {
            // Inserir dados na planilha central
            const range = abaCentral.getRange(
              ultimaLinha, 1, 
              resultado.dados.length, 
              CABECALHOS_CENTRAL.length
            );
            range.setValues(resultado.dados);
            
            ultimaLinha += resultado.dados.length;
            relatorioExecucao.registrosImportados += resultado.dados.length;
            
            Logger.log(`✅ Secretaria ${resultado.siglaSecretaria}: ${resultado.dados.length} registros importados`);
          }
          
          if (resultado.duplicatasEvitadas > 0) {
            relatorioExecucao.duplicatasEvitadas += resultado.duplicatasEvitadas;
            Logger.log(`🔍 Duplicatas evitadas: ${resultado.duplicatasEvitadas}`);
          }
          
          relatorioExecucao.secretariasProcessadas++;
          
        } catch (erro) {
          const mensagemErro = `Erro na planilha ${indice + 1} (${planilhaId}): ${erro.toString()}`;
          Logger.log(`❌ ${mensagemErro}`);
          relatorioExecucao.erros.push(mensagemErro);
        }
        
        // Pequena pausa para evitar limite de API
        Utilities.sleep(500);
      });
      
      // Finalizar relatório
      relatorioExecucao.fim = new Date();
      relatorioExecucao.duracao = Math.round((relatorioExecucao.fim - relatorioExecucao.inicio) / 1000);
      
      // Aplicar formatação à planilha
      aplicarFormatacao(abaCentral, ultimaLinha - 1);
      
      // Exibir resultados
      exibirResultadoImportacao(relatorioExecucao);
      
      Logger.log("🎉 === IMPORTAÇÃO CONCLUÍDA COM SUCESSO ===");
      
    } catch (erro) {
      Logger.log("💥 ERRO CRÍTICO NA IMPORTAÇÃO: " + erro.toString());
      SpreadsheetApp.getUi().alert(
        "❌ Erro Crítico",
        `Ocorreu um erro crítico durante a importação:\n\n${erro.toString()}\n\nVerifique os logs para mais detalhes.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  // ========================================================================
  // FUNÇÕES DE PROCESSAMENTO DE DADOS
  // ========================================================================
  
  /**
   * Processa uma planilha individual de secretaria
   * @param {string} planilhaId - ID da planilha no Google Drive
   * @param {Set} registrosExistentes - Set para controle de duplicatas
   * @returns {Object} Resultado do processamento
   */
  function processarPlanilhaSecretaria(planilhaId, registrosExistentes) {
    let tentativa = 1;
    
    while (tentativa <= CONFIG.MAX_TENTATIVAS) {
      try {
        // Abrir planilha externa
        const planilhaExterna = SpreadsheetApp.openById(planilhaId);
        const nomeCompleto = planilhaExterna.getName();
        
        // Extrair sigla da secretaria do nome da planilha
        const siglaSecretaria = extrairSiglaSecretaria(nomeCompleto);
        
        // Obter aba de dados
        const abaOrigem = planilhaExterna.getSheetByName(CONFIG.ABA_ORIGEM);
        if (!abaOrigem) {
          Logger.log(`⚠️ Aba "${CONFIG.ABA_ORIGEM}" não encontrada em: ${nomeCompleto}`);
          return { sucesso: false, erro: "Aba não encontrada" };
        }
        
        // Obter dados (ignorar linhas 1-4, colunas B-O exceto M)
        const ultimaLinha = abaOrigem.getLastRow();
        const ultimaColuna = abaOrigem.getLastColumn();
        
        if (ultimaLinha <= CONFIG.LINHA_INICIO_DADOS) {
          Logger.log(`ℹ️ Nenhum dado encontrado em: ${nomeCompleto}`);
          return { sucesso: true, dados: [], siglaSecretaria, duplicatasEvitadas: 0 };
        }
        
        // Ler dados das colunas B a O (índices 2 a 15)
        const dadosRange = abaOrigem.getRange(
          CONFIG.LINHA_INICIO_DADOS + 1, 2, // Linha 5, coluna B
          ultimaLinha - CONFIG.LINHA_INICIO_DADOS, 14 // Até coluna O
        );
        
        const dadosBrutos = dadosRange.getValues();
        const dadosProcessados = [];
        let duplicatasEvitadas = 0;
        
        // Processar cada linha de dados
        dadosBrutos.forEach((linha, indice) => {
          // Verificar se linha não está vazia
          if (linha.some(valor => valor !== "" && valor !== null && valor !== undefined)) {
            
            // Extrair dados relevantes (B-L, N-O, pular M)
            const nome = (linha[0] || "").toString().trim(); // Coluna B
            const prontuario = (linha[1] || "").toString().trim(); // Coluna C
            
            // Controle de duplicatas por nome + prontuário
            const chaveUnica = `${nome}|${prontuario}`;
            if (registrosExistentes.has(chaveUnica) && nome && prontuario) {
              duplicatasEvitadas++;
              Logger.log(`🔍 Duplicata evitada: ${nome} (${prontuario})`);
              return; // Pular este registro
            }
            
            // Marcar como processado
            if (nome && prontuario) {
              registrosExistentes.add(chaveUnica);
            }
            
            // Montar linha para planilha central
            const linhaCentral = [
              siglaSecretaria,              // A - Secretaria
              nome,                         // B - Nome
              prontuario,                   // C - Prontuário
              linha[2] || "",              // D - Formação Acadêmica
              linha[3] || "",              // E - Área de Formação
              linha[4] || "",              // F - Cargo Concurso
              linha[5] || "",              // G - CC / FE
              linha[6] || "",              // H - Função Gratificada
              linha[7] || "",              // I - Readaptado
              linha[8] || "",              // J - Justificativa
              linha[9] || "",              // K - Ação (o que)
              linha[10] || "",             // L - Condicionalidade
              linha[13] || ""              // M - Data da Liberação (era coluna O)
            ];
            
            dadosProcessados.push(linhaCentral);
          }
        });
        
        return {
          sucesso: true,
          dados: dadosProcessados,
          siglaSecretaria,
          duplicatasEvitadas
        };
        
      } catch (erro) {
        Logger.log(`⚠️ Tentativa ${tentativa}/${CONFIG.MAX_TENTATIVAS} falhou para ${planilhaId}: ${erro.toString()}`);
        tentativa++;
        
        if (tentativa <= CONFIG.MAX_TENTATIVAS) {
          Utilities.sleep(2000); // Aguardar 2 segundos antes de tentar novamente
        }
      }
    }
    
    return { 
      sucesso: false, 
      erro: `Falha após ${CONFIG.MAX_TENTATIVAS} tentativas` 
    };
  }
  
  /**
   * Extrai a sigla da secretaria do nome da planilha
   * @param {string} nomeCompleto - Nome completo da planilha
   * @returns {string} Sigla da secretaria
   */
  function extrairSiglaSecretaria(nomeCompleto) {
    // Padrão: "PGE - Descrição - SIGLA"
    const partes = nomeCompleto.split(" - ");
    
    if (partes.length >= 3) {
      // Última parte é a sigla
      return partes[partes.length - 1].trim().toUpperCase();
    } else if (partes.length === 2) {
      // Se só tem 2 partes, a segunda é a sigla
      return partes[1].trim().toUpperCase();
    } else {
      // Fallback: usar nome inteiro limitado
      return nomeCompleto.substring(0, 20).toUpperCase();
    }
  }
  
  // ========================================================================
  // FUNÇÕES DE CONFIGURAÇÃO DA PLANILHA CENTRAL
  // ========================================================================
  
  /**
   * Obtém ou cria a aba central
   * @param {Spreadsheet} planilhaCentral - Objeto da planilha central
   * @returns {Sheet} Aba central
   */
  function obterOuCriarAbaCentral(planilhaCentral) {
    let abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
    
    if (!abaCentral) {
      Logger.log(`📋 Criando aba "${CONFIG.ABA_CENTRAL}"`);
      abaCentral = planilhaCentral.insertSheet(CONFIG.ABA_CENTRAL);
    }
    
    return abaCentral;
  }
  
  /**
   * Configura a aba central com cabeçalhos
   * @param {Sheet} abaCentral - Aba central
   */
  function configurarAbaCentral(abaCentral) {
    // Limpar conteúdo existente
    abaCentral.clear();
    
    // Configurar cabeçalhos
    const rangeCabecalho = abaCentral.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
    rangeCabecalho.setValues([CABECALHOS_CENTRAL]);
    
    // Formatação do cabeçalho
    rangeCabecalho
      .setBackground("#1f4e79")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true);
    
    // Congelar primeira linha
    abaCentral.setFrozenRows(1);
    
    Logger.log("✅ Aba central configurada com sucesso");
  }
  
  /**
   * Aplica formatação à planilha central
   * @param {Sheet} abaCentral - Aba central
   * @param {number} totalLinhas - Total de linhas com dados
   */
  function aplicarFormatacao(abaCentral, totalLinhas) {
    if (totalLinhas <= 1) return;
    
    try {
      // Range de dados (sem cabeçalho)
      const rangeDados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length);
      
      // Formatação geral
      rangeDados
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      
      // Formatação alternada de linhas
      const regraFormatacao = SpreadsheetApp.newConditionalFormatRule()
        .setRanges([rangeDados])
        .whenFormulaSatisfied('=MOD(ROW(),2)=0')
        .setBackground("#f8f9fa")
        .build();
      
      abaCentral.setConditionalFormatRules([regraFormatacao]);
      
      // Ajustar largura das colunas
      abaCentral.autoResizeColumns(1, CABECALHOS_CENTRAL.length);
      
      // Destacar coluna de secretaria
      const colunaSecretaria = abaCentral.getRange(2, 1, totalLinhas - 1, 1);
      colunaSecretaria
        .setBackground("#e8f4fd")
        .setFontWeight("bold");
      
      Logger.log("✅ Formatação aplicada com sucesso");
      
    } catch (erro) {
      Logger.log("⚠️ Erro na formatação: " + erro.toString());
    }
  }
  
  // ========================================================================
  // FUNÇÕES DE INTERFACE E RELATÓRIOS
  // ========================================================================
  
  /**
   * Exibe resultados da importação
   * @param {Object} relatorio - Dados do relatório de execução
   */
  function exibirResultadoImportacao(relatorio) {
    const mensagem = `
  🎉 IMPORTAÇÃO CONCLUÍDA!
  
  📊 RESUMO DA EXECUÇÃO:
  • Secretarias processadas: ${relatorio.secretariasProcessadas}/23
  • Registros importados: ${relatorio.registrosImportados}
  • Duplicatas evitadas: ${relatorio.duplicatasEvitadas}
  • Duração: ${relatorio.duracao}s
  • Erros: ${relatorio.erros.length}
  
  ${relatorio.erros.length > 0 ? '⚠️ ATENÇÃO: Alguns erros ocorreram. Verifique os logs para detalhes.' : '✅ Processo executado sem erros!'}
  
  📅 Última atualização: ${relatorio.fim.toLocaleString('pt-BR')}
    `;
    
    SpreadsheetApp.getUi().alert(
      "🏛️ Sistema Banco de Talentos", 
      mensagem,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  /**
   * Exibe informações sobre o sistema
   */
  function exibirSobre() {
    const sobre = `
  🏛️ SISTEMA BANCO DE TALENTOS
  Programa Governo Eficaz - Santana de Parnaíba
  
  ✨ MISSÃO DO PROGRAMA:
  "O Programa Governo Eficaz não é apenas uma ferramenta de gestão - é um investimento estratégico no futuro do nosso governo e no bem-estar e desenvolvimento dos nossos servidores."
  
  🎯 FILOSOFIA:
  "Um governo eficaz começa com pessoas certas nos lugares certos"
  
  🔧 FUNCIONALIDADES DO SISTEMA:
  • Consolidação automática de dados das 23 secretarias
  • Prevenção inteligente de duplicatas
  • Interface intuitiva e relatórios detalhados
  • Atualizações em tempo real
  
  📞 SUPORTE E CONTATO:
  Dúvidas, suporte técnico ou orientações:
  
  📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br
  📱 4622-7500 - 8819 / 8644 / 7574
  
  👥 EQUIPE DE DESENVOLVIMENTO:
  Sistema desenvolvido para otimizar a gestão de talentos municipais, promovendo eficiência e transparência na administração pública.
  
  🚀 VERSÃO: 2.0 - Otimizada para Google Apps Script
  📅 Última atualização: ${new Date().toLocaleDateString('pt-BR')}
    `;
    
    SpreadsheetApp.getUi().alert(
      "ℹ️ Sobre o Sistema", 
      sobre,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  /**
   * Gera relatório detalhado de execução
   */
  function gerarRelatorioExecucao() {
    try {
      const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
      const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
      
      if (!abaCentral) {
        SpreadsheetApp.getUi().alert(
          "⚠️ Relatório não disponível",
          "Execute primeiro a importação dos dados.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const totalLinhas = abaCentral.getLastRow();
      const totalRegistros = Math.max(0, totalLinhas - 1); // Excluir cabeçalho
      
      // Contar registros por secretaria
      const estatisticasPorSecretaria = {};
      
      if (totalRegistros > 0) {
        const dadosSecretaria = abaCentral.getRange(2, 1, totalRegistros, 1).getValues();
        
        dadosSecretaria.forEach(linha => {
          const secretaria = linha[0] || "NÃO IDENTIFICADA";
          estatisticasPorSecretaria[secretaria] = (estatisticasPorSecretaria[secretaria] || 0) + 1;
        });
      }
      
      // Montar relatório
      let relatorioDetalhado = `
  📊 RELATÓRIO DETALHADO DE EXECUÇÃO
  
  📈 ESTATÍSTICAS GERAIS:
  • Total de registros: ${totalRegistros}
  • Total de secretarias: ${Object.keys(estatisticasPorSecretaria).length}
  • Planilhas configuradas: ${PLANILHAS_SECRETARIAS.length}
  
  🏢 DISTRIBUIÇÃO POR SECRETARIA:
  `;
      
      // Ordenar secretarias por quantidade de registros
      const secretariasOrdenadas = Object.entries(estatisticasPorSecretaria)
        .sort((a, b) => b[1] - a[1]);
      
      secretariasOrdenadas.forEach(([secretaria, quantidade]) => {
        relatorioDetalhado += `• ${secretaria}: ${quantidade} registros\n`;
      });
      
      relatorioDetalhado += `
  📅 Data do relatório: ${new Date().toLocaleString('pt-BR')}
  🔄 Para atualizar os dados, use o menu "Banco de Talentos > Atualizar Dados"
      `;
      
      SpreadsheetApp.getUi().alert(
        "📊 Relatório de Execução",
        relatorioDetalhado,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (erro) {
      Logger.log("❌ Erro ao gerar relatório: " + erro.toString());
      SpreadsheetApp.getUi().alert(
        "❌ Erro no Relatório",
        "Ocorreu um erro ao gerar o relatório. Verifique os logs.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  /**
   * Limpa todos os dados da planilha central (mantém cabeçalhos)
   */
  function limparDadosCentral() {
    const resposta = SpreadsheetApp.getUi().alert(
      "🧹 Confirmar Limpeza",
      "Tem certeza que deseja limpar todos os dados da planilha central?\n\nEsta ação não pode ser desfeita.",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (resposta === SpreadsheetApp.getUi().Button.YES) {
      try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (abaCentral) {
          // Limpar apenas dados (manter cabeçalhos)
          const ultimaLinha = abaCentral.getLastRow();
          if (ultimaLinha > 1) {
            abaCentral.getRange(2, 1, ultimaLinha - 1, abaCentral.getLastColumn()).clearContent();
          }
          
          SpreadsheetApp.getUi().alert(
            "✅ Limpeza Concluída",
            "Todos os dados foram removidos. Os cabeçalhos foram mantidos.",
            SpreadsheetApp.getUi().ButtonSet.OK
          );
          
          Logger.log("🧹 Dados da planilha central limpos com sucesso");
        }
        
      } catch (erro) {
        Logger.log("❌ Erro na limpeza: " + erro.toString());
        SpreadsheetApp.getUi().alert(
          "❌ Erro na Limpeza",
          "Ocorreu um erro durante a limpeza dos dados.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
  }
  
  // ========================================================================
  // FUNÇÕES UTILITÁRIAS
  // ========================================================================
  
  /**
   * Função de teste para verificar conectividade
   */
  function testarConectividade() {
    Logger.log("🧪 === TESTE DE CONECTIVIDADE ===");
    
    let sucessos = 0;
    let erros = 0;
    
    PLANILHAS_SECRETARIAS.slice(0, 5).forEach((id, indice) => {
      try {
        const planilha = SpreadsheetApp.openById(id);
        const nome = planilha.getName();
        Logger.log(`✅ Planilha ${indice + 1}: ${nome}`);
        sucessos++;
      } catch (erro) {
        Logger.log(`❌ Planilha ${indice + 1} (${id}): ${erro.toString()}`);
        erros++;
      }
    });
    
    Logger.log(`📊 Resultado: ${sucessos} sucessos, ${erros} erros`);
    
    SpreadsheetApp.getUi().alert(
      "🧪 Teste de Conectividade",
      `Teste realizado em 5 planilhas:\n\n✅ Sucessos: ${sucessos}\n❌ Erros: ${erros}\n\nVerifique os logs para detalhes.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  /**
   * Função para backup dos dados atuais
   */
  function criarBackup() {
    try {
      const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
      const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
      
      if (!abaCentral || abaCentral.getLastRow() <= 1) {
        SpreadsheetApp.getUi().alert(
          "⚠️ Backup não necessário",
          "Não há dados para fazer backup.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const dataHora = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd_HH-mm");
      const nomeBackup = `Backup_BT_${dataHora}`;
      
      // Criar cópia da aba
      const abaBackup = abaCentral.copyTo(planilhaCentral);
      abaBackup.setName(nomeBackup);
      
      SpreadsheetApp.getUi().alert(
        "💾 Backup Criado",
        `Backup criado com sucesso!\n\nNome da aba: ${nomeBackup}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      Logger.log(`💾 Backup criado: ${nomeBackup}`);
      
    } catch (erro) {
      Logger.log("❌ Erro no backup: " + erro.toString());
      SpreadsheetApp.getUi().alert(
        "❌ Erro no Backup",
        "Não foi possível criar o backup dos dados.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  // ========================================================================
  // LOGS E DEBUGGING
  // ========================================================================
  
  /**
   * Limpa logs antigos (executar manualmente se necessário)
   */
  function limparLogs() {
    console.clear();
    Logger.log("🧹 Logs limpos - " + new Date().toLocaleString('pt-BR'));
  }
  
  /**
   * Exibe informações do sistema
   */
  function informacoesSistema() {
    Logger.log("🔧 === INFORMAÇÕES DO SISTEMA ===");
    Logger.log(`📊 Planilhas configuradas: ${PLANILHAS_SECRETARIAS.length}`);
    Logger.log(`📋 Aba central: ${CONFIG.ABA_CENTRAL}`);
    Logger.log(`📂 Aba origem: ${CONFIG.ABA_ORIGEM}`);
    Logger.log(`🚀 Versão: 2.0`);
    Logger.log(`📅 Data: ${new Date().toLocaleString('pt-BR')}`);
  }
  
  // ========================================================================
  // FIM DO CÓDIGO
  // ========================================================================
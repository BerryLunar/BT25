/**
 * ========================================================================
 * SISTEMA BANCO DE TALENTOS - VERSÃO CORRIGIDA
 * Santana de Parnaíba - SP
 * ========================================================================
 * 
 * CORREÇÕES IMPLEMENTADAS:
 * 1. Extração correta da sigla das secretarias (usar array ao invés do nome)
 * 2. Mapeamento correto das colunas conforme especificação
 * 3. Ordenação alfabética consistente das secretarias
 * 4. Dados começam na linha 5 (LINHA_INICIO_DADOS = 4)
 * 
 * ========================================================================
 */

// ========================================================================
// CONFIGURAÇÕES GLOBAIS CORRIGIDAS
// ========================================================================

/**
 * IDs das 23 planilhas das secretarias municipais - CORRIGIDO
 * Agora usando o nome definido no array como sigla oficial
 */
const PLANILHAS_SECRETARIAS = [
    { id: "1Z6rfDo09m2nUQIjjAS0S3iuweToHAtIXK9idRb-RzuA", nome: "SECOM" },
    { id: "14W2ewZBf-MgoKUYewj7--FLjDC1lVL6-5EUpHTHDoqU", nome: "SEMEDES" },
    { id: "15ztEc0wlK1mkrRd-DYkfXUn-Vq-k2JT0yLV-JaMeYV4", nome: "SEMOP" },
    { id: "1u7L6Qh57zFqQRCPNTJLGGReuEBlykCnD65SwwNFeRNA", nome: "SEMUTRANS" },
    { id: "1Nc9O1Ha038gKY5LcfxUVClhTq6rsR0zghdSJqfScI6k", nome: "SMA" },
    { id: "1QrRoRoOsyrKgFIitAYWl1g53zYyxymWnoQCVBsRyHsM", nome: "SMAFEL" },
    { id: "1rlyWJDE3srgUMyJy2eDlNrA4Sp57Njca9oypUkRSpP8", nome: "SMCC" },
    { id: "1_4d9POGUbjKHHGPCcbQGx3-wLQuSmThKlLsFG98vDpw", nome: "SMCL" },
    { id: "1n5JfTcpy8EtSlBY-bT3JLqV5kECrN_QeSM79TXbEZjk", nome: "SMCT" },
    { id: "1giGaLo8jtOJ2VCSFcwhULFR5aYZcgPdifSfIKHXAm1A", nome: "SMDS" },
    { id: "1WBOuLGg7hwFY7ehP1qKmuzN8-ZJR12GhMdYcNDCYHz0", nome: "SME" },
    { id: "1cWASt4ldQbIEFm0XW1xPd7zJJZ9kNjGCylgc5XQIbPg", nome: "SMF" },
    { id: "1n2UuXYvKzz1Dau32aFwVhHLWzt2brClbJ_CFcpVa3Ks", nome: "SMGAED" },
    { id: "16SHJbhAb7XVmEhREFX_cVFeR2xtczizXjvpaUAwpFAk", nome: "SMH" },
    { id: "1xIWtAH9P7HZzjzroFG01KhtvRmy-sKV1YxCCBmOBxJo", nome: "SMMAP" },
    { id: "1qWSF0f7wVmJPM7Ht2_dtBPDBztt8PyyFwdL9rwwGTBM", nome: "SMMF" },
    { id: "1eSp1C9K-AO2ZJr3ApFuDDlhGhsJXcvbxZZUn04VildM", nome: "SMNJ" },
    { id: "1AhWPtTgLqF_VEBNM6HxcxxJl5bkaRG5MLHPIX3Pa-3Y", nome: "SMOP" },
    { id: "1EObh9xVjRqrPY1Fdz_ji7Nzq-2lcHXZKJysF6SJf66w", nome: "SMOU" },
    { id: "17dEVkFJNNGanitiJYnT7Lwqz3DWsz8qODZA0K7Quggs", nome: "SMS" },
    { id: "1klqCbpMJVyCXTdpBkNrsrVfZozEOFWi1VT0xZjjSf8g", nome: "SMSD" },
    { id: "1SjSlad1XQTPA0PqPUrpB7WMfPO0WxWV1f7SpHvhubrQ", nome: "SMSM" },
    { id: "1LPaScCjVYVK6OVA5ZTeUXj3PXiM3nq_1GUdyiTru0jk", nome: "SMSU" }
];

/**
* Configurações do sistema - CORRIGIDAS
*/
const CONFIG = {
    ABA_CENTRAL: "BT 2025",
    ABA_ORIGEM: "Banco de Talentos (Externo)",
    LINHA_INICIO_DADOS: 4, // Dados começam na linha 5 (índice 4)
    LOTE_SIZE: 5,
    TIMEOUT_POR_LOTE: 30000,
    DELAY_ENTRE_LOTES: 2000,
    MAX_TENTATIVAS: 2
};

/**
* Cabeçalhos padronizados - CONFIRMADOS
*/
const CABECALHOS_CENTRAL = [
    "Secretaria",           // A
    "Nome",                 // B  
    "Prontuário",          // C
    "Formação Acadêmica",   // D
    "Área de Formação",     // E
    "Cargo Concurso",       // F
    "CC / FE",             // G
    "Função Gratificada",   // H
    "Readaptado",          // I
    "Justificativa",       // J
    "Ação (o que)",        // K
    "Condicionalidade",    // L
    "Data da Inclusão"     // M
];

// ========================================================================
// FUNÇÃO PRINCIPAL CORRIGIDA - ORDENAÇÃO ALFABÉTICA
// ========================================================================

/**
* Função principal otimizada para importação - CORRIGIDA
*/
function importarBancoDeTalentosOtimizado() {
    const inicioExecucao = new Date();
    let relatorio = {
        inicio: inicioExecucao,
        secretariasProcessadas: 0,
        registrosImportados: 0,
        erros: [],
        lotes: []
    };
    
    try {
        Logger.log("🚀 === INICIANDO IMPORTAÇÃO CORRIGIDA ===");
        
        // Preparar planilha central
        const { planilhaCentral, abaCentral } = prepararPlanilhaCentral();
        
        // ORDENAR SECRETARIAS ALFABETICAMENTE ANTES DO PROCESSAMENTO
        const secretariasOrdenadas = [...PLANILHAS_SECRETARIAS].sort((a, b) => 
            a.nome.localeCompare(b.nome)
        );
        
        Logger.log("🔤 Secretarias ordenadas alfabeticamente:");
        secretariasOrdenadas.forEach((s, i) => {
            Logger.log(`  ${i + 1}. ${s.nome}`);
        });
        
        // Coletar todos os dados primeiro
        const todosOsDados = [];
        
        // Processar em lotes (usando secretarias ordenadas)
        const lotes = criarLotes(secretariasOrdenadas, CONFIG.LOTE_SIZE);
        
        for (let i = 0; i < lotes.length; i++) {
            const lote = lotes[i];
            const numeroLote = i + 1;
            const totalLotes = lotes.length;
            
            // Mostrar progresso
            mostrarProgresso(numeroLote, totalLotes, lote.length);
            
            const resultadoLote = processarLoteSecretarias(lote, numeroLote);
            
            // Adicionar dados do lote aos dados totais (já em ordem alfabética)
            todosOsDados.push(...resultadoLote.dados);
            
            // Atualizar relatório
            relatorio.secretariasProcessadas += resultadoLote.processadas;
            relatorio.erros.push(...resultadoLote.erros);
            relatorio.lotes.push(resultadoLote);
            
            // Pausa entre lotes
            if (i < lotes.length - 1) {
                Utilities.sleep(CONFIG.DELAY_ENTRE_LOTES);
            }
        }
        
        // Inserir todos os dados (já ordenados por secretaria)
        if (todosOsDados.length > 0) {
            Logger.log(`📝 Inserindo ${todosOsDados.length} registros ordenados...`);
            
            const range = abaCentral.getRange(2, 1, todosOsDados.length, CABECALHOS_CENTRAL.length);
            range.setValues(todosOsDados);
            
            relatorio.registrosImportados = todosOsDados.length;
            
            // Aplicar formatação
            aplicarFormatacaoOtimizada(abaCentral, todosOsDados.length + 1);
        }
        
        // Finalizar
        relatorio.fim = new Date();
        relatorio.duracao = Math.round((relatorio.fim - relatorio.inicio) / 1000);
        
        exibirResultadoOtimizado(relatorio);
        
        Logger.log("🎉 === IMPORTAÇÃO CORRIGIDA CONCLUÍDA ===");
        
    } catch (erro) {
        Logger.log("💥 ERRO CRÍTICO: " + erro.toString());
        
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Importação",
            `Erro crítico durante a importação:\n\n${erro.toString()}\n\n📋 Alguns dados podem ter sido preservados.\nVerifique a planilha e os logs.`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

// ========================================================================
// PROCESSAMENTO INDIVIDUAL DE SECRETARIA - CORRIGIDO
// ========================================================================

/**
* Processa uma secretaria individual - VERSÃO CORRIGIDA
*/
function processarSecretariaOtimizada(secretaria) {
    try {
        // USAR A SIGLA DIRETAMENTE DO ARRAY - SEM EXTRAÇÃO DO NOME DA PLANILHA
        const siglaSecretaria = secretaria.nome; // Usar diretamente a sigla do array
        
        Logger.log(`📂 Processando: ${siglaSecretaria} (ID: ${secretaria.id.substring(0, 10)}...)`);
        
        // Abrir planilha
        const planilhaExterna = SpreadsheetApp.openById(secretaria.id);
        const nomeCompletoPlanilha = planilhaExterna.getName();
        
        Logger.log(`📋 Nome da planilha: ${nomeCompletoPlanilha}`);
        Logger.log(`🏷️ Sigla usada: ${siglaSecretaria}`);
        
        // Verificar aba
        const abaOrigem = planilhaExterna.getSheetByName(CONFIG.ABA_ORIGEM);
        if (!abaOrigem) {
            return { 
                sucesso: false, 
                erro: `Aba "${CONFIG.ABA_ORIGEM}" não encontrada`,
                siglaSecretaria 
            };
        }
        
        // Obter dados de forma otimizada
        const ultimaLinha = abaOrigem.getLastRow();
        
        if (ultimaLinha <= CONFIG.LINHA_INICIO_DADOS) {
            Logger.log(`ℹ️ Sem dados: ${siglaSecretaria}`);
            return { 
                sucesso: true, 
                dados: [], 
                siglaSecretaria 
            };
        }
        
        // ============================================================================
        // MAPEAMENTO CORRETO DAS COLUNAS - CONFIRMADO:
        // DADOS COMEÇAM NA LINHA 5 (CONFIG.LINHA_INICIO_DADOS = 4)
        // B = Nome | C = Prontuário | D = Formação | E = Área | F = Cargo 
        // G = CC/FE | H = Função | I = Readaptado | J = Justificativa 
        // K = Ação | L = Condicionalidade | N = Data da Inclusão
        // ============================================================================
        
        // Calcular linhas de dados disponíveis
        const totalLinhas = ultimaLinha - CONFIG.LINHA_INICIO_DADOS;
        
        Logger.log(`📊 ${siglaSecretaria}: Linha ${CONFIG.LINHA_INICIO_DADOS + 1} até ${ultimaLinha} (${totalLinhas} linhas)`);
        
        // Ler dados das colunas B até N 
        // B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14
        const dadosRange = abaOrigem.getRange(
            CONFIG.LINHA_INICIO_DADOS + 1, // Linha 5 (índice 4 + 1)
            2, // Coluna B (Nome) = índice 2
            totalLinhas, 
            13 // Colunas B até N (B=2 até N=14 = 13 colunas)
        );
        
        const dadosBrutos = dadosRange.getValues();
        
        Logger.log(`📖 ${siglaSecretaria}: Lidas ${dadosBrutos.length} linhas de dados brutos`);
        
        // Processar dados com mapeamento correto
        const dadosProcessados = [];
        
        dadosBrutos.forEach((linha, indiceLinhaArray) => {
            const linhaReal = CONFIG.LINHA_INICIO_DADOS + 1 + indiceLinhaArray;
            
            // Verificar se linha tem dados (verificar pelo menos nome)
            const nome = (linha[0] || "").toString().trim(); // linha[0] = Nome (coluna B)
            
            if (nome) { // Se tem nome, processar a linha
                
                // Mapear corretamente conforme especificação:
                const linhaCentral = [
                    siglaSecretaria,                           // A - Secretaria (usar sigla do array)
                    nome,                                      // B - Nome (B na origem, índice 0)
                    (linha[1] || "").toString().trim(),        // C - Prontuário (C na origem, índice 1)  
                    (linha[2] || "").toString().trim(),        // D - Formação Acadêmica (D na origem, índice 2)
                    (linha[3] || "").toString().trim(),        // E - Área de Formação (E na origem, índice 3)
                    (linha[4] || "").toString().trim(),        // F - Cargo Concurso (F na origem, índice 4)
                    (linha[5] || "").toString().trim(),        // G - CC / FE (G na origem, índice 5)
                    (linha[6] || "").toString().trim(),        // H - Função Gratificada (H na origem, índice 6)
                    (linha[7] || "").toString().trim(),        // I - Readaptado (I na origem, índice 7)
                    (linha[8] || "").toString().trim(),        // J - Justificativa (J na origem, índice 8)
                    (linha[9] || "").toString().trim(),        // K - Ação (o que) (K na origem, índice 9)
                    (linha[10] || "").toString().trim(),       // L - Condicionalidade (L na origem, índice 10)
                    formatarDataBrasileira(linha[12] || "")     // M - Data da Inclusão (N na origem, índice 12)
                ];
                
                dadosProcessados.push(linhaCentral);
                
                // Log detalhado para primeira linha de cada secretaria (debug)
                if (dadosProcessados.length === 1) {
                    Logger.log(`🔍 ${siglaSecretaria} - Primeira linha (${linhaReal}): ${nome} | ${linhaCentral[2]} | ${linhaCentral[3]}`);
                }
            }
        });
        
        Logger.log(`✅ ${siglaSecretaria}: ${dadosProcessados.length} registros processados`);
        
        return {
            sucesso: true,
            dados: dadosProcessados,
            siglaSecretaria: siglaSecretaria
        };
        
    } catch (erro) {
        Logger.log(`❌ Erro em ${secretaria.nome}: ${erro.toString()}`);
        return {
            sucesso: false,
            erro: erro.toString(),
            siglaSecretaria: secretaria.nome
        };
    }
}

// ========================================================================
// FUNÇÃO AUXILIAR PARA BUSCAR SIGLA POR ID (CASO NECESSÁRIO)
// ========================================================================

/**
* Busca sigla da secretaria pelo ID da planilha
*/
function buscarSiglaPorId(idPlanilha) {
    const secretaria = PLANILHAS_SECRETARIAS.find(s => s.id === idPlanilha);
    return secretaria ? secretaria.nome : "DESCONHECIDA";
}

/**
* Lista todas as secretarias em ordem alfabética (para debug)
*/
function listarSecretariasOrdenadas() {
    const secretariasOrdenadas = [...PLANILHAS_SECRETARIAS].sort((a, b) => 
        a.nome.localeCompare(b.nome)
    );
    
    Logger.log("🔤 === SECRETARIAS EM ORDEM ALFABÉTICA ===");
    secretariasOrdenadas.forEach((secretaria, indice) => {
        Logger.log(`${String(indice + 1).padStart(2, '0')}. ${secretaria.nome}`);
    });
    
    return secretariasOrdenadas;
}

// ========================================================================
// TESTE E VALIDAÇÃO DAS CORREÇÕES
// ========================================================================

/**
* Teste ESPECÍFICO do mapeamento com UMA secretaria
*/
function testeMapemantoColunas() {
    Logger.log("🧪 === TESTE DO MAPEAMENTO DE COLUNAS ===");
    
    // Testar com a primeira secretaria da lista (SECOM)
    const secretariaTeste = PLANILHAS_SECRETARIAS[0]; 
    
    try {
        Logger.log(`📋 Testando: ${secretariaTeste.nome}`);
        
        const planilha = SpreadsheetApp.openById(secretariaTeste.id);
        const aba = planilha.getSheetByName(CONFIG.ABA_ORIGEM);
        
        if (!aba) {
            Logger.log(`❌ Aba não encontrada: ${CONFIG.ABA_ORIGEM}`);
            return;
        }
        
        const ultimaLinha = aba.getLastRow();
        Logger.log(`📊 Última linha: ${ultimaLinha}`);
        
        if (ultimaLinha > CONFIG.LINHA_INICIO_DADOS) {
            // Ler cabeçalhos da planilha de origem
            const cabecalhos = aba.getRange(1, 1, 1, 20).getValues()[0];
            Logger.log("📋 Cabeçalhos encontrados:");
            cabecalhos.forEach((cab, indice) => {
                const coluna = String.fromCharCode(65 + indice); // A=65
                if (cab && cab.toString().trim()) {
                    Logger.log(`  ${coluna}${indice + 1}: ${cab}`);
                }
            });
            
            // Ler primeira linha de dados
            const primeiraLinhaDados = aba.getRange(CONFIG.LINHA_INICIO_DADOS + 1, 1, 1, 20).getValues()[0];
            Logger.log("📋 Primeira linha de dados:");
            primeiraLinhaDados.forEach((valor, indice) => {
                const coluna = String.fromCharCode(65 + indice);
                if (valor && valor.toString().trim()) {
                    Logger.log(`  ${coluna}: ${valor}`);
                }
            });
            
            // Testar processamento completo
            Logger.log("\n🔍 === TESTE DE PROCESSAMENTO ===");
            const resultado = processarSecretariaOtimizada(secretariaTeste);
            
            if (resultado.sucesso && resultado.dados.length > 0) {
                const primeiroRegistro = resultado.dados[0];
                Logger.log("✅ Primeiro registro processado:");
                CABECALHOS_CENTRAL.forEach((cabecalho, indice) => {
                    Logger.log(`  ${cabecalho}: ${primeiroRegistro[indice]}`);
                });
            }
        }
        
        const relatorio = `
🧪 TESTE DE MAPEAMENTO CONCLUÍDO

📋 Secretaria: ${secretariaTeste.nome}
📊 Última linha: ${ultimaLinha}

🔧 AJUSTE APLICADO:
• Leitura dos dados começa na COLUNA B (Nome)
• Ignora coluna A (pode ter numeração ou estar vazia)

✅ Verifique os logs do Apps Script para ver:
• Cabeçalhos encontrados (B até T)
• Primeira linha de dados (B até T)
• Mapeamento aplicado
• Primeiro registro processado

📝 PRÓXIMO PASSO:
Se os dados estiverem corretos nos logs, execute "Importar Dados"
        `;
        
        SpreadsheetApp.getUi().alert(
            "🧪 Teste de Mapeamento",
            relatorio,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro no teste: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro no Teste",
            `Erro no teste de mapeamento:\n${erro.toString()}`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Validar estrutura de todas as secretarias
*/
function validarEstruturaSecretarias() {
    Logger.log("🔍 === VALIDAÇÃO DE ESTRUTURA DAS SECRETARIAS ===");
    
    let sucessos = 0;
    let erros = [];
    
    PLANILHAS_SECRETARIAS.forEach((secretaria, indice) => {
        try {
            const planilha = SpreadsheetApp.openById(secretaria.id);
            const aba = planilha.getSheetByName(CONFIG.ABA_ORIGEM);
            
            if (aba) {
                const ultimaLinha = aba.getLastRow();
                Logger.log(`✅ ${String(indice + 1).padStart(2, '0')}. ${secretaria.nome}: ${ultimaLinha} linhas`);
                sucessos++;
            } else {
                const erro = `Aba "${CONFIG.ABA_ORIGEM}" não encontrada`;
                Logger.log(`❌ ${String(indice + 1).padStart(2, '0')}. ${secretaria.nome}: ${erro}`);
                erros.push(`${secretaria.nome}: ${erro}`);
            }
            
        } catch (erro) {
            Logger.log(`💥 ${String(indice + 1).padStart(2, '0')}. ${secretaria.nome}: ${erro.toString()}`);
            erros.push(`${secretaria.nome}: ${erro.toString()}`);
        }
        
        // Pausa para evitar timeout
        Utilities.sleep(100);
    });
    
    const resumo = `
🔍 VALIDAÇÃO CONCLUÍDA:
✅ Sucessos: ${sucessos}/${PLANILHAS_SECRETARIAS.length}
❌ Erros: ${erros.length}
📊 Taxa de sucesso: ${Math.round(sucessos/PLANILHAS_SECRETARIAS.length*100)}%
    `;
    
    Logger.log(resumo);
    
    if (erros.length > 0) {
        Logger.log("❌ Lista de erros:");
        erros.forEach(erro => Logger.log(`  • ${erro}`));
    }
    
    return { sucessos, erros, total: PLANILHAS_SECRETARIAS.length };
}

// ========================================================================
// MENU ATUALIZADO COM FUNÇÕES DE TESTE
// ========================================================================

/**
* Cria menu com opções de teste
*/
function criarMenuPersonalizadoCorrigido() {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu("🏛️ Banco de Talentos v3.1")
        .addItem("🔄 Importar Dados", "iniciarImportacaoManual")
        .addSeparator()
        .addItem("📊 Atualizar Secretaria Específica", "atualizarSecretariaEspecifica")
        .addItem("🔍 Verificar Dados Existentes", "verificarDadosExistentes")
        .addSeparator()
        .addSubMenu(ui.createMenu("🧪 Testes e Debug")
            .addItem("🧪 Testar Mapeamento", "testeMapemantoColunas")
            .addItem("🔍 Validar Estruturas", "validarEstruturaSecretarias")
            .addItem("🔤 Listar Secretarias Ordenadas", "listarSecretariasOrdenadas")
            .addItem("🧹 Limpar Logs", "limparLogs"))
        .addSeparator()
        .addItem("📈 Relatório Completo", "gerarRelatorioCompleto")
        .addItem("🧹 Limpar e Reiniciar", "limparEReiniciar")
        .addSeparator()
        .addItem("ℹ️ Sobre o Sistema", "exibirSobre")
        .addToUi();
}

// ========================================================================
// INICIALIZAÇÃO CORRIGIDA
// ========================================================================

/**
* Função executada ao abrir - VERSÃO CORRIGIDA
*/
function onOpen() {
    try {
        criarMenuPersonalizadoCorrigido();
        Logger.log("✅ Menu personalizado corrigido criado");
        
        // Mostrar informações da versão corrigida
        mostrarInstrucoesCorrigidas();
        
    } catch (erro) {
        Logger.log("❌ Erro na inicialização: " + erro.toString());
    }
}

/**
* Mostra instruções da versão corrigida
*/
function mostrarInstrucoesCorrigidas() {
    const instrucoes = `
🏛️ SISTEMA BANCO DE TALENTOS - VERSÃO 3.1 CORRIGIDA

🔧 CORREÇÕES IMPLEMENTADAS:
• ✅ Sigla das secretarias: agora usa o array PLANILHAS_SECRETARIAS
• ✅ Mapeamento de colunas: corrigido conforme especificação
• ✅ Ordenação alfabética: secretarias sempre ordenadas
• ✅ Início dos dados: confirmado linha 5 (índice 4)

🧪 NOVAS OPÇÕES DE TESTE:
• "Testar Mapeamento" - verifica se as colunas estão corretas
• "Validar Estruturas" - testa todas as secretarias
• "Listar Secretarias Ordenadas" - mostra ordem alfabética

⚡ COMO USAR:
1. Execute "Testar Mapeamento" primeiro para verificar
2. Use "Importar Dados" para processamento completo
3. Dados ficarão ordenados alfabeticamente por secretaria

🎯 VERSÃO 3.1 - PROBLEMAS CORRIGIDOS!
    `;
    
    SpreadsheetApp.getUi().alert(
        "🎉 Sistema Corrigido - v3.1!", 
        instrucoes,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

/**
 * ========================================================================
 * 🎉 RESUMO DAS CORREÇÕES IMPLEMENTADAS:
 * 
 * 1. 🏷️ SIGLA DA SECRETARIA:
 *    ❌ Antes: extraía do nome da planilha (pegava parte errada)
 *    ✅ Agora: usa diretamente secretaria.nome do array
 * 
 * 2. 🔤 ORDENAÇÃO ALFABÉTICA:
 *    ❌ Antes: ordenava os dados depois de coletar
 *    ✅ Agora: ordena as secretarias ANTES do processamento
 * 
 * 3. 📊 MAPEAMENTO DE COLUNAS:
 *    ✅ Confirmado: B→B, C→C, D→D, E→E, F→F, G→G, H→H, I→I, J→J, K→K, L→L, N→M
 *    ✅ Início dos dados: linha 5 (CONFIG.LINHA_INICIO_DADOS = 4)
 *    ✅ Data da Inclusão: coluna N (índice 12) da origem
 * 
 * 4. 🧪 NOVAS FUNÇÕES DE TESTE:
 *    • testeMapemantoColunas() - verifica estrutura
 *    • validarEstruturaSecretarias() - testa conectividade
 *    • listarSecretariasOrdenadas() - mostra ordem alfabética
 * 
 * 5. 🔍 VALIDAÇÃO DE DADOS:
 *    ✅ Verificação baseada no campo Nome (não linha vazia)
 *    ✅ Logs detalhados para debug
 *    ✅ Contagem precisa de registros por secretaria
 * ========================================================================
 */

// ========================================================================
// FUNÇÕES COMPLEMENTARES PARA FINALIZAR O SISTEMA
// ========================================================================

/**
* Função para testar uma secretaria específica
*/
function testarSecretariaEspecifica() {
    const opcoes = PLANILHAS_SECRETARIAS.map((s, i) => `${i + 1} - ${s.nome}`);
    
    const resposta = SpreadsheetApp.getUi().prompt(
        "🧪 Testar Secretaria Específica",
        `Digite o número da secretaria (1-${PLANILHAS_SECRETARIAS.length}):\n\n` + opcoes.join("\n"),
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    
    if (resposta.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        const numero = parseInt(resposta.getResponseText());
        
        if (numero >= 1 && numero <= PLANILHAS_SECRETARIAS.length) {
            const secretaria = PLANILHAS_SECRETARIAS[numero - 1];
            executarTesteDetalhado(secretaria);
        } else {
            SpreadsheetApp.getUi().alert("⚠️ Número inválido", "Digite um número entre 1 e " + PLANILHAS_SECRETARIAS.length);
        }
    }
}

/**
* Executa teste detalhado de uma secretaria
*/
function executarTesteDetalhado(secretaria) {
    Logger.log(`🔍 === TESTE DETALHADO: ${secretaria.nome} ===`);
    
    try {
        // Testar processamento
        const resultado = processarSecretariaOtimizada(secretaria);
        
        let relatorio = `
🧪 TESTE DETALHADO - ${secretaria.nome}

📋 INFORMAÇÕES BÁSICAS:
• ID da Planilha: ${secretaria.id}
• Sigla Oficial: ${secretaria.nome}
• Status: ${resultado.sucesso ? '✅ Sucesso' : '❌ Erro'}

`;

        if (resultado.sucesso) {
            relatorio += `
📊 DADOS PROCESSADOS:
• Registros encontrados: ${resultado.dados.length}
• Sigla usada: ${resultado.siglaSecretaria}

`;

            if (resultado.dados.length > 0) {
                const primeiroRegistro = resultado.dados[0];
                relatorio += `
🔍 PRIMEIRO REGISTRO (EXEMPLO):
• Nome: ${primeiroRegistro[1]}
• Prontuário: ${primeiroRegistro[2]}
• Formação: ${primeiroRegistro[3]}
• Área: ${primeiroRegistro[4]}
• Cargo: ${primeiroRegistro[5]}

`;
            }
        } else {
            relatorio += `
❌ ERRO ENCONTRADO:
${resultado.erro}

`;
        }

        relatorio += `📅 Teste realizado em: ${new Date().toLocaleString('pt-BR')}`;

        SpreadsheetApp.getUi().alert(
            `🧪 Teste - ${secretaria.nome}`,
            relatorio,
            SpreadsheetApp.getUi().ButtonSet.OK
        );

    } catch (erro) {
        Logger.log(`❌ Erro no teste detalhado: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro no Teste",
            `Erro ao testar ${secretaria.nome}:\n${erro.toString()}`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Comparar dados antes e depois (para validação)
*/
function compararDadosAntesDepois() {
    const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
    const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
    
    if (!abaCentral || abaCentral.getLastRow() <= 1) {
        SpreadsheetApp.getUi().alert(
            "ℹ️ Sem Dados para Comparar",
            "Execute uma importação primeiro.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
    }

    const totalLinhas = abaCentral.getLastRow();
    const totalRegistros = totalLinhas - 1;
    
    // Analisar primeira e última secretaria para verificar ordenação
    const primeiraLinha = abaCentral.getRange(2, 1, 1, 13).getValues()[0];
    const ultimaLinha = abaCentral.getRange(totalLinhas, 1, 1, 13).getValues()[0];
    
    // Contar registros por secretaria
    const dadosSecretaria = abaCentral.getRange(2, 1, totalRegistros, 1).getValues();
    const contadorSecretarias = {};
    
    dadosSecretaria.forEach(linha => {
        const secretaria = linha[0] || "NÃO IDENTIFICADA";
        contadorSecretarias[secretaria] = (contadorSecretarias[secretaria] || 0) + 1;
    });

    const secretariasEncontradas = Object.keys(contadorSecretarias).sort();

    let relatorio = `
🔍 ANÁLISE DE DADOS ATUAIS

📊 ESTATÍSTICAS:
• Total de registros: ${totalRegistros}
• Secretarias encontradas: ${secretariasEncontradas.length}

🔤 VERIFICAÇÃO DE ORDENAÇÃO:
• Primeira secretaria: ${primeiraLinha[0]}
• Última secretaria: ${ultimaLinha[0]}
• Ordenação alfabética: ${secretariasEncontradas[0] === primeiraLinha[0] ? '✅ Correta' : '⚠️ Verificar'}

🏢 SECRETARIAS ENCONTRADAS (${secretariasEncontradas.length}):
${secretariasEncontradas.map(s => `• ${s}: ${contadorSecretarias[s]} registros`).join('\n')}

📋 EXEMPLO DO PRIMEIRO REGISTRO:
• Secretaria: ${primeiraLinha[0]}
• Nome: ${primeiraLinha[1]}
• Prontuário: ${primeiraLinha[2]}
• Formação: ${primeiraLinha[3]}

`;

    Logger.log(relatorio);
    
    SpreadsheetApp.getUi().alert(
        "🔍 Análise de Dados",
        relatorio,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

/**
* Função para corrigir dados existentes (se necessário)
*/
function corrigirDadosExistentes() {
    const resposta = SpreadsheetApp.getUi().alert(
        "🔧 Corrigir Dados Existentes",
        "Esta função vai:\n\n• Verificar dados atuais\n• Identificar problemas\n• Aplicar correções necessárias\n\nDeseja continuar?",
        SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (resposta === SpreadsheetApp.getUi().Button.YES) {
        try {
            const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
            const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
            
            if (!abaCentral || abaCentral.getLastRow() <= 1) {
                SpreadsheetApp.getUi().alert("ℹ️ Nenhum dado para corrigir");
                return;
            }

            const totalLinhas = abaCentral.getLastRow();
            const totalRegistros = totalLinhas - 1;
            
            // Ler todos os dados
            const todosDados = abaCentral.getRange(2, 1, totalRegistros, CABECALHOS_CENTRAL.length).getValues();
            
            // Ordenar alfabeticamente por secretaria
            todosDados.sort((a, b) => {
                const secretariaA = (a[0] || "").toString().toUpperCase();
                const secretariaB = (b[0] || "").toString().toUpperCase();
                return secretariaA.localeCompare(secretariaB);
            });
            
            // Reescrever dados ordenados
            abaCentral.getRange(2, 1, totalRegistros, CABECALHOS_CENTRAL.length).setValues(todosDados);
            
            // Aplicar formatação brasileira
            aplicarFormatacaoOtimizada(abaCentral, totalLinhas);
            
            SpreadsheetApp.getUi().alert(
                "✅ Correção Concluída",
                `Dados reordenados alfabeticamente!\n\n📊 ${totalRegistros} registros processados\n🔤 Ordenação por secretaria aplicada`,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
        } catch (erro) {
            Logger.log(`❌ Erro na correção: ${erro.toString()}`);
            SpreadsheetApp.getUi().alert(
                "❌ Erro na Correção",
                "Ocorreu um erro durante a correção dos dados.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }
}

/**
* Menu atualizado com todas as funções
*/
function criarMenuCompletoCorrigido() {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu("🏛️ Banco de Talentos v3.1")
        .addItem("🔄 Importar Dados", "iniciarImportacaoManual")
        .addSeparator()
        .addItem("📊 Atualizar Secretaria Específica", "atualizarSecretariaEspecifica")
        .addItem("🔍 Verificar Dados Existentes", "verificarDadosExistentes")
        .addItem("🔧 Corrigir Dados Existentes", "corrigirDadosExistentes")
        .addItem("🎨 Aplicar Formatação Brasileira", "aplicarFormatacaoBrasileira")
        .addItem("📅 Corrigir Formato de Datas", "corrigirFormatoDatas")
        .addItem("🔲 Aplicar Cores e Bordas", "aplicarCoresEBordas")
        .addItem("📝 Aplicar Ajuste de Texto", "aplicarAjusteTexto")
        .addSeparator()
        .addSubMenu(ui.createMenu("🧪 Testes e Debug")
            .addItem("🧪 Testar Mapeamento", "testeMapemantoColunas")
            .addItem("🔍 Validar Estruturas", "validarEstruturaSecretarias")
            .addItem("🎯 Testar Secretaria Específica", "testarSecretariaEspecifica")
            .addItem("📊 Comparar Dados", "compararDadosAntesDepois")
            .addItem("🔤 Listar Secretarias Ordenadas", "listarSecretariasOrdenadas")
            .addItem("🧹 Limpar Logs", "limparLogs"))
        .addSeparator()
        .addItem("📈 Relatório Completo", "gerarRelatorioCompleto")
        .addItem("🧹 Limpar e Reiniciar", "limparEReiniciar")
        .addSeparator()
        .addItem("ℹ️ Sobre o Sistema", "exibirSobreCorrigido")
        .addToUi();
}

/**
* Informações sobre a versão corrigida
*/
function exibirSobreCorrigido() {
    const sobre = `
🏛️ SISTEMA BANCO DE TALENTOS v3.1 CORRIGIDO
Programa Governo Eficaz - Santana de Parnaíba

🔧 CORREÇÕES IMPLEMENTADAS:
• ✅ Sigla das secretarias: usa array PLANILHAS_SECRETARIAS
• ✅ Mapeamento correto: B→B, C→C, D→D, etc.
• ✅ Ordenação alfabética: antes do processamento
• ✅ Data da Inclusão: coluna N corretamente mapeada
• ✅ Início dos dados: linha 5 confirmada
• ✅ Formatação brasileira: DD/MM/YYYY, Calibri 10, cores alternadas, bordas, texto ajustado

🧪 FERRAMENTAS DE TESTE:
• 🧪 Testar Mapeamento - verifica estrutura de colunas
• 🔍 Validar Estruturas - testa conectividade com secretarias
• 🎯 Testar Secretaria Específica - análise individual
• 📊 Comparar Dados - analisa dados existentes
• 🔧 Corrigir Dados Existentes - reordena se necessário

⚡ PERFORMANCE OTIMIZADA:
• Processamento em lotes de 5 secretarias
• Ordenação prévia das secretarias
• Logs detalhados para debug
• Recuperação automática de erros

🎨 FORMATAÇÃO BRASILEIRA:
• Fonte: Calibri 10 em toda a planilha
• Data: DD/MM/YYYY (formato brasileiro)
• Alinhamento: centralizado vertical e horizontal
• Texto: quebra de linha ativada, altura ajustada automaticamente
• Cores alternadas: #ffffff (linhas pares) e #ebeff1 (linhas ímpares)
• Bordas: todas as bordas, cor #ffffff, estilo SOLID_MEDIUM
• Larguras personalizadas: A(66), B(257), C(68), D(108), E(103), F(168), G(97), H(88), I(76), J(256), K(125), L(170), M(88)

📊 MAPEAMENTO DE COLUNAS CONFIRMADO:
Origem → Destino
B (Nome) → B (Nome)
C (Prontuário) → C (Prontuário) 
D (Formação) → D (Formação)
E (Área) → E (Área)
F (Cargo) → F (Cargo)
G (CC/FE) → G (CC/FE)
H (Função) → H (Função)
I (Readaptado) → I (Readaptado)
J (Justificativa) → J (Justificativa)
K (Ação) → K (Ação)
L (Condicionalidade) → L (Condicionalidade)
N (Data da Inclusão) → M (Data da Inclusão)

🎯 COMO USAR A VERSÃO CORRIGIDA:
1. Execute "Testar Mapeamento" para verificar estrutura
2. Use "Validar Estruturas" para testar conectividade
3. Execute "Importar Dados" para processamento completo
4. Use "Aplicar Formatação Brasileira" para formatação personalizada
5. Use "Comparar Dados" para validar resultados

📞 SUPORTE TÉCNICO:
📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br
📱 4622-7500 - 8819 / 8644 / 7574

🚀 Versão 3.1 - Problemas Corrigidos!
📅 ${new Date().toLocaleDateString('pt-BR')}
    `;
    
    SpreadsheetApp.getUi().alert(
        "ℹ️ Sistema v3.1 - Corrigido",
        sobre,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

// ========================================================================
// FUNÇÕES PRINCIPAIS DO MENU (QUE ESTAVAM FALTANDO)
// ========================================================================

/**
* Inicia importação manual com confirmação
*/
function iniciarImportacaoManual() {
    const resposta = SpreadsheetApp.getUi().alert(
        "🔄 Iniciar Importação",
        `Deseja importar dados de todas as ${PLANILHAS_SECRETARIAS.length} secretarias?\n\n⏱️ Tempo estimado: 3-5 minutos\n📊 Os dados serão ordenados alfabeticamente por secretaria`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (resposta === SpreadsheetApp.getUi().Button.YES) {
        importarBancoDeTalentosOtimizado();
    }
}

/**
* Cria lotes de planilhas para processamento
*/
function criarLotes(planilhas, tamanhoLote) {
    const lotes = [];
    for (let i = 0; i < planilhas.length; i += tamanhoLote) {
        lotes.push(planilhas.slice(i, i + tamanhoLote));
    }
    return lotes;
}

/**
* Processa um lote de secretarias
*/
function processarLoteSecretarias(lote, numeroLote) {
    Logger.log(`📦 Processando lote ${numeroLote} com ${lote.length} planilhas`);
    
    const dadosLote = [];
    const errosLote = [];
    let processadas = 0;
    
    lote.forEach((secretaria, indice) => {
        try {
            const resultado = processarSecretariaOtimizada(secretaria);
            
            if (resultado.sucesso) {
                dadosLote.push(...resultado.dados);
                processadas++;
                Logger.log(`✅ ${resultado.siglaSecretaria}: ${resultado.dados.length} registros`);
            } else {
                errosLote.push(`${secretaria.nome}: ${resultado.erro}`);
                Logger.log(`❌ ${secretaria.nome}: ${resultado.erro}`);
            }
            
        } catch (erro) {
            const mensagemErro = `${secretaria.nome}: ${erro.toString()}`;
            errosLote.push(mensagemErro);
            Logger.log(`💥 ${mensagemErro}`);
        }
        
        // Micro pausa entre planilhas do mesmo lote
        Utilities.sleep(200);
    });
    
    return {
        dados: dadosLote,
        processadas: processadas,
        erros: errosLote,
        numeroLote: numeroLote
    };
}

/**
* Prepara planilha central
*/
function prepararPlanilhaCentral() {
    const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
    let abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
    
    if (!abaCentral) {
        Logger.log(`📋 Criando aba "${CONFIG.ABA_CENTRAL}"`);
        abaCentral = planilhaCentral.insertSheet(CONFIG.ABA_CENTRAL);
    }
    
    // Limpar e configurar cabeçalhos
    abaCentral.clear();
    const rangeCabecalho = abaCentral.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
    rangeCabecalho.setValues([CABECALHOS_CENTRAL]);
    
    // Formatação do cabeçalho com padrões brasileiros
    rangeCabecalho
        .setBackground("#1f4e79")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontFamily("Calibri")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    
    abaCentral.setFrozenRows(1);
    
    return { planilhaCentral, abaCentral };
}

/**
* Mostra progresso durante processamento
*/
function mostrarProgresso(loteAtual, totalLotes, tamanheLote) {
    const progresso = `
🔄 PROCESSANDO DADOS

📊 Progresso: Lote ${loteAtual} de ${totalLotes}
📁 Planilhas neste lote: ${tamanheLote}
⏱️ Aguarde...

${loteAtual === 1 ? '🚀 Iniciando processamento...' : ''}
${loteAtual === totalLotes ? '🏁 Lote final - quase pronto!' : ''}
    `;
    
    // Usar toast para não interromper
    SpreadsheetApp.getActive().toast(
        `Processando lote ${loteAtual}/${totalLotes}...`, 
        "🔄 Importando Dados", 
        5
    );
    
    Logger.log(`📊 Progresso: ${loteAtual}/${totalLotes} - ${tamanheLote} planilhas`);
}

/**
* Converte data para formato brasileiro DD/MM/YYYY
*/
function formatarDataBrasileira(data) {
    if (!data || data === "" || data === null) {
        return "";
    }
    
    try {
        // Se já é uma string no formato correto, retorna
        if (typeof data === "string" && data.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            return data;
        }
        
        // Se é um objeto Date, converte
        let dataObj;
        if (data instanceof Date) {
            dataObj = data;
        } else {
            // Tenta converter string para Date
            dataObj = new Date(data);
        }
        
        // Verifica se a data é válida
        if (isNaN(dataObj.getTime())) {
            return "";
        }
        
        // Formata para DD/MM/YYYY
        const dia = String(dataObj.getDate()).padStart(2, '0');
        const mes = String(dataObj.getMonth() + 1).padStart(2, '0');
        const ano = dataObj.getFullYear();
        
        return `${dia}/${mes}/${ano}`;
        
    } catch (erro) {
        Logger.log(`⚠️ Erro ao formatar data: ${erro.toString()}`);
        return "";
    }
}

/**
* Aplica formatação otimizada com padrões brasileiros
*/
function aplicarFormatacaoOtimizada(abaCentral, totalLinhas) {
    if (totalLinhas <= 1) return;
    
    try {
        Logger.log("🎨 Aplicando formatação brasileira...");
        
        // ========================================================================
        // CONFIGURAÇÕES DE LARGURA DAS COLUNAS (em pixels)
        // ========================================================================
        const largurasColunas = [66, 257, 68, 108, 103, 168, 97, 88, 76, 256, 125, 170, 88];
        
        // Aplicar larguras específicas
        largurasColunas.forEach((largura, indice) => {
            abaCentral.setColumnWidth(indice + 1, largura);
        });
        
        // ========================================================================
        // FORMATAÇÃO GERAL DA PLANILHA
        // ========================================================================
        
        // Formatação do cabeçalho (linha 1)
        const rangeCabecalho = abaCentral.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
        rangeCabecalho
            .setBackground("#1f4e79")
            .setFontColor("#ffffff")
            .setFontWeight("bold")
            .setFontFamily("Calibri")
            .setFontSize(10)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        
        // Formatação dos dados (linhas 2 em diante)
        const rangeDados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length);
        rangeDados
            .setFontFamily("Calibri")
            .setFontSize(10)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setWrap(true)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        
        // ========================================================================
        // FORMATAÇÃO ESPECÍFICA DA COLUNA DE DATA (M)
        // ========================================================================
        const colunaData = abaCentral.getRange(2, 13, totalLinhas - 1, 1); // Coluna M
        colunaData.setNumberFormat("dd/mm/yyyy"); // Formato brasileiro DD/MM/YYYY
        
        // ========================================================================
        // FORMATAÇÃO CONDICIONAL E BORDAS
        // ========================================================================
        
        // Aplicar formatação condicional e bordas usando função auxiliar
        aplicarFormatacaoCondicionalEBordas(abaCentral, totalLinhas);
        
        // ========================================================================
        // CONFIGURAÇÕES ADICIONAIS
        // ========================================================================
        
        // Congelar primeira linha
        abaCentral.setFrozenRows(1);
        
        // Ajustar altura das linhas automaticamente
        for (let i = 2; i <= totalLinhas; i++) {
            abaCentral.autoResizeRows(i);
        }
        
        // Manter altura fixa para o cabeçalho
        abaCentral.setRowHeight(1, 25);
        
        Logger.log("✅ Formatação brasileira aplicada com sucesso!");
        Logger.log(`📊 Colunas configuradas: ${largurasColunas.join(", ")} pixels`);
        Logger.log("📅 Formato de data: DD/MM/YYYY (brasileiro)");
        Logger.log("🔤 Fonte: Calibri 10");
        Logger.log("🎨 Cores alternadas: #ffffff e #ebeff1");
        Logger.log("🔲 Bordas: todas as bordas, cor #ffffff, estilo SOLID_MEDIUM");
        Logger.log("📝 Texto: quebra de linha ativada, altura ajustada automaticamente");
        
    } catch (erro) {
        Logger.log("⚠️ Erro na formatação: " + erro.toString());
    }
}

/**
* Exibe resultado otimizado
*/
function exibirResultadoOtimizado(relatorio) {
    const porcentagemSucesso = Math.round((relatorio.secretariasProcessadas / PLANILHAS_SECRETARIAS.length) * 100);
    
    const mensagem = `
🎉 IMPORTAÇÃO CONCLUÍDA!

📊 RESULTADOS:
• Secretarias processadas: ${relatorio.secretariasProcessadas}/${PLANILHAS_SECRETARIAS.length} (${porcentagemSucesso}%)
• Registros importados: ${relatorio.registrosImportados}
• Lotes processados: ${relatorio.lotes.length}
• Duração total: ${relatorio.duracao}s

✨ DADOS ORDENADOS ALFABETICAMENTE POR SECRETARIA

${relatorio.erros.length > 0 ? `⚠️ Erros encontrados: ${relatorio.erros.length}\nConsulte os logs para detalhes.` : '✅ Processo executado sem erros!'}

📅 Concluído em: ${relatorio.fim.toLocaleString('pt-BR')}
    `;
    
    SpreadsheetApp.getUi().alert(
        "🏛️ Banco de Talentos - Sucesso!", 
        mensagem,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

/**
* Atualiza apenas uma secretaria específica
*/
function atualizarSecretariaEspecifica() {
    // Criar lista de opções
    const opcoes = PLANILHAS_SECRETARIAS.map((s, i) => `${i + 1} - ${s.nome}`);
    
    const resposta = SpreadsheetApp.getUi().prompt(
        "📂 Atualizar Secretaria Específica",
        `Digite o número da secretaria (1-${PLANILHAS_SECRETARIAS.length}):\n\n` + opcoes.join("\n"),
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    
    if (resposta.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        const numero = parseInt(resposta.getResponseText());
        
        if (numero >= 1 && numero <= PLANILHAS_SECRETARIAS.length) {
            const secretaria = PLANILHAS_SECRETARIAS[numero - 1];
            atualizarUmaSecretaria(secretaria);
        } else {
            SpreadsheetApp.getUi().alert("⚠️ Número inválido", "Digite um número entre 1 e " + PLANILHAS_SECRETARIAS.length);
        }
    }
}

/**
* Atualiza dados de uma secretaria específica
*/
function atualizarUmaSecretaria(secretaria) {
    try {
        Logger.log(`🔄 Atualizando secretaria: ${secretaria.nome}`);
        
        const resultado = processarSecretariaOtimizada(secretaria);
        
        if (resultado.sucesso) {
            SpreadsheetApp.getUi().alert(
                "✅ Secretaria Atualizada",
                `${resultado.siglaSecretaria}: ${resultado.dados.length} registros processados`,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        } else {
            SpreadsheetApp.getUi().alert(
                "❌ Erro na Atualização",
                `Erro ao processar ${secretaria.nome}:\n${resultado.erro}`,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
        
    } catch (erro) {
        Logger.log(`❌ Erro ao atualizar ${secretaria.nome}: ${erro.toString()}`);
    }
}

/**
* Verifica dados existentes na planilha
*/
function verificarDadosExistentes() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Nenhum Dado Encontrado",
                "A aba central ainda não foi criada.\nUse 'Importar Dados' primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        const totalRegistros = Math.max(0, totalLinhas - 1);
        
        if (totalRegistros === 0) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Planilha Vazia",
                "A planilha central não possui dados.\nUse 'Importar Dados' para começar.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        // Contar registros por secretaria
        const dadosSecretaria = abaCentral.getRange(2, 1, totalRegistros, 1).getValues();
        const contadorSecretarias = {};
        
        dadosSecretaria.forEach(linha => {
            const secretaria = linha[0] || "NÃO IDENTIFICADA";
            contadorSecretarias[secretaria] = (contadorSecretarias[secretaria] || 0) + 1;
        });
        
        // Ordenar secretarias alfabeticamente
        const secretariasOrdenadas = Object.entries(contadorSecretarias)
            .sort((a, b) => a[0].localeCompare(b[0]));
        
        let relatorio = `
📊 DADOS EXISTENTES NA PLANILHA

📈 RESUMO GERAL:
• Total de registros: ${totalRegistros}
• Secretarias identificadas: ${Object.keys(contadorSecretarias).length}

🏢 DISTRIBUIÇÃO POR SECRETARIA (em ordem alfabética):
`;
        
        secretariasOrdenadas.forEach(([secretaria, quantidade]) => {
            relatorio += `• ${secretaria}: ${quantidade} registros\n`;
        });
        
        relatorio += `
📅 Verificação realizada em: ${new Date().toLocaleString('pt-BR')}
        `;
        
        SpreadsheetApp.getUi().alert(
            "📊 Dados Existentes",
            relatorio,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro na verificação: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Verificação",
            "Ocorreu um erro ao verificar os dados existentes.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Gera relatório completo do sistema
*/
function gerarRelatorioCompleto() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral || abaCentral.getLastRow() <= 1) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Relatório Não Disponível",
                "Execute a importação de dados primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        const totalRegistros = totalLinhas - 1;
        
        // Análise detalhada
        const todosOsDados = abaCentral.getRange(2, 1, totalRegistros, CABECALHOS_CENTRAL.length).getValues();
        
        const analise = {
            secretarias: {},
            camposVazios: {},
            estatisticas: {
                comNome: 0,
                comProntuario: 0,
                comFormacao: 0,
                readaptados: 0
            }
        };
        
        // Processar cada registro
        todosOsDados.forEach(linha => {
            const secretaria = linha[0] || "NÃO IDENTIFICADA";
            const nome = linha[1] || "";
            const prontuario = linha[2] || "";
            const formacao = linha[3] || "";
            const readaptado = linha[8] || "";
            
            // Contar por secretaria
            analise.secretarias[secretaria] = (analise.secretarias[secretaria] || 0) + 1;
            
            // Estatísticas de preenchimento
            if (nome.toString().trim()) analise.estatisticas.comNome++;
            if (prontuario.toString().trim()) analise.estatisticas.comProntuario++;
            if (formacao.toString().trim()) analise.estatisticas.comFormacao++;
            if (readaptado.toString().trim().toLowerCase().includes('sim')) analise.estatisticas.readaptados++;
            
            // Campos vazios por coluna
            CABECALHOS_CENTRAL.forEach((cabecalho, indice) => {
                if (!linha[indice] || !linha[indice].toString().trim()) {
                    analise.camposVazios[cabecalho] = (analise.camposVazios[cabecalho] || 0) + 1;
                }
            });
        });
        
        // Montar relatório
        const secretariasOrdenadas = Object.entries(analise.secretarias)
            .sort((a, b) => b[1] - a[1]); // Ordenar por quantidade (maior primeiro)
        
        let relatorioCompleto = `
📊 RELATÓRIO COMPLETO DO SISTEMA

📈 ESTATÍSTICAS GERAIS:
• Total de registros: ${totalRegistros}
• Secretarias ativas: ${Object.keys(analise.secretarias).length}
• Registros com nome: ${analise.estatisticas.comNome} (${Math.round(analise.estatisticas.comNome/totalRegistros*100)}%)
• Registros com prontuário: ${analise.estatisticas.comProntuario} (${Math.round(analise.estatisticas.comProntuario/totalRegistros*100)}%)
• Servidores readaptados: ${analise.estatisticas.readaptados}

🏢 RANKING DE SECRETARIAS (por quantidade de registros):
`;
        
        secretariasOrdenadas.forEach(([secretaria, quantidade], indice) => {
            const porcentagem = Math.round(quantidade/totalRegistros*100);
            relatorioCompleto += `${indice + 1}. ${secretaria}: ${quantidade} (${porcentagem}%)\n`;
        });
        
        relatorioCompleto += `
📊 QUALIDADE DOS DADOS (campos com maior índice de preenchimento):
`;
        
        const camposOrdenados = Object.entries(analise.camposVazios)
            .sort((a, b) => a[1] - b[1]) // Menos vazios primeiro
            .slice(0, 5);
        
        camposOrdenados.forEach(([campo, vazios]) => {
            const preenchidos = totalRegistros - vazios;
            const porcentagem = Math.round(preenchidos/totalRegistros*100);
            relatorioCompleto += `• ${campo}: ${porcentagem}% preenchido\n`;
        });
        
        relatorioCompleto += `
📅 Relatório gerado em: ${new Date().toLocaleString('pt-BR')}
🔄 Para dados atualizados, execute nova importação
        `;
        
        SpreadsheetApp.getUi().alert(
            "📊 Relatório Completo",
            relatorioCompleto,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro no relatório: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro no Relatório",
            "Ocorreu um erro ao gerar o relatório completo.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Limpa dados e reinicia sistema
*/
function limparEReiniciar() {
    const resposta = SpreadsheetApp.getUi().alert(
        "⚠️ Confirmar Limpeza Total",
        "Esta ação irá:\n\n• Remover TODOS os dados importados\n• Manter apenas os cabeçalhos\n• Permitir uma importação limpa\n\n❗ Esta ação NÃO pode ser desfeita!\n\nDeseja continuar?",
        SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (resposta === SpreadsheetApp.getUi().Button.YES) {
        try {
            const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
            const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
            
            if (abaCentral) {
                // Limpar tudo exceto cabeçalhos
                const ultimaLinha = abaCentral.getLastRow();
                if (ultimaLinha > 1) {
                    abaCentral.getRange(2, 1, ultimaLinha - 1, abaCentral.getLastColumn()).clearContent();
                    abaCentral.getRange(2, 1, ultimaLinha - 1, abaCentral.getLastColumn()).clearFormat();
                }
                
                // Reconfigurar cabeçalhos com formatação brasileira
                const rangeCabecalho = abaCentral.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
                rangeCabecalho.setValues([CABECALHOS_CENTRAL]);
                rangeCabecalho
                    .setBackground("#1f4e79")
                    .setFontColor("#ffffff")
                    .setFontWeight("bold")
                    .setFontFamily("Calibri")
                    .setFontSize(10)
                    .setHorizontalAlignment("center")
                    .setVerticalAlignment("middle");
            }
            
            SpreadsheetApp.getUi().alert(
                "✅ Sistema Reiniciado",
                "Todos os dados foram removidos com sucesso!\n\n🚀 Sistema pronto para nova importação.\n\nUse 'Importar Dados' quando necessário.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
            Logger.log("🧹 Sistema reiniciado - dados limpos");
            
        } catch (erro) {
            Logger.log(`❌ Erro na limpeza: ${erro.toString()}`);
            SpreadsheetApp.getUi().alert(
                "❌ Erro na Limpeza",
                "Ocorreu um erro durante a limpeza dos dados.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }
}

/**
* Corrigir formato de datas em planilha existente
*/
function corrigirFormatoDatas() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Aba Não Encontrada",
                "A aba central não foi encontrada.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        
        if (totalLinhas <= 1) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem Dados",
                "Não há dados para corrigir.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        // Ler dados da coluna M (Data da Inclusão)
        const colunaData = abaCentral.getRange(2, 13, totalLinhas - 1, 1);
        const dadosData = colunaData.getValues();
        
        // Converter cada data para formato brasileiro
        const datasCorrigidas = dadosData.map(linha => {
            const dataOriginal = linha[0];
            return [formatarDataBrasileira(dataOriginal)];
        });
        
        // Aplicar as datas corrigidas
        colunaData.setValues(datasCorrigidas);
        
        // Aplicar formatação de data
        colunaData.setNumberFormat("dd/mm/yyyy");
        
        SpreadsheetApp.getUi().alert(
            "✅ Datas Corrigidas",
            `Formato de datas corrigido com sucesso!\n\n📊 ${totalLinhas - 1} registros processados\n📅 Formato: DD/MM/YYYY (brasileiro)`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro na correção de datas: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Correção",
            "Ocorreu um erro ao corrigir o formato das datas.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Aplicar formatação condicional e bordas (função auxiliar)
*/
function aplicarFormatacaoCondicionalEBordas(abaCentral, totalLinhas) {
    try {
        Logger.log("🎨 Aplicando formatação condicional e bordas...");
        
        // ========================================================================
        // FORMATAÇÃO CONDICIONAL COM CORES ALTERNADAS
        // ========================================================================
        
        // Aplicar cores alternadas nas linhas de dados
        for (let i = 2; i <= totalLinhas; i++) {
            const linha = abaCentral.getRange(i, 1, 1, CABECALHOS_CENTRAL.length);
            const corFundo = (i % 2 === 0) ? "#ffffff" : "#ebeff1";
            linha.setBackground(corFundo);
        }
        
        // Garantir que a coluna secretaria mantenha sua cor especial
        const colunaSecretaria = abaCentral.getRange(2, 1, totalLinhas - 1, 1);
        colunaSecretaria.setBackground("#e8f4fd");
        colunaSecretaria.setFontWeight("bold");
        
        // ========================================================================
        // BORDAS E ESTILO
        // ========================================================================
        
        // Aplicar bordas em toda a planilha
        const rangeCompleto = abaCentral.getRange(1, 1, totalLinhas, CABECALHOS_CENTRAL.length);
        
        // Aplicar bordas externas
        rangeCompleto.setBorder(
            true, true, true, true, false, false, // bordas externas
            "#ffffff", // cor da borda
            SpreadsheetApp.BorderStyle.SOLID_MEDIUM // estilo
        );
        
        // Aplicar bordas internas (grade)
        rangeCompleto.setBorder(
            false, false, false, false, true, true, // bordas internas
            "#ffffff", // cor da borda
            SpreadsheetApp.BorderStyle.SOLID // estilo mais fino para grade
        );
        
        Logger.log("✅ Formatação condicional e bordas aplicadas!");
        
    } catch (erro) {
        Logger.log("⚠️ Erro na formatação condicional: " + erro.toString());
    }
}

/**
* Aplicar ajuste de texto e quebra de linha
*/
function aplicarAjusteTexto() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Aba Não Encontrada",
                "A aba central não foi encontrada.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        
        if (totalLinhas <= 1) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem Dados",
                "Não há dados para formatar.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        Logger.log("📝 Aplicando ajuste de texto e quebra de linha...");
        
        // Formatação dos dados (linhas 2 em diante)
        const rangeDados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length);
        rangeDados
            .setWrap(true)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        
        // Ajustar altura das linhas automaticamente
        for (let i = 2; i <= totalLinhas; i++) {
            abaCentral.autoResizeRows(i);
        }
        
        SpreadsheetApp.getUi().alert(
            "✅ Ajuste de Texto Aplicado",
            `Ajuste de texto e quebra de linha aplicados com sucesso!\n\n📊 ${totalLinhas - 1} registros formatados\n📝 Quebra de linha: ativada\n📏 Altura das linhas: ajustada automaticamente`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro no ajuste de texto: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro no Ajuste",
            "Ocorreu um erro ao aplicar o ajuste de texto.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Aplicar apenas cores alternadas e bordas
*/
function aplicarCoresEBordas() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Aba Não Encontrada",
                "A aba central não foi encontrada.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        
        if (totalLinhas <= 1) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem Dados",
                "Não há dados para formatar.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        // Aplicar formatação condicional e bordas
        aplicarFormatacaoCondicionalEBordas(abaCentral, totalLinhas);
        
        SpreadsheetApp.getUi().alert(
            "✅ Cores e Bordas Aplicadas",
            `Formatação condicional e bordas aplicadas com sucesso!\n\n📊 ${totalLinhas - 1} registros formatados\n🎨 Cores alternadas: #ffffff e #ebeff1\n🔲 Bordas: todas as bordas, cor #ffffff`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro na formatação: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Formatação",
            "Ocorreu um erro ao aplicar cores e bordas.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Aplicar formatação brasileira em planilha existente
*/
function aplicarFormatacaoBrasileira() {
    try {
        const planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
        const abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaCentral) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Aba Não Encontrada",
                "A aba central não foi encontrada.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        const totalLinhas = abaCentral.getLastRow();
        
        if (totalLinhas <= 1) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem Dados",
                "Não há dados para formatar.\nExecute a importação primeiro.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return;
        }
        
        // Aplicar formatação brasileira
        aplicarFormatacaoOtimizada(abaCentral, totalLinhas);
        
        SpreadsheetApp.getUi().alert(
            "✅ Formatação Aplicada",
            `Formatação brasileira aplicada com sucesso!\n\n📊 ${totalLinhas - 1} registros formatados\n🔤 Fonte: Calibri 10\n📅 Data: DD/MM/YYYY\n📝 Texto: ajustado automaticamente\n🎨 Cores alternadas: #ffffff e #ebeff1\n🔲 Bordas: todas as bordas, cor #ffffff\n📏 Colunas: larguras personalizadas`,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        
    } catch (erro) {
        Logger.log(`❌ Erro na formatação: ${erro.toString()}`);
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Formatação",
            "Ocorreu um erro ao aplicar a formatação brasileira.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
* Limpar logs para debug
*/
function limparLogs() {
    console.clear();
    Logger.log("🧹 Logs limpos - " + new Date().toLocaleString('pt-BR'));
    SpreadsheetApp.getActive().toast("Logs limpos!", "🧹 Debug", 2);
}

/**
* Validação da configuração do sistema
*/
function validarConfiguracao() {
    Logger.log("🔍 === VALIDAÇÃO DO SISTEMA ===");
    
    const problemas = [];
    
    // Validar IDs das planilhas
    if (PLANILHAS_SECRETARIAS.length !== 23) {
        problemas.push(`❌ Esperadas 23 planilhas, encontradas ${PLANILHAS_SECRETARIAS.length}`);
    }
    
    // Validar cabeçalhos
    if (CABECALHOS_CENTRAL.length !== 13) {
        problemas.push(`❌ Esperados 13 cabeçalhos, encontrados ${CABECALHOS_CENTRAL.length}`);
    }
    
    // Validar configurações
    if (CONFIG.LOTE_SIZE < 1 || CONFIG.LOTE_SIZE > 10) {
        problemas.push(`❌ Tamanho do lote inválido: ${CONFIG.LOTE_SIZE}`);
    }
    
    if (problemas.length === 0) {
        Logger.log("✅ Sistema validado - configuração correta");
        return true;
    } else {
        Logger.log("❌ Problemas encontrados:");
        problemas.forEach(p => Logger.log(p));
        return false;
    }
}

/**
* Função de inicialização atualizada
*/
function onOpen() {
    try {
        criarMenuCompletoCorrigido();
        Logger.log("✅ Menu completo corrigido criado");
        
        // Mostrar instruções da versão corrigida
        mostrarInstrucoesCorrigidas();
        
    } catch (erro) {
        Logger.log("❌ Erro na inicialização v3.1: " + erro.toString());
    }
}

/**
* Mostra instruções da versão corrigida
*/
function mostrarInstrucoesCorrigidas() {
    const instrucoes = `
🏛️ SISTEMA BANCO DE TALENTOS - VERSÃO 3.1 CORRIGIDA

🔧 CORREÇÕES IMPLEMENTADAS:
• ✅ Sigla das secretarias: agora usa o array PLANILHAS_SECRETARIAS
• ✅ Mapeamento de colunas: corrigido conforme especificação
• ✅ Ordenação alfabética: secretarias sempre ordenadas
• ✅ Início dos dados: confirmado linha 5 (índice 4)

🧪 NOVAS OPÇÕES DE TESTE:
• "Testar Mapeamento" - verifica se as colunas estão corretas
• "Validar Estruturas" - testa todas as secretarias
• "Listar Secretarias Ordenadas" - mostra ordem alfabética

⚡ COMO USAR:
1. Execute "Testar Mapeamento" primeiro para verificar
2. Use "Importar Dados" para processamento completo
3. Dados ficarão ordenados alfabeticamente por secretaria

🎯 VERSÃO 3.1 - PROBLEMAS CORRIGIDOS!
    `;
    
    SpreadsheetApp.getUi().alert(
        "🎉 Sistema Corrigido - v3.1!", 
        instrucoes,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}
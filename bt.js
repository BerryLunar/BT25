/**
 * ========================================================================
 * SISTEMA BANCO DE TALENTOS - VERSÃO FINAL CORRIGIDA
 * Santana de Parnaíba - SP
 * ========================================================================
 * 
 * FUNCIONALIDADES IMPLEMENTADAS:
 * 1. Função "Atualizar Secretaria Específica" corrigida
 * 2. Menu simplificado com apenas 3 opções
 * 3. Autofill na aba "Movimentações 2025"
 * 4. Sincronização de status com PGE das secretarias
 * 5. Envio automático de e-mails para secretários e pontos focais
 * 
 * ========================================================================
 */

// ========================================================================
// CONFIGURAÇÕES GLOBAIS
// ========================================================================

var PLANILHAS_SECRETARIAS = [
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

var CONFIG = {
    ABA_CENTRAL: "BT 2025",
    ABA_ORIGEM: "Banco de Talentos (Externo)",
    ABA_MOVIMENTACOES: "Movimentações 2025",
    LINHA_INICIO_DADOS: 4,
    LOTE_SIZE: 5,
    TIMEOUT_POR_LOTE: 30000,
    DELAY_ENTRE_LOTES: 2000,
    MAX_TENTATIVAS: 2
};

var CABECALHOS_CENTRAL = [
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
    "Data da Inclusão",     // M
    "Status da Movimentação", // N
    "Interesse do Servidor", // O
];

// E-mails dos secretários para notificação
var EMAILS_SECRETARIOS = {
    "SECOM": "rogerio.05735@santanadeparnaiba.sp.gov.br",
    "SEMEDES": "rosalia.37356@santanadeparnaiba.sp.gov.br",
    "SEMOP": "willian.35778@santanadeparnaiba.sp.gov.br",
    "SEMUTRANS": "mauriceia.13547@santanadeparnaiba.sp.gov.br",
    "SMA": "joao.37097@santanadeparnaiba.sp.gov.br",
    "SMAFEL": "wellisson.41377@santanadeparnaiba.sp.gov.br",
    "SMCC": "moises.32666@santanadeparnaiba.sp.gov.br",
    "SMCL": "jose.45849@santanadeparnaiba.sp.gov.br",
    "SMCT": "ricardo.29338@santanadeparnaiba.sp.gov.br",
    "SMDS": "camila.42179@santanadeparnaiba.sp.gov.br",
    "SME": "denise.16870@edu.santanadeparnaiba.sp.gov.br",
    "SMF": "olga.28375@santanadeparnaiba.sp.gov.br",
    "SMGAED": "pedro.41937@santanadeparnaiba.sp.gov.br",
    "SMH": "angela.29303@santanadeparnaiba.sp.gov.br",
    "SMMAP": "juliana.35797@santanadeparnaiba.sp.gov.br",
    "SMMF": "mariana.37113@santanadeparnaiba.sp.gov.br",
    "SMNJ": "albaneide.32343@santanadeparnaiba.sp.gov.br",
    "SMOP": "gerlaine.40923@santanadeparnaiba.sp.gov.br",
    "SMOU": "simone.43610@santanadeparnaiba.sp.gov.br",
    "SMS": "wilson.45853@santanadeparnaiba.sp.gov.br",
    "SMSD": "viviane.26822@santanadeparnaiba.sp.gov.br",
    "SMSM": "maria.42819@santanadeparnaiba.sp.gov.br",
    "SMSU": "ricardo.02732@santanadeparnaiba.sp.gov.br"
};

// E-mails dos pontos focais para notificação
var EMAILS_PONTOS_FOCAIS = {
    "SECOM": "carolina.38338@santanadeparnaiba.sp.gov.br",
    "SEMEDES": "ana.44313@santanadeparnaiba.sp.gov.br",
    "SEMOP": "rosangela.20158@santanadeparnaiba.sp.gov.br",
    "SEMUTRANS": "marcio.37806@santanadeparnaiba.sp.gov.br",
    "SMA": "libian.34565@santanadeparnaiba.sp.gov.br",
    "SMAFEL": "vitoria.40868@santanadeparnaiba.sp.gov.br",
    "SMCC": "jailton.34100@santanadeparnaiba.sp.gov.br",
    "SMCL": "rubens.26653@santanadeparnaiba.sp.gov.br",
    "SMCT": "diego.35011@santanadeparnaiba.sp.gov.br",
    "SMDS": "andre.26547@santanadeparnaiba.sp.gov.br",
    "SME": "tania.03067@edu.santanadeparnaiba.sp.gov.br",
    "SMF": "elza.40028@santanadeparnaiba.sp.gov.br",
    "SMGAED": "vaumil.46330@santanadeparnaiba.sp.gov.br",
    "SMH": "mauricio.29797@santanadeparnaiba.sp.gov.br",
    "SMMAP": "diego.28488@santanadeparnaiba.sp.gov.br",
    "SMMF": "veruska.32203@santanadeparnaiba.sp.gov.br",
    "SMNJ": "selma.001ff@santanadeparnaiba.sp.gov.br",
    "SMOP": "veronica.32196@santanadeparnaiba.sp.gov.br",
    "SMOU": "vivian.29442@santanadeparnaiba.sp.gov.br",
    "SMS": "raquel.41575@santanadeparnaiba.sp.gov.br",
    "SMSD": "carla.23199@santanadeparnaiba.sp.gov.br",
    "SMSM": "vera.27405@santanadeparnaiba.sp.gov.br",
    "SMSU": "felipe.42463@santanadeparnaiba.sp.gov.br"
};

// ========================================================================
// MENU PRINCIPAL SIMPLIFICADO
// ========================================================================

function onOpen() {
    try {
        criarMenuPersonalizado();
        Logger.log("✅ Menu personalizado criado");
    } catch (erro) {
        Logger.log("❌ Erro na inicialização: " + erro.toString());
    }
}

function criarMenuPersonalizado() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("🏛️ Banco de Talentos")
        .addItem("🔄 Atualizar Banco", "importarBancoDeTalentosOtimizado")
        .addItem("📊 Atualizar Secretaria Específica", "atualizarSecretariaEspecifica")
        .addSeparator()
        .addItem("ℹ️ Sobre o Sistema", "exibirSobre")
        .addToUi();
}

function exibirSobre() {
    var sobre = 
        "🏛️ SISTEMA BANCO DE TALENTOS\n" +
        "Programa Governo Eficaz - Santana de Parnaíba\n\n" +
        "🎯 COMO USAR:\n" +
        "• Utilize a função \"Atualizar Secretaria Específica\" para adição mais rápida de dados\n" +
        "• Para atualização completa dos dados, utilize a função \"Atualizar Banco\"\n" +
        "• Toda movimentação e seus detalhes devem ser registrados na aba \"Movimentações 2025\"\n" +
        "• Notificações automáticas serão enviadas para secretários e pontos focais\n\n" +
        "📞 SUPORTE TÉCNICO:\n" +
        "📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br\n" +
        "📱 4622-7500 - 8819 / 8644 / 7574\n\n" +
        "🚀 Versão 3.2 - Sistema de Notificações Duplas\n" +
        "📅 10/09/2025";
    
    SpreadsheetApp.getUi().alert("ℹ️ Sobre o Sistema", sobre, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================================================
// FUNÇÃO PRINCIPAL DE IMPORTAÇÃO (ATUALIZAR BANCO)
// ========================================================================

function importarBancoDeTalentosOtimizado() {
    var inicioExecucao = new Date();
    var relatorio = {
        inicio: inicioExecucao,
        secretariasProcessadas: 0,
        registrosImportados: 0,
        erros: [],
        lotes: []
    };
    
    try {
        Logger.log("🚀 === INICIANDO ATUALIZAÇÃO COMPLETA ===");
        
        // Preparar planilha central
        var planilhaCentral = prepararPlanilhaCentral();
        var abaCentral = planilhaCentral.abaCentral;
        
        // Ordenar secretarias alfabeticamente
        var secretariasOrdenadas = PLANILHAS_SECRETARIAS.slice().sort(function(a, b) {
            return a.nome.localeCompare(b.nome);
        });
        
        // Coletar todos os dados
        var todosOsDados = [];
        
        // Processar em lotes
        var lotes = criarLotes(secretariasOrdenadas, CONFIG.LOTE_SIZE);
        
        for (var i = 0; i < lotes.length; i++) {
            var lote = lotes[i];
            var numeroLote = i + 1;
            var totalLotes = lotes.length;
            
            // Mostrar progresso
            SpreadsheetApp.getActive().toast(
                "Processando lote " + numeroLote + "/" + totalLotes + "...", 
                "🔄 Atualizando Banco", 
                5
            );
            
            var resultadoLote = processarLoteSecretarias(lote, numeroLote);
            
            // Adicionar dados do lote
            todosOsDados = todosOsDados.concat(resultadoLote.dados);
            
            // Atualizar relatório
            relatorio.secretariasProcessadas += resultadoLote.processadas;
            relatorio.erros = relatorio.erros.concat(resultadoLote.erros);
            relatorio.lotes.push(resultadoLote);
            
            // Pausa entre lotes
            if (i < lotes.length - 1) {
                Utilities.sleep(CONFIG.DELAY_ENTRE_LOTES);
            }
        }
        
        // Inserir todos os dados
        if (todosOsDados.length > 0) {
            Logger.log("📝 Inserindo " + todosOsDados.length + " registros ordenados...");
            
            var range = abaCentral.getRange(2, 1, todosOsDados.length, CABECALHOS_CENTRAL.length);
            range.setValues(todosOsDados);
            
            relatorio.registrosImportados = todosOsDados.length;
            
            // Aplicar formatação
            aplicarFormatacaoOtimizada(abaCentral, todosOsDados.length + 1);
        }
        
        // Finalizar
        relatorio.fim = new Date();
        relatorio.duracao = Math.round((relatorio.fim - relatorio.inicio) / 1000);
        
        exibirResultadoOtimizado(relatorio);
        
        Logger.log("🎉 === ATUALIZAÇÃO COMPLETA CONCLUÍDA ===");
        
    } catch (erro) {
        Logger.log("💥 ERRO CRÍTICO: " + erro.toString());
        
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Atualização",
            "Erro crítico durante a atualização:\n\n" + erro.toString(),
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

// ========================================================================
// FUNÇÃO ATUALIZAR SECRETARIA ESPECÍFICA - CORRIGIDA
// ========================================================================

function atualizarSecretariaEspecifica() {
    var secretariasOrdenadas = PLANILHAS_SECRETARIAS.slice().sort(function(a, b) {
        return a.nome.localeCompare(b.nome);
    });
    
    var opcoes = [];
    for (var i = 0; i < secretariasOrdenadas.length; i++) {
        opcoes.push((i + 1) + " - " + secretariasOrdenadas[i].nome);
    }
    
    var resposta = SpreadsheetApp.getUi().prompt(
        "📂 Atualizar Secretaria Específica",
        "Digite o número da secretaria (1-" + PLANILHAS_SECRETARIAS.length + "):\n\n" + opcoes.join("\n"),
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );

    if (resposta.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        var numero = parseInt(resposta.getResponseText());
        
        if (numero >= 1 && numero <= PLANILHAS_SECRETARIAS.length) {
            var secretaria = secretariasOrdenadas[numero - 1];
            atualizarUmaSecretaria(secretaria);
        } else {
            SpreadsheetApp.getUi().alert("⚠️ Número inválido", "Digite um número entre 1 e " + PLANILHAS_SECRETARIAS.length);
        }
    }
}

function atualizarUmaSecretaria(secretaria) {
    try {
        Logger.log("📂 Atualizando secretaria: " + secretaria.nome);
        
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var abaCentral = ss.getSheetByName(CONFIG.ABA_CENTRAL);
        
        // Criar aba central se não existir
        if (!abaCentral) {
            var planilhaCentral = prepararPlanilhaCentral();
            abaCentral = planilhaCentral.abaCentral;
        }

        // Processar dados da secretaria
        var resultado = processarSecretariaOtimizada(secretaria);

        if (resultado.sucesso && resultado.dados.length > 0) {
            // Remover dados antigos da secretaria
            removerDadosSecretaria(abaCentral, secretaria.nome);
            
            // Inserir novos dados no final
            var ultimaLinha = abaCentral.getLastRow();
            var novaLinha = ultimaLinha + 1;
            
            abaCentral.getRange(novaLinha, 1, resultado.dados.length, CABECALHOS_CENTRAL.length)
                .setValues(resultado.dados);
            
            // Aplicar formatação
            aplicarFormatacaoOtimizada(abaCentral, novaLinha + resultado.dados.length - 1);
            
            // Reordenar dados alfabeticamente
            reordenarDadosAlfabeticamente(abaCentral);

            SpreadsheetApp.getUi().alert(
                "✅ Secretaria Atualizada",
                resultado.siglaSecretaria + ": " + resultado.dados.length + " registros processados e inseridos com sucesso!",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
            Logger.log("✅ " + secretaria.nome + ": " + resultado.dados.length + " registros atualizados");
            
        } else if (resultado.sucesso && resultado.dados.length === 0) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem Dados",
                secretaria.nome + ": Nenhum registro encontrado para atualização.",
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        } else {
            SpreadsheetApp.getUi().alert(
                "❌ Erro na Atualização",
                "Erro ao processar " + secretaria.nome + ":\n" + resultado.erro,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
        
    } catch (erro) {
        Logger.log("❌ Erro ao atualizar " + secretaria.nome + ": " + erro.toString());
        SpreadsheetApp.getUi().alert(
            "❌ Erro Crítico",
            "Erro inesperado ao atualizar " + secretaria.nome + ". Consulte os logs.",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

function removerDadosSecretaria(abaCentral, nomeSecretaria) {
    try {
        var totalLinhas = abaCentral.getLastRow();
        if (totalLinhas <= 1) return;
        
        // Obter todos os dados
        var dados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        
        // Filtrar dados, removendo a secretaria específica
        var dadosFiltrados = [];
        for (var i = 0; i < dados.length; i++) {
            var secretariaLinha = (dados[i][0] || "").toString().trim();
            if (secretariaLinha !== nomeSecretaria) {
                dadosFiltrados.push(dados[i]);
            }
        }
        
        // Limpar área de dados
        if (totalLinhas > 1) {
            abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).clearContent();
        }
        
        // Reescrever dados filtrados
        if (dadosFiltrados.length > 0) {
            abaCentral.getRange(2, 1, dadosFiltrados.length, CABECALHOS_CENTRAL.length).setValues(dadosFiltrados);
        }
        
        Logger.log("📝 Dados antigos da " + nomeSecretaria + " removidos");
        
    } catch (erro) {
        Logger.log("⚠️ Erro ao remover dados da " + nomeSecretaria + ": " + erro.toString());
    }
}

function reordenarDadosAlfabeticamente(abaCentral) {
    try {
        var totalLinhas = abaCentral.getLastRow();
        if (totalLinhas <= 2) return;
        
        // Obter todos os dados
        var dados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        
        // Ordenar alfabeticamente por secretaria
        dados.sort(function(a, b) {
            var secretariaA = (a[0] || "").toString().toUpperCase();
            var secretariaB = (b[0] || "").toString().toUpperCase();
            return secretariaA.localeCompare(secretariaB);
        });
        
        // Reescrever dados ordenados
        abaCentral.getRange(2, 1, dados.length, CABECALHOS_CENTRAL.length).setValues(dados);
        
        Logger.log("📤 Dados reordenados alfabeticamente");
        
    } catch (erro) {
        Logger.log("⚠️ Erro ao reordenar dados: " + erro.toString());
    }
}

// ========================================================================
// FUNÇÃO ONEDIT - AUTOFILL E SINCRONIZAÇÃO
// ========================================================================

function onEdit(e) {
    try {
        var aba = e.range.getSheet();
        var nomeAba = aba.getName();
        
        if (nomeAba !== CONFIG.ABA_MOVIMENTACOES) return;

        var colunaEditada = e.range.getColumn();
        var linhaEditada = e.range.getRow();

        // Autofill baseado no Prontuário (coluna C)
        if (colunaEditada === 3 && linhaEditada >= 3) {
            executarAutofill(aba, linhaEditada, e.range.getValue());
        }

        // Sincronização de status (coluna F)
        if (colunaEditada === 6 && linhaEditada >= 3) {
            sincronizarStatus(aba, linhaEditada);
        }

    } catch (erro) {
        Logger.log("❌ Erro no onEdit Movimentações: " + erro.toString());
    }
}

function executarAutofill(aba, linha, prontuario) {
    try {
        var prontuarioLimpo = prontuario.toString().trim();
        if (!prontuarioLimpo) return;

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var abaBT = ss.getSheetByName(CONFIG.ABA_CENTRAL);
        
        if (!abaBT) {
            Logger.log("⚠️ Aba BT 2025 não encontrada para autofill");
            return;
        }

        var totalLinhas = abaBT.getLastRow();
        if (totalLinhas <= 1) return;

        // Buscar dados na aba BT 2025
        var dadosBT = abaBT.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        
        var encontrado = null;
        for (var i = 0; i < dadosBT.length; i++) {
            var prontuarioBT = (dadosBT[i][2] || "").toString().trim(); // Coluna C (Prontuário)
            if (prontuarioBT === prontuarioLimpo) {
                encontrado = dadosBT[i];
                break;
            }
        }

        if (encontrado) {
            // Preencher automaticamente
            aba.getRange(linha, 1).setValue(encontrado[0]);  // Secretaria (A)
            aba.getRange(linha, 2).setValue(encontrado[1]);  // Nome (B)
            aba.getRange(linha, 4).setValue(encontrado[5]);  // Cargo (D) - coluna F do BT
            aba.getRange(linha, 5).setValue(encontrado[10]); // Ação (E) - coluna K do BT
            
            Logger.log("✅ Autofill executado para prontuário " + prontuarioLimpo);
            
            // Mostrar confirmação visual
            SpreadsheetApp.getActive().toast(
                "Dados preenchidos para " + encontrado[1] + " (" + encontrado[0] + ")", 
                "📋 Autofill", 
                3
            );
        } else {
            Logger.log("⚠️ Prontuário " + prontuarioLimpo + " não encontrado na base de dados");
            SpreadsheetApp.getActive().toast(
                "Prontuário " + prontuarioLimpo + " não encontrado na base de dados", 
                "⚠️ Autofill", 
                3
            );
        }

    } catch (erro) {
        Logger.log("❌ Erro no autofill: " + erro.toString());
    }
}

// ========================================================================
// SINCRONIZAÇÃO DE STATUS E ENVIO DE E-MAIL
// ========================================================================

function sincronizarStatus(aba, linha) {
    try {
        var secretaria = aba.getRange(linha, 1).getValue();     // Coluna A
        var nome = aba.getRange(linha, 2).getValue();           // Coluna B
        var prontuario = aba.getRange(linha, 3).getValue();     // Coluna C
        var statusNovo = aba.getRange(linha, 6).getValue();     // Coluna F

        if (!secretaria || !prontuario || !statusNovo) {
            Logger.log("⚠️ Dados insuficientes para sincronização de status");
            return;
        }

        Logger.log("🔄 Sincronizando status: " + secretaria + " - " + prontuario + " - " + statusNovo);

        // Encontrar a secretaria correspondente
        var secretariaInfo = null;
        for (var i = 0; i < PLANILHAS_SECRETARIAS.length; i++) {
            if (PLANILHAS_SECRETARIAS[i].nome === secretaria.toString().trim()) {
                secretariaInfo = PLANILHAS_SECRETARIAS[i];
                break;
            }
        }
        
        if (!secretariaInfo) {
            Logger.log("⚠️ Secretaria " + secretaria + " não encontrada no array");
            return;
        }

        // Atualizar PGE da secretaria
        try {
            var planilhaPGE = SpreadsheetApp.openById(secretariaInfo.id);
            var abaPGE = planilhaPGE.getSheetByName("Planejamento de Gestão Estratégica");
            
            if (!abaPGE) {
                Logger.log("⚠️ Aba PGE não encontrada na planilha da " + secretaria);
                return;
            }

            var totalLinhasPGE = abaPGE.getLastRow();
            if (totalLinhasPGE < 4) return;

            // Buscar o prontuário na planilha PGE (dados começam na linha 4, coluna C)
            var dadosPGE = abaPGE.getRange(4, 1, totalLinhasPGE - 3, abaPGE.getLastColumn()).getValues();

            for (var i = 0; i < dadosPGE.length; i++) {
                var prontuarioPGE = (dadosPGE[i][2] || "").toString().trim(); // Coluna C
                
                if (prontuarioPGE === prontuario.toString().trim()) {
                    // Atualizar status na coluna Q (índice 16)
                    abaPGE.getRange(i + 4, 17).setValue(statusNovo); // Coluna Q = índice 17
                    Logger.log("✅ Status atualizado no PGE da " + secretaria);
                    break;
                }
            }

        } catch (erroPGE) {
            Logger.log("❌ Erro ao atualizar PGE da " + secretaria + ": " + erroPGE.toString());
        }

        // Enviar e-mail se necessário
        var statusParaEmail = ["Em Andamento", "Concluído"];
        var enviarEmail = false;
        for (var j = 0; j < statusParaEmail.length; j++) {
            if (statusNovo.toString() === statusParaEmail[j]) {
                enviarEmail = true;
                break;
            }
        }
        
        if (enviarEmail) {
            enviarEmailNotificacao(secretaria, nome, prontuario, statusNovo);
        }

        // Confirmação visual
        SpreadsheetApp.getActive().toast(
            "Status \"" + statusNovo + "\" sincronizado para " + secretaria, 
            "🔄 Sincronização", 
            3
        );

    } catch (erro) {
        Logger.log("❌ Erro ao sincronizar status: " + erro.toString());
    }
}

function enviarEmailNotificacao(secretaria, nome, prontuario, status) {
    try {
        var emailSecretario = EMAILS_SECRETARIOS[secretaria];
        var emailPontoFocal = EMAILS_PONTOS_FOCAIS[secretaria];
        
        if (!emailSecretario && !emailPontoFocal) {
            Logger.log("⚠️ Nenhum e-mail encontrado para a secretaria: " + secretaria);
            return;
        }

        // Preparar lista de destinatários
        var destinatarios = [];
        if (emailSecretario) destinatarios.push(emailSecretario);
        if (emailPontoFocal) destinatarios.push(emailPontoFocal);

        var assunto = "Banco de Talentos - Movimentação de Servidor (" + status + ")";
        var corpo = criarCorpoEmailNotificacao(secretaria, nome, prontuario, status);

        // Enviar para todos os destinatários
        for (var i = 0; i < destinatarios.length; i++) {
            try {
                MailApp.sendEmail({
                    to: destinatarios[i],
                    subject: assunto,
                    htmlBody: corpo
                });
                
                Logger.log("📧 E-mail enviado para " + destinatarios[i] + " (" + secretaria + ")");
                
            } catch (erroEmail) {
                Logger.log("❌ Erro ao enviar e-mail para " + destinatarios[i] + ": " + erroEmail.toString());
            }
        }

        // Confirmação visual consolidada
        var tipoDestinatario = "";
        if (emailSecretario && emailPontoFocal) {
            tipoDestinatario = "secretário e ponto focal";
        } else if (emailSecretario) {
            tipoDestinatario = "secretário";
        } else {
            tipoDestinatario = "ponto focal";
        }

        SpreadsheetApp.getActive().toast(
            "E-mail enviado para " + tipoDestinatario + " da " + secretaria, 
            "📧 Notificação", 
            3
        );

    } catch (erro) {
        Logger.log("❌ Erro ao enviar notificação: " + erro.toString());
    }
}

function criarCorpoEmailNotificacao(secretaria, nome, prontuario, status) {
    return "<div style=\"font-family: Calibri, Arial, sans-serif; line-height: 1.6; color: #333;\">" +
        "<h2 style=\"color: #1f4e79;\">🏛️ Banco de Talentos - Prefeitura de Santana de Parnaíba</h2>" +
        "<p><strong>Prezado(a) Gestor(a),</strong></p>" +
        "<p>Informamos que um servidor vinculado à <strong>" + secretaria + "</strong> encontra-se em processo de movimentação através do Banco de Talentos.</p>" +
        "<div style=\"background-color: #f8f9fa; padding: 15px; border-left: 4px solid #1f4e79; margin: 20px 0;\">" +
        "<h3 style=\"margin-top: 0; color: #1f4e79;\">📋 Dados da Movimentação:</h3>" +
        "<p><strong>• Nome do Servidor:</strong> " + nome + "</p>" +
        "<p><strong>• Prontuário:</strong> " + prontuario + "</p>" +
        "<p><strong>• Status Atual:</strong> " + status + "</p>" +
        "<p><strong>• Secretaria de Origem:</strong> " + secretaria + "</p>" +
        "</div>" +
        "<p>Este é um aviso automático do Sistema Banco de Talentos do Programa Governo Eficaz.</p>" +
        "<p>Para mais informações ou esclarecimentos, entre em contato conosco:</p>" +
        "<p>📧 <strong>sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br</strong><br>" +
        "📱 <strong>4622-7500 - 8819 / 8644 / 7574</strong></p>" +
        "<hr style=\"border: none; border-top: 1px solid #ddd; margin: 30px 0;\">" +
        "<p style=\"font-size: 12px; color: #666;\">" +
        "<strong>Atenciosamente,</strong><br>" +
        "Programa Governo Eficaz<br>" +
        "Prefeitura de Santana de Parnaíba<br>" +
        "<em>Sistema automatizado - não responda este e-mail</em>" +
        "</p>" +
        "</div>";
}

// ========================================================================
// FUNÇÕES AUXILIARES E DE PROCESSAMENTO
// ========================================================================

function processarSecretariaOtimizada(secretaria) {
    try {
        var siglaSecretaria = secretaria.nome;
        
        Logger.log("📂 Processando: " + siglaSecretaria + " (ID: " + secretaria.id.substring(0, 10) + "...)");
        
        // Abrir planilha
        var planilhaExterna = SpreadsheetApp.openById(secretaria.id);
        
        // Verificar aba
        var abaOrigem = planilhaExterna.getSheetByName(CONFIG.ABA_ORIGEM);
        if (!abaOrigem) {
            return { 
                sucesso: false, 
                erro: "Aba \"" + CONFIG.ABA_ORIGEM + "\" não encontrada",
                siglaSecretaria: siglaSecretaria 
            };
        }
        
        // Obter dados
        var ultimaLinha = abaOrigem.getLastRow();
        
        if (ultimaLinha <= CONFIG.LINHA_INICIO_DADOS) {
            Logger.log("ℹ️ Sem dados: " + siglaSecretaria);
            return { 
                sucesso: true, 
                dados: [], 
                siglaSecretaria: siglaSecretaria 
            };
        }
        
        // Calcular linhas de dados disponíveis
        var totalLinhas = ultimaLinha - CONFIG.LINHA_INICIO_DADOS;
        
        Logger.log("📊 " + siglaSecretaria + ": Linha " + (CONFIG.LINHA_INICIO_DADOS + 1) + " até " + ultimaLinha + " (" + totalLinhas + " linhas)");
        
        // Ler dados das colunas B até R
        var dadosRange = abaOrigem.getRange(
            CONFIG.LINHA_INICIO_DADOS + 1,
            2, // Coluna B (Nome)
            totalLinhas, 
            17 // Colunas B até R
        );
        
        var dadosBrutos = dadosRange.getValues();
        
        Logger.log("📖 " + siglaSecretaria + ": Lidas " + dadosBrutos.length + " linhas de dados brutos");
        
        // Processar dados
        var dadosProcessados = [];
        
        for (var i = 0; i < dadosBrutos.length; i++) {
            var linha = dadosBrutos[i];
            var nome = (linha[0] || "").toString().trim();
            
            if (nome) {
                var linhaCentral = [
                    siglaSecretaria,                           // A - Secretaria
                    nome,                                      // B - Nome
                    (linha[1] || "").toString().trim(),        // C - Prontuário
                    (linha[2] || "").toString().trim(),        // D - Formação Acadêmica
                    (linha[3] || "").toString().trim(),        // E - Área de Formação
                    (linha[4] || "").toString().trim(),        // F - Cargo Concurso
                    (linha[5] || "").toString().trim(),        // G - CC / FE
                    (linha[6] || "").toString().trim(),        // H - Função Gratificada
                    (linha[7] || "").toString().trim(),        // I - Readaptado
                    (linha[8] || "").toString().trim(),        // J - Justificativa
                    (linha[9] || "").toString().trim(),        // K - Ação
                    (linha[10] || "").toString().trim(),       // L - Condicionalidade
                    formatarDataBrasileira(linha[11] || ""),   // M - Data da Inclusão
                    (linha[14] || "").toString().trim(),       // N - Status da Movimentação
                    (linha[15] || "").toString().trim(),       // O - Interesse do Servidor
                ];
            
                dadosProcessados.push(linhaCentral);
            }
        }

        Logger.log("✅ " + siglaSecretaria + ": " + dadosProcessados.length + " registros processados");
        
        return {
            sucesso: true,
            dados: dadosProcessados,
            siglaSecretaria: siglaSecretaria
        };
        
    } catch (erro) {
        Logger.log("❌ Erro em " + secretaria.nome + ": " + erro.toString());
        return {
            sucesso: false,
            erro: erro.toString(),
            siglaSecretaria: secretaria.nome
        };
    }
}

function prepararPlanilhaCentral() {
    var planilhaCentral = SpreadsheetApp.getActiveSpreadsheet();
    var abaCentral = planilhaCentral.getSheetByName(CONFIG.ABA_CENTRAL);
    
    if (!abaCentral) {
        Logger.log("📋 Criando aba \"" + CONFIG.ABA_CENTRAL + "\"");
        abaCentral = planilhaCentral.insertSheet(CONFIG.ABA_CENTRAL);
    }
    
    // Limpar e configurar cabeçalhos
    abaCentral.clear();
    var rangeCabecalho = abaCentral.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
    rangeCabecalho.setValues([CABECALHOS_CENTRAL]);
    
    // Formatação do cabeçalho
    rangeCabecalho
        .setBackground("#1f4e79")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontFamily("Calibri")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    
    abaCentral.setFrozenRows(1);
    
    return { planilhaCentral: planilhaCentral, abaCentral: abaCentral };
}

function criarLotes(planilhas, tamanhoLote) {
    var lotes = [];
    for (var i = 0; i < planilhas.length; i += tamanhoLote) {
        lotes.push(planilhas.slice(i, i + tamanhoLote));
    }
    return lotes;
}

function processarLoteSecretarias(lote, numeroLote) {
    Logger.log("📦 Processando lote " + numeroLote + " com " + lote.length + " planilhas");
    
    var dadosLote = [];
    var errosLote = [];
    var processadas = 0;
    
    for (var i = 0; i < lote.length; i++) {
        var secretaria = lote[i];
        try {
            var resultado = processarSecretariaOtimizada(secretaria);
            
            if (resultado.sucesso) {
                dadosLote = dadosLote.concat(resultado.dados);
                processadas++;
                Logger.log("✅ " + resultado.siglaSecretaria + ": " + resultado.dados.length + " registros");
            } else {
                errosLote.push(secretaria.nome + ": " + resultado.erro);
                Logger.log("❌ " + secretaria.nome + ": " + resultado.erro);
            }
            
        } catch (erro) {
            var mensagemErro = secretaria.nome + ": " + erro.toString();
            errosLote.push(mensagemErro);
            Logger.log("💥 " + mensagemErro);
        }
        
        Utilities.sleep(200);
    }
    
    return {
        dados: dadosLote,
        processadas: processadas,
        erros: errosLote,
        numeroLote: numeroLote
    };
}

function formatarDataBrasileira(data) {
    if (!data || data === "" || data === null) {
        return "";
    }
    
    try {
        if (typeof data === "string" && data.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            return data;
        }
        
        var dataObj;
        if (data instanceof Date) {
            dataObj = data;
        } else {
            dataObj = new Date(data);
        }
        
        if (isNaN(dataObj.getTime())) {
            return "";
        }
        
        var dia = String(dataObj.getDate()).padStart(2, '0');
        var mes = String(dataObj.getMonth() + 1).padStart(2, '0');
        var ano = dataObj.getFullYear();
        
        return dia + "/" + mes + "/" + ano;
        
    } catch (erro) {
        Logger.log("⚠️ Erro ao formatar data: " + erro.toString());
        return "";
    }
}

function aplicarFormatacaoOtimizada(abaCentral, totalLinhas) {
    if (totalLinhas <= 1) return;
    
    try {
        Logger.log("🎨 Aplicando formatação brasileira...");
        
        // Larguras das colunas (em pixels)
        var largurasColunas = [66, 257, 68, 108, 103, 168, 97, 88, 76, 256, 125, 170, 88, 120, 150];
        
        for (var i = 0; i < largurasColunas.length; i++) {
            abaCentral.setColumnWidth(i + 1, largurasColunas[i]);
        }
        
        // Formatação dos dados
        var rangeDados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length);
        rangeDados
            .setFontFamily("Calibri")
            .setFontSize(10)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setWrap(true);
        
        // Formatação da coluna de data (M)
        if (totalLinhas > 1) {
            var colunaData = abaCentral.getRange(2, 13, totalLinhas - 1, 1);
            colunaData.setNumberFormat("dd/mm/yyyy");
        }
        
        // Aplicar cores alternadas
        for (var i = 2; i <= totalLinhas; i++) {
            var linha = abaCentral.getRange(i, 1, 1, CABECALHOS_CENTRAL.length);
            var corFundo = (i % 2 === 0) ? "#ffffff" : "#f8f9fa";
            linha.setBackground(corFundo);
        }
        
        // Destaque na coluna secretaria
        var colunaSecretaria = abaCentral.getRange(2, 1, totalLinhas - 1, 1);
        colunaSecretaria.setBackground("#e3f2fd");
        colunaSecretaria.setFontWeight("bold");
        
        // Aplicar bordas
        var rangeCompleto = abaCentral.getRange(1, 1, totalLinhas, CABECALHOS_CENTRAL.length);
        rangeCompleto.setBorder(
            true, true, true, true, true, true,
            "#cccccc",
            SpreadsheetApp.BorderStyle.SOLID
        );
        
        Logger.log("✅ Formatação brasileira aplicada com sucesso!");
        
    } catch (erro) {
        Logger.log("⚠️ Erro na formatação: " + erro.toString());
    }
}

function exibirResultadoOtimizado(relatorio) {
    var porcentagemSucesso = Math.round((relatorio.secretariasProcessadas / PLANILHAS_SECRETARIAS.length) * 100);
    
    var mensagem = 
        "🎉 ATUALIZAÇÃO CONCLUÍDA!\n\n" +
        "📊 RESULTADOS:\n" +
        "• Secretarias processadas: " + relatorio.secretariasProcessadas + "/" + PLANILHAS_SECRETARIAS.length + " (" + porcentagemSucesso + "%)\n" +
        "• Registros importados: " + relatorio.registrosImportados + "\n" +
        "• Duração total: " + relatorio.duracao + "s\n\n" +
        "✨ DADOS ORDENADOS ALFABETICAMENTE POR SECRETARIA\n\n" +
        (relatorio.erros.length > 0 ? 
            "⚠️ Erros encontrados: " + relatorio.erros.length + "\nConsulte os logs para detalhes." : 
            "✅ Processo executado sem erros!") + 
        "\n\n📅 Concluído em: " + relatorio.fim.toLocaleString('pt-BR');
    
    SpreadsheetApp.getUi().alert(
        "🏛️ Banco de Talentos - Sucesso!", 
        mensagem,
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}
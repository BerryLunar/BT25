/**
 * ========================================================================
 * SISTEMA BANCO DE TALENTOS
 * Santana de Parnaíba - SP
 * 
 * Funcionalidades:
 * • Atualização completa ou por secretaria
 * • Autofill na aba "Movimentações 2025"
 * • Sincronização automática com PGE
 * • Envio automático de e-mails ao alterar status
 * 
 * Notificação é enviada quando o status muda para:
 *   • Em Andamento
 *   • Concluído
 *   • Cancelado
 * 
 * ⚠️ IMPORTANTE: Execute "Instalar Gatilhos" uma vez!
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

// E-mails dos secretários
var EMAILS_SECRETARIOS = {
    "ARAT": "rosalia.37356@santanadeparnaiba.sp.gov.br",
    "DEFESA CIVIL": "rogerio.05735@santanadeparnaiba.sp.gov.br",
    "SECOM": "marcio.37806@santanadeparnaiba.sp.gov.br",
    "SEMEDES": "joao.37097@santanadeparnaiba.sp.gov.br",
    "SEMOP (PRIVADA)": "wellisson.41377@santanadeparnaiba.sp.gov.br",
    "SEMUTTRANS": "moises.32666@santanadeparnaiba.sp.gov.br",
    "SMA": "jose.45849@santanadeparnaiba.sp.gov.br",
    "SMAFEL": "ricardo.29338@santanadeparnaiba.sp.gov.br",
    "SMCC": "helio.37195@santanadeparnaiba.sp.gov.br",
    "SMCL": "cleusa.27102@santanadeparnaiba.sp.gov.br",
    "SMCT": "valmir.37361@santanadeparnaiba.sp.gov.br",
    "SMDS": "marcos.45852@santanadeparnaiba.sp.gov.br",
    "SME": "denise.16870@edu.santanadeparnaiba.sp.gov.br",
    "SMF": "vaumil.46330@santanadeparnaiba.sp.gov.br",
    "SMGAED": "mauricio.29797@santanadeparnaiba.sp.gov.br",
    "SMH": "diego.28488@santanadeparnaiba.sp.gov.br",
    "SMMAP": "veruska.32203@santanadeparnaiba.sp.gov.br",
    "SMMF": "selma.001ff@santanadeparnaiba.sp.gov.br",
    "SMNJ": "veronica.32196@santanadeparnaiba.sp.gov.br",
    "SMOP (PÚBLICA)": "vivian.29442@santanadeparnaiba.sp.gov.br",
    "SMOU": "wilson.45853@santanadeparnaiba.sp.gov.br",
    "SMS": "maria.42819@santanadeparnaiba.sp.gov.br",
    "SMSD (TI)": "ricardo.02732@santanadeparnaiba.sp.gov.br",
    "SMSM": "mario.29793@santanadeparnaiba.sp.gov.br",
    "SMSU": "eduardo.43317@santanadeparnaiba.sp.gov.br"
};

// E-mails dos pontos focais
var EMAILS_PONTOS_FOCAIS = {
    "ARAT": ["ana.44313@santanadeparnaiba.sp.gov.br"],
    "DEFESA CIVIL": ["carolina.38338@santanadeparnaiba.sp.gov.br"],
    "SECOM": [
      "rosangela.20158@santanadeparnaiba.sp.gov.br",
      "willian.35778@santanadeparnaiba.sp.gov.br"
    ],
    "SEMEDES": [
      "mauriceia.13547@santanadeparnaiba.sp.gov.br",
      "libian.34565@santanadeparnaiba.sp.gov.br"
    ],
    "SEMOP (PRIVADA)": ["vitoria.40868@santanadeparnaiba.sp.gov.br"],
    "SEMUTTRANS": ["jailton.34100@santanadeparnaiba.sp.gov.br"],
    "SMA": [
      "rubens.26653@santanadeparnaiba.sp.gov.br",
      "vitoria.44738@santanadeparnaiba.sp.gov.br"
    ],
    "SMAFEL": ["diego.35011@santanadeparnaiba.sp.gov.br"],
    "SMCC": ["izabel.30143@santanadeparnaiba.sp.gov.br"],
    "SMCL": [
      "andre.26547@santanadeparnaiba.sp.gov.br",
      "camila.42179@santanadeparnaiba.sp.gov.br",
      "cintia.09595@santanadeparnaiba.sp.gov.br"
    ],
    "SMCT": ["sandra.45791@santanadeparnaiba.sp.gov.br"],
    "SMDS": [
      "maria.10508@santanadeparnaiba.sp.gov.br",
      "bruna.31191@santanadeparnaiba.sp.gov.br"
    ],
    "SME": ["tania.03067@edu.santanadeparnaiba.sp.gov.br"],
    "SMF": [
      "elza.40028@santanadeparnaiba.sp.gov.br",
      "olga.28375@santanadeparnaiba.sp.gov.br"
    ],
    "SMGAED": ["pedro.41937@santanadeparnaiba.sp.gov.br"],
    "SMH": ["angela.29303@santanadeparnaiba.sp.gov.br"],
    "SMMAP": ["juliana.35797@santanadeparnaiba.sp.gov.br"],
    "SMMF": ["mariana.37113@santanadeparnaiba.sp.gov.br"],
    "SMNJ": ["albaneide.32343@santanadeparnaiba.sp.gov.br"],
    "SMOP (PÚBLICA)": ["gerlaine.40923@santanadeparnaiba.sp.gov.br"],
    "SMOU": [
      "simone.43610@santanadeparnaiba.sp.gov.br",
      "raquel.41575@santanadeparnaiba.sp.gov.br"
    ],
    "SMS": [
      "carla.23199@santanadeparnaiba.sp.gov.br",
      "viviane.26822@santanadeparnaiba.sp.gov.br",
      "vera.27405@santanadeparnaiba.sp.gov.br"
    ],
    "SMSD (TI)": ["felipe.42463@santanadeparnaiba.sp.gov.br"],
    "SMSM": ["william.14340@santanadeparnaiba.sp.gov.br"],
    "SMSU": ["ana.39251@santanadeparnaiba.sp.gov.br"]
};

// ========================================================================
// MENU PRINCIPAL
// ========================================================================

function onOpen() {
    criarMenuPersonalizado();
}

function criarMenuPersonalizado() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("🏛️ Banco de Talentos")
        .addItem("🔄 Atualizar Banco", "importarBancoDeTalentosOtimizado")
        .addItem("📊 Atualizar Secretaria Específica", "atualizarSecretariaEspecifica")
        .addSeparator()
        .addItem("📝 Gerar Encaminhamento", "gerarEncaminhamentoSelecionado")
        .addItem("📄 Gerar Ciência", "gerarCienciaSelecionado")
        .addSeparator()
        .addItem("⚙️ Instalar Gatilhos", "instalarGatilhos")
        .addSeparator()
        .addItem("ℹ️ Sobre o Sistema", "exibirSobre")
        .addToUi();
}

function exibirSobre() {
    var sobre = 
        "🏛️ SISTEMA BANCO DE TALENTOS\n" +
        "Programa Governo Eficaz - Santana de Parnaíba\n\n" +
        "🎯 COMO USAR:\n" +
        "• Use \"Atualizar Secretaria\" para atualizações rápidas\n" +
        "• Use \"Atualizar Banco\" para sincronização completa\n" +
        "• Retire os filtros antes de clicar em Atualizar\n" +
        "• Registre movimentações na aba \"Movimentações 2025\"\n\n" +
        "📞 SUPORTE TÉCNICO:\n" +
        "📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br\n" +
        "📱 4622-7500 - 8819 / 8644 / 7574\n\n" +
        "🚀 Versão Final\n" +
        "📅 03/11/2025";
    
    SpreadsheetApp.getUi().alert("ℹ️ Sobre o Sistema", sobre, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================================================
// INSTALAR GATILHOS AUTOMÁTICOS
// ========================================================================

function instalarGatilhos() {
    try {
        // Remove gatilhos antigos
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            if (triggers[i].getHandlerFunction() === 'onEdit') {
                ScriptApp.deleteTrigger(triggers[i]);
            }
        }

        // Instala novo gatilho
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        ScriptApp.newTrigger('onEdit')
            .forSpreadsheet(ss)
            .onEdit()
            .create();

        SpreadsheetApp.getUi().alert(
            "✅ Gatilhos Instalados",
            "As funções automáticas foram ativadas com sucesso!\n\n" +
            "• Alterações no status serão sincronizadas\n" +
            "• E-mails serão enviados automaticamente",
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    } catch (erro) {
        SpreadsheetApp.getUi().alert(
            "❌ Erro ao instalar gatilhos",
            "Erro: " + erro.toString(),
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}
// ========================================================================
// FUNÇÃO ONEDIT - AUTOFILL E SINCRONIZAÇÃO DE STATUS
// ========================================================================

function onEdit(e) {
    try {
        if (!e || !e.range) return;
        
        var aba = e.range.getSheet();
        var nomeAba = aba.getName();

        // Apenas reage na aba de movimentações
        if (nomeAba !== CONFIG.ABA_MOVIMENTACOES) return;

        var linha = e.range.getRow();
        var coluna = e.range.getColumn();

        // Coluna C (Prontuário) → preenche dados automaticamente
        if (coluna === 3 && linha >= 3) {
            executarAutofill(aba, linha, e.value);
        }

       // Coluna G (Status da Movimentação) → sincroniza e notifica
          if (coluna === 7 && linha >= 3) {
              sincronizarStatus(aba, linha);
          }
    } catch (erro) {
        Logger.log("❌ Erro no onEdit: " + erro.toString());
    }
}

// ========================================================================
// AUTOFILL: PREENCHE DADOS AO DIGITAR O PRONTUÁRIO
// ========================================================================

function executarAutofill(aba, linha, prontuario) {
    try {
        var prontuarioLimpo = (prontuario || "").toString().trim();
        if (!prontuarioLimpo) return;

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var abaBT = ss.getSheetByName(CONFIG.ABA_CENTRAL);
        if (!abaBT) return;

        var totalLinhas = abaBT.getLastRow();
        if (totalLinhas <= 1) return;

        var dadosBT = abaBT.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        var encontrado = null;

        for (var i = 0; i < dadosBT.length; i++) {
            var prontuarioBT = (dadosBT[i][2] || "").toString().trim(); // Coluna C
            if (prontuarioBT === prontuarioLimpo) {
                encontrado = dadosBT[i];
                break;
            }
        }

        if (encontrado) {
            aba.getRange(linha, 1).setValue(encontrado[0]); // A: Secretaria
            aba.getRange(linha, 2).setValue(encontrado[1]); // B: Nome
            aba.getRange(linha, 4).setValue(encontrado[5]); // D: Cargo Concurso
            aba.getRange(linha, 5).setValue(encontrado[6]); // E: CC/FE
            aba.getRange(linha, 6).setValue(encontrado[10]); // F: Ação

            SpreadsheetApp.getActive().toast(
                "Preenchido: " + encontrado[1],
                "📋 Autofill",
                3
            );
        } else {
            SpreadsheetApp.getActive().toast(
                "Prontuário não encontrado",
                "⚠️",
                3
            );
        }
    } catch (erro) {
        Logger.log("❌ Erro no autofill: " + erro.toString());
    }
}

// ========================================================================
// SINCRONIZAR STATUS COM PGE E ENVIAR NOTIFICAÇÃO
// ========================================================================

function sincronizarStatus(aba, linha) {
    try {
        var secretaria = (aba.getRange(linha, 1).getValue() || "").toString().trim();     // A
        var nome = (aba.getRange(linha, 2).getValue() || "").toString().trim();           // B
        var prontuario = (aba.getRange(linha, 3).getValue() || "").toString().trim();     // C
        var statusNovo = (aba.getRange(linha, 7).getValue() || "").toString().trim(); // G

        Logger.log("🔄 Sincronizando status:");
        Logger.log("Secretaria: " + secretaria);
        Logger.log("Nome: " + nome);
        Logger.log("Prontuário: " + prontuario);
        Logger.log("Status Novo: " + statusNovo);

        if (!secretaria || !prontuario || !statusNovo) {
            Logger.log("❌ Dados incompletos - cancelando sincronização");
            return;
        }

        // Status que disparam notificação - CORRIGIDO
        var statusComEmail = ["Em Andamento", "EM MOVIMENTAÇÃO", "Concluído", "CONCLUÍDO", "Cancelado", "CANCELADO"];
        if (!statusComEmail.includes(statusNovo)) {
            Logger.log("⚠️ Status não requer notificação: " + statusNovo);
            return;
        }

        // Encontra a planilha da secretaria - MELHORADO
        var infoSecretaria = PLANILHAS_SECRETARIAS.filter(function(s) { 
            return s.nome.toUpperCase() === secretaria.toUpperCase(); 
        })[0];
        
        if (!infoSecretaria) {
            Logger.log("❌ Secretaria não encontrada no mapeamento: " + secretaria);
            SpreadsheetApp.getActive().toast(
                "Secretaria '" + secretaria + "' não encontrada no sistema",
                "⚠️ Aviso",
                5
            );
            return;
        }

        Logger.log("📋 Secretaria mapeada: " + infoSecretaria.nome + " (ID: " + infoSecretaria.id + ")");

        // Atualiza status na planilha PGE - CORRIGIDO
        var statusAtualizado = false;
        try {
            Logger.log("🔗 Abrindo planilha externa...");
            var planilhaPGE = SpreadsheetApp.openById(infoSecretaria.id);
            
            // Tenta diferentes nomes de aba - FLEXIBILIZADO
            var possiveisAbas = [
                "Banco de Talentos (Externo)",
                "Planejamento de Gestão Estratégica", 
                "PGE",
                "Dados"
            ];
            
            var abaPGE = null;
            for (var i = 0; i < possiveisAbas.length; i++) {
                try {
                    abaPGE = planilhaPGE.getSheetByName(possiveisAbas[i]);
                    if (abaPGE) {
                        Logger.log("✅ Aba encontrada: " + possiveisAbas[i]);
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!abaPGE) {
                Logger.log("❌ Nenhuma aba válida encontrada na planilha");
                SpreadsheetApp.getActive().toast(
                    "Aba não encontrada na planilha da " + secretaria,
                    "⚠️ Aviso",
                    5
                );
                // Continua para enviar e-mail mesmo sem atualizar PGE
            } else {
                // Procura o prontuário e atualiza status - MELHORADO
                var ultimaLinha = abaPGE.getLastRow();
                Logger.log("📊 Última linha da planilha PGE: " + ultimaLinha);
                
                if (ultimaLinha >= 4) {
                    // Busca na coluna C (prontuário) a partir da linha 4
                    var dadosProntuarios = abaPGE.getRange(4, 3, ultimaLinha - 3, 1).getValues();
                    
                    for (var j = 0; j < dadosProntuarios.length; j++) {
                        var prontuarioPGE = (dadosProntuarios[j][0] || "").toString().trim();
                        
                        if (prontuarioPGE === prontuario) {
                            var linhaAtualizar = j + 4;
                            
                            // Verifica se existe coluna de status (tenta várias posições)
                            var colunasStatus = [14, 15, 16, 17]; // N, O, P, Q
                            var colunaStatusEncontrada = false;
                            
                            for (var k = 0; k < colunasStatus.length; k++) {
                                try {
                                    var cabecalho = abaPGE.getRange(1, colunasStatus[k]).getValue();
                                    if (cabecalho && cabecalho.toString().toLowerCase().includes("status")) {
                                        abaPGE.getRange(linhaAtualizar, colunasStatus[k]).setValue(statusNovo);
                                        Logger.log("✅ Status atualizado na coluna " + colunasStatus[k] + " (linha " + linhaAtualizar + ")");
                                        statusAtualizado = true;
                                        colunaStatusEncontrada = true;
                                        break;
                                    }
                                } catch (e) {
                                    continue;
                                }
                            }
                            
                            if (!colunaStatusEncontrada) {
                                // Se não encontrou coluna de status, usa coluna Q (17) como padrão
                                abaPGE.getRange(linhaAtualizar, 17).setValue(statusNovo);
                                Logger.log("✅ Status atualizado na coluna padrão Q (linha " + linhaAtualizar + ")");
                                statusAtualizado = true;
                            }
                            break;
                        }
                    }
                    
                    if (!statusAtualizado) {
                        Logger.log("❌ Prontuário " + prontuario + " não encontrado na planilha PGE");
                    }
                }
            }
        } catch (erroPGE) {
            Logger.log("❌ Erro ao atualizar PGE: " + erroPGE.toString());
            SpreadsheetApp.getActive().toast(
                "Erro ao acessar planilha da " + secretaria,
                "⚠️ Aviso",
                5
            );
        }

        // Envia e-mail de notificação - SEMPRE TENTA
        Logger.log("📧 Iniciando envio de e-mail...");
        enviarEmailNotificacao(secretaria, nome, prontuario, statusNovo);

        // Feedback para o usuário - MELHORADO
        var mensagem = statusAtualizado ? 
            "Status \"" + statusNovo + "\" sincronizado com PGE" :
            "E-mail enviado (PGE não atualizado)";
            
        SpreadsheetApp.getActive().toast(
            mensagem,
            "🔄 Sincronização",
            4
        );

    } catch (erro) {
        Logger.log("💥 Erro em sincronizarStatus: " + erro.toString());
        SpreadsheetApp.getActive().toast(
            "Erro na sincronização: " + erro.message,
            "❌ Erro",
            5
        );
    }
}
// ========================================================================
// ENVIO DE E-MAIL PARA SECRETÁRIOS E PONTOS FOCAIS
// ========================================================================

function enviarEmailNotificacao(secretaria, nome, prontuario, status) {
    try {
        Logger.log("📨 Preparando envio de e-mail para: " + secretaria);
        
        // Normaliza nome da secretaria para busca
        var secretariaNormalizada = secretaria.toUpperCase().trim();
        
        // Busca e-mails
        var emailSecretario = null;
        var emailPontoFocal = null;
        
        // Procura por correspondência exata ou parcial
        for (var chave in EMAILS_SECRETARIOS) {
            if (chave.toUpperCase() === secretariaNormalizada || 
                secretariaNormalizada.includes(chave.toUpperCase()) ||
                chave.toUpperCase().includes(secretariaNormalizada)) {
                emailSecretario = EMAILS_SECRETARIOS[chave];
                break;
            }
        }
        
        for (var chave in EMAILS_PONTOS_FOCAIS) {
            if (chave.toUpperCase() === secretariaNormalizada || 
                secretariaNormalizada.includes(chave.toUpperCase()) ||
                chave.toUpperCase().includes(secretariaNormalizada)) {
                emailPontoFocal = EMAILS_PONTOS_FOCAIS[chave];
                break;
            }
        }

        Logger.log("📧 E-mail secretário encontrado: " + (emailSecretario || "NÃO"));
        Logger.log("📧 E-mail ponto focal encontrado: " + (emailPontoFocal ? emailPontoFocal.length + " endereço(s)" : "NÃO"));

        if (!emailSecretario && !emailPontoFocal) {
            Logger.log("⚠️ Nenhum e-mail encontrado para: " + secretaria);
            SpreadsheetApp.getActive().toast(
                "E-mails não cadastrados para " + secretaria,
                "⚠️ Aviso",
                5
            );
            return;
        }

        var destinatarios = [];
        if (emailSecretario) destinatarios.push(emailSecretario);
        if (emailPontoFocal) destinatarios = destinatarios.concat(emailPontoFocal);

        var assunto = "Banco de Talentos - Movimentação de Servidor (" + status + ")";
        var corpo = criarCorpoEmailNotificacao(secretaria, nome, prontuario, status);

        var emailsEnviados = 0;
        var emailsFalharam = 0;

        for (var i = 0; i < destinatarios.length; i++) {
            try {
                Logger.log("📤 Enviando para: " + destinatarios[i]);
                
                MailApp.sendEmail({
                    to: destinatarios[i],
                    subject: assunto,
                    htmlBody: corpo
                });
                
                Logger.log("✅ E-mail enviado para " + destinatarios[i]);
                emailsEnviados++;
                
                // Pequena pausa entre envios
                Utilities.sleep(500);
                
            } catch (erroEmail) {
                Logger.log("❌ Falha ao enviar para " + destinatarios[i] + ": " + erroEmail.toString());
                emailsFalharam++;
            }
        }

        // Feedback final
        var mensagemFinal = "";
        if (emailsEnviados > 0) {
            mensagemFinal = "E-mail enviado para " + emailsEnviados + " destinatário(s)";
        }
        if (emailsFalharam > 0) {
            mensagemFinal += (mensagemFinal ? " (" + emailsFalharam + " falharam)" : emailsFalharam + " e-mails falharam");
        }

        SpreadsheetApp.getActive().toast(
            mensagemFinal || "Processo concluído",
            "📧 Notificação",
            4
        );

    } catch (erro) {
        Logger.log("💥 Erro ao enviar e-mail: " + erro.toString());
        SpreadsheetApp.getActive().toast(
            "Erro no envio de e-mail: " + erro.message,
            "❌ Erro",
            5
        );
    }
}

// ========================================================================
// FUNÇÃO DE TESTE PARA DEBUG
// ========================================================================

function testarSincronizacao() {
    // Use esta função para testar manualmente
    // Substitua os valores pelos dados do seu teste
    
    var secretaria = "SMA";
    var nome = "TESTE A";
    var prontuario = "123321";
    var status = "EM MOVIMENTAÇÃO";
    
    Logger.log("🧪 === TESTE MANUAL DE SINCRONIZAÇÃO ===");
    Logger.log("Dados de teste: " + secretaria + " | " + nome + " | " + prontuario + " | " + status);
    
    // Simula a sincronização
    enviarEmailNotificacao(secretaria, nome, prontuario, status);
    
    Logger.log("🧪 === FIM DO TESTE ===");
}

// ========================================================================
// FUNÇÃO PARA VERIFICAR ESTRUTURA DA PLANILHA (DEBUG)
// ========================================================================

function verificarEstruturaPGE() {
    try {
        // ID da SMA
        var idSMA = "1Nc9O1Ha038gKY5LcfxUVClhTq6rsR0zghdSJqfScI6k";
        
        var planilha = SpreadsheetApp.openById(idSMA);
        var aba = planilha.getSheetByName("Planejamento de Gestão Estratégica");
        
        if (!aba) {
            Logger.log("❌ Aba não encontrada!");
            return;
        }
        
        Logger.log("📋 === ESTRUTURA DA PLANILHA PGE-SMA ===");
        
        // Verifica cabeçalhos (linha 1)
        var ultimaColuna = aba.getLastColumn();
        var cabecalhos = aba.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        
        Logger.log("🏷️ CABEÇALHOS:");
        for (var i = 0; i < cabecalhos.length; i++) {
            var letra = String.fromCharCode(65 + i); // A, B, C, etc.
            Logger.log(letra + " (índice " + (i + 1) + "): " + cabecalhos[i]);
        }
        
        // Verifica dados de teste (a partir da linha 4)
        var ultimaLinha = aba.getLastRow();
        Logger.log("📊 Última linha com dados: " + ultimaLinha);
        
        if (ultimaLinha >= 4) {
            Logger.log("🔍 DADOS LINHA 4 (primeira linha de dados):");
            var dadosLinha4 = aba.getRange(4, 1, 1, ultimaColuna).getValues()[0];
            
            for (var j = 0; j < Math.min(dadosLinha4.length, 18); j++) { // Até coluna R
                var letra = String.fromCharCode(65 + j);
                Logger.log(letra + ": '" + dadosLinha4[j] + "'");
            }
            
            // Foca no prontuário (coluna C)
            var prontuario = dadosLinha4[2]; // índice 2 = coluna C
            Logger.log("👤 Prontuário encontrado: '" + prontuario + "'");
            
            // Foca no status atual (coluna Q)
            if (dadosLinha4.length > 16) {
                var statusAtual = dadosLinha4[16]; // índice 16 = coluna Q
                Logger.log("📊 Status atual: '" + statusAtual + "'");
            }
        }
        
        Logger.log("📋 === FIM DA VERIFICAÇÃO ===");
        
    } catch (erro) {
        Logger.log("❌ Erro ao verificar estrutura: " + erro.toString());
    }
}
// ========================================================================
// CORPO DO E-MAIL DE NOTIFICAÇÃO
// ========================================================================

function criarCorpoEmailNotificacao(secretaria, nome, prontuario, status) {
    var corStatus = "#333";
    if (status === "Em Andamento") corStatus = "#d9534f";   // vermelho
    if (status === "Concluído") corStatus = "#5cb85c";      // verde
    if (status === "Cancelado") corStatus = "#f0ad4e";      // laranja

    return `
    <div style="max-width:600px; margin:auto; font-family: Calibri, Arial, sans-serif; line-height: 1.6; color: #333; font-size:14px;">
        <h2 style="color: #1f4e79; text-align:center;">🏛️ Banco de Talentos</h2>
        <p><strong>Prezado(a) Gestor(a),</strong></p>
        <p>Informamos que um servidor vinculado à <strong>${secretaria}</strong> teve seu status atualizado.</p>

        <div style="background-color: #f8f9fa; padding: 15px 20px; border-left: 4px solid #1f4e79; border-radius: 6px; margin: 20px 0;">
            <h3 style="margin-top:0; color: #1f4e79;">📌 Detalhes:</h3>
            <p>📛 <strong>Nome:</strong> ${nome}</p>
            <p>🆔 <strong>Prontuário:</strong> ${prontuario}</p>
            <p>⚡ <strong>Status:</strong> <span style="color:${corStatus}; font-weight:bold;">${status}</span></p>
            <p>🏢 <strong>Secretaria:</strong> ${secretaria}</p>
        </div>

        <p>Este é um aviso automático do Programa Governo Eficaz.</p>
        <p>📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br | 📱 4622-7500</p>

        <p style="font-size:11px; color:#666; text-align:center;">
            <em>Sistema automatizado - não responda este e-mail</em>
        </p>
    </div>`;
}

// ========================================================================
// FUNÇÃO PRINCIPAL: ATUALIZAR BANCO COMPLETO
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

        // Prepara aba central
        var planilhaCentral = prepararPlanilhaCentral();
        var abaCentral = planilhaCentral.abaCentral;

        // Ordena secretarias alfabeticamente
        var secretariasOrdenadas = PLANILHAS_SECRETARIAS.slice().sort(function(a, b) {
            return a.nome.localeCompare(b.nome);
        });

        var todosOsDados = [];
        var lotes = criarLotes(secretariasOrdenadas, CONFIG.LOTE_SIZE);

        for (var i = 0; i < lotes.length; i++) {
            var lote = lotes[i];
            var numeroLote = i + 1;
            var totalLotes = lotes.length;

            SpreadsheetApp.getActive().toast(
                "Processando lote " + numeroLote + "/" + totalLotes + "...",
                "🔄 Atualizando Banco",
                5
            );

            var resultadoLote = processarLoteSecretarias(lote, numeroLote);
            todosOsDados = todosOsDados.concat(resultadoLote.dados);

            relatorio.secretariasProcessadas += resultadoLote.processadas;
            relatorio.erros = relatorio.erros.concat(resultadoLote.erros);

            if (i < lotes.length - 1) {
                Utilities.sleep(CONFIG.DELAY_ENTRE_LOTES);
            }
        }

        // Insere todos os dados
        if (todosOsDados.length > 0) {
    Logger.log("📝 Inserindo " + todosOsDados.length + " registros...");
    var range = abaCentral.getRange(2, 1, todosOsDados.length, CABECALHOS_CENTRAL.length);
    range.setValues(todosOsDados);
    relatorio.registrosImportados = todosOsDados.length;
    reordenarDadosAlfabeticamente(abaCentral);
    aplicarFormatacaoOtimizada(abaCentral, abaCentral.getLastRow());
        }
            // Atualiza aba Movimentações 2025 com base na aba BT 2025
            atualizarMovimentacoesAutomatico();
        // Finaliza
        relatorio.fim = new Date();
        relatorio.duracao = Math.round((relatorio.fim - relatorio.inicio) / 1000);
        exibirResultadoOtimizado(relatorio);

        Logger.log("🎉 === ATUALIZAÇÃO CONCLUÍDA ===");
    } catch (erro) {
        Logger.log("💥 ERRO CRÍTICO: " + erro.toString());
        SpreadsheetApp.getUi().alert(
            "❌ Erro na Atualização",
            "Erro crítico:\n\n" + erro.toString(),
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

// ========================================================================
// ATUALIZAR UMA SECRETARIA ESPECÍFICA
// ========================================================================

function atualizarSecretariaEspecifica() {
    var secretariasOrdenadas = PLANILHAS_SECRETARIAS.slice().sort((a, b) => a.nome.localeCompare(b.nome));
    var opcoes = secretariasOrdenadas.map((s, i) => `${i + 1} - ${s.nome}`);

    var resposta = SpreadsheetApp.getUi().prompt(
        "📂 Atualizar Secretaria Específica",
        "Digite o número da secretaria (1-" + PLANILHAS_SECRETARIAS.length + "):\n\n" + opcoes.join("\n"),
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );

    if (resposta.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        var numero = parseInt(resposta.getResponseText());
        if (Number.isInteger(numero) && numero >= 1 && numero <= PLANILHAS_SECRETARIAS.length) {
            var secretaria = secretariasOrdenadas[numero - 1];
            atualizarUmaSecretaria(secretaria);
        } else {
            SpreadsheetApp.getUi().alert(
                "⚠️ Número inválido",
                "Digite um número entre 1 e " + PLANILHAS_SECRETARIAS.length
            );
        }
    }
    if (resultado.sucesso && resultado.dados.length > 0) {
  removerDadosSecretaria(abaCentral, secretaria.nome);
  var novaLinha = abaCentral.getLastRow() + 1;
  abaCentral.getRange(novaLinha, 1, resultado.dados.length, CABECALHOS_CENTRAL.length)
           .setValues(resultado.dados);
  aplicarFormatacaoOtimizada(abaCentral, novaLinha + resultado.dados.length - 1);
  reordenarDadosAlfabeticamente(abaCentral);

  // Atualiza Movimentações após atualização da secretaria
  atualizarMovimentacoesAutomatico();

  SpreadsheetApp.getUi().alert(
    "✅ Sucesso",
    secretaria.nome + ": " + resultado.dados.length + " registros atualizados.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
}

function atualizarUmaSecretaria(secretaria) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var abaCentral = ss.getSheetByName(CONFIG.ABA_CENTRAL) || prepararPlanilhaCentral().abaCentral;

        var resultado = processarSecretariaOtimizada(secretaria);

       if (resultado.sucesso && resultado.dados.length > 0) {
    removerDadosSecretaria(abaCentral, secretaria.nome);
    var novaLinha = abaCentral.getLastRow() + 1;
    abaCentral.getRange(novaLinha, 1, resultado.dados.length, CABECALHOS_CENTRAL.length)
             .setValues(resultado.dados);

    reordenarDadosAlfabeticamente(abaCentral);

    aplicarFormatacaoOtimizada(abaCentral, abaCentral.getLastRow());

    SpreadsheetApp.getUi().alert(
        "✅ Sucesso",
        secretaria.nome + ": " + resultado.dados.length + " registros atualizados.",
        SpreadsheetApp.getUi().ButtonSet.OK
    );


        } else if (resultado.dados.length === 0) {
            SpreadsheetApp.getUi().alert(
                "ℹ️ Sem dados",
                secretaria.nome + ": nenhum registro encontrado."
            );
        } else {
            SpreadsheetApp.getUi().alert(
                "❌ Erro",
                "Falha ao processar " + secretaria.nome + ":\n" + resultado.erro
            );
        }
    } catch (erro) {
        SpreadsheetApp.getUi().alert(
            "❌ Erro Crítico",
            "Erro inesperado: " + erro.toString()
        );
    }
}

// ========================================================================
// REMOVER DADOS ANTIGOS DE UMA SECRETARIA
// ========================================================================

function removerDadosSecretaria(abaCentral, nomeSecretaria) {
    try {
        var totalLinhas = abaCentral.getLastRow();
        if (totalLinhas <= 1) return;

        var dados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        var dadosFiltrados = dados.filter(function(linha) {
            return (linha[0] || "").toString().trim() !== nomeSecretaria;
        });

        if (totalLinhas > 1) {
            abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).clearContent();
        }

        if (dadosFiltrados.length > 0) {
            abaCentral.getRange(2, 1, dadosFiltrados.length, CABECALHOS_CENTRAL.length)
                     .setValues(dadosFiltrados);
        }
    } catch (erro) {
        Logger.log("⚠️ Erro ao remover dados: " + erro.toString());
    }
}

// ========================================================================
// REORDENAR DADOS POR SECRETARIA (A-Z)
// ========================================================================

function reordenarDadosAlfabeticamente(abaCentral) {
    try {
        var totalLinhas = abaCentral.getLastRow();
        if (totalLinhas <= 2) return;

        var dados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length).getValues();
        dados.sort(function(a, b) {
            var sA = (a[0] || "").toString().toUpperCase();
            var sB = (b[0] || "").toString().toUpperCase();
            return sA.localeCompare(sB);
        });
        abaCentral.getRange(2, 1, dados.length, CABECALHOS_CENTRAL.length).setValues(dados);
    } catch (erro) {
        Logger.log("⚠️ Erro ao reordenar: " + erro.toString());
    }
}

// ========================================================================
// PROCESSAR DADOS DE UMA SECRETARIA EXTERNA
// ========================================================================

function processarSecretariaOtimizada(secretaria) {
    try {
        var sigla = secretaria.nome;
        Logger.log("📂 Processando: " + sigla);

        var planilha = SpreadsheetApp.openById(secretaria.id);
        var abaOrigem = planilha.getSheetByName(CONFIG.ABA_ORIGEM);
        if (!abaOrigem) {
            return { sucesso: false, erro: "Aba não encontrada", siglaSecretaria: sigla };
        }

        var ultimaLinha = abaOrigem.getLastRow();
        if (ultimaLinha <= CONFIG.LINHA_INICIO_DADOS) {
            return { sucesso: true, dados: [], siglaSecretaria: sigla };
        }

        var dadosBrutos = abaOrigem.getRange(
            CONFIG.LINHA_INICIO_DADOS + 1,
            2, // Coluna B
            ultimaLinha - CONFIG.LINHA_INICIO_DADOS,
            17 // Colunas B até R
        ).getValues();

        var dadosProcessados = [];
        for (var i = 0; i < dadosBrutos.length; i++) {
            var linha = dadosBrutos[i];
            if ((linha[0] || "").toString().trim()) {
                dadosProcessados.push([
                    sigla,                                  // A
                    linha[0],                               // B: Nome
                    linha[1],                               // C: Prontuário
                    linha[2],                               // D: Formação
                    linha[3],                               // E: Área
                    linha[4],                               // F: Cargo
                    linha[5],                               // G: CC/FE
                    linha[6],                               // H: Função Gratificada
                    linha[7],                               // I: Readaptado
                    linha[8],                               // J: Justificativa
                    linha[9],                               // K: Ação
                    linha[10],                              // L: Condicionalidade
                    formatarDataBrasileira(linha[11]),      // M: Data
                    (linha[15] || "").toString().trim(),   // N: Status (coluna Q)
                    (linha[16] || "").toString().trim()    // O: Interesse (coluna R)
                ]);
            }
        }

        return { sucesso: true, dados: dadosProcessados, siglaSecretaria: sigla };
    } catch (erro) {
        return { sucesso: false, erro: erro.toString(), siglaSecretaria: secretaria.nome };
    }
}

// ========================================================================
// PREPARAR ABA CENTRAL (cria, limpa, cabeçalho)
// ========================================================================

function prepararPlanilhaCentral() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(CONFIG.ABA_CENTRAL);

    if (!aba) {
        aba = ss.insertSheet(CONFIG.ABA_CENTRAL);
    }

    aba.clear();
    var cabecalho = aba.getRange(1, 1, 1, CABECALHOS_CENTRAL.length);
    cabecalho.setValues([CABECALHOS_CENTRAL]);
    cabecalho.setBackground("#1f4e79")
           .setFontColor("#ffffff")
           .setFontWeight("bold")
           .setFontFamily("Calibri")
           .setFontSize(10)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");

    aba.setFrozenRows(1);
    return { abaCentral: aba };
}

// ========================================================================
// CRIAR LOTES PARA PROCESSAMENTO SEQUENCIAL
// ========================================================================

function criarLotes(planilhas, tamanho) {
    var lotes = [];
    for (var i = 0; i < planilhas.length; i += tamanho) {
        lotes.push(planilhas.slice(i, i + tamanho));
    }
    return lotes;
}

// ========================================================================
// PROCESSAR UM LOTE DE SECRETARIAS
// ========================================================================

function processarLoteSecretarias(lote, numeroLote) {
    var dadosLote = [];
    var errosLote = [];
    var processadas = 0;

    for (var i = 0; i < lote.length; i++) {
        var resultado = processarSecretariaOtimizada(lote[i]);
        if (resultado.sucesso) {
            dadosLote = dadosLote.concat(resultado.dados);
            processadas++;
        } else {
            errosLote.push(lote[i].nome + ": " + resultado.erro);
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

// ========================================================================
// FORMATAR DATA PARA PADRÃO BRASILEIRO (dd/mm/yyyy)
// ========================================================================

function formatarDataBrasileira(data) {
    if (!data || typeof data === 'string') return "";

    if (data instanceof Date && !isNaN(data.getTime())) {
        var dia = String(data.getDate()).padStart(2, '0');
        var mes = String(data.getMonth() + 1).padStart(2, '0');
        var ano = data.getFullYear();
        return dia + "/" + mes + "/" + ano;
    }
    return "";
}

// ========================================================================
// APLICAR FORMATAÇÃO VISUAL NA ABA CENTRAL
// ========================================================================

function aplicarFormatacaoOtimizada(abaCentral, totalLinhas) {
    if (totalLinhas <= 1) return;

    var larguras = [66, 257, 68, 108, 103, 168, 97, 88, 76, 256, 125, 170, 88, 120, 150];
    for (var i = 0; i < larguras.length; i++) {
        abaCentral.setColumnWidth(i + 1, larguras[i]);
    }

    var rangeDados = abaCentral.getRange(2, 1, totalLinhas - 1, CABECALHOS_CENTRAL.length);
    rangeDados.setFontFamily("Calibri")
               .setFontSize(10)
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .setWrap(true);

    // Coluna M (Data da Inclusão)
    if (totalLinhas > 1) {
        abaCentral.getRange(2, 13, totalLinhas - 1, 1).setNumberFormat("dd/mm/yyyy");
    }

    // Cores alternadas
    for (var i = 2; i <= totalLinhas; i++) {
        abaCentral.getRange(i, 1, 1, CABECALHOS_CENTRAL.length)
                 .setBackground(i % 2 === 0 ? "#ffffff" : "#f8f9fa");
    }

    // Destaque na coluna Secretaria
    abaCentral.getRange(2, 1, totalLinhas - 1, 1)
             .setBackground("#e3f2fd")
             .setFontWeight("bold");

    // Bordas
    abaCentral.getRange(1, 1, totalLinhas, CABECALHOS_CENTRAL.length)
             .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
}

// ========================================================================
// EXIBIR RESULTADO DA ATUALIZAÇÃO
// ========================================================================

function exibirResultadoOtimizado(relatorio) {
    var porcentagem = Math.round((relatorio.secretariasProcessadas / PLANILHAS_SECRETARIAS.length) * 100);
    var mensagem = 
        "🎉 ATUALIZAÇÃO CONCLUÍDA!\n\n" +
        "📊 RESULTADOS:\n" +
        "• Secretarias: " + relatorio.secretariasProcessadas + "/" + PLANILHAS_SECRETARIAS.length + " (" + porcentagem + "%)\n" +
        "• Registros: " + relatorio.registrosImportados + "\n" +
        "• Tempo: " + relatorio.duracao + " segundos\n\n" +
        (relatorio.erros.length > 0 ? 
            "⚠️ Erros: " + relatorio.erros.length + "\nVerifique os logs." : 
            "✅ Tudo certo!") + 
        "\n\n📅 " + relatorio.fim.toLocaleString('pt-BR');

    SpreadsheetApp.getUi().alert("✅ Sucesso!", mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
}
// ========================================================================
// PREENCHER MOVIMENTAÇÕES AUTOMATICAMENTE
// ========================================================================

function atualizarMovimentacoesAutomatico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaBT = ss.getSheetByName(CONFIG.ABA_CENTRAL); // "BT 2025"
  var abaMov = ss.getSheetByName(CONFIG.ABA_MOVIMENTACOES); // "Movimentações 2025"

  if (!abaBT || !abaMov) {
    Logger.log("❌ Abas não encontradas!");
    return;
  }

  // Pega todos os dados da aba BT 2025
  var totalLinhasBT = abaBT.getLastRow();
  if (totalLinhasBT <= 1) return;
  var dadosBT = abaBT.getRange(2, 1, totalLinhasBT - 1, CABECALHOS_CENTRAL.length).getValues();

  // Pega prontuários já existentes na aba Movimentações 2025
  var totalLinhasMov = abaMov.getLastRow();
  var prontuariosExistentes = [];
  if (totalLinhasMov > 2) {
    prontuariosExistentes = abaMov.getRange(3, 3, totalLinhasMov - 2, 1).getValues()
      .map(function(r) { return (r[0] || "").toString().trim(); });
  }

  var novosRegistros = [];

  for (var i = 0; i < dadosBT.length; i++) {
    var linha = dadosBT[i];
    var secretaria = (linha[0] || "").toString().trim();
    var nome = (linha[1] || "").toString().trim();
    var prontuario = (linha[2] || "").toString().trim();
    var cargo = (linha[5] || "").toString().trim();
    var ccfe = (linha[6] || "").toString().trim(); // Coluna G no BT 2025
    var acao = (linha[10] || "").toString().trim();
    var status = (linha[13] || "").toString().trim();
    
    // Condição: só entra se ação for LIBERAÇÃO IMEDIATA e prontuário ainda não existir
    if (acao.toUpperCase() === "LIBERAÇÃO IMEDIATA" && prontuario && !prontuariosExistentes.includes(prontuario)) {
      novosRegistros.push([secretaria, nome, prontuario, cargo, ccfe, acao, status || ""]);
    }
  }

  // Insere os novos registros
  if (novosRegistros.length > 0) {
    var primeiraLinhaLivre = abaMov.getLastRow() + 1;
    abaMov.getRange(primeiraLinhaLivre, 1, novosRegistros.length, novosRegistros[0].length).setValues(novosRegistros);

    SpreadsheetApp.getActive().toast(
      novosRegistros.length + " registros adicionados em Movimentações 2025",
      "📥 Atualização Automática",
      5
    );
  } else {
    Logger.log("⚠️ Nenhum registro novo encontrado.");
  }
}
// ========================================================================
// GERAR CIÊNCIA / ENCAMINHAMENTO - INTEGRAÇÃO COM "Movimentações 2025"
// ========================================================================

// IDs e modelos fornecidos (ajuste se necessário)
var MODELO_CIENCIA_ID = "15L8QmJn56zrkmAvuSvZ2F_LAPMjcHwvICw4X9gJUfps"; // ID do modelo Ciência (do link que você mandou)
var MODELO_ENCAMINHAMENTO_ID = "1KSneV2-clDw67Qx67_Y12-RBuaEZWp-f6ShBdNXDRec"; // ID do modelo Encaminhamento (link que você mandou)
var PASTA_CIENCIA_ID = "1ZMjslt15pHmkkHHRTdqKqeQt3ZLKSfO3"; // pasta para salvar Ciências
var PASTA_ENCAMINHAMENTO_ID = "1o45nMcyPDDqB009s_VcvFrTpJShiaAub"; // pasta para salvar Encaminhamentos

// Planilha de numeração (para pegar próximo número e pintar de amarelo)
var PLANILHA_NUMERACAO_ENCAMINHAMENTO_ID = "1vdQa93PB1CyZP0PSAN9AHc5ukmc6L09WUP2AnLagaUE"; // ID que você passou

// Mapeamento secretários (tabela que você forneceu)
var SECRETARIOS = {
  "SECOM": "Márcio Augusto Rossone",
  "SEMEDES": "João Marcos Dolabani Port",
  "SEMOP": "Wellison Ivanildo Da Silva",
  "SEMUTTRANS": "Moises Arruda",
  "SMA": "José Roberto Martins",
  "SMAFEL": "Ricardo Souza Paixão",
  "SMCC": "Helio De Souza Silva",
  "SMCL": "Clueusa Carvalho",
  "SMCT": "Valmir Baptista Damas",
  "SMDS": "Marcos Pestana Corrêa",
  "SME": "Denise Marques Da Silva",
  "SMF": "Vaumil Antonio Pontes",
  "SMGAED": "Maurício Ribeiro Nunes",
  "SMH": "Diego Oliveira Dias",
  "SMMAP": "Veruska Ticiana Franklin De Carvalho",
  "SMMF": "Selma Oliveira Cezar",
  "SMNJ": "Veronica Mutti Calderaro Teixeira Koishi",
  "SMOP": "Vivian Cristina Matiassi Do Carmo",
  "SMOU": "Wilson Felipe Dorácio",
  "SMS": "Maria Silvia De Almeida Mello",
  "SMSD": "Ricardo Cordeiro Branco De Souza",
  "SMSM": "Mário Cesar Da Silva",
  "SMSU": "Eduardo Espósito"
};

// UTIL: retorna "dd", "mesExtenso", "ano" para uma Date
function partesDataBR(data) {
  var dataCorrigida = new Date(Utilities.formatDate(data, "GMT-3", "yyyy/MM/dd HH:mm:ss"));
  var d = dataCorrigida.getDate().toString().padStart(2, "0");
  var meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  var m = meses[dataCorrigida.getMonth()];
  var a = dataCorrigida.getFullYear().toString();
  return { dia: d, mes: m, ano: a };
}

// UTIL: pega próximo número em branco numa planilha de numeração e pinta de amarelo (retorna string)
function pegarProximoNumeroMemoEmPlanilha(planilhaId, nomeAba) {
  try {
    var planilha = SpreadsheetApp.openById(planilhaId);
    var sheet = planilha.getSheetByName(nomeAba || sheet.getSheets()[0].getName());
  } catch (e) {
    // tenta abrir sem nome de aba — busca primeira
    var planilha = SpreadsheetApp.openById(planilhaId);
    var sheet = planilha.getSheets()[0];
  }
  if (!sheet) throw new Error("Aba de numeração não encontrada.");

  var range = sheet.getDataRange();
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var numero = values[i][j];
      var cor = (backgrounds[i] && backgrounds[i][j]) ? backgrounds[i][j] : "";
      if (numero && numero.toString().trim() !== "") {
        var celulaSemCor = (
          cor === "" ||
          cor === "#ffffff" ||
          cor === "#FFFFFF" ||
          cor === null ||
          cor === undefined ||
          (typeof cor === "string" && cor.toLowerCase() === "#ffffff")
        );
        if (celulaSemCor) {
          var cell = sheet.getRange(i + 1, j + 1);
          cell.setBackground("#FFFF00");
          return numero.toString().trim();
        }
      }
    }
  }
  throw new Error("Nenhum número disponível (todos usados).");
}

// UTIL: substituir tokens simples no documento (texto puro)
function substituirTokensNoDoc(doc, mapaTokens) {
  var body = doc.getBody();
  for (var token in mapaTokens) {
    var valor = mapaTokens[token] !== undefined && mapaTokens[token] !== null ? mapaTokens[token] : "";
    body.replaceText(token, valor);
  }
  doc.saveAndClose();
}

// UTIL: cria documento a partir de template, aplica substituições, move para pasta e retorna URL e fileId
function criarDocumentoPorTemplate(templateId, nomeArquivo, pastaId, mapaTokens) {
  var copia = DriveApp.getFileById(templateId).makeCopy(nomeArquivo);
  if (pastaId) {
    var pasta = DriveApp.getFolderById(pastaId);
    pasta.addFile(copia);
    // remover da pasta raiz para não duplicar visualmente (opcional)
    DriveApp.getRootFolder().removeFile(copia);
  }
  var doc = DocumentApp.openById(copia.getId());
  substituirTokensNoDoc(doc, mapaTokens);
  return { url: doc.getUrl(), id: copia.getId() };
}
function abrirDocumento(url) {
  var html = HtmlService.createHtmlOutput(`
    <script>
      window.open('${url}', '_blank');
      google.script.host.close();
    </script>
  `);
  SpreadsheetApp.getUi().showModalDialog(html, "Documento Gerado com Sucesso!");
}
// ADICIONA LINK NA ABA "Controle de Memos"
// A coluna layout assumido na aba Controle de Memos (C..I = col 3..9):
// 3 = Encaminhamento (hyperlink), 4 = Ciência (hyperlink), 5 = Data, 6 = Secretaria, 7 = Cargo, 8 = Observações, 9 = Processo
function adicionarLinkControleMemos(tipo, url, dadosMov) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaControle = ss.getSheetByName("Controle de Memos");
  if (!abaControle) {
    SpreadsheetApp.getUi().alert("Aba 'Controle de Memos' não encontrada.");
    return;
  }

  // tenta achar linha já existente para MESMA PESSOA (nome + processo)
  var ultimaLinha = abaControle.getLastRow();
  var linhaDestino = null;

  if (ultimaLinha >= 3) {
    var dadosControle = abaControle.getRange(3, 4, ultimaLinha - 2, 6).getValues(); // C..I = 3..9
    for (var i = 0; i < dadosControle.length; i++) {
      var linha = i + 3;
      var nomeLinkEnc = abaControle.getRange(linha, 3).getFormula() || "";
      var nomeLinkCie = abaControle.getRange(linha, 4).getFormula() || "";
      var processoCel = (abaControle.getRange(linha, 9).getValue() || "").toString().trim();
      var secretariaCel = (abaControle.getRange(linha, 6).getValue() || "").toString().trim();
      var nomeCel = (abaControle.getRange(linha, 7).getValue() || "").toString().trim(); // cargo coluna 7
      if (
        processoCel === dadosMov.processo ||
        nomeCel.includes(dadosMov.cargoEscolhido) ||
        secretariaCel === dadosMov.secretariaDestino
      ) {
        linhaDestino = linha;
        break;
      }
    }
  }

  // se não encontrou, cria nova linha
  if (!linhaDestino) linhaDestino = abaControle.getLastRow() + 1;

  // criar hiperlink sem erro de aspas
  var linkFormula = '=HYPERLINK("' + url.replace(/"/g, '""') + '";"' +
                    (tipo === "encaminhamento" ? "Encaminhamento" : "Ciência") + '")';

  if (tipo === "encaminhamento") {
    abaControle.getRange(linhaDestino, 3).setFormula(linkFormula);
  } else if (tipo === "ciencia") {
    abaControle.getRange(linhaDestino, 4).setFormula(linkFormula);
  }

  // Data
  abaControle.getRange(linhaDestino, 5).setValue(dadosMov.dataMovimentacao || "");
  // Secretaria
  abaControle.getRange(linhaDestino, 6).setValue(dadosMov.secretariaDestino || "");
  // Cargo
  abaControle.getRange(linhaDestino, 7).setValue(dadosMov.cargoEscolhido || "");
  // Observações (deixa em branco — não coloca o nome do doc aqui)
  abaControle.getRange(linhaDestino, 8).setValue("");
  // Processo
  abaControle.getRange(linhaDestino, 9).setValue(dadosMov.processo || "");
}

// Extrai dados da linha ativa na aba "Movimentações 2025"
function extrairDadosMovimentacao(aba, linha) {
  // Colunas importantes no seu layout:
  // A = 1 Secretaria
  // B = 2 Nome
  // C = 3 Prontuário
  // D = 4 Cargo Concurso
  // E = 5 CC / FE
  // H = 8 Secretaria Destino
  // J = 10 Data da Movimentação
  // K = 11 Processo
  var secretaria = (aba.getRange(linha, 1).getValue() || "").toString().trim();
  var nome = (aba.getRange(linha, 2).getValue() || "").toString().trim();
  var prontuario = (aba.getRange(linha, 3).getValue() || "").toString().trim();
  var cargoConcurso = (aba.getRange(linha, 4).getValue() || "").toString().trim();
  var ccfe = (aba.getRange(linha, 5).getValue() || "").toString().trim();
  var secretariaDestino = (aba.getRange(linha, 8).getValue() || "").toString().trim();
  var dataMov = aba.getRange(linha, 10).getValue();
  var processo = (aba.getRange(linha, 11).getValue() || "").toString().trim();

  // Regra condicional: se Cargo Concurso é N/A -> usar CC / FE; se Cargo Concurso preenchido, usa ele
  var cargoEscolhido = "";
  if (cargoConcurso && cargoConcurso.toString().trim().toUpperCase() !== "N/A") {
    cargoEscolhido = cargoConcurso;
  } else {
    cargoEscolhido = ccfe;
  }

  // normalizar data para dd/mm/yyyy se for Date
  var dataStr = "";
  if (dataMov instanceof Date && !isNaN(dataMov.getTime())) {
    dataStr = Utilities.formatDate(dataMov, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } else if (dataMov) {
    dataStr = dataMov.toString();
  }

  return {
    secretaria: secretaria,
    nome: nome,
    prontuario: prontuario,
    cargoConcurso: cargoConcurso,
    ccfe: ccfe,
    secretariaDestino: secretariaDestino,
    dataMovimentacao: dataStr,
    processo: processo,
    cargoEscolhido: cargoEscolhido
  };
}

// FUNÇÃO: Gerar Ciência (menu)
function gerarCienciaSelecionado() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Movimentações 2025");
    if (!aba) { SpreadsheetApp.getUi().alert("Aba 'Movimentações 2025' não encontrada."); return; }

    var linha = aba.getActiveRange().getRow();
    if (linha < 3) {
      SpreadsheetApp.getUi().alert("Selecione uma linha válida em 'Movimentações 2025' (a partir da linha 3).");
      return;
    }

    var dados = extrairDadosMovimentacao(aba, linha);
    var siglaDestino = (dados.secretariaDestino || "").toUpperCase().trim();

    var SECRETARIAS_EXTENSO = {
      "SECOM": "Secretaria Municipal de Comunicação",
      "SEMEDES": "Secretaria Municipal de Emprego, Desenvolvimento, Ciência e Tecnologia",
      "SEMOP": "Secretaria Municipal de Obras Privadas",
      "SEMUTTRANS": "Secretaria Municipal de Transporte e Trânsito",
      "SMA": "Secretaria Municipal de Administração",
      "SMAFEL": "Secretaria Municipal de Esporte e Lazer",
      "SMCC": "Secretaria Municipal da Casa Civil",
      "SMCL": "Secretaria Municipal de Compras e Licitações",
      "SMCT": "Secretaria Municipal de Cultura",
      "SMDS": "Secretaria Municipal de Desenvolvimento Social",
      "SME": "Secretaria Municipal de Educação",
      "SMF": "Secretaria Municipal de Finanças",
      "SMGAED": "Secretaria Municipal de Gestão, Assuntos Estratégicos e Desenvolvimento",
      "SMH": "Secretaria Municipal de Habitação",
      "SMMAP": "Secretaria Municipal de Meio Ambiente e Planejamento",
      "SMMF": "Secretaria Municipal da Mulher e da Família",
      "SMNJ": "Secretaria Municipal de Negócios Jurídicos",
      "SMOP": "Secretaria Municipal de Obras Públicas",
      "SMOU": "Secretaria Municipal de Operações Urbanas",
      "SMS": "Secretaria Municipal de Saúde",
      "SMSD": "Secretaria Municipal de Serviços Digitais",
      "SMSM": "Secretaria Municipal de Serviços Municipais",
      "SMSU": "Secretaria Municipal de Segurança Urbana"
    };
   var secretariaDestinoExtenso = SECRETARIAS_EXTENSO[siglaDestino] || siglaDestino;

    // 🔹 Data da movimentação (coluna J)
    var dataMov = dados.dataMovimentacao && dados.dataMovimentacao !== ""
      ? Utilities.parseDate(dados.dataMovimentacao, "GMT-3", "dd/MM/yyyy")
      : new Date();

    // 🔹 Data da movimentação formatada (para [DATAMOV])
    var dataMovFormatada = Utilities.formatDate(dataMov, "GMT-3", "dd/MM/yyyy");

    // 🔹 Data da movimentação por extenso (para [DIA], [MES], [ANO])
    var partesMov = partesDataBR(dataMov);

    // 🔹 Data de hoje (para assinatura)
    var dataHoje = new Date();
    var partesHoje = partesDataBR(new Date(Utilities.formatDate(dataHoje, "GMT-3", "yyyy/MM/dd")));

    // 🔹 Tokens para o documento
    var tokens = {
      "\\[DIA\\]": partesMov.dia,
      "\\[MES\\]": partesMov.mes,
      "\\[ANO\\]": partesMov.ano,
      "\\[DATAMOV\\]": dataMovFormatada,
      "\\[SECRETARIADESTINO\\]": secretariaDestinoExtenso,
      "\\[NOME\\]": dados.nome || "",
      "\\[DIAHOJE\\]": partesHoje.dia,
      "\\[MESHOJE\\]": partesHoje.mes,
      "\\[ANOHOJE\\]": partesHoje.ano
    };

    // 🔹 Cria o documento
    var nomeArquivo = dados.nome + " - Programa Governo Eficaz 2025";
    var resultado = criarDocumentoPorTemplate(MODELO_CIENCIA_ID, nomeArquivo, PASTA_CIENCIA_ID, tokens);

    // 🔹 Adiciona link no controle
    adicionarLinkControleMemos("ciencia", resultado.url, {
      dataMovimentacao: dados.dataMovimentacao,
      secretariaDestino: siglaDestino,
      cargoEscolhido: dados.cargoEscolhido,
      processo: dados.processo
    });

   
    abrirDocumento(resultado.url);
    SpreadsheetApp.getUi().alert("Ciência gerada com sucesso!");
  } catch (erro) {
    SpreadsheetApp.getUi().alert("Erro ao gerar Ciência: " + erro.toString());
  }
}

// FUNÇÃO: Gerar Encaminhamento (menu)
function gerarEncaminhamentoSelecionado() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Movimentações 2025");
    if (!aba) { SpreadsheetApp.getUi().alert("Aba 'Movimentações 2025' não encontrada."); return; }

    var linha = aba.getActiveRange().getRow();
    if (linha < 3) {
      SpreadsheetApp.getUi().alert("Selecione uma linha válida em 'Movimentações 2025' (a partir da linha 3).");
      return;
    }

    var dados = extrairDadosMovimentacao(aba, linha);

    // IDs e tabelas auxiliares
    var siglaDestino = (dados.secretariaDestino || "").toUpperCase().trim();
    var nomeSecretario = SECRETARIOS[siglaDestino] || "";
    var SECRETARIAS_EXTENSO = {
      "SECOM": "Secretaria Municipal de Comunicação",
      "SEMEDES": "Secretaria Municipal de Emprego, Desenvolvimento, Ciência e Tecnologia",
      "SEMOP": "Secretaria Municipal de Obras Privadas",
      "SEMUTTRANS": "Secretaria Municipal de Transporte e Trânsito",
      "SMA": "Secretaria Municipal de Administração",
      "SMAFEL": "Secretaria Municipal de Esporte e Lazer",
      "SMCC": "Secretaria Municipal da Casa Civil",
      "SMCL": "Secretaria Municipal de Compras e Licitações",
      "SMCT": "Secretaria Municipal de Cultura",
      "SMDS": "Secretaria Municipal de Desenvolvimento Social",
      "SME": "Secretaria Municipal de Educação",
      "SMF": "Secretaria Municipal de Finanças",
      "SMGAED": "Secretaria Municipal de Gestão, Assuntos Estratégicos e Desenvolvimento",
      "SMH": "Secretaria Municipal de Habitação",
      "SMMAP": "Secretaria Municipal de Meio Ambiente e Planejamento",
      "SMMF": "Secretaria Municipal da Mulher e da Família",
      "SMNJ": "Secretaria Municipal de Negócios Jurídicos",
      "SMOP": "Secretaria Municipal de Obras Públicas",
      "SMOU": "Secretaria Municipal de Operações Urbanas",
      "SMS": "Secretaria Municipal de Saúde",
      "SMSD": "Secretaria Municipal de Serviços Digitais",
      "SMSM": "Secretaria Municipal de Serviços Municipais",
      "SMSU": "Secretaria Municipal de Segurança Urbana"
    };
    var secretariaDestinoExtenso = SECRETARIAS_EXTENSO[siglaDestino] || siglaDestino;

    // Pega a data da movimentação (coluna J)
    var dataMov = dados.dataMovimentacao && dados.dataMovimentacao !== ""
      ? Utilities.parseDate(dados.dataMovimentacao, "GMT-3", "dd/MM/yyyy")
      : new Date();
    var dataHoje = new Date();

    var partesHoje = partesDataBR(new Date(Utilities.formatDate(dataHoje, "GMT-3", "yyyy/MM/dd")));
    var dataMovFormatada = Utilities.formatDate(dataMov, "GMT-3", "dd/MM/yyyy");

    // Número do memorando
    var numeroMemo = "";
    try {
      numeroMemo = pegarProximoNumeroMemoEmPlanilha(PLANILHA_NUMERACAO_ENCAMINHAMENTO_ID);
    } catch (e) {
      numeroMemo = "BACKUP-" + Math.floor(Math.random() * 9000 + 1000);
    }

    // Preenchimento dos tokens
    var tokensEnc = {
      "\\[NUMERO\\]": numeroMemo,
      "\\[DIA\\]": partesHoje.dia,
      "\\[MES\\]": partesHoje.mes,
      "\\[ANO\\]": partesHoje.ano,
      "\\[PRONTUÁRIO\\]": dados.prontuario || "",
      "\\[SECRETARIADESTINO\\]": secretariaDestinoExtenso,
      "\\[SECRETARIO\\]": nomeSecretario,
      "\\[NOME\\]": dados.nome || "",
      "\\[CARGO\\]": dados.cargoEscolhido || "",
      "\\[PROCESSO\\]": dados.processo || "",
      "\\[DATAMOV\\]": dataMovFormatada // para usar no corpo do texto
    };

    var nomeArquivo = dados.nome + " - " + dados.prontuario + " - " + dados.cargoEscolhido + " - " + (dados.secretaria || "") + " x " + (siglaDestino || "") + " - PGE 2025";
    var resultado = criarDocumentoPorTemplate(MODELO_ENCAMINHAMENTO_ID, nomeArquivo, PASTA_ENCAMINHAMENTO_ID, tokensEnc);

    adicionarLinkControleMemos("encaminhamento", resultado.url, {
      dataMovimentacao: dados.dataMovimentacao,
      secretariaDestino: siglaDestino,
      cargoEscolhido: dados.cargoEscolhido,
      processo: dados.processo
    });

    abrirDocumento(resultado.url);
    SpreadsheetApp.getUi().alert("Encaminhamento gerado com sucesso!");
  } catch (erro) {
    SpreadsheetApp.getUi().alert("Erro ao gerar Encaminhamento: " + erro.toString());
  }
}

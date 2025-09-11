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
    "SMA": "natalice.36293@santanadeparnaiba.sp.gov.br",
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
      "luana.41331@santanadeparnaiba.sp.gov.br"
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
        "• Execute \"Instalar Gatilhos\" uma vez após abrir a planilha\n" +
        "• Use \"Atualizar Secretaria\" para atualizações rápidas\n" +
        "• Use \"Atualizar Banco\" para sincronização completa\n" +
        "• Registre movimentações na aba \"Movimentações 2025\"\n\n" +
        "📞 SUPORTE TÉCNICO:\n" +
        "📧 sma.programagovernoeficaz@santanadeparnaiba.sp.gov.br\n" +
        "📱 4622-7500 - 8819 / 8644 / 7574\n\n" +
        "🚀 Versão Final - Funcionalidade Garantida\n" +
        "📅 11/09/2025";
    
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

        // Coluna F (Status da Movimentação) → sincroniza e notifica
        if (coluna === 6 && linha >= 3) {
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
            aba.getRange(linha, 5).setValue(encontrado[10]); // E: Ação

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
        var secretaria = aba.getRange(linha, 1).getValue();     // A
        var nome = aba.getRange(linha, 2).getValue();           // B
        var prontuario = aba.getRange(linha, 3).getValue();     // C
        var statusNovo = aba.getRange(linha, 6).getValue();     // F

        if (!secretaria || !prontuario || !statusNovo) return;

        // Status que disparam notificação
        var statusComEmail = ["Em Andamento", "Concluído", "Cancelado"];
        if (!statusComEmail.includes(statusNovo)) return;

        // Encontra a planilha da secretaria
        var infoSecretaria = PLANILHAS_SECRETARIAS.filter(s => s.nome === secretaria)[0];
        if (!infoSecretaria) return;

        // Atualiza status no PGE (coluna Q = índice 17)
        try {
            var planilhaPGE = SpreadsheetApp.openById(infoSecretaria.id);
            var abaPGE = planilhaPGE.getSheetByName("Planejamento de Gestão Estratégica");
            if (!abaPGE) return;

            var dadosPGE = abaPGE.getRange(4, 3, abaPGE.getLastRow() - 3, 1).getValues(); // Coluna C

            for (var i = 0; i < dadosPGE.length; i++) {
                if ((dadosPGE[i][0] || "").toString().trim() === prontuario.toString().trim()) {
                    abaPGE.getRange(i + 4, 17).setValue(statusNovo); // Coluna Q
                    break;
                }
            }
        } catch (erro) {
            Logger.log("❌ Erro ao atualizar PGE: " + erro.toString());
        }

        // Envia e-mail de notificação
        enviarEmailNotificacao(secretaria, nome, prontuario, statusNovo);

        SpreadsheetApp.getActive().toast(
            "Status \"" + statusNovo + "\" sincronizado",
            "🔄",
            3
        );
    } catch (erro) {
        Logger.log("❌ Erro em sincronizarStatus: " + erro.toString());
    }
}

// ========================================================================
// ENVIO DE E-MAIL PARA SECRETÁRIOS E PONTOS FOCAIS
// ========================================================================

function enviarEmailNotificacao(secretaria, nome, prontuario, status) {
    try {
        var emailSecretario = EMAILS_SECRETARIOS[secretaria];
        var emailPontoFocal = EMAILS_PONTOS_FOCAIS[secretaria];

        if (!emailSecretario && !emailPontoFocal) return;

        var destinatarios = [];
        if (emailSecretario) destinatarios.push(emailSecretario);
        if (emailPontoFocal) destinatarios = destinatarios.concat(emailPontoFocal);

        var assunto = "Banco de Talentos - Movimentação de Servidor (" + status + ")";
        var corpo = criarCorpoEmailNotificacao(secretaria, nome, prontuario, status);

        for (var i = 0; i < destinatarios.length; i++) {
            try {
                MailApp.sendEmail({
                    to: destinatarios[i],
                    subject: assunto,
                    htmlBody: corpo
                });
                Logger.log("📧 E-mail enviado para " + destinatarios[i]);
            } catch (erro) {
                Logger.log("❌ Falha ao enviar para " + destinatarios[i] + ": " + erro.toString());
            }
        }

        var tipoDestinatario = emailSecretario && emailPontoFocal ? "secretário e ponto focal" :
                              emailSecretario ? "secretário" : "ponto focal";

        SpreadsheetApp.getActive().toast(
            "E-mail enviado para " + tipoDestinatario,
            "📧 Notificação",
            3
        );
    } catch (erro) {
        Logger.log("❌ Erro ao enviar e-mail: " + erro.toString());
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
            aplicarFormatacaoOtimizada(abaCentral, todosOsDados.length + 1);
        }

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
            aplicarFormatacaoOtimizada(abaCentral, novaLinha + resultado.dados.length - 1);
            reordenarDadosAlfabeticamente(abaCentral);

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
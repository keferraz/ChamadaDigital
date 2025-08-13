const PLANILHA_ID = "1cXZ5Jh6stvygynOnWbzWopHjdiw3xVliqrYBjJ5H8as";
const ABA_DADOS_EQUIPE = "dadosEquipe_3";

const HISTORICO_ID = "1hSRUlLJkc8iSZc3h7Rdd2tfB4sciamImemyjEySQYBM";
const ABA_HISTORICO = "Historico_Gerado_pelo_APP";

/***** UI *****/
function doGet(e) {
  const pagina = e.parameter.page || "index";
  try {
    let template = HtmlService.createTemplateFromFile(pagina);
    return template.evaluate()
      .setTitle("Lançamento Futuros")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("Erro: arquivo '" + pagina + "' não encontrado.");
  }
}

function incluirRecursosExternos() {
  return `
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
    <script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
  `;
}

function incluirHTML(arquivo) {
  return HtmlService.createHtmlOutputFromFile(arquivo).getContent();
}

/***** Utils de data *****/
function tz() { return Session.getScriptTimeZone() || "America/Sao_Paulo"; }

function parseDDMMYYYYToDateBR(s) {
  const [dd, mm, yyyy] = s.split("/").map(n => parseInt(n, 10));
  const d = new Date(yyyy, mm - 1, dd);
  d.setHours(0,0,0,0);
  return d;
}
function hoje00() {
  const h = new Date();
  h.setHours(0,0,0,0);
  return h;
}
function isHojeOuFuturo(ddmmyyyy) {
  return parseDDMMYYYYToDateBR(ddmmyyyy) >= hoje00();
}

/***** Dados base *****/
function obterTeamLeaders() {
  const sheet = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_DADOS_EQUIPE);
  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  const teamLeaders = new Set();
  const desligadosInvalidos = new Set();

  for (let i = 1; i < dados.length; i++) {
    const teamLeader = dados[i][3];
    const colaborador = dados[i][5];
    const statusHC = dados[i][11];
    const dataDesligamento = dados[i][12];

    if (teamLeader) teamLeaders.add(teamLeader);

    if (colaborador === teamLeader && statusHC === "DESLIGADO" && dataDesligamento instanceof Date) {
      const dias = (hoje - dataDesligamento) / (1000 * 60 * 60 * 24);
      if (dias > 30) desligadosInvalidos.add(colaborador.trim().toUpperCase());
    }
  }

  return [...teamLeaders]
    .filter(nome => !desligadosInvalidos.has(nome.trim().toUpperCase()))
    .map(nome => ({ nome }));
}

/***** Carrega equipe para a tela *****/
function responderNome(nomeSelecionado) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const abaEquipe = planilha.getSheetByName(ABA_DADOS_EQUIPE);

  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);

  const dadosEquipe = abaEquipe.getDataRange().getValues();
  const dadosHistorico = abaHistorico.getDataRange().getValues();

  const hoje = new Date();
  const resultado = [];

  const statusMultiplo = [
    "Transferido","OWNboarding","Atestado","Banco-de-Horas","Afastamento","Licenca","Atestado-Acd-Trab",
    "Ferias","Treinamento-Ext","Treinamento-Int","Treinamento-REP-III","Fretado",
    "Afastamento-Acd-Trab","Sinergia SP014","Sinergia CX","Sinergia INB","Sinergia LOS",
    "Sinergia MG01","Sinergia MWH","Sinergia OUT","Sinergia QUA","Sinergia RC01","Sinergia RET",
    "Sinergia SP011","Sinergia SP02","Sinergia SP03","Sinergia SP04","Sinergia SP05","Sinergia SP06",
    "Sinergia SP09","Sinergia INV","Sinergia SVC","Sinergia RC-SP10","Sinergia Sortation","Sinergia Insumo",
    "Atividade-Externa","Suspensao","Desligado"
  ];

  for (let i = 1; i < dadosEquipe.length; i++) {
    const [areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada] = dadosEquipe[i];

    if ((teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase()) continue;

    const entrada = (dadosEquipe[i][17] || "").toUpperCase().trim(); // Coluna R
    let statusAbs = "";
    let statusAtivo = false;
    let diasRestantes = 0;
    let registradoHoje = false;
    let possuiFuturo = false;
    let statusFuturo = "";

    const hojeFmt = Utilities.formatDate(hoje, tz(), "dd/MM/yyyy");

    for (let j = 1; j < dadosHistorico.length; j++) {
      const [, , , , , , , colaboradorHist, turnoHist, , , empresaHist, processoHist,
        statusPresenca, dataHora, qtdDiasStr, dataFuturaStr] = dadosHistorico[j];

      const nomeHistNormalizado = (colaboradorHist || "").trim().toUpperCase();
      const nomeColaboradorNormalizado = (colaborador || "").trim().toUpperCase();

      if (
        nomeHistNormalizado === nomeColaboradorNormalizado &&
        turnoHist?.toString().trim().toUpperCase() === (turno || "").toUpperCase().trim() &&
        empresaHist?.toString().trim().toUpperCase() === (empresa || "").toUpperCase().trim() &&
        processoHist?.toString().trim().toUpperCase() === (processo || "").toUpperCase().trim()
      ) {
        let dataHist = dataHora instanceof Date ? dataHora : new Date(dataHora);
        const dataHistFmt = Utilities.formatDate(dataHist, tz(), "dd/MM/yyyy");

        if (dataHistFmt === hojeFmt) {
          statusAbs = statusPresenca;
          statusAtivo = statusMultiplo.includes(statusAbs);
          registradoHoje = true;

          if (statusMultiplo.includes(statusAbs)) {
            const qtdDias = parseInt(qtdDiasStr || "1", 10);
            const dataFutura = new Date(dataFuturaStr);
            diasRestantes = Math.ceil((dataFutura - hoje) / (1000 * 60 * 60 * 24));
          }
        } else if (dataHist > hoje && statusMultiplo.includes(statusPresenca)) {
          possuiFuturo = true;
          statusFuturo = statusPresenca;
        }
      }
    }

    if (!registradoHoje) {
      if (entrada === "JUSTIFICAR") statusAbs = "";
      else statusAbs = "Selecione";
    }

    // auto preencher com status futuro (exceto Banco-de-Horas)
    if (!registradoHoje && possuiFuturo && statusFuturo && statusFuturo !== "Banco-de-Horas") {
      statusAbs = statusFuturo;
    }

    resultado.push({
      AREA_HEAD: areaHead,
      IDGROOT: idgroot,
      LDAP: ldap,
      TEAM_LEADER: teamLeader,
      SUPERVISOR: supervisor,
      COLABORADOR: colaborador,
      TURNO: turno,
      ESCALA: escala,
      TURMA_ESCALA: turmaEscala,
      EMPRESA: empresa,
      PROCESSO: processo,
      STATUS_HC: statusHC,
      DATA_DEMISSAO: dataDemissao,
      STATUS_abs: statusAbs,
      BANCO_ATIVO: statusAtivo,
      DIAS_RESTANTES: diasRestantes,
      REGISTRADO_HOJE: registradoHoje,
      POSSUI_FUTURO: possuiFuturo,
      STATUS_FUTURO: statusFuturo
    });
  }
  return resultado;
}

/***** Registrar no histórico (com trava server-side para hoje/futuro) *****/
function registrarPresenca(nomeSelecionado, listaStatus, datasSelecionadas) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const abaDados = planilha.getSheetByName(ABA_DADOS_EQUIPE);

  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  if (!abaHistorico) return `Erro: Aba ${ABA_HISTORICO} não encontrada.`;

  const dados = abaDados.getDataRange().getValues();
  const registros = [];

  // TRAVA: somente hoje ou futuro
  let datas = (datasSelecionadas || []).map(s => s && s.trim()).filter(Boolean);
  const removidas = [];
  datas = datas.filter(ds => {
    const ok = isHojeOuFuturo(ds);
    if (!ok) removidas.push(ds);
    return ok;
  });
  if (datas.length === 0) {
    return "Nenhuma data válida (somente hoje ou futuras).";
  }

  const desligadosComunicados = [];
  const usuarioEmail = Session.getActiveUser().getEmail() || "indefinido";
  const hoje = new Date();

  const statusMultiplo = [
    "Transferido", "OWNboarding", "Atestado", "Banco-de-Horas", "Afastamento", "Licenca",
    "Atestado-Acd-Trab", "Ferias", "Treinamento-Ext", "Treinamento-Int", "Treinamento-REP-III",
    "Fretado", "Afastamento-Acd-Trab", "Sinergia SP014", "Sinergia CX", "Sinergia INB",
    "Sinergia LOS", "Sinergia MG01", "Sinergia MWH", "Sinergia OUT", "Sinergia QUA", "Sinergia RC01",
    "Sinergia RET", "Sinergia SP011", "Sinergia SP02", "Sinergia SP03", "Sinergia SP04", "Sinergia SP05",
    "Sinergia SP06", "Sinergia SP09", "Sinergia INV", "Sinergia SVC", "Sinergia RC-SP10",
    "Sinergia Sortation", "Sinergia Insumo", "Atividade-Externa", "Suspensao", "Desligado"
  ];

  for (let i = 1; i < dados.length; i++) {
    const [areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada] = dados[i];

    if ((teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase()) continue;

    const index = listaStatus.findIndex(item =>
      item.COLABORADOR?.toUpperCase().trim() === (colaborador || "").toUpperCase().trim() &&
      item.TURNO?.toUpperCase().trim() === (turno || "").toUpperCase().trim() &&
      item.EMPRESA?.toUpperCase().trim() === (empresa || "").toUpperCase().trim() &&
      item.PROCESSO?.toUpperCase().trim() === (processo || "").toUpperCase().trim()
    );

    if (
      index === -1 ||
      !listaStatus[index].STATUS_abs ||
      listaStatus[index].STATUS_abs === "Selecione" ||
      listaStatus[index].STATUS_abs === "Folga-Escala"
    ) continue;

    const statusPresenca = listaStatus[index].STATUS_abs;
    let qtdDias = 1;
    let motivoDesligamento = "";

    if (statusMultiplo.includes(statusPresenca)) {
      qtdDias = statusPresenca === "Desligado"
        ? 31
        : parseInt(listaStatus[index].QTD_DIAS || "1", 10);
    }

    if (statusPresenca === "Desligado") {
      motivoDesligamento = listaStatus[index].MOTIVO_DESLIGAMENTO || "";
      const dataFormatada = Utilities.formatDate(hoje, tz(), "dd/MM/yyyy");
      enviarEmailDesligamento(colaborador, dataFormatada, teamLeader, motivoDesligamento);
      desligadosComunicados.push(colaborador);
    }

    for (const dataBase of datas) {
      for (let d = 0; d < qtdDias; d++) {
        let dataStr = dataBase;
        if (qtdDias > 1) {
          const [dd, mm, yyyy] = dataBase.split("/").map(Number);
          const novaData = new Date(yyyy, mm - 1, dd);
          novaData.setDate(novaData.getDate() + d);
          dataStr = Utilities.formatDate(novaData, tz(), "dd/MM/yyyy");
        }

        registros.push([
          dataDemissao || "",                 // A
          statusHC || "",                     // B
          areaHead || "",                     // C
          supervisor || "",                   // D
          teamLeader || "",                   // E
          idgroot || "",                      // F
          ldap || "",                         // G
          colaborador || "",                  // H
          turno || "",                        // I
          escala || "",                       // J
          turmaEscala || "",                  // K
          empresa || "",                      // L
          processo || "",                     // M
          statusPresenca,                     // N
          dataStr,                            // O
          statusMultiplo.includes(statusPresenca) && statusPresenca !== "Desligado" ? (qtdDias - d - 1) : "", // P
          "",                                 // Q (pode ser usado pra log, se quiser)
          statusPresenca === "Desligado" ? motivoDesligamento : "", // R
          statusPresenca === "Desligado" ? (qtdDias - d - 1) : "",  // S
          '', '', '',                         // T,U,V (reservas)
          usuarioEmail                        // W (quem registrou)
        ]);
      }
    }
  }

  if (registros.length > 0) {
    const linhaInicial = Math.max(abaHistorico.getLastRow() + 1, 2);
    abaHistorico.getRange(linhaInicial, 1, registros.length, registros[0].length).setValues(registros);
    let mensagemFinal = `Presença registrada com sucesso para ${nomeSelecionado}.`;
    if (desligadosComunicados.length > 0) {
      mensagemFinal += `\n\n📧 E-mail(s) enviado(s) sobre o desligamento de: ${desligadosComunicados.join(", ")}`;
    }
    if (removidas && removidas.length) {
      mensagemFinal += `\n\n(Algumas datas passadas foram ignoradas: ${removidas.join(", ")})`;
    }
    return mensagemFinal;
  } else {
    return `Nenhum registro encontrado para salvar para ${nomeSelecionado}.`;
  }
}

/***** E-mail desligamento *****/
function enviarEmailDesligamento(nomeColaborador, dataAtual, teamLeader, motivoDesligamento) {
  const destinatarios = "Flow.brsp10@mercadolivre.com, acesso.brsp10@mercadolivre.com, lossprevention.brsp10@mercadolivre.com, losspreventionsp05@mercadolivre.com";
  const emcopia = "";
  const assunto = "Solicitação de desligamento de colaborador";

  const mensagem = `
    <p>Prezados,</p>
    <p>Boa tarde,</p>
    <p>Venho por meio deste solicitar o desligamento do colaborador(a) abaixo, bem como a atualização do status na (HC) de "Ativo" para "Desligado":</p>
    <p>
      📅 <b>Data da solicitação de desligamento:</b> ${dataAtual}<br>
      👤 <b>Nome completo:</b> ${nomeColaborador}<br>
      👨‍💼 <b>Team Leader:</b> ${teamLeader}<br>
      📝 <b>Motivo do desligamento:</b> ${motivoDesligamento}
    </p>
    <p>Solicito, por gentileza, que a base seja atualizada o quanto antes para mantermos os registros operacionais em conformidade.</p>
    <p>Fico à disposição para quaisquer dúvidas ou esclarecimentos adicionais.</p>
    <p>Atenciosamente,<br>${teamLeader}</p>
  `;

  MailApp.sendEmail({
    to: destinatarios,
    bcc: emcopia,
    subject: assunto,
    htmlBody: mensagem
  });
}

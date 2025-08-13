const PLANILHA_ID = "1cXZ5Jh6stvygynOnWbzWopHjdiw3xVliqrYBjJ5H8as";
const ABA_DADOS_EQUIPE = "dadosEquipe_3";
const HISTORICO_ID = "1hSRUlLJkc8iSZc3h7Rdd2tfB4sciamImemyjEySQYBM";
//const ABA_HISTORICO = "testes do app"; // <= Aba que recebe os dados dos testes realizados
const ABA_HISTORICO = "Historico_Gerado_pelo_APP"; // <= Aba que recebe os dados oficiais

function doGet(e) {
  const pagina = e.parameter.page || "index";
  try {
    const template = HtmlService.createTemplateFromFile(pagina);
    return template.evaluate()
      .setTitle("Controle ABS")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h3>Erro ao carregar página "${pagina}": ${error.message}</h3>`);
  }
}


function obterTeamLeadersPorTurno(turnoSelecionado) {
  const sheet = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_DADOS_EQUIPE);
  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  const teamLeaders = new Set();
  const desligadosInvalidos = new Set();

  for (let i = 1; i < dados.length; i++) {
    const teamLeader = dados[i][3];
    const colaborador = dados[i][5];
    const turno = dados[i][6];
    const statusHC = dados[i][11];
    const dataDesligamento = dados[i][12];

    if (teamLeader && turno === turnoSelecionado) teamLeaders.add(teamLeader);

    if (colaborador === teamLeader && statusHC === "DESLIGADO" && dataDesligamento instanceof Date) {
      const dias = (hoje - dataDesligamento) / (1000 * 60 * 60 * 24);
      if (dias > 30) desligadosInvalidos.add(colaborador.trim().toUpperCase());
    }
  }
  return [...teamLeaders]
    .filter(nome => !desligadosInvalidos.has(nome.trim().toUpperCase()))
    .map(nome => ({ nome }));
}


function responderNome(nomeSelecionado, turnoSelecionado) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const abaEquipe = planilha.getSheetByName(ABA_DADOS_EQUIPE);
  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  const dadosEquipe = abaEquipe.getDataRange().getValues();
  const dadosHistorico = abaHistorico.getDataRange().getValues();
  const hoje = new Date();
  const dataConsultaFmt = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const resultado = [];

  const statusMultiplo = [
    "Transferido","OWNboarding","Atestado","Banco-de-Horas","Afastamento","Licenca","Atestado-Acd-Trab","Ferias",
    "Treinamento-Ext","Treinamento-Int","Treinamento-REP-III","Ferias","Fretado","Afastamento","Afastamento-Acd-Trab",
    "Licenca","Sinergia SP014","Sinergia CX","Sinergia INB","Sinergia LOS","Sinergia MG01","Sinergia MWH","Sinergia OUT",
    "Sinergia QUA","Sinergia RC01","Sinergia RET","Sinergia SP011","Sinergia SP02","Sinergia SP03","Sinergia SP04",
    "Sinergia SP05","Sinergia SP06","Sinergia SP09","Sinergia INV","Sinergia RET","Sinergia QUA","Sinergia SVC",
    "Sinergia RC-SP10","Sinergia Sortation","Sinergia Insumo","Atividade-Externa","Suspensao","Desligado"
  ];

  for (let i = 1; i < dadosEquipe.length; i++) {
    const [
      areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada
    ] = dadosEquipe[i];

    if ((teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase()) continue;
    if ((turno || "").trim().toUpperCase() !== (turnoSelecionado || "").trim().toUpperCase()) continue;

    const entrada = (dadosEquipe[i][17] || "").toUpperCase().trim();
    let statusAbs = "";
    let statusAtivo = false;
    let diasRestantes = 0;
    let registradoNaData = false;

    // Busca mais inteligente: VARRE o histórico, conferindo nome, turno, empresa e processo
    for (let j = 1; j < dadosHistorico.length; j++) {
      const [
        , , , , , , , colaboradorHist, turnoHist, , , empresaHist, processoHist,
        statusPresenca, dataHora, qtdDiasStr
      ] = dadosHistorico[j];

      // Normaliza dados
      const nomeHist = (colaboradorHist || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
      const nomeColab = (colaborador || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();

      if (
        nomeHist === nomeColab &&
        (turnoHist || "").toString().trim().toUpperCase() === (turno || "").toString().trim().toUpperCase() &&
        (empresaHist || "").toString().trim().toUpperCase() === (empresa || "").toString().trim().toUpperCase() &&
        (processoHist || "").toString().trim().toUpperCase() === (processo || "").toString().trim().toUpperCase()
      ) {
        // Interpreta a data do registro do histórico
        let dataRegistro;
        if (dataHora instanceof Date) {
          dataRegistro = dataHora;
        } else if (typeof dataHora === "string" && dataHora.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
          // Está como dd/MM/yyyy
          let partes = dataHora.split("/");
          dataRegistro = new Date(`${partes[2]}-${partes[1]}-${partes[0]}T00:00:00`);
        } else {
          dataRegistro = new Date(dataHora);
        }

        let qtdDias = parseInt(qtdDiasStr || "1", 10);
        for (let d = 0; d < qtdDias; d++) {
          let dataDia = new Date(dataRegistro);
          dataDia.setDate(dataDia.getDate() + d);
          let dataDiaFmt = Utilities.formatDate(dataDia, Session.getScriptTimeZone(), "dd/MM/yyyy");

          if (dataDiaFmt === dataConsultaFmt) {
            statusAbs = statusPresenca;
            statusAtivo = statusMultiplo.includes(statusAbs);
            registradoNaData = true;
            diasRestantes = qtdDias - d - 1;
            break;
          }
        }
        if (registradoNaData) break;
      }
    }

    // Lógica padrão se não achou no histórico para a data
    if (!registradoNaData) {
      if (entrada === "JUSTIFICAR") {
        statusAbs = "";
      } else if (["OK", "NOK", "N/A"].includes(entrada)) {
        statusAbs = "Presente";
      } else if (entrada === "FOLGA-ESCALADA") {
        statusAbs = "Folga-Escala";
      } else if (entrada === "PRESENCA-HE") {
        statusAbs = "Presenca-HE";
      } else {
        statusAbs = "Selecione";
      }
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
      REGISTRADO_HOJE: registradoNaData,
      POSSUI_FUTURO: false,
      STATUS_FUTURO: ""
    });
  }
  return resultado;
}

/*
function responderNome(nomeSelecionado, turnoSelecionado) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const abaEquipe = planilha.getSheetByName(ABA_DADOS_EQUIPE);
  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  const dadosEquipe = abaEquipe.getDataRange().getValues();
  const dadosHistorico = abaHistorico.getDataRange().getValues();
  const hoje = new Date();
  const dataConsultaFmt = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const resultado = [];

  const statusMultiplo = [
    "Transferido","OWNboarding","Atestado","Banco-de-Horas","Afastamento","Licenca","Atestado-Acd-Trab","Ferias",
    "Treinamento-Ext","Treinamento-Int","Treinamento-REP-III","Ferias","Fretado","Afastamento","Afastamento-Acd-Trab",
    "Licenca","Sinergia SP014","Sinergia CX","Sinergia INB","Sinergia LOS","Sinergia MG01","Sinergia MWH","Sinergia OUT",
    "Sinergia QUA","Sinergia RC01","Sinergia RET","Sinergia SP011","Sinergia SP02","Sinergia SP03","Sinergia SP04",
    "Sinergia SP05","Sinergia SP06","Sinergia SP09","Sinergia INV","Sinergia RET","Sinergia QUA","Sinergia SVC",
    "Sinergia RC-SP10","Sinergia Sortation","Sinergia Insumo","Atividade-Externa","Suspensao","Desligado"
  ];

  // 1. Cria um dicionário rápido dos registros do histórico
  const mapaHistorico = {};
  for (let j = 1; j < dadosHistorico.length; j++) {
    const colaboradorHist = (dadosHistorico[j][7] || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase(); // COLABORADOR
    const idgrootHist = (dadosHistorico[j][5] || "").toString().trim().toUpperCase(); // IDGROOT
    // Coluna O = 14: dataHora
    const dataRegistro = (() => {
      const valor = dadosHistorico[j][14];
      if (valor instanceof Date) return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
      if (typeof valor === "string" && valor.match(/^\d{2}\/\d{2}\/\d{4}$/)) return valor;
      return "";
    })();

    const qtdDias = parseInt(dadosHistorico[j][15] || "1", 10); // QTD DIAS (coluna P)
    // Gera chave para cada dia do range de afastamento (ex: banco de horas de vários dias)
    for (let d = 0; d < qtdDias; d++) {
      let dataExpandida = "";
      if (dataRegistro) {
        const partes = dataRegistro.split("/");
        const dt = new Date(`${partes[2]}-${partes[1]}-${partes[0]}T00:00:00`);
        dt.setDate(dt.getDate() + d);
        dataExpandida = Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      const chave = `${colaboradorHist}|${idgrootHist}|${dataExpandida}`;
      mapaHistorico[chave] = {
        statusPresenca: dadosHistorico[j][13], // STATUS (coluna N)
        qtdDias,
        indexDia: d // Para calcular diasRestantes depois
      };
    }
  }

  // 2. Percorre equipe e busca pelo dicionário simplificado
  for (let i = 1; i < dadosEquipe.length; i++) {
    const [
      areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada
    ] = dadosEquipe[i];

    if ((teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase()) continue;
    if ((turno || "").trim().toUpperCase() !== (turnoSelecionado || "").trim().toUpperCase()) continue;

    const entrada = (dadosEquipe[i][17] || "").toUpperCase().trim();
    let statusAbs = "";
    let statusAtivo = false;
    let diasRestantes = 0;
    let registradoNaData = false;

    // Chave de busca rápida
    const colaboradorNorm = (colaborador || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
    const idgrootNorm = (idgroot || "").toString().trim().toUpperCase();
    const chaveBusca = `${colaboradorNorm}|${idgrootNorm}|${dataConsultaFmt}`;
    const registroHist = mapaHistorico[chaveBusca];

    if (registroHist) {
      statusAbs = registroHist.statusPresenca;
      statusAtivo = statusMultiplo.includes(statusAbs);
      registradoNaData = true;
      diasRestantes = registroHist.qtdDias - registroHist.indexDia - 1;
    }

    // Lógica padrão se não achou no histórico para a data
    if (!registradoNaData) {
      if (entrada === "JUSTIFICAR") {
        statusAbs = "";
      } else if (["OK", "NOK", "N/A"].includes(entrada)) {
        statusAbs = "Presente";
      } else if (entrada === "FOLGA-ESCALADA") {
        statusAbs = "Folga-Escala";
      } else if (entrada === "PRESENCA-HE") {
        statusAbs = "Presenca-HE";
      } else {
        statusAbs = "Selecione";
      }
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
      REGISTRADO_HOJE: registradoNaData,
      POSSUI_FUTURO: false,
      STATUS_FUTURO: ""
    });
  }
  return resultado;
}
*/

function responderNomeMulti(nomesSelecionados, turnoSelecionado) {
  if (!Array.isArray(nomesSelecionados)) return [];
  let resultadoFinal = [];
  nomesSelecionados.forEach(nome => {
    let resultadoTL = responderNome(nome, turnoSelecionado);
    if (Array.isArray(resultadoTL)) resultadoFinal = resultadoFinal.concat(resultadoTL);
  });
  return resultadoFinal;
}
function registrarPresencaMultiTL(registrosTodosTLs, versao, turnoSelecionado) {
  if (!Array.isArray(registrosTodosTLs) || !registrosTodosTLs.length) return "Nenhum registro informado.";

  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  if (!abaHistorico) return "Erro: Aba não encontrada.";

  // == INÍCIO: Lock ANTES da leitura de getLastRow() ==
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Garante exclusividade global!

    // SÓ BUSQUE getLastRow() aqui, já dentro do lock!
    const startRow = Math.max(abaHistorico.getLastRow() + 1, 2);
    abaHistorico.getRange(startRow, 1, registrosTodosTLs.length, registrosTodosTLs[0].length)
                .setValues(registrosTodosTLs);

    return `Presenças registradas com sucesso para ${registrosTodosTLs.length} colaboradores.`;
  } catch (e) {
    return "Erro ao registrar presença (lock): " + e.message;
  } finally {
    lock.releaseLock();
  }
}
/*
function registrarPresenca(nomeSelecionado, listaStatus, versao, turnoSelecionado) {
  const planilha       = SpreadsheetApp.openById(PLANILHA_ID);
  const abaDados       = planilha.getSheetByName(ABA_DADOS_EQUIPE);
  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  const dados          = abaDados.getDataRange().getValues();
  const registros      = [];
  const hoje           = new Date();
  const desligadosCom  = [];
  if (!abaHistorico) return "Erro: Aba 'Historico_3' não encontrada.";
  const usuarioEmail = Session.getActiveUser().getEmail() || "indefinido";
  const statusMultiplo = [
    "Transferido", "OWNboarding", "Atestado", "Banco-de-Horas", "Afastamento", "Licenca", "Atestado-Acd-Trab", "Ferias",
    "Treinamento-Ext", "Treinamento-Int", "Treinamento-REP-III", "Ferias", "Fretado", "Afastamento", "Afastamento-Acd-Trab",
    "Licenca", "Sinergia SP014", "Sinergia CX", "Sinergia INB", "Sinergia LOS", "Sinergia MG01", "Sinergia MWH", "Sinergia OUT",
    "Sinergia QUA", "Sinergia RC01", "Sinergia RET", "Sinergia SP011", "Sinergia SP02", "Sinergia SP03", "Sinergia SP04",
    "Sinergia SP05", "Sinergia SP06", "Sinergia SP09", "Sinergia INV", "Sinergia RET", "Sinergia QUA", "Sinergia SVC",
    "Sinergia RC-SP10", "Sinergia Sortation", "Sinergia Insumo", "Atividade-Externa", "Suspensao", "Desligado"
  ];

  for (let i = 1; i < dados.length; i++) {
    const [
      areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada
    ] = dados[i];

  if (
      (teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase() ||
      (turno || "").trim().toUpperCase() !== (turnoSelecionado || "").trim().toUpperCase()
    ) continue;

    // estado inicial
    let statusPresenca    = "NÃO DEFINIDO";
    let qtdDias           = 1;
    let motivoDesligamento= "";
    const entrada         = (dados[i][17]||"").toUpperCase().trim();  // Col R

    // busca índice no array vindo do select
    const index = listaStatus.findIndex(item =>
      item.COLABORADOR?.toUpperCase().trim() === colaborador?.toUpperCase().trim() &&
      item.TURNO?.       toUpperCase().trim() === turno?.toUpperCase().trim() &&
      item.EMPRESA?.     toUpperCase().trim() === empresa?.toUpperCase().trim() &&
      item.PROCESSO?.    toUpperCase().trim() === processo?.toUpperCase().trim()
    );

    if (index !== -1) {
      // ————— VEIO DO SELECT —————
      statusPresenca = listaStatus[index].STATUS_abs || "NÃO DEFINIDO";

      if (statusMultiplo.includes(statusPresenca)) {
        qtdDias = (statusPresenca === "Desligado")
          ? 31
          : parseInt(listaStatus[index].QTD_DIAS || "1", 10);
      }

      if (
        statusPresenca === "Desligado" ||
        statusPresenca === "Transferido" ||
        statusPresenca === "Afastamento" ||
        statusPresenca === "Afastamento-Acd-Trab" ||
        statusPresenca === "Licenca"
      ) {
        motivoDesligamento = listaStatus[index].MOTIVO_DESLIGAMENTO || "";
        const dataFmt = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

        // Só envia email se o colaborador está como ATIVO na planilha!
        if ((statusHC || "").toUpperCase().trim() === "ATIVO") {
          enviarEmailDesligamento(colaborador, dataFmt, teamLeader, motivoDesligamento, statusPresenca);
          desligadosCom.push(colaborador);
        }
      }


      // ————— fim do bloco “select” —————
    }
    else {
      // ————— só quem NÃO veio do select —————
      if (entrada === "FOLGA-ESCALADA") {
        statusPresenca = "Folga-Escala";
      }
      else if (entrada === "PRESENCA-HE") {
        statusPresenca = "Presenca-HE";
      }
      // outros casos de entrada podem ir aqui…
    }

    // agora empacota a gravação, repetindo por qtdDias
    for (let d = 0; d < qtdDias; d++) {
      const dataHora = Utilities.formatDate(
        new Date(hoje.getTime() + d*86400000),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
      );
      const diasRest = qtdDias - d - 1;

      registros.push([
        dataDemissao||"",  // A
        statusHC||"",      // B
        areaHead||"",      // C
        supervisor||"",    // D
        teamLeader||"",    // E
        idgroot||"",       // F
        ldap||"",          // G
        colaborador||"",   // H
        turno||"",         // I
        escala||"",        // J
        turmaEscala||"",   // K
        empresa||"",       // L
        processo||"",      // M
        statusPresenca,    // N
        dataHora,          // O
        (statusMultiplo.includes(statusPresenca) && statusPresenca!=="Desligado")
          ? diasRest : "", // P
        "",                // Q
        statusPresenca==="Desligado"? motivoDesligamento : "", // R
        statusPresenca==="Desligado"? diasRest : "",           // S
        Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy"), // T
        "",                // U
        versao,            // V
        usuarioEmail
      ]);
    }
  }

  if (registros.length) {
 let ultimaLinha = abaHistorico.getLastRow() + 1;
  abaHistorico.getRange(ultimaLinha, 1, registros.length, registros[0].length).setValues(registros);
    let msg = `Presença registrada com sucesso para ${nomeSelecionado}.`;
    if (desligadosCom.length)
      msg += `\n\n📧 E-mail(s) enviado(s) para: ${desligadosCom.join(", ")}`;
    return msg;
  }
}
*/
function registrarPresenca(nomeSelecionado, listaStatus, versao, turnoSelecionado) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const abaDados = planilha.getSheetByName(ABA_DADOS_EQUIPE);
  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const abaHistorico = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  const dados = abaDados.getDataRange().getValues();
  const registros = [];
  const hoje = new Date();
  const desligadosCom = [];
  if (!abaHistorico) return "Erro: Aba 'Historico_3' não encontrada.";
  const usuarioEmail = Session.getActiveUser().getEmail() || "indefinido";
  const statusMultiplo = [
    "Transferido", "OWNboarding", "Atestado", "Banco-de-Horas", "Afastamento", "Licenca", "Atestado-Acd-Trab", "Ferias",
    "Treinamento-Ext", "Treinamento-Int", "Treinamento-REP-III", "Ferias", "Fretado", "Afastamento", "Afastamento-Acd-Trab",
    "Licenca", "Sinergia SP014", "Sinergia CX", "Sinergia INB", "Sinergia LOS", "Sinergia MG01", "Sinergia MWH", "Sinergia OUT",
    "Sinergia QUA", "Sinergia RC01", "Sinergia RET", "Sinergia SP011", "Sinergia SP02", "Sinergia SP03", "Sinergia SP04",
    "Sinergia SP05", "Sinergia SP06", "Sinergia SP09", "Sinergia INV", "Sinergia RET", "Sinergia QUA", "Sinergia SVC",
    "Sinergia RC-SP10", "Sinergia Sortation", "Sinergia Insumo", "Atividade-Externa", "Suspensao", "Desligado"
  ];

  for (let i = 1; i < dados.length; i++) {
    const [
      areaHead, idgroot, ldap, teamLeader, supervisor, colaborador, turno,
      escala, turmaEscala, empresa, processo, statusHC, dataDemissao, , statusEntrada
    ] = dados[i];

    if (
      (teamLeader || "").trim().toUpperCase() !== nomeSelecionado.trim().toUpperCase() ||
      (turno || "").trim().toUpperCase() !== (turnoSelecionado || "").trim().toUpperCase()
    ) continue;

    // estado inicial
    let statusPresenca = "NÃO DEFINIDO";
    let qtdDias = 1;
    let motivoDesligamento = "";
    const entrada = (dados[i][17] || "").toUpperCase().trim(); // Col R

    // busca índice no array vindo do select
    const index = listaStatus.findIndex(item =>
      item.COLABORADOR?.toUpperCase().trim() === colaborador?.toUpperCase().trim() &&
      item.TURNO?.toUpperCase().trim() === turno?.toUpperCase().trim() &&
      item.EMPRESA?.toUpperCase().trim() === empresa?.toUpperCase().trim() &&
      item.PROCESSO?.toUpperCase().trim() === processo?.toUpperCase().trim()
    );

    if (index !== -1) {
      // ————— VEIO DO SELECT —————
      statusPresenca = listaStatus[index].STATUS_abs || "NÃO DEFINIDO";

      if (statusMultiplo.includes(statusPresenca)) {
        qtdDias = (statusPresenca === "Desligado")
          ? 31
          : parseInt(listaStatus[index].QTD_DIAS || "1", 10);
      }

      if (
        statusPresenca === "Desligado" ||
        statusPresenca === "Transferido" ||
        statusPresenca === "Afastamento" ||
        statusPresenca === "Afastamento-Acd-Trab" ||
        statusPresenca === "Licenca"
      ) {
        motivoDesligamento = listaStatus[index].MOTIVO_DESLIGAMENTO || "";
        const dataFmt = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

        // Só envia email se o colaborador está como ATIVO na planilha!
        if ((statusHC || "").toUpperCase().trim() === "ATIVO") {
          enviarEmailDesligamento(colaborador, dataFmt, teamLeader, motivoDesligamento, statusPresenca);
          desligadosCom.push(colaborador);
        }
      }

      // ————— fim do bloco “select” —————
    }
    else {
      // ————— só quem NÃO veio do select —————
      if (entrada === "FOLGA-ESCALADA") {
        statusPresenca = "Folga-Escala";
      }
      else if (entrada === "PRESENCA-HE") {
        statusPresenca = "Presenca-HE";
      }
      // outros casos de entrada podem ir aqui…
    }

    // agora empacota a gravação, repetindo por qtdDias
    for (let d = 0; d < qtdDias; d++) {
      const dataHora = Utilities.formatDate(
        new Date(hoje.getTime() + d * 86400000),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
      );
      const diasRest = qtdDias - d - 1;

      registros.push([
        dataDemissao || "",  // A
        statusHC || "",      // B
        areaHead || "",      // C
        supervisor || "",    // D
        teamLeader || "",    // E
        idgroot || "",       // F
        ldap || "",          // G
        colaborador || "",   // H
        turno || "",         // I
        escala || "",        // J
        turmaEscala || "",   // K
        empresa || "",       // L
        processo || "",      // M
        statusPresenca,      // N
        dataHora,            // O
        (statusMultiplo.includes(statusPresenca) && statusPresenca !== "Desligado")
          ? diasRest : "",   // P
        "",                  // Q
        statusPresenca === "Desligado" ? motivoDesligamento : "", // R
        statusPresenca === "Desligado" ? diasRest : "",           // S
        Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy"), // T
        "",                  // U
        versao,              // V
        usuarioEmail
      ]);
    }
  }

  if (registros.length) {
    // BLOCO DE GRAVAÇÃO PROTEGIDO POR LOCK E APPENDROW
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000); // Exclusividade global!
        const ultimaLinha = abaHistorico.getLastRow() + 1;
        abaHistorico.getRange(ultimaLinha, 1, registros.length, registros[0].length)
          .setValues(registros);
      let msg = `Presença registrada com sucesso para ${nomeSelecionado}, ${registros.length} linhas salvas.`;
      if (desligadosCom.length)
        msg += `\n\n📧 E-mail(s) enviado(s) para: ${desligadosCom.join(", ")}`;
      return msg;
    } catch (e) {
      return "Erro ao registrar presença (lock): " + e.message;
    } finally {
      lock.releaseLock();
    }
  }
}

function enviarEmailDesligamento(nomeColaborador, dataAtual, teamLeader, motivoDesligamento, statusPresenca) {
  // Destinatários conforme o status
  let destinatarios = "";
  if (statusPresenca === "Desligado") {
    destinatarios = "Flow.brsp10@mercadolivre.com, acesso.brsp10@mercadolivre.com, lossprevention.brsp10@mercadolivre.com, lms-brsp10@mercadolivre.com";
  } else {
    destinatarios = "Flow.brsp10@mercadolivre.com, lms-brsp10@mercadolivre.com";//
  }
  const emcopia = "";

  let assunto = "";
  let mensagem = "";

  if (statusPresenca === "Desligado") {
    assunto = "Solicitação de desligamento de colaborador";
    mensagem = `
      <p>Prezados,</p>
      <p>Boa tarde,</p>
      <p>
        Venho por meio deste solicitar o desligamento do colaborador(a) abaixo, bem como a atualização do status na (HC) de "Ativo" para "Desligado":
      </p>
      <p>
        📅 <b>Data da solicitação de desligamento:</b> ${dataAtual}<br>
        👤 <b>Nome completo:</b> ${nomeColaborador}<br>
        👨‍💼 <b>Team Leader:</b> ${teamLeader}<br>
        📝 <b>Motivo do desligamento:</b> ${motivoDesligamento}
      </p>
      <p>
        Solicito, por gentileza, que a base seja atualizada o quanto antes para mantermos os registros operacionais em conformidade.
      </p>
      <p>
        Fico à disposição para quaisquer dúvidas ou esclarecimentos adicionais.
      </p>
      <p>
        Atenciosamente,<br>
        ${teamLeader}
      </p>
    `;
  } else if (statusPresenca === "Transferido") {
    assunto = "Solicitação de transferência de colaborador";
    mensagem = `
      <p>Prezados,</p>
      <p>Boa tarde,</p>
      <p>
        Venho por meio deste solicitar a transferência do colaborador(a) abaixo:
      </p>
      <p>
        📅 <b>Data da solicitação de transferência:</b> ${dataAtual}<br>
        👤 <b>Nome completo:</b> ${nomeColaborador}<br>
        👨‍💼 <b>Team Leader:</b> ${teamLeader}<br>
        📝 <b>Motivo da transferência:</b> ${motivoDesligamento}
      </p>
      <p>
        Solicito, por gentileza, que a base seja atualizada o quanto antes para mantermos os registros operacionais em conformidade.
      </p>
      <p>
        Fico à disposição para quaisquer dúvidas ou esclarecimentos adicionais.
      </p>
      <p>
        Atenciosamente,<br>
        ${teamLeader}
      </p>
    `;
  } else if (
    statusPresenca === "Afastamento" ||
    statusPresenca === "Afastamento-Acd-Trab"
  ) {
    assunto = `Solicitação de ${statusPresenca.toLowerCase().replace("-", " ")} de colaborador`;
    mensagem = `
      <p>Prezados,</p>
      <p>Boa tarde,</p>
      <p>
        Venho por meio deste solicitar a atualização do status do colaborador(a) abaixo para "${statusPresenca.replace("-", " ")}":
      </p>
      <p>
        📅 <b>Data da solicitação de afastamento/licença:</b> ${dataAtual}<br>
        👤 <b>Nome completo:</b> ${nomeColaborador}<br>
        👨‍💼 <b>Team Leader:</b> ${teamLeader}<br>
        📝 <b>Motivo:</b> ${motivoDesligamento}
      </p>
      <p>
        Solicito, por gentileza, que a base seja atualizada o quanto antes para mantermos os registros operacionais em conformidade.
      </p>
      <p>
        Fico à disposição para quaisquer dúvidas ou esclarecimentos adicionais.
      </p>
      <p>
        Atenciosamente,<br>
        ${teamLeader}
      </p>
    `;
  }

  MailApp.sendEmail({
    to: destinatarios,
    bcc: emcopia,
    subject: assunto,
    htmlBody: mensagem
  });
}



function verificarRegistroHoje(teamLeader, turnoSelecionado) {
  const planilhaHistorico = SpreadsheetApp.openById(HISTORICO_ID);
  const aba = planilhaHistorico.getSheetByName(ABA_HISTORICO);
  const dados = aba.getDataRange().getValues();
  const hoje = new Date();
  const dataHojeFmt = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

  for (let i = 1; i < dados.length; i++) {
    const tl = (dados[i][4] || "").toString().trim().toUpperCase();   // Coluna E - Team Leader
    const turno = (dados[i][8] || "").toString().trim().toUpperCase(); // Coluna I - TURNO
    const dataExecucaoStr = (dados[i][19] || "").toString().trim();   // Coluna T - Data Execução

    if (
      tl === teamLeader.trim().toUpperCase() &&
      turno === (turnoSelecionado || "").trim().toUpperCase() &&
      dataExecucaoStr === dataHojeFmt
    ) {
      // Encontrou lançamento para o mesmo TL, TURNO e DATA
      return `⚠️ Atenção, já existe dados da sua equipe registrados para o turno "${turnoSelecionado}" na data ${dataHojeFmt}!⚠️\n\nAcesse a planilha ABS Control Tower para visualizar os dados:\nhttps://docs.google.com/spreadsheets/d/1ugAXKQWod5Vk7DSulEo53H-Wwox_B4bRaQZjA-TgO4U/edit?gid=0#gid=0`;
    }
  }
  return ""; // Nenhum registro encontrado
}
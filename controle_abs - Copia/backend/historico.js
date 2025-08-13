/***** CONFIG *****/
const PLANILHA_ID   = "1cXZ5Jh6stvygynOnWbzWopHjdiw3xVliqrYBjJ5H8as";
const ABA_DADOS_EQUIPE = "dadosEquipe_3";
const HISTORICO_ID  = "1-GRyDj6BUBjRnO2QqMmihxVCZxw3JvJLIrhFHSmgpbI";
const ABA_HISTORICO = "Historico_agosto_teste";

/***** UI *****/
function doGet(e) {
  const pagina = e.parameter.page || "index";
  try {
    let template = HtmlService.createTemplateFromFile(pagina);
    return template.evaluate()
      .setTitle("Historico de registros")
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
    <script src="https://cdn.jsdelivr.net/npm/flatpickr/dist/l10n/pt.js"></script>
  `;
}

function incluirHTML(arquivo) {
  return HtmlService.createHtmlOutputFromFile(arquivo).getContent();
}

/***** UTIL *****/
function tz() { return Session.getScriptTimeZone() || "America/Sao_Paulo"; }

function normalizarTextoABS(v) {
  return (v || "")
    .toString()
    .replace(/^'+|'+$/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizarIDGROOT(v) {
  return (v || "").toString().replace(/^'+/, "").replace(/\s+/g, "").trim().toUpperCase();
}

function normalizarDataABS(dataCelula) {
  // Aceita Date, "'dd/MM/yyyy", "dd/MM/yyyy", "yyyy-mm-dd"
  if (!dataCelula) return "";
  if (dataCelula instanceof Date) {
    return Utilities.formatDate(dataCelula, tz(), "dd/MM/yyyy");
  }
  let d = dataCelula.toString().replace(/^'+|'+$/g, "").replace(/\s+/g, "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
    const [y, m, dia] = d.split("-");
    return `${dia}/${m}/${y}`;
  }
  return d;
}

function parseDDMMYYYYToDateBR(s) {
  // "dd/MM/yyyy" -> Date (00:00 na TZ do script)
  const [dd, mm, yyyy] = s.split("/").map(x => parseInt(x, 10));
  const d = new Date(yyyy, mm - 1, dd);
  d.setHours(0,0,0,0);
  return d;
}

function hoje00() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0,0,0,0);
}

function dataMinimaPermitida() {
  // hoje - 2 dias (00:00)
  const h = hoje00();
  return new Date(h.getFullYear(), h.getMonth(), h.getDate() - 2, 0,0,0,0);
}

function isDataPermitida(ddmmyyyy) {
  // Permitido: >= (hoje - 2)  (passado limitado)  |  Futuro: liberado
  const d = parseDDMMYYYYToDateBR(ddmmyyyy);
  const min = dataMinimaPermitida();
  return d >= min;
}

/***** DADOS BASE *****/
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

/***** BUSCA HISTÓRICO PARA EDIÇÃO (TL + múltiplas datas) *****/
function obterHistoricoPorDatas(nomeTeamLeader, datasSelecionadas) {
  if (!nomeTeamLeader) throw new Error("Team Leader não informado.");

  // Normaliza para dd/MM/yyyy e aplica regra de bloqueio (passado só até -2; futuro livre)
  const datasRaw = (datasSelecionadas || [])
    .map(d => d && d.toString().trim())
    .filter(Boolean)
    .map(normalizarDataABS); // dd/MM/yyyy

  if (datasRaw.length === 0) throw new Error("Nenhuma data válida informada.");

  const datasValidas = [];
  const datasBloqueadas = [];
  for (const ds of datasRaw) {
    if (isDataPermitida(ds)) datasValidas.push(ds);
    else datasBloqueadas.push(ds);
  }

  if (datasValidas.length === 0) {
    throw new Error("As datas selecionadas são anteriores ao limite (apenas até 2 dias antes de hoje).");
  }

  const plan = SpreadsheetApp.openById(HISTORICO_ID);
  const aba  = plan.getSheetByName(ABA_HISTORICO);
  if (!aba) throw new Error(`Aba '${ABA_HISTORICO}' não encontrada.`);

  const valores = aba.getDataRange().getValues();
  const res = [];
  for (let i = 1; i < valores.length; i++) {
    const linha = valores[i];
    const teamLeader = normalizarTextoABS(linha[4]);  // E
    const dataRef    = normalizarDataABS(linha[14]);  // O
    if (teamLeader !== normalizarTextoABS(nomeTeamLeader)) continue;
    if (!datasValidas.includes(dataRef)) continue;

    res.push({
      _rowIndex: i + 1,           // linha real na planilha (1-based)
      AREA_HEAD: linha[2],        // C
      IDGROOT:   linha[5],        // F
      LDAP:      linha[6],        // G
      COLABORADOR: linha[7],      // H
      TURNO:       linha[8],      // I
      ESCALA:      linha[9],      // J
      TURMA_ESCALA:linha[10],     // K
      EMPRESA:     linha[11],     // L
      PROCESSO:    linha[12],     // M
      STATUS_HC:   linha[1],      // B
      STATUS_abs:  linha[13],     // N
      QTD_DIAS:    linha[15],     // P
      DATA:        dataRef,       // O
      MOTIVO:      linha[17] || "",// R
      USUARIO:     linha[22]    //W
    });
  }

  // Opcional: informar ao front que houve datas removidas (via throw ou simplesmente seguir).
  // Aqui sigo retornando normalmente; o front já filtra via flatpickr/onChange.
  return res;
}

/***** SALVAR EDIÇÕES (UPDATE EM LINHA EXISTENTE) *****/
function salvarEdicoesHistorico(alteracoes) {
  if (!Array.isArray(alteracoes) || alteracoes.length === 0) return "Nada a salvar.";

  const plan = SpreadsheetApp.openById(HISTORICO_ID);
  const aba  = plan.getSheetByName(ABA_HISTORICO);
  if (!aba) throw new Error(`Aba '${ABA_HISTORICO}' não encontrada.`);

  const agora = new Date();
  const carimbo = Utilities.formatDate(agora, tz(), "dd/MM/yyyy HH:mm:ss");

  // 👇 tenta obter o e-mail do usuário que acessa o app
  // (em alguns cenários pode vir vazio; ver observação abaixo)
  const usuarioEmail =
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    "usuário-desconhecido";

  // Mapa chave -> linha
  const valores = aba.getDataRange().getValues();
  const mapa = new Map();
  for (let i = 1; i < valores.length; i++) {
    const l = valores[i];
    const key = [
      normalizarTextoABS(l[7]),   // H COLABORADOR
      normalizarTextoABS(l[8]),   // I TURNO
      normalizarTextoABS(l[11]),  // L EMPRESA
      normalizarTextoABS(l[12]),  // M PROCESSO
      normalizarDataABS(l[14])    // O DATA (dd/MM/yyyy)
    ].join("|");
    mapa.set(key, i + 1);
  }

  const statusComDias = new Set([
    "TRANSFERIDO","OWNBOARDING","ATESTADO","BANCO-DE-HORAS","AFASTAMENTO","LICENCA","ATESTADO-ACD-TRAB","FERIAS",
    "TREINAMENTO-EXT","TREINAMENTO-INT","TREINAMENTO-REP-III","FRETADO","AFASTAMENTO-ACD-TRAB",
    "SINERGIA SP014","SINERGIA CX","SINERGIA INB","SINERGIA LOS","SINERGIA MG01","SINERGIA MWH","SINERGIA OUT",
    "SINERGIA QUA","SINERGIA RC01","SINERGIA RET","SINERGIA SP011","SINERGIA SP02","SINERGIA SP03","SINERGIA SP04",
    "SINERGIA SP05","SINERGIA SP06","SINERGIA SP09","SINERGIA INV","SINERGIA SVC","SINERGIA RC-SP10",
    "SINERGIA SORTATION","SINERGIA INSUMO","ATIVIDADE-EXTERNA","SUSPENSAO"
  ]);

  const porLinha = new Map(); // row -> [{col,value},...]

  alteracoes.forEach(a => {
    const key = [
      normalizarTextoABS(a.nome),
      normalizarTextoABS(a.turno),
      normalizarTextoABS(a.empresa),
      normalizarTextoABS(a.processo),
      normalizarDataABS(a.dataRef)
    ].join("|");

    const rowIdx = mapa.get(key);
    if (!rowIdx) return; // só edita se registro existir

    const novoStatus = (a.novoStatus || "").toString().trim();
    const motivo     = (a.motivoDesligamento || "").toString().trim();

    const edits = [];
    // N (14): STATUS
    edits.push({ col: 14, value: novoStatus });

    // P (16): QTD_DIAS
    if (statusComDias.has(normalizarTextoABS(novoStatus))) {
      const n = parseInt(a.qtdDias, 10);
      edits.push({ col: 16, value: Number.isFinite(n) ? n : "" });
    } else if (normalizarTextoABS(novoStatus) === "DESLIGADO") {
      edits.push({ col: 16, value: 31 });
    } else {
      edits.push({ col: 16, value: "" });
    }

    // Q (17): carimbo + e-mail do editor
    edits.push({ col: 17, value: `Alterado em: ${carimbo} por: ${usuarioEmail}` });

    // R (18): motivo (se desligado)
    if (normalizarTextoABS(novoStatus) === "DESLIGADO") {
      edits.push({ col: 18, value: motivo });
    }

    if (!porLinha.has(rowIdx)) porLinha.set(rowIdx, []);
    porLinha.get(rowIdx).push(...edits);
  });

  // aplica em lote por linha
  porLinha.forEach((lst, row) => {
    const linhaVals = aba.getRange(row, 1, 1, aba.getLastColumn()).getValues()[0];
    lst.forEach(({ col, value }) => { linhaVals[col - 1] = value; });
    aba.getRange(row, 1, 1, linhaVals.length).setValues([linhaVals]);
  });

  return "✅ Alterações salvas.";
}

function preencherStatusABS() {
  // 1. Abre as planilhas e abas necessárias
  const ss = SpreadsheetApp.openById('1ugAXKQWod5Vk7DSulEo53H-Wwox_B4bRaQZjA-TgO4U'); // Planilha destino
  const abaJulho = ss.getSheetByName('Agosto_2025'); // Aba de destino
  const ssHistorico = SpreadsheetApp.openById('140f04559QKqzqpOWHwcN5NZhwf38IHH6nSfJuWuriCY'); // Planilha origem histórico
  const abaHistorico = ssHistorico.getSheetByName('Historico_agosto'); // Aba histórico
  const abaEscala = ss.getSheetByName('Escala'); // Aba escala (folga)

  // 2. Valida se todas as abas existem
  if (!abaHistorico) throw new Error('Aba "Historico_agosto" não encontrada!');
  if (!abaJulho) throw new Error('Aba "Agosto_2025" não encontrada!');
  if (!abaEscala) throw new Error('Aba "Escala" não encontrada!');

  // 3. Carrega todos os dados relevantes das abas
  const dadosJulho = abaJulho.getDataRange().getValues();            // Matriz completa da aba Julho_2025
  const headersDatas = dadosJulho[1].slice(16, 47);                  // Datas do mês (cabeçalho das colunas Q:AU, linha 2)
  const dadosHistorico = abaHistorico.getDataRange().getValues();    // Histórico completo
  const dadosEscala = abaEscala.getDataRange().getValues();          // Folgas escala

  // 4. Define os índices das colunas (base 0) de Julho_2025
  const COL_JULHO_STATUS_RH = 0;    // A (Status RH)
  const COL_JULHO_TURMA = 6;        // G (Turma/letra)
  const COL_JULHO_NOME = 10;        // K (Nome)
  const COL_JULHO_DATA_DESLIG = 12; // M (Data desligamento)
  const COL_JULHO_DATA_ADMISSAO = 13; // N (Data admissão)

  // 5. Aqui serão montadas todas as linhas para preencher Q:AU
  const linhasParaPreencher = [];

  // 6. Processa cada colaborador a partir da linha 3 (índice 2)
  for (let i = 2; i < dadosJulho.length; i++) {
    const statusRH = dadosJulho[i][COL_JULHO_STATUS_RH];
    const turma = dadosJulho[i][COL_JULHO_TURMA];
    const nome = dadosJulho[i][COL_JULHO_NOME];
    const dataDesligamento = dadosJulho[i][COL_JULHO_DATA_DESLIG];
    const dataAdmissao = dadosJulho[i][COL_JULHO_DATA_ADMISSAO];

    // 7. Se coluna A está vazia, preenche toda a linha com ""
    if (!statusRH || String(statusRH).trim() === "") {
      linhasParaPreencher.push(Array(headersDatas.length).fill(""));
      continue; // pula para o próximo colaborador
    }

    // 8. Monta o status de cada dia do mês (Q:AU)
    const linhaStatus = [];
    for (let j = 0; j < headersDatas.length; j++) {
      const dataDia = headersDatas[j];

      // 8.1. Antes da data de admissão não preenche nada
      if (dataAdmissao && compararDatasTexto(dataDia, dataAdmissao) < 0) {
        linhaStatus.push("");
        continue;
      }

      // 8.2. Após data de desligamento: "Desligado"
      if (statusRH == "DESLIGADO" && dataDesligamento && compararDatasTexto(dataDia, dataDesligamento) >= 0) {
        linhaStatus.push("Desligado");
        continue;
      }

      // 8.3. Busca status exato no histórico para o dia/nome (sem considerar turno)
      let status = buscarStatusPorNomeData(
        dadosHistorico,
        nome,
        dataDia
      );
      if (status) {
        linhaStatus.push(status);
        continue;
      }

      // 8.4. Se não há status no histórico, valida se está de folga escala
      let folgaEscala = verificarFolgaEscalaPorTurma(dadosEscala, turma, dataDia);
      if (folgaEscala) {
        linhaStatus.push("Folga-Escala");
      } else {
        linhaStatus.push("(Preencher)");
      }
    }
    // 9. Adiciona a linha finalizada ao array principal
    linhasParaPreencher.push(linhaStatus);
  }

  // 10. Escreve todas as linhas no range Q3:AU (mantendo cabeçalho nas duas primeiras linhas)
  abaJulho.getRange(3, 17, linhasParaPreencher.length, headersDatas.length).setValues(linhasParaPreencher);
}

/**
 * Busca status no histórico apenas por nome e data (independente do turno)
 */
function buscarStatusPorNomeData(historico, nome, data) {
  for (let i = 1; i < historico.length; i++) {
    if (
      historico[i][7] == nome &&       // COLABORADOR (H)
      historico[i][14] == data         // DATA (O)
    ) {
      return historico[i][13];         // Status (N)
    }
  }
  return "";
}

/**
 * Verifica se a turma/letra está de folga escala na data (Escala!A:A = data, B:F = turma)
 */
function verificarFolgaEscalaPorTurma(dadosEscala, turma, dataDia) {
  for (let i = 1; i < dadosEscala.length; i++) {
    if (String(dadosEscala[i][0]).trim() == String(dataDia).trim()) {
      for (let j = 1; j <= 5; j++) { // Colunas B até F
        if (String(dadosEscala[i][j]).trim().toUpperCase() == String(turma).trim().toUpperCase()) {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Compara datas no formato texto dd/MM/yyyy
 * Retorna:
 *   1 se data1 > data2
 *   0 se data1 == data2
 *  -1 se data1 < data2
 */
function compararDatasTexto(data1, data2) {
  function parseData(str) {
    if (!str || typeof str !== 'string') return null;
    const p = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (p) return new Date(+p[3], +p[2] - 1, +p[1]);
    return null;
  }
  const d1 = parseData(data1);
  const d2 = parseData(data2);
  if (!d1 || !d2) return 0;
  if (d1.getTime() > d2.getTime()) return 1;
  if (d1.getTime() < d2.getTime()) return -1;
  return 0;
}

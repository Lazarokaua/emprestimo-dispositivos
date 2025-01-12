/**
 * Verifica se uma máquina está cadastrada no sistema
 * @param {string|number} idMaquina - ID da máquina a ser verificada
 * @returns {boolean} True se a máquina estiver cadastrada, False caso contrário
 */
function verificarMaquinaCadastrada(idMaquina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaPatio = ss.getSheetByName("Patio");
  const planilhaDispositivos = ss.getSheetByName("Dispositivos");

  const dadosPatio = planilhaPatio.getRange("A2:C").getValues();
  const dadosDispositivos = planilhaDispositivos.getRange("A2:C").getValues();

  for (let dados of [dadosPatio, dadosDispositivos]) {
    for (let row of dados) {
      if (row[0] == idMaquina) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Verifica se uma máquina está atualmente em uso
 * @param {string|number} idMaquina - ID da máquina a ser verificada
 * @returns {boolean} True se a máquina estiver em uso, False caso contrário
 */
function verificarMaquinaEmUso(idMaquina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaControle = ss.getSheetByName("Controle");
  const dadosControle = planilhaControle.getRange("A2:I").getValues();

  for (let row of dadosControle) {
    if (row[0] == idMaquina && row[8] !== "Disponível") {
      return true;
    }
  }

  return false;
}

/**
 * Obtém informações detalhadas de uma máquina
 * @param {string|number} idMaquina - ID da máquina
 * @returns {Object|null} Objeto com tipo e status da máquina ou null se não encontrada
 */
function obterInformacoesMaquina(idMaquina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaPatio = ss.getSheetByName("Patio");
  const planilhaDispositivos = ss.getSheetByName("Dispositivos");

  const dadosPatio = planilhaPatio.getRange("A2:C").getValues();
  const dadosDispositivos = planilhaDispositivos.getRange("A2:C").getValues();

  for (let dados of [dadosPatio, dadosDispositivos]) {
    for (let row of dados) {
      if (row[0] == idMaquina) {
        return { tipo: row[1], status: row[2] };
      }
    }
  }

  return null;
}

/**
 * Atualiza o status de uma máquina no sistema
 * @param {string|number} idMaquina - ID da máquina
 * @param {string} novoStatus - Novo status a ser definido
 */
function atualizarStatusMaquina(idMaquina, novoStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaPatio = ss.getSheetByName("Patio");
  const planilhaDispositivos = ss.getSheetByName("Dispositivos");

  const dadosPatio = planilhaPatio.getRange("A2:C").getValues();
  const dadosDispositivos = planilhaDispositivos.getRange("A2:C").getValues();

  for (let [index, dados] of [dadosPatio, dadosDispositivos].entries()) {
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] == idMaquina) {
        const planilha = index === 0 ? planilhaPatio : planilhaDispositivos;
        planilha.getRange(i + 2, 3).setValue(novoStatus);
        return;
      }
    }
  }
}

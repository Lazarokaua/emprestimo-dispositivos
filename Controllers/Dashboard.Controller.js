/**
 * Atualiza todas as informações do dashboard com dados atuais do sistema
 * Inclui estatísticas gerais, por tipo de máquina e por setor
 */
function atualizarDashboard() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilhaDashboard = ss.getSheetByName("DASHBOARD");
    const planilhaControle = ss.getSheetByName("Controle");
    const planilhaPatio = ss.getSheetByName("Patio");
    const planilhaDispositivos = ss.getSheetByName("Dispositivos");
  
    // Limpa o dashboard atual para nova atualização
    planilhaDashboard.clear();
  
    // Configura o título do dashboard
    planilhaDashboard.getRange("A1").setValue("Dashboard de Gerenciamento de Máquinas");
    planilhaDashboard.getRange("A1:B1").merge().setFontWeight("bold").setHorizontalAlignment("center");
  
    // Obtém dados de todas as planilhas relevantes
    const dadosControle = planilhaControle.getDataRange().getValues();
    const dadosPatio = planilhaPatio.getDataRange().getValues();
    const dadosDispositivos = planilhaDispositivos.getDataRange().getValues();
  
    // Calcula estatísticas gerais do sistema
    const totalMaquinas = dadosPatio.length + dadosDispositivos.length - 2; // Subtrai cabeçalhos
    const maquinasEmUso = dadosControle.filter(row => row[8] === "Em Uso").length;
    const maquinasDisponiveis = totalMaquinas - maquinasEmUso;
    const maquinasPendentes = dadosControle.filter(row => row[8] === "Pendente").length;
  
    // Adiciona seção de estatísticas gerais
    let row = 3;
    planilhaDashboard.getRange(`A${row}`).setValue("Estatísticas Gerais");
    row++;
    planilhaDashboard.getRange(`A${row}`).setValue("Total de Máquinas:");
    planilhaDashboard.getRange(`B${row}`).setValue(totalMaquinas);
    row++;
    planilhaDashboard.getRange(`A${row}`).setValue("Máquinas em Uso:");
    planilhaDashboard.getRange(`B${row}`).setValue(maquinasEmUso);
    row++;
    planilhaDashboard.getRange(`A${row}`).setValue("Máquinas Disponíveis:");
    planilhaDashboard.getRange(`B${row}`).setValue(maquinasDisponiveis);
    row++; // Incrementa a linha para adicionar a informação de máquinas pendentes
    planilhaDashboard.getRange(`A${row}`).setValue("Máquinas Pendentes:"); // Adiciona o label "Máquinas Pendentes"
    planilhaDashboard.getRange(`B${row}`).setValue(maquinasPendentes); // Adiciona o valor de máquinas pendentes
    row += 2; // pula uma linha a mais
  
    // Calcula e adiciona estatísticas por tipo de máquina
    const tiposMaquinas = {};
    [...dadosPatio.slice(1), ...dadosDispositivos.slice(1)].forEach(linha => {
      const tipo = linha[1];
      tiposMaquinas[tipo] = (tiposMaquinas[tipo] || 0) + 1;
    });
  
    // Adiciona seção de estatísticas por tipo
    planilhaDashboard.getRange(`A${row}`).setValue("Estatísticas por Tipo de Máquina");
    row++;
    for (const [tipo, quantidade] of Object.entries(tiposMaquinas)) {
      planilhaDashboard.getRange(`A${row}`).setValue(tipo);
      planilhaDashboard.getRange(`B${row}`).setValue(quantidade);
      row++;
    }
    row += 2;
  
    // Processa informações específicas por tipo de dispositivo
    const tiposDispositivos = ["Paleteira Manual", "Empilhadeira", "Coletor"];
    
    // Adiciona cards para cada tipo de dispositivo
    tiposDispositivos.forEach(tipo => {
      let total = 0;
      let emUso = 0;
      let noPatio = 0;
      let noDispositivos = 0;
  
      // Calcula total e disponíveis por localização
      [...dadosPatio.slice(1), ...dadosDispositivos.slice(1)].forEach(linha => {
        if (linha[1].toLowerCase().includes(tipo.toLowerCase())) {
          total++;
          if (linha[2] === "Disponível") {
            if (tipo === "Paleteira Manual" || tipo === "Empilhadeira") {
              noPatio++;
            } else {
              noDispositivos++;
            }
          }
        }
      });

      // Calcula máquinas em uso a partir da planilha de operação
      dadosControle.slice(1).forEach(linha => {
        if (linha[1].toLowerCase().includes(tipo.toLowerCase()) && linha[8] === "Em Uso") {
          emUso++;
        }
      });
  
      // Adiciona informações ao dashboard
      planilhaDashboard.getRange(`A${row}`).setValue(tipo);
      row++;
      planilhaDashboard.getRange(`A${row}`).setValue("Total:");
      planilhaDashboard.getRange(`B${row}`).setValue(total);
      row++;
      planilhaDashboard.getRange(`A${row}`).setValue("Em uso:");
      planilhaDashboard.getRange(`B${row}`).setValue(emUso);
      row++;
      if (tipo === "Paleteira Manual Manual" || tipo === "Empilhadeira") {
        planilhaDashboard.getRange(`A${row}`).setValue("No Patio:");
        planilhaDashboard.getRange(`B${row}`).setValue(noPatio);
      } else {
        planilhaDashboard.getRange(`A${row}`).setValue("Na Sala Coletor:");
        planilhaDashboard.getRange(`B${row}`).setValue(noDispositivos);
      }
      row += 2;
    });
  
    // Calcula e adiciona estatísticas por setor
    const maquinasPorSetor = {};
    dadosControle.slice(1).forEach(linha => {
      if (linha[8] === "Em Uso") { // Verifica se está em uso
        const setor = linha[4]; // Coluna E - Setor
        const tipo = linha[1];  // Coluna B - Tipo de máquina
        
        if (!maquinasPorSetor[setor]) {
          maquinasPorSetor[setor] = {};
        }
        if (!maquinasPorSetor[setor][tipo]) {
          maquinasPorSetor[setor][tipo] = 0;
        }
        maquinasPorSetor[setor][tipo]++;
      }
    });
  
    // Preparar dados para o dashboard
    const dadosMaquinasPorSetor = [];
    for (const [setor, tipos] of Object.entries(maquinasPorSetor)) {
      const total = Object.values(tipos).reduce((sum, count) => sum + count, 0);
      dadosMaquinasPorSetor.push([setor, total]);
    }
  
    // Adicionar dados ao dashboard
    let rowSetor = row;
    planilhaDashboard.getRange(`A${rowSetor}`).setValue("Máquinas por Setor");
    rowSetor++;
    dadosMaquinasPorSetor.forEach(([setor, quantidade]) => {
      planilhaDashboard.getRange(`A${rowSetor}`).setValue(setor);
      planilhaDashboard.getRange(`B${rowSetor}`).setValue(quantidade);
      rowSetor++;
    });
  
    // Adicionar informações das máquinas em uso por setor
    row += 2; // Adiciona duas linhas em branco para separação
    planilhaDashboard.getRange(`A${row}`).setValue("Máquinas em Uso por Setor");
    row++;
  
    const maquinasEmUsoPorSetor = {};
    dadosControle.slice(1).forEach(linha => {
      const status = linha[8]; // Assumindo que o status está na coluna I (índice 8)
      if (status === "Em Uso") {
        const setor = linha[4]; // Assumindo que o setor está na coluna E (índice 4)
        const tipo = linha[1]; // Assumindo que o tipo de máquina está na coluna B (índice 1)
        if (!maquinasEmUsoPorSetor[setor]) {
          maquinasEmUsoPorSetor[setor] = {};
        }
        maquinasEmUsoPorSetor[setor][tipo] = (maquinasEmUsoPorSetor[setor][tipo] || 0) + 1;
      }
    });
  
    for (const [setor, tipos] of Object.entries(maquinasEmUsoPorSetor)) {
      planilhaDashboard.getRange(`A${row}`).setValue(setor);
      row++;
      for (const [tipo, quantidade] of Object.entries(tipos)) {
        planilhaDashboard.getRange(`A${row}`).setValue(`  ${tipo}:`);
        planilhaDashboard.getRange(`B${row}`).setValue(quantidade);
        row++;
      }
      row++; // Adiciona uma linha em branco entre setores
    }
  
    // Adicionar dados para o gráfico de máquinas por setor
    row += 2;
    planilhaDashboard.getRange(`A${row}`).setValue("Dados para Gráfico de Máquinas por Setor");
    row++;
  
    // Usar o Setor.Service para obter a lista oficial de setores
    const setoresFixos = obterSetores();
    
    // Preparar dados para o gráfico
    const dadosGrafico = [["Setor", "Quantidade"]];
    setoresFixos.forEach(setor => {
      const quantidade = maquinasEmUsoPorSetor[setor] ? 
        Object.values(maquinasEmUsoPorSetor[setor]).reduce((a, b) => a + b, 0) : 0;
      dadosGrafico.push([setor, quantidade]);
    });
  
    // Inserir dados na planilha
    const rangeGrafico = planilhaDashboard.getRange(row, 1, dadosGrafico.length, 2);
    rangeGrafico.setValues(dadosGrafico);
  
    // Atualizar a linha atual após inserir os dados
    row += dadosGrafico.length + 2;
  
    // Formatar a planilha
    planilhaDashboard.autoResizeColumns(1, 2);
  }
  
  /**
   * Cria ou atualiza o gatilho para atualização automática do dashboard
   * Configura para atualizar a cada 5 minutos
   */
  function criarGatilhoAtualizacaoDashboard() {
    // Remove gatilhos existentes para evitar duplicação
    var gatilhos = ScriptApp.getProjectTriggers();
    for (var i = 0; i < gatilhos.length; i++) {
      if (gatilhos[i].getHandlerFunction() == 'atualizarDashboard') {
        ScriptApp.deleteTrigger(gatilhos[i]);
      }
    }
    
    // Cria novo gatilho para atualização a cada 5 minutos
    ScriptApp.newTrigger('atualizarDashboard')
      .timeBased()
      .everyMinutes(5)
      .create();
  }
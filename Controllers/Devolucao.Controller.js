/**
 * Registra a devolução de uma máquina no sistema
 * @param {string|number} idMaquina - ID da máquina a ser devolvida
 * @returns {Object} Objeto com tipo e mensagem do resultado da operação
 */
function registrarDevolucao(idMaquina) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilhaOperacao = ss.getSheetByName("OPERACAO");
    const planilhaRegistros = ss.getSheetByName("REGISTROS");
  
    // Obtém dados da planilha de operação
    const dadosOperacao = planilhaOperacao.getRange("A2:I").getValues();
    let linhaMaquina = -1;
  
    // Localiza a máquina em uso ou pendente
    for (let i = 0; i < dadosOperacao.length; i++) {
      if (dadosOperacao[i][0] == idMaquina && 
          (dadosOperacao[i][8] === "Em Uso" || dadosOperacao[i][8] === "Pendente")) {
        linhaMaquina = i + 2;
        break;
      }
    }
  
    // Verifica se encontrou a máquina
    if (linhaMaquina === -1) {
      return { 
        tipo: "erro", 
        mensagem: "Erro: Máquina com ID " + idMaquina + " não encontrada em uso ou pendente!" 
      };
    }
  
    const dataAtual = new Date();
    const dataDevolucao = new Date();
  
    // Atualiza informações na planilha de operação
    planilhaOperacao.getRange(linhaMaquina, 8)
                   .setValue(dataDevolucao)
                   .setNumberFormat("dd/MM/yyyy HH:mm:ss");
    planilhaOperacao.getRange(linhaMaquina, 9).setValue("Disponível");
  
    // Prepara dados para histórico
    const dadosSaida = planilhaOperacao.getRange(linhaMaquina, 1, 1, 9).getValues()[0];
  
    // Organiza dados para registro histórico
    const registro = [
      dadosSaida[0], // ID MAQUINA
      dadosSaida[1], // TIPO DE DISPOSITIVO
      dadosSaida[2], // MATRICULA
      dadosSaida[3], // COLABORADOR
      dadosSaida[4], // SETOR
      dadosSaida[5], // DATA EMPRESTIMO
      dadosSaida[6], // HORA EMPRESTIMO
      dataDevolucao, // DATA DEVOLUCAO
      dataAtual,     // HORA DEVOLUCAO
      "Disponível"   // STATUS
    ];
  
    // Adiciona ao histórico e remove da operação atual
    planilhaRegistros.appendRow(registro);
    planilhaOperacao.deleteRow(linhaMaquina);
  
    // Atualiza status da máquina
    atualizarStatusMaquina(idMaquina, "Disponível");
  
    // Atualiza controle de máquinas por colaborador
    const scriptProperties = PropertiesService.getScriptProperties();
    let maquinasEmUsoPorColaborador = JSON.parse(
      scriptProperties.getProperty('maquinasEmUsoPorColaborador') || '{}'
    );
  
    // Atualiza registro do colaborador
    const matricula = dadosSaida[2];
    if (maquinasEmUsoPorColaborador[matricula]) {
      const tipoMaquina = dadosSaida[1];
      
      // Remove máquina da lista do colaborador
      const index = maquinasEmUsoPorColaborador[matricula].indexOf(tipoMaquina);
      if (index > -1) {
        maquinasEmUsoPorColaborador[matricula].splice(index, 1);
      }
  
      // Remove colaborador se não tiver mais máquinas
      if (maquinasEmUsoPorColaborador[matricula].length === 0) {
        delete maquinasEmUsoPorColaborador[matricula];
      }
  
      // Atualiza propriedades do script
      scriptProperties.setProperty(
        'maquinasEmUsoPorColaborador', 
        JSON.stringify(maquinasEmUsoPorColaborador)
      );
    }
  
    return { 
      tipo: "sucesso", 
      mensagem: `Devolução da máquina ${idMaquina} registrada com sucesso para o colaborador: ${dadosSaida[3]}` 
    };
  }
  
  /**
   * Atualiza as propriedades do script após uma devolução
   * @param {string} matricula - Matrícula do colaborador
   * @param {string} tipoMaquina - Tipo da máquina devolvida
   */
  function atualizarPropriedadesDevolucao(matricula, tipoMaquina) {
    const scriptProperties = PropertiesService.getScriptProperties();
    let maquinasEmUsoPorColaborador = JSON.parse(scriptProperties.getProperty('maquinasEmUsoPorColaborador') || '{}');
  
    if (maquinasEmUsoPorColaborador[matricula]) {
      // Remove o tipo da máquina da lista
      maquinasEmUsoPorColaborador[matricula] = maquinasEmUsoPorColaborador[matricula].filter(tipo => tipo !== tipoMaquina);
      
      // Se não houver mais máquinas emprestadas, remova a matrícula
      if (maquinasEmUsoPorColaborador[matricula].length === 0) {
        delete maquinasEmUsoPorColaborador[matricula];
      }
      
      scriptProperties.setProperty('maquinasEmUsoPorColaborador', JSON.stringify(maquinasEmUsoPorColaborador));
    }
  }
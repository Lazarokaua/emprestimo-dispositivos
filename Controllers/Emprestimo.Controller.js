/**
 * Sincroniza o objeto maquinasEmUsoPorColaborador com os dados reais da planilha
 */
function sincronizarMaquinasEmUso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaControle = ss.getSheetByName("Controle");
  const dadosControle = planilhaControle.getRange("A2:I").getValues();
  
  // Reinicia o objeto de Controle
  let maquinasEmUsoPorColaborador = {};
  
  // Percorre todas as máquinas em uso na planilha
  dadosControle.forEach(row => {
    if (row[8] === "Em Uso") { // Se status é "Em Uso"
      const matricula = row[2];  // Matrícula do colaborador
      const tipoMaquina = row[1]; // Tipo da máquina
      
      if (!maquinasEmUsoPorColaborador[matricula]) {
        maquinasEmUsoPorColaborador[matricula] = [];
      }
      
      if (!maquinasEmUsoPorColaborador[matricula].includes(tipoMaquina)) {
        maquinasEmUsoPorColaborador[matricula].push(tipoMaquina);
      }
    }
  });
  
  // Atualiza as propriedades do script
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('maquinasEmUsoPorColaborador', JSON.stringify(maquinasEmUsoPorColaborador));
  
  return maquinasEmUsoPorColaborador;
}

/**
 * Registra o empréstimo de uma máquina para um colaborador
 * @param {string|number} matricula - Matrícula do colaborador
 * @param {string|number} idMaquina - ID da máquina a ser emprestada
 * @returns {Object} Objeto com tipo e mensagem do resultado da operação
 */
function registrarEmprestimo(matricula, idMaquina) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      if (!matricula || !idMaquina) {
        return { 
          tipo: "erro", 
          mensagem: "Os campos Matrícula e ID Máquina precisam ser preenchidos!" 
        };
      }

      matricula = String(matricula).padStart(4, '0');
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const planilhaControle = ss.getSheetByName("Controle");
      const scriptProperties = PropertiesService.getScriptProperties();
      
      // Verifica se a planilha de operação existe
      if (!planilhaControle) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: Planilha de operação não encontrada!" 
        };
      }

      // Verificações de máquina e disponibilidade
      if (!verificarMaquinaCadastrada(idMaquina)) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: Máquina com ID " + idMaquina + " não cadastrada!" 
        };
      }

      if (verificarMaquinaEmUso(idMaquina)) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: Máquina com ID " + idMaquina + " já está em uso!" 
        };
      }

      // Obtém informações da máquina e do colaborador
      const infoMaquina = obterInformacoesMaquina(idMaquina);
      if (!infoMaquina) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: Informações da máquina não encontradas!" 
        };
      }

      const infoColaborador = obterInformacoesColaborador(matricula);
      if (!infoColaborador) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: Colaborador com matrícula " + matricula + " não encontrado!" 
        };
      }

      // Sincroniza o objeto com os dados reais da planilha
      const maquinasEmUsoPorColaborador = sincronizarMaquinasEmUso();

      if (maquinasEmUsoPorColaborador[matricula] && 
          maquinasEmUsoPorColaborador[matricula].includes(infoMaquina.tipo)) {
        return { 
          tipo: "erro", 
          mensagem: "Erro: O colaborador já possui uma máquina do tipo " + infoMaquina.tipo + " emprestada!" 
        };
      }

      const dataAtual = new Date();

      // Prepara os dados do novo registro
      const novoRegistro = [
        idMaquina,           // ID da máquina
        infoMaquina.tipo,    // Tipo da máquina
        matricula,           // Matrícula do colaborador
        infoColaborador.nome,// Nome do colaborador
        infoColaborador.setor,// Setor do colaborador
        dataAtual,           // Data do empréstimo
        dataAtual,           // Hora do empréstimo
        "",                  // Hora da devolução (vazio inicialmente)
        "Em Uso"             // Status inicial
      ];

      // Insere o novo registro na planilha
      planilhaControle.insertRowsBefore(2, 1);
      const range = planilhaControle.getRange(2, 1, 1, 9);
      range.setValues([novoRegistro]);

      // Atualiza o status da máquina
      atualizarStatusMaquina(idMaquina, "Em Uso");

      // Atualiza o registro de máquinas emprestadas pelo colaborador
      if (!maquinasEmUsoPorColaborador[matricula]) {
        maquinasEmUsoPorColaborador[matricula] = [];
      }
      maquinasEmUsoPorColaborador[matricula].push(infoMaquina.tipo);
      scriptProperties.setProperty(
        'maquinasEmUsoPorColaborador', 
        JSON.stringify(maquinasEmUsoPorColaborador)
      );

      // Retorna sucesso com informações do empréstimo
      return { 
        tipo: "sucesso", 
        mensagem: `Empréstimo da máquina ${idMaquina} (${infoMaquina.tipo}) registrado com sucesso para o colaborador: ${infoColaborador.nome}`,
        nomeColaborador: infoColaborador.nome
      };

    } catch (error) {
      // Captura e retorna qualquer erro ocorrido
      return { 
        tipo: "erro", 
        mensagem: "Erro ao processar empréstimo: " + error.toString() 
      };
    } finally {
      // Sempre libera o lock ao finalizar
      lock.releaseLock();
    }
}

function testarSincronizacao() {
  const resultado = sincronizarMaquinasEmUso();
  console.log('Estado atual das máquinas por colaborador:', resultado);
}

function testarDessincronizacao() {
  // Força uma dessincronização
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('maquinasEmUsoPorColaborador', JSON.stringify({
    "1234567": ["Paleteira"] // Força um registro falso
  }));
  
  // Tenta emprestar uma Paleteira
  const resultado = registrarEmprestimo("1234567", "1234"); // ID de uma Paleteira
  console.log('Resultado do teste:', resultado);
}

function verificarEstadoAtual() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const estado = scriptProperties.getProperty('maquinasEmUsoPorColaborador');
  Logger.log('Estado atual: ' + estado); // Isso vai aparecer nos logs da planilha
  return estado;
}

function limparEstado() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('maquinasEmUsoPorColaborador', '{}');
  Logger.log('Estado limpo com sucesso');
  return 'Estado limpo';
}

function verificarInconsistencias() {
  // Pega o estado atual da propriedade do script
  const scriptProperties = PropertiesService.getScriptProperties();
  const estadoAtual = JSON.parse(scriptProperties.getProperty('maquinasEmUsoPorColaborador') || '{}');
  
  // Pega os dados reais da planilha
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaControle = ss.getSheetByName("Controle");
  const dadosControle = planilhaControle.getRange("A2:I").getValues();
  
  // Monta objeto com dados reais da planilha
  const estadoReal = {};
  dadosControle.forEach(row => {
    if (row[8] === "Em Uso") {
      const matricula = row[2];
      const tipoMaquina = row[1];
      
      if (!estadoReal[matricula]) {
        estadoReal[matricula] = [];
      }
      if (!estadoReal[matricula].includes(tipoMaquina)) {
        estadoReal[matricula].push(tipoMaquina);
      }
    }
  });
  
  // Procura inconsistências
  const inconsistencias = {
    registrosExtras: {}, // Registros que existem na propriedade mas não na planilha
    registrosFaltando: {} // Registros que existem na planilha mas não na propriedade
  };
  
  // Verifica registros extras (na propriedade mas não na planilha)
  Object.keys(estadoAtual).forEach(matricula => {
    if (!estadoReal[matricula]) {
      inconsistencias.registrosExtras[matricula] = estadoAtual[matricula];
    } else {
      estadoAtual[matricula].forEach(tipo => {
        if (!estadoReal[matricula].includes(tipo)) {
          if (!inconsistencias.registrosExtras[matricula]) {
            inconsistencias.registrosExtras[matricula] = [];
          }
          inconsistencias.registrosExtras[matricula].push(tipo);
        }
      });
    }
  });
  
  // Verifica registros faltando (na planilha mas não na propriedade)
  Object.keys(estadoReal).forEach(matricula => {
    if (!estadoAtual[matricula]) {
      inconsistencias.registrosFaltando[matricula] = estadoReal[matricula];
    } else {
      estadoReal[matricula].forEach(tipo => {
        if (!estadoAtual[matricula].includes(tipo)) {
          if (!inconsistencias.registrosFaltando[matricula]) {
            inconsistencias.registrosFaltando[matricula] = [];
          }
          inconsistencias.registrosFaltando[matricula].push(tipo);
        }
      });
    }
  });
  
  Logger.log('Estado na propriedade do script:');
  Logger.log(estadoAtual);
  Logger.log('Estado real na planilha:');
  Logger.log(estadoReal);
  Logger.log('Inconsistências encontradas:');
  Logger.log(inconsistencias);
  
  return {
    estadoAtual: estadoAtual,
    estadoReal: estadoReal,
    inconsistencias: inconsistencias
  };
}

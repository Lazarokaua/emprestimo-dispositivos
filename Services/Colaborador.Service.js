/**
 * Busca e retorna as informações de um colaborador com base na matrícula
 * @param {string|number} matricula - Matrícula do colaborador a ser buscado
 * @returns {Object|null} Objeto com informações do colaborador ou null se não encontrado
 */
function obterInformacoesColaborador(matricula) {
    // Obtém a planilha ativa e a aba de colaboradores
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilhaColaboradores = ss.getSheetByName("COLABORADORES");
    
    // Obtém todos os dados dos colaboradores (a partir da linha 2 até o final)
    const dadosColaboradores = planilhaColaboradores.getRange("A2:C").getValues();
    
    // Padroniza a matrícula recebida para ter 7 dígitos (adiciona zeros à esquerda)
    matricula = String(matricula).padStart(7, '0');
    
    // Percorre todos os registros procurando pela matrícula
    for (let row of dadosColaboradores) {
      // Padroniza a matrícula da planilha para comparação
      let matriculaPlanilha = String(row[0]).padStart(7, '0');
      
      // Se encontrar a matrícula, retorna os dados do colaborador
      if (matriculaPlanilha === matricula) {
        return { 
          nome: row[1],        // Nome do colaborador
          setor: row[2],       // Setor do colaborador
          matricula: matriculaPlanilha // Matrícula padronizada
        };
      }
    }
    
    // Retorna null se não encontrar o colaborador
    return null;
}
/**
 * Busca e retorna as informações de um colaborador com base na matrícula
 * @param {string|number} matricula - Matrícula do colaborador a ser buscado
 * @returns {Object|null} Objeto com informações do colaborador ou null se não encontrado
 */

function obterInformacoesColaborador(matricula){
    // obtendo a planilha ativa
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // aba Colaboradores 
    const planilhaColaboradores = ss.getSheetByName("Colaboradores");

   // Obtém todos os dados dos Colaboradores (a partir da linha 2 até o final)
   const dataColaboradores = planilhaColaboradores.getRange("A2:C").getValues();

   // padroniza a matricula recebida
   matricula = String(matricula).padStart(4, '0');

   // Percorre todos os registros procurando pela matrícula
   for (let row of dataColaboradores){
    let matriculaPlanilha = String(row[0]).padStart(4, '0');

    if (matricula === matriculaPlanilha){
      return {
        nome: row[1],
        setor: row[2],
        matricula: matriculaPlanilha

      };
    }
   }
}
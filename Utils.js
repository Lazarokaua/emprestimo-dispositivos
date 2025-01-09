/**
 * Calcula a diferença em horas entre duas datas
 * @param {Date} dataInicial - Data inicial
 * @param {Date} dataFinal - Data final
 * @returns {number} Diferença em horas
 */
function calcularDiferencaEmHoras(dataInicial, dataFinal) {
  const diferencaEmMilissegundos = dataFinal - dataInicial;
  const diferencaEmHoras = diferencaEmMilissegundos / (1000 * 60 * 60);
  return diferencaEmHoras;
}

/**
 * Verifica e atualiza o status de máquinas que estão em uso por mais de 12 horas
 * @param {Date} dataInicial - Data inicial para verificação
 * @param {Date} dataFinal - Data final para verificação
 */
function verificarMaquinasPendentes(dataInicial, dataFinal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaOperacao = ss.getSheetByName("OPERACAO"); // Planilha de Operação
  const ultimaLinha = planilhaOperacao.getLastRow(); // Última linha da planilha

  for (let i = 2; i <= ultimaLinha; i++) {
    // Início do loop a partir da segunda linha
    const dataEmprestimo = planilhaOperacao.getRange(i, 7).getValue(); // Coluna 7: Hora Empréstimo
    const status = planilhaOperacao.getRange(i, 9).getValue(); // Coluna 9: Status

    if (status === "Em Uso") {
      // Verifica se o status é "Em Uso"
      const dataAtual = new Date();
      const diferencaEmHoras = calcularDiferencaEmHoras(
        dataEmprestimo,
        dataAtual
      );

      if (diferencaEmHoras >= 12) {
        // Verifica se a diferença em horas é maior ou igual a 12
        planilhaOperacao.getRange(i, 9).setValue("Pendente");
      }
    }
  }
}

/**
 * Gera um relatório em Excel dos registros do sistema
 * Cria uma planilha temporária, converte para Excel e disponibiliza para download
 */
function gerarRelatorioExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaRegistros = ss.getSheetByName("REGISTROS");
  
  // Cria uma nova planilha temporária para o relatório
  const tempSpreadsheet = SpreadsheetApp.create("Relatório_Registros");
  
  // Copia os dados da planilha de registros para a planilha temporária
  const dados = planilhaRegistros.getDataRange().getValues();
  const tempSheet = tempSpreadsheet.getSheets()[0];
  tempSheet.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
  
  // Converte para Excel usando o MimeType correto
  const url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
  
  // Cria uma interface HTML com um link de download
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <html>
        <body>
          <script>
            function downloadExcel() {
              window.open('${url}', '_blank');
              google.script.host.close();
            }
            window.onload = downloadExcel;
          </script>
          <p>Se o download não iniciar automaticamente, 
             <a href="#" onclick="downloadExcel(); return false;">clique aqui</a>.</p>
        </body>
      </html>
    `)
    .setWidth(300)
    .setHeight(100);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download do Relatório');
  
  // Exclui a planilha temporária do Google Sheets após um breve delay
  Utilities.sleep(1000);
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
}

function getDownloadUrl(fileId) {
  return DriveApp.getFileById(fileId).getDownloadUrl();
}

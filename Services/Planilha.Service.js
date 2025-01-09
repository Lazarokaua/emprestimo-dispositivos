/**
 * Cria o menu personalizado na interface do Google Sheets quando a planilha é aberta
 * Adiciona opções para empréstimo/devolução, atualização do dashboard e download de relatório
 */
function onOpen() {
  
  // Cria o menu
  SpreadsheetApp.getUi()
    .createMenu("Gerenciamento de Dispositivos")
    .addItem("Empréstimo/Devolução", "abrirInterface")
    .addItem("Abrir Interface Web", "abrirInterfaceWeb")
    .addItem("Atualizar Dashboard", "atualizarDashboard")
    .addItem("Baixar Relatório Excel", "gerarRelatorioExcel")
    .addSeparator() // Adiciona uma linha separadora
    .addItem("Verificar e Corrigir Inconsistências", "verificarECorrigirInconsistencias")
    .addSeparator()
    .addItem("Enviar Feedback/Sugestão", "abrirFormularioFeedback")
    .addToUi();
}

/**
 * Abre a interface modal para empréstimo e devolução de máquinas
 * Define dimensões e título da janela modal
 */
function abrirInterface() {
    const html = HtmlService.createHtmlOutputFromFile("interface")
      .setWidth(800)
      .setHeight(600)
      .setTitle("Gerenciamento de Máquinas")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, "Gerenciamento de Máquinas");
}

/**
 * Remove registros com status "Disponível" da planilha de operação
 * Limpa registros desnecessários mantendo apenas máquinas em uso
 */
function limparConteudo() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilha = ss.getSheetByName("OPERACAO");
    const valores = planilha.getRange("I:I").getValues();

    // Remove linhas com status "Disponível" de baixo para cima
    for (let i = valores.length - 1; i >= 1; i--) {
      if (valores[i][0] === "Disponível") {
        planilha.deleteRow(i + 1);
      }
    }
}

/**
 * Remove registros vazios ou com erros da planilha de operação
 * Solicita confirmação do usuário antes de executar a limpeza
 */
function limparConteudoTodo() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilha = ss.getSheetByName("OPERACAO");

    // Solicita confirmação do usuário
    const resposta = Browser.msgBox(
      "CONFIRMAÇÃO",
      "Tem certeza que quer apagar os erros da tabela?",
      Browser.Buttons.YES_NO
    );

    if (resposta === "yes") {
      const dados = planilha.getRange("A2:I").getValues();
      // Remove linhas vazias ou com erros de baixo para cima
      for (let i = dados.length - 1; i >= 0; i--) {
        if (dados[i][2] === "" && dados[i][0] === "" && 
            dados[i][7] === "" && dados[i][8] === "") {
          planilha.deleteRow(i + 2);
        }
      }
    }
}


// Adicionar nova função para abrir a interface web
function abrirInterfaceWeb() {
  const url = 'https://script.google.com/macros/s/AKfycbyQeOnfYN5Rar2VhoM5BgzkAxNFyYsoGfWARMfFtYN6b8pEIfztlbjvCmx5ANwxyOI/exec';
  const html = HtmlService.createHtmlOutput(
    `<script>
      window.open('${url}', '_blank');
    </script>`
  )
  .setWidth(100)
  .setHeight(50);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo interface...');
}

function abrirFormularioFeedback() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; }
      input, textarea { width: 100%; padding: 8px; margin-bottom: 10px; }
      textarea { height: 150px; }
      button { 
        background-color: #4CAF50; 
        color: white; 
        padding: 10px 20px; 
        border: none; 
        cursor: pointer; 
      }
      button:hover { background-color: #45a049; }
    </style>
    <div class="form-group">
      <label for="assunto">Assunto:</label>
      <input type="text" id="assunto" placeholder="Bug, Sugestão, Dúvida, etc.">
    </div>
    <div class="form-group">
      <label for="mensagem">Mensagem:</label>
      <textarea id="mensagem" placeholder="Descreva em detalhes sua sugestão ou o problema encontrado..."></textarea>
    </div>
    <button onclick="enviarFeedback()">Enviar Feedback</button>
    
    <script>
      function enviarFeedback() {
        const assunto = document.getElementById('assunto').value;
        const mensagem = document.getElementById('mensagem').value;
        
        if (!assunto || !mensagem) {
          alert('Por favor, preencha todos os campos.');
          return;
        }
        
        google.script.run
          .withSuccessHandler(function() {
            alert('Feedback enviado com sucesso! Obrigado pela sua contribuição.');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Erro ao enviar feedback: ' + error);
          })
          .enviarEmailFeedback(assunto, mensagem);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(400)
  .setTitle('Enviar Feedback');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Feedback');
}

function enviarEmailFeedback(assunto, mensagem) {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const developerEmail = 'lazarokaua22@gmail.com';
    
    const emailBody = `
      Novo feedback recebido:
      
      De: ${userName} (${userEmail})
      Assunto: ${assunto}
      
      Mensagem:
      ${mensagem}
      
      Data: ${new Date().toLocaleString('pt-BR')}
    `;
    
    GmailApp.sendEmail(
      developerEmail,
      `[Feedback] ${assunto} - Gerenciamento de Dispositivos`,
      emailBody
    );
  } catch (error) {
    Logger.log('Erro ao enviar email: ' + error.toString());
    throw new Error('Não foi possível enviar o feedback. Por favor, tente novamente mais tarde.');
  }
}

// function verificarAcessoUsuario() {
//   const usuarioAtual = Session.getEffectiveUser().getEmail();
//   const usuariosAutorizados = [
//     'usuario1@empresa.com',
//     'usuario2@empresa.com'
//     // ... adicione mais emails autorizados
//   ];
  
//   return usuariosAutorizados.includes(usuarioAtual);
// }

function verificarECorrigirInconsistencias() {
  const ui = SpreadsheetApp.getUi();
  
  // Verifica inconsistências
  const resultado = verificarInconsistencias();
  const {estadoAtual, estadoReal, inconsistencias} = resultado;
  
  // Se não houver inconsistências
  if (Object.keys(inconsistencias.registrosExtras).length === 0 && 
      Object.keys(inconsistencias.registrosFaltando).length === 0) {
    ui.alert(
      'Verificação Concluída',
      'Não foram encontradas inconsistências no sistema.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Mostra as inconsistências encontradas
  let mensagem = 'Foram encontradas as seguintes inconsistências:\n\n';
  
  if (Object.keys(inconsistencias.registrosExtras).length > 0) {
    mensagem += 'Registros extras (serão removidos):\n';
    Object.entries(inconsistencias.registrosExtras).forEach(([matricula, tipos]) => {
      mensagem += `- Matrícula ${matricula}: ${tipos.join(', ')}\n`;
    });
  }
  
  // Pergunta se deseja corrigir
  const resposta = ui.alert(
    'Inconsistências Encontradas',
    mensagem + '\nDeseja corrigir estas inconsistências agora?',
    ui.ButtonSet.YES_NO
  );
  
  if (resposta === ui.Button.YES) {
    // Corrige sincronizando com o estado real
    sincronizarMaquinasEmUso();
    
    ui.alert(
      'Correção Concluída',
      'As inconsistências foram corrigidas com sucesso.',
      ui.ButtonSet.OK
    );
  }
}
function doGet(e) {
    try {
      const template = HtmlService.createTemplateFromFile('interface');
      const html = template.evaluate()
        .setTitle('Gerenciamento de Máquinas')
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');
      
      return html;
    } catch (error) {
      Logger.log('Erro em doGet: ' + error.toString());
      return HtmlService.createHtmlOutput(
        '<h1>Erro ao carregar a aplicação</h1>' +
        '<p>Por favor, contate o administrador.</p>' +
        '<p>Erro: ' + error.toString() + '</p>'
      );
    }
  }
  
  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } 
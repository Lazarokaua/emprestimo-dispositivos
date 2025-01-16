function  doGet(){
  return HtmlService.createTemplateFromFile('interface').evaluate();
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
    .addItem('interface', 'abrirInterface')
    .addToUi()
}

function abrirInterface() {
  const html = HtmlService.createHtmlOutputFromFile("interface")
    .setWidth(800)
    .setHeight(600)
    .setTitle("Gerenciador de Dispositivos")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, "Gerenciador de Dispositivos");
}


function getProjectId() {
  const projectId = PropertiesService.getScriptProperties().getProperty('PROJECT_ID');
  return projectId;
}
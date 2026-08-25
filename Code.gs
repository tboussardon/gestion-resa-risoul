function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestion Risoul')
    .addItem('💰 Gérer les Tarifs', 'afficherInterfaceTarifs')
    .addItem('📅 Gérer les Réservations', 'afficherInterfaceResa')
    .addToUi();
}

function afficherInterfaceTarifs() {
  const html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestion des Tarifs');
}

function afficherInterfaceResa() {
  // Charge un nouveau fichier que nous allons créer : Resa.html
  const html = HtmlService.createTemplateFromFile('Resa')
      .evaluate()
      .setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestion des Réservations');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
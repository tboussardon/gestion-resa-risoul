function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestion Risoul')
    .addItem('💰 Gérer les Tarifs', 'afficherInterfaceTarifs')
    .addItem('📅 Gérer les Réservations', 'afficherInterfaceResa')
    .addItem('🏔️ Gérer la Station', 'afficherInterfaceStation')
    .addSeparator()
    .addItem('📊 Synthèse & Rapports', 'afficherInterfaceSynthese') // NOUVEAU
    .addToUi();
}

function afficherInterfaceSynthese() {
  const html = HtmlService.createTemplateFromFile('Synthese')
      .evaluate()
      .setWidth(1250).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Synthèse & Rapports');
}

function afficherInterfaceTarifs() {
  const html = HtmlService.createTemplateFromFile('Tarifs')
      .evaluate()
      .setWidth(1150).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 Gérer les Tarifs');
}

function afficherInterfaceResa() {
  const html = HtmlService.createTemplateFromFile('Resa')
      .evaluate()
      .setWidth(1150).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Gérer les Réservations');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getConciergeries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleListes = ss.getSheetByName('listes');
  if (!feuilleListes) return [];
  
  const data = feuilleListes.getDataRange().getValues();
  let conciergeries = [];
  
  // Parcours à partir de la ligne 2 (index 1) pour ignorer l'en-tête
  for (let i = 1; i < data.length; i++) {
    let val = data[i][0]; // Colonne A = index 0
    if (val) {
      conciergeries.push(val.toString().trim());
    }
  }
  return [...new Set(conciergeries)]; // Déduplication
}

function afficherInterfaceResa() {
  const html = HtmlService.createTemplateFromFile('Resa')
      .evaluate()
      .setWidth(1150).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Gérer les Réservations');
}
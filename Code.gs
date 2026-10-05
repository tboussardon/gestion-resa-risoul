function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestion Risoul')
    .addItem('💰 Gérer les Tarifs', 'afficherInterfaceTarifs')
    .addItem('📅 Gérer les Réservations', 'afficherInterfaceResa')
    .addItem('🏔️ Gérer la Station', 'afficherInterfaceStation')
    .addSeparator()
    .addItem('📊 Synthèse & Rapports', 'afficherInterfaceSynthese')
    .addToUi();
}

function afficherInterfaceStation() {
  const html = HtmlService.createTemplateFromFile('Station')
      .evaluate()
      .setWidth(1150).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '🏔️ Gérer la Station');
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
  
  for (let i = 1; i < data.length; i++) {
    let val = data[i][0];
    if (val) {
      conciergeries.push(val.toString().trim());
    }
  }
  return [...new Set(conciergeries)];
}

/**
 * Fonction appelée automatiquement lors de l'ouverture de l'URL de la Web App.
 */
function doGet(e) {
  // 1. Définir la page par défaut à afficher (Index = Tarifs)
  let pageAfficher = 'Index'; 
  
  // 2. Si l'URL contient un paramètre "page" (ex: URL_DU_SCRIPT?page=Resa), on le récupère
  if (e && e.parameter && e.parameter.page) {
    pageAfficher = e.parameter.page;
  }

  // 3. Sécurité : Vérifier que la page demandée existe bien parmi vos fichiers
  const pagesValides = ['Index', 'Resa', 'Station', 'Synthese'];
  if (!pagesValides.includes(pageAfficher)) {
    pageAfficher = 'Index'; // Retour à l'accueil en cas d'erreur
  }

  // 4. Générer le HTML depuis le fichier correspondant
  let htmlOutput = HtmlService.createTemplateFromFile(pageAfficher).evaluate();
  
  // 5. Paramétrages de la page web (Titre de l'onglet et adaptation aux smartphones)
  htmlOutput.setTitle('Gestion Risoul - ' + pageAfficher)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
            
  return htmlOutput;
}

/**
 * Fonction utilitaire pour récupérer l'URL de notre propre Web App
 * Cela permet de générer les liens du menu de navigation.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
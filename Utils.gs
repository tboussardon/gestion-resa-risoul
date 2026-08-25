function genererIdUnique() {
  const caracteres = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let id = '';
  for (let i = 0; i < 5; i++) {
    id += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }
  return id;
}

function getListesDeroulantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleListes = ss.getSheetByName('listes');
  const data = feuilleListes.getDataRange().getValues();
  
  let saisons = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) saisons.push(data[i][3]);
  }
  return {
    saisons: [...new Set(saisons)] 
  };
}
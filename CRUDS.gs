function getSheetData(nomFeuille) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille).getDataRange().getValues();
}

function rechercherEnregistrement(nomFeuille, idColIndex, idValue) {
  const data = getSheetData(nomFeuille);
  const entetes = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColIndex] === idValue) {
      let objetResultat = {};
      entetes.forEach((entete, index) => {
        objetResultat[entete] = data[i][index];
      });
      return { row: i + 1, data: objetResultat };
    }
  }
  return null; 
}

function creerEnregistrement(nomFeuille, idColName, formData) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  const entetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  
  const nouvelId = genererIdUnique();
  formData[idColName.trim()] = nouvelId; 
  
  let nouvelleLigne = entetes.map(entete => {
    let nomColonnePropre = entete.toString().trim();
    if (nomColonnePropre === idColName.trim()) {
      return nouvelId;
    }
    return formData[nomColonnePropre] !== undefined ? formData[nomColonnePropre] : "";
  });
  
  feuille.appendRow(nouvelleLigne);
  return nouvelId; 
}

function modifierEnregistrement(nomFeuille, rowIndex, formData) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  const entetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  
  let ligneMaj = entetes.map(entete => {
    let nomColonnePropre = entete.toString().trim(); 
    return formData[nomColonnePropre] !== undefined ? formData[nomColonnePropre] : "";
  });
  
  feuille.getRange(rowIndex, 1, 1, ligneMaj.length).setValues([ligneMaj]);
  return true;
}

function supprimerEnregistrement(nomFeuille, rowIndex) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  feuille.deleteRow(rowIndex);
  return true;
}

function rechercherEnregistrementParDate(nomFeuille, nomColonne, dateSearchStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName(nomFeuille);
  const data = feuille.getDataRange().getValues();
  const entetes = data[0];
  
  let colIndex = -1;
  for (let c = 0; c < entetes.length; c++) {
    if (entetes[c].toString().trim() === nomColonne) {
      colIndex = c;
      break;
    }
  }
  
  if (colIndex === -1) {
    return { erreurDebug: "La colonne '" + nomColonne + "' est introuvable." };
  }

  const timeZone = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    let cellValue = data[i][colIndex];
    if (!cellValue) continue;
    
    let dateStr = "";
    if (Object.prototype.toString.call(cellValue) === '[object Date]') {
      dateStr = Utilities.formatDate(cellValue, timeZone, "yyyy-MM-dd");
    } else {
      let str = cellValue.toString().trim();
      if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
        dateStr = str.substring(0, 10);
      } else if (str.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        let parts = str.substring(0, 10).split('/');
        dateStr = parts[2] + "-" + parts[1] + "-" + parts[0];
      }
    }
    
    if (dateStr === dateSearchStr) {
      let objetResultat = {};
      entetes.forEach((entete, index) => {
        let val = data[i][index];
        if (Object.prototype.toString.call(val) === '[object Date]') {
          objetResultat[entete] = Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
        } else {
          objetResultat[entete] = val;
        }
      });
      return { row: i + 1, data: objetResultat };
    }
  }
  return null;
}

function getInfoCalendrierParDate(dateSearchStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleCal = ss.getSheetByName('cal_26-50');
  if (!feuilleCal) return null;
  
  const data = feuilleCal.getDataRange().getValues();
  const displayData = feuilleCal.getDataRange().getDisplayValues();
  const timeZone = ss.getSpreadsheetTimeZone();
  
  const entetes = data[0];
  let idxSaison = -1;
  for (let c = 0; c < entetes.length; c++) {
    if (entetes[c].toString().trim().toLowerCase() === 'saison') {
      idxSaison = c;
      break;
    }
  }
  
  for (let i = 1; i < data.length; i++) {
    let cellText = displayData[i][3].trim(); 
    if (!cellText) continue;
    
    let datePart = cellText.split(' ')[0];
    let formattedCellText = "";
    
    if (datePart.includes('/')) {
      let parts = datePart.split('/');
      if (parts.length === 3) {
        formattedCellText = parts[2] + "-" + parts[1].padStart(2, '0') + "-" + parts[0].padStart(2, '0');
      }
    } else if (datePart.includes('-')) {
      let parts = datePart.split('-');
      if (parts.length === 3 && parts[0].length === 4) {
        formattedCellText = parts[0] + "-" + parts[1].padStart(2, '0') + "-" + parts[2].padStart(2, '0');
      }
    }
    
    if (!formattedCellText) {
      let cellValue = data[i][3];
      if (Object.prototype.toString.call(cellValue) === '[object Date]') {
        let d = new Date(cellValue.getTime());
        d.setHours(12, 0, 0, 0);
        let year = d.getFullYear();
        let month = String(d.getMonth() + 1).padStart(2, '0');
        let day = String(d.getDate()).padStart(2, '0');
        formattedCellText = year + "-" + month + "-" + day;
      }
    }
    
    if (formattedCellText === dateSearchStr) {
      let rowVals = data[i];
      let dateFinVal = rowVals[4]; 
      let dateFinStr = "";
      
      if (Object.prototype.toString.call(dateFinVal) === '[object Date]') {
        dateFinStr = Utilities.formatDate(dateFinVal, timeZone, "yyyy-MM-dd");
      } else {
        dateFinStr = dateFinVal.toString().substring(0, 10);
      }
      
      return {
        annee: rowVals[1],     
        semaine: rowVals[2],   
        dateFin: dateFinStr,   
        saison: idxSaison !== -1 ? rowVals[idxSaison] : "" 
      };
    }
  }
  return null;
}

function dateExisteDeja(nomFeuille, nomColonne, dateSearchStr, currentRowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName(nomFeuille);
  const data = feuille.getDataRange().getValues();
  const entetes = data[0];
  
  let colIndex = -1;
  for (let c = 0; c < entetes.length; c++) {
    if (entetes[c].toString().trim() === nomColonne) {
      colIndex = c;
      break;
    }
  }
  if (colIndex === -1) return false;

  const timeZone = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    if (currentRowIndex && (i + 1) === parseInt(currentRowIndex)) continue;

    let cellValue = data[i][colIndex];
    if (!cellValue) continue;
    
    let dateStr = "";
    if (Object.prototype.toString.call(cellValue) === '[object Date]') {
      dateStr = Utilities.formatDate(cellValue, timeZone, "yyyy-MM-dd");
    } else {
      let str = cellValue.toString().trim();
      if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
        dateStr = str.substring(0, 10);
      } else if (str.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        let parts = str.substring(0, 10).split('/');
        dateStr = parts[2] + "-" + parts[1] + "-" + parts[0];
      }
    }
    
    if (dateStr === dateSearchStr) return true; 
  }
  return false;
}
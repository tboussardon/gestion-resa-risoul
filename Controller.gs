const CONFIG_STATION = {
  feuille: 'cal_station',
  nomId: 'ID_cal_station',
  colonneRechercheDate: 'ouverture_station'
};

/**
 * Recherche une période station par date d'ouverture
 */
function traiterRechercheStation(dateSaisie) {
  return rechercherEnregistrementParDate(CONFIG_STATION.feuille, CONFIG_STATION.colonneRechercheDate, dateSaisie);
}

/**
 * Sauvegarde (Création ou Modification) d'une période station
 */
function traiterSauvegardeStation(formData, rowIndex) {
  let dateOuverture = formData['ouverture_station'];
  let dateFermeture = formData['fermeture_station'];

  if (!dateOuverture || !dateFermeture) {
    return "ERREUR : Les dates d'ouverture et de fermeture sont obligatoires.";
  }

  // 1. Contrôle : fermeture postérieure ou égale à l'ouverture
  if (dateFermeture < dateOuverture) {
    return "ERREUR : La date de fermeture ne peut pas être antérieure à la date d'ouverture.";
  }

  let infoO = determinerSaisonEtAnnee(dateOuverture);
  let infoF = determinerSaisonEtAnnee(dateFermeture);

  // 2. Contrôle : ouverture dans les créneaux
  if (!infoO) {
    return "ERREUR : La date d'ouverture (" + dateOuverture + ") est hors créneau (Hiver : 01/12-25/04, Été : 20/06-30/09).";
  }

  // 3. Contrôle : fermeture dans les créneaux
  if (!infoF) {
    return "ERREUR : La date de fermeture (" + dateFermeture + ") est hors créneau (Hiver : 01/12-25/04, Été : 20/06-30/09).";
  }

  // 4. Contrôle : même saison et même période
  if (infoO.saison !== infoF.saison || infoO.anneeDebut !== infoF.anneeDebut) {
    return "ERREUR : Les dates d'ouverture et de fermeture doivent appartenir à la même saison.";
  }

  let saisonCalculee = infoO.saison;
  let etatCalcule = "Ouverte";
  let anneesStation = infoO.anneeDebut + '/' + infoO.anneeFin;

  formData['saison_station'] = saisonCalculee;
  formData['saison'] = saisonCalculee;
  formData['etat_station'] = etatCalcule;
  formData['état_station'] = etatCalcule;
  formData['années_station'] = anneesStation;

  // Vérification des doublons sur la date d'ouverture
  let estUnDoublon = dateExisteDeja(CONFIG_STATION.feuille, CONFIG_STATION.colonneRechercheDate, dateOuverture, rowIndex);
  if (estUnDoublon) {
    return "ERREUR : Une période d'ouverture existe déjà pour la date du " + dateOuverture + ".";
  }

  if (rowIndex) {
    if (!formData[CONFIG_STATION.nomId] || formData[CONFIG_STATION.nomId].toString().trim() === "") {
      formData[CONFIG_STATION.nomId] = genererIdUnique();
    }
    modifierEnregistrement(CONFIG_STATION.feuille, rowIndex, formData);
    return "Modification de la période station effectuée avec succès.";
  } else {
    let nouvelId = creerEnregistrement(CONFIG_STATION.feuille, CONFIG_STATION.nomId, formData);
    return "Période station créée avec succès (ID : " + nouvelId + ").";
  }
}

/**
 * Détermine si une date appartient à une saison valide et renvoie les années associées
 */
function determinerSaisonEtAnnee(dateStr) {
  if (!dateStr) return null;
  let parts = dateStr.split('-');
  let a = parseInt(parts[0], 10);
  let m = parseInt(parts[1], 10);
  let j = parseInt(parts[2], 10);

  // Saison Été : 20 juin au 30 septembre
  if ((m === 6 && j >= 20) || (m === 7 || m === 8) || (m === 9 && j <= 30)) {
    return { saison: "Été", anneeDebut: a, anneeFin: a };
  }

  // Saison Hiver : 1er décembre au 25 avril
  if (m === 12 || (m >= 1 && m <= 3) || (m === 4 && j <= 25)) {
    let anneeDebut = (m <= 4) ? a - 1 : a;
    return { saison: "Hiver", anneeDebut: anneeDebut, anneeFin: anneeDebut + 1 };
  }

  return null;
}

/**
 * Suppression d'une période station
 */
function traiterSuppressionStation(rowIndex) {
  supprimerEnregistrement(CONFIG_STATION.feuille, rowIndex);
  return "Période d'ouverture supprimée définitivement.";
}

/**
 * Récupère la liste des saisons depuis la feuille 'listes' (colonne D)
 */
function getSaisonsStation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleListes = ss.getSheetByName('listes');
  if (!feuilleListes) return [];
  
  const data = feuilleListes.getDataRange().getValues();
  let saisons = [];
  
  for (let i = 1; i < data.length; i++) {
    let val = data[i][3]; // Colonne D = index 3
    if (val) {
      saisons.push(val.toString().trim());
    }
  }
  return [...new Set(saisons)];
}

const CONFIG_TARIFS = {
  feuille: 'cal_tarifs_26-50',
  nomId: 'ID_cal_tarifs_26-50',
  colonneRechercheDate: 'Date Début (Samedi)'
};

function traiterRechercheTarif(dateSaisie) {
  return rechercherEnregistrementParDate(CONFIG_TARIFS.feuille, CONFIG_TARIFS.colonneRechercheDate, dateSaisie);
}

function recupererInfosCalendrier(dateSaisie) {
  return getInfoCalendrierParDate(dateSaisie);
}

function traiterSauvegardeTarif(formData, rowIndex) {
  let dateDebut = formData['Date Début (Samedi)'];
  if (!dateDebut) return "ERREUR : La date de début est obligatoire.";

  let estUnDoublon = dateExisteDeja(CONFIG_TARIFS.feuille, CONFIG_TARIFS.colonneRechercheDate, dateDebut, rowIndex);
  if (estUnDoublon) {
    return "ERREUR : Un tarif existe déjà pour cette date de début (" + dateDebut + "). Impossible de créer un doublon.";
  }

  let tarifOt = parseFloat(formData['tarif_ot']) || 0;
  let nomChampLibre = formData.hasOwnProperty('tarifs_libres') ? 'tarifs_libres' : 'tarif_libre';
  let tarifLibre = parseFloat(formData[nomChampLibre]) || 0;

  if (tarifOt > 0 && tarifLibre > 0) {
    return "ERREUR : Un Tarif OT et un Tarif Libre ne peuvent pas être saisis simultanément.";
  }

  let commission = 0;
  let montantRestant = 0;
  
  if (tarifOt > 0) {
    commission = tarifOt * 0.12;
    montantRestant = tarifOt - commission;
  } else if (tarifLibre > 0) {
    commission = 0;
    montantRestant = tarifLibre;
  }

  formData['commission_ot'] = commission.toFixed(2);
  formData['montant_restant'] = montantRestant.toFixed(2);

  let infosCal = getInfoCalendrierParDate(dateDebut);
  if (infosCal) {
    if (!formData['Année']) formData['Année'] = infosCal.annee;
    
    let numSem = formData['num_semaine'] || formData['N° Semaine'];
    if (!numSem) {
      formData['N° Semaine'] = infosCal.semaine;
    } else {
      formData['N° Semaine'] = numSem;
    }
    
    if (!formData['Date Fin (Samedi)']) formData['Date Fin (Samedi)'] = infosCal.dateFin;
  }
  
  delete formData['num_semaine'];

  if (rowIndex) {
    modifierEnregistrement(CONFIG_TARIFS.feuille, rowIndex, formData);
    return "Modification effectuée avec succès.";
  } else {
    let nouvelId = creerEnregistrement(CONFIG_TARIFS.feuille, CONFIG_TARIFS.nomId, formData);
    return "Création réussie (ID généré : " + nouvelId + ").";
  }
}

function traiterSuppressionTarif(rowIndex) {
  supprimerEnregistrement(CONFIG_TARIFS.feuille, rowIndex);
  return "Suppression confirmée.";
}

const CONFIG_RESA = {
  feuille: 'cal_reservation_26-50',
  nomId: 'ID_cal_reservation_26-50',
  colonneRechercheDate: 'Date Début (Samedi)'
};

function traiterRechercheReservation(dateSaisie) {
  return rechercherEnregistrementParDate(CONFIG_RESA.feuille, CONFIG_RESA.colonneRechercheDate, dateSaisie);
}

function traiterSauvegardeReservation(formData, rowIndex) {
  let dateDebut = formData['Date Début (Samedi)'];
  if (!dateDebut) return "ERREUR : La date de début est obligatoire.";

  let estUnDoublon = dateExisteDeja(CONFIG_RESA.feuille, CONFIG_RESA.colonneRechercheDate, dateDebut, rowIndex);
  if (estUnDoublon) {
    return "ERREUR : Une réservation existe déjà pour la semaine du " + dateDebut + ".";
  }

  let infosCal = getInfoCalendrierParDate(dateDebut);
  if (infosCal) {
    if (!formData['Année']) formData['Année'] = infosCal.annee;
    let numSem = formData['num_semaine'] || formData['N° Semaine'];
    if (!numSem) formData['N° Semaine'] = infosCal.semaine;
    else formData['N° Semaine'] = numSem;
    if (!formData['Date Fin (Samedi)']) formData['Date Fin (Samedi)'] = infosCal.dateFin;
    if (infosCal.saison) formData['saison'] = infosCal.saison; 
  }
  delete formData['num_semaine'];

  let infosStation = getInfosStationParDate(dateDebut);
  formData['etat_station'] = infosStation.etat;
  formData['saison'] = infosStation.saison;

  let tarifAssocie = rechercherEnregistrementParDate(CONFIG_TARIFS.feuille, CONFIG_TARIFS.colonneRechercheDate, dateDebut);
  if (tarifAssocie && tarifAssocie.data && tarifAssocie.data['ID_cal_tarifs_26-50']) {
    formData['ID_cal_tarifs_26-50'] = tarifAssocie.data['ID_cal_tarifs_26-50'];
  } else {
    formData['ID_cal_tarifs_26-50'] = "";
  }

  if (rowIndex) {
    if (!formData[CONFIG_RESA.nomId] || formData[CONFIG_RESA.nomId].toString().trim() === "") {
      formData[CONFIG_RESA.nomId] = genererIdUnique();
    }
    modifierEnregistrement(CONFIG_RESA.feuille, rowIndex, formData);
    return "Modification de la réservation effectuée.";
  } else {
    let nouvelId = creerEnregistrement(CONFIG_RESA.feuille, CONFIG_RESA.nomId, formData);
    return "Réservation créée avec succès (ID : " + nouvelId + ").";
  }
}

function traiterSuppressionReservation(rowIndex) {
  supprimerEnregistrement(CONFIG_RESA.feuille, rowIndex);
  return "Réservation supprimée définitivement.";
}

function getOriginesReservation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleListes = ss.getSheetByName('listes');
  if (!feuilleListes) return [];
  
  const data = feuilleListes.getDataRange().getValues();
  let origines = [];
  
  for (let i = 1; i < data.length; i++) {
    let val = data[i][2]; 
    if (val) {
      origines.push(val.toString().trim());
    }
  }
  return [...new Set(origines)];
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

function getInfosStationParDate(dateSearchStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleStation = ss.getSheetByName('cal_station');
  
  if (!feuilleStation) return { etat: "Fermée", saison: "" }; 

  const data = feuilleStation.getDataRange().getValues();
  const entetes = data[0];

  let idxOuverture = -1, idxFermeture = -1, idxEtat = -1, idxSaison = -1;

  for (let c = 0; c < entetes.length; c++) {
    let nomCol = entetes[c].toString().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
    if (nomCol === 'ouverturestation') idxOuverture = c;
    if (nomCol === 'fermeturestation') idxFermeture = c;
    if (nomCol === 'etatstation') idxEtat = c;
    if (nomCol === 'saisonstation' || nomCol === 'saison') idxSaison = c;
  }

  if (idxOuverture === -1 || idxFermeture === -1) return { etat: "Fermée", saison: "" };

  let parts = dateSearchStr.split('-');
  let targetDate = new Date(parts[0], parts[1] - 1, parts[2]);
  targetDate.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    let dOuverture = data[i][idxOuverture];
    let dFermeture = data[i][idxFermeture];

    if (!dOuverture || !dFermeture) continue;

    let dateO = (Object.prototype.toString.call(dOuverture) === '[object Date]') ? new Date(dOuverture.getTime()) : new Date(dOuverture);
    let dateF = (Object.prototype.toString.call(dFermeture) === '[object Date]') ? new Date(dFermeture.getTime()) : new Date(dFermeture);
    
    dateO.setHours(0,0,0,0);
    dateF.setHours(0,0,0,0);

    if (targetDate >= dateO && targetDate <= dateF) {
      let etatTrouve = idxEtat !== -1 ? data[i][idxEtat].toString().trim() : "Fermée";
      let saisonTrouvee = idxSaison !== -1 ? data[i][idxSaison].toString().trim() : "";
      return { etat: etatTrouve, saison: saisonTrouvee };
    }
  }
  
  return { etat: "Fermée", saison: "" };
}
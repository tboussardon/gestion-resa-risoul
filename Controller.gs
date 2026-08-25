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

  // Détection des doublons de dates
  let estUnDoublon = dateExisteDeja(CONFIG_TARIFS.feuille, CONFIG_TARIFS.colonneRechercheDate, dateDebut, rowIndex);
  if (estUnDoublon) {
    return "ERREUR : Un tarif existe déjà pour cette date de début (" + dateDebut + "). Impossible de créer un doublon.";
  }

  // --- NOUVEAU : Exclusivité OT / Libre et calcul serveur ---
  let tarifOt = parseFloat(formData['tarif_ot']) || 0;
  // Gère automatiquement tarif_libre ou tarifs_libres selon l'attribut name du HTML
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
    commission = 0; // Aucune commission sur un tarif direct
    montantRestant = tarifLibre;
  }

  formData['commission_ot'] = commission.toFixed(2);
  formData['montant_restant'] = montantRestant.toFixed(2);

  // Remplissage automatique de l'Année, Semaine et Date Fin
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

  // Enregistrement
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

// --- CONFIGURATION RÉSERVATIONS ---
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

  // 1. Détection des doublons (Une seule réservation par date de début)
  let estUnDoublon = dateExisteDeja(CONFIG_RESA.feuille, CONFIG_RESA.colonneRechercheDate, dateDebut, rowIndex);
  if (estUnDoublon) {
    return "ERREUR : Une réservation existe déjà pour la semaine du " + dateDebut + ".";
  }

  // 2. Remplissage automatique (Année, Semaine, Date Fin)
  let infosCal = getInfoCalendrierParDate(dateDebut);
  if (infosCal) {
    if (!formData['Année']) formData['Année'] = infosCal.annee;
    let numSem = formData['num_semaine'] || formData['N° Semaine'];
    if (!numSem) formData['N° Semaine'] = infosCal.semaine;
    else formData['N° Semaine'] = numSem;
    if (!formData['Date Fin (Samedi)']) formData['Date Fin (Samedi)'] = infosCal.dateFin;
    
    // --- NOUVEAU : Sauvegarde automatique de la saison ---
    if (infosCal.saison) formData['saison'] = infosCal.saison; 
  }
  delete formData['num_semaine'];

  // --- NOUVEAU : Récupération automatique de l'état ET de la saison depuis cal_station ---
  let infosStation = getInfosStationParDate(dateDebut);
  formData['etat_station'] = infosStation.etat;
  formData['saison'] = infosStation.saison;

  // 3. LA RELATION MAGIQUE AVEC LES TARIFS
  // On cherche si un tarif existe à cette même date dans l'autre feuille
  let tarifAssocie = rechercherEnregistrementParDate(CONFIG_TARIFS.feuille, CONFIG_TARIFS.colonneRechercheDate, dateDebut);
  if (tarifAssocie && tarifAssocie.data && tarifAssocie.data['ID_cal_tarifs_26-50']) {
    formData['ID_cal_tarifs_26-50'] = tarifAssocie.data['ID_cal_tarifs_26-50'];
  } else {
    formData['ID_cal_tarifs_26-50'] = ""; // Aucun tarif lié
  }

  // 4. Enregistrement
  if (rowIndex) {
    // CORRECTION : Si la ligne existante n'a pas d'ID, on lui en génère un
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

// --- FONCTION POUR ALIMENTER LE MENU DÉROULANT ---
function getOriginesReservation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleListes = ss.getSheetByName('listes');
  if (!feuilleListes) return [];
  
  const data = feuilleListes.getDataRange().getValues();
  let origines = [];
  
  // On boucle à partir de 1 pour ignorer la ligne d'en-tête
  for (let i = 1; i < data.length; i++) {
    // La colonne C correspond à l'index 2 (A=0, B=1, C=2)
    let val = data[i][2]; 
    if (val) {
      origines.push(val.toString().trim());
    }
  }
  
  // Retourne uniquement les valeurs uniques
  return [...new Set(origines)];
}

// --- NOUVEAU : Déduire l'état de la station ET la saison selon la date ---
function getInfosStationParDate(dateSearchStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleStation = ss.getSheetByName('cal_station');
  
  if (!feuilleStation) return { etat: "Fermée", saison: "" }; 

  const data = feuilleStation.getDataRange().getValues();
  const entetes = data[0];

  let idxOuverture = -1, idxFermeture = -1, idxEtat = -1;
  let idxSaison = 5; // On force la lecture sur la Colonne F (index 5)

  for (let c = 0; c < entetes.length; c++) {
    let nomCol = entetes[c].toString().trim();
    if (nomCol === 'ouverture_station') idxOuverture = c;
    if (nomCol === 'fermeture_station') idxFermeture = c;
    if (nomCol === 'etat_station') idxEtat = c;
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

    // Si la date est dans la période, on renvoie l'état ET la saison (colonne F)
    if (targetDate >= dateO && targetDate <= dateF) {
      let etatTrouve = idxEtat !== -1 ? data[i][idxEtat].toString().trim() : "Fermée";
      let saisonTrouvee = data[i][idxSaison] ? data[i][idxSaison].toString().trim() : "";
      
      return { etat: etatTrouve, saison: saisonTrouvee };
    }
  }
  
  return { etat: "Fermée", saison: "" };
}
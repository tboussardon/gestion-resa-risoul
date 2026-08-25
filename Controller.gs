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
  let etatCalcule = "ouverte";
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

  // Conversion en vrais objets Date pour un formatage correct dans Sheets
  if (formData['Date Début (Samedi)'] && typeof formData['Date Début (Samedi)'] === 'string') {
    let p = formData['Date Début (Samedi)'].split('-');
    if (p.length === 3) formData['Date Début (Samedi)'] = new Date(p[0], p[1] - 1, p[2]);
  }
  if (formData['Date Fin (Samedi)'] && typeof formData['Date Fin (Samedi)'] === 'string') {
    let p = formData['Date Fin (Samedi)'].split('-');
    if (p.length === 3) formData['Date Fin (Samedi)'] = new Date(p[0], p[1] - 1, p[2]);
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

function parseDateFlexible(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    let d = new Date(val.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  let str = val.toString().trim();
  // Format AAAA-MM-JJ
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
    let p = str.substring(0, 10).split('-');
    return new Date(p[0], p[1] - 1, p[2]);
  }
  // Format JJ/MM/AAAA
  if (str.match(/^\d{2}\/\d{2}\/\d{4}/)) {
    let p = str.substring(0, 10).split('/');
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return null;
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

  let targetDate = parseDateFlexible(dateSearchStr);
  if (!targetDate) return { etat: "Fermée", saison: "" };

  for (let i = 1; i < data.length; i++) {
    let dateO = parseDateFlexible(data[i][idxOuverture]);
    let dateF = parseDateFlexible(data[i][idxFermeture]);

    if (!dateO || !dateF) continue;

    if (targetDate >= dateO && targetDate <= dateF) {
      let etatTrouve = idxEtat !== -1 ? data[i][idxEtat].toString().trim() : "Fermée";
      let saisonTrouvee = idxSaison !== -1 ? data[i][idxSaison].toString().trim() : "";
      return { etat: etatTrouve, saison: saisonTrouvee };
    }
  }
  
  return { etat: "Fermée", saison: "" };
}

/**
 * Récupère les données croisées et filtrées (Réservations + Tarifs)
 */
function obtenirDonneesSynthese(filtres) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fResa = ss.getSheetByName('cal_reservation_26-50');
  const fTarifs = ss.getSheetByName('cal_tarifs_26-50');
  
  if (!fResa) return { erreur: "Feuille 'cal_reservation_26-50' introuvable." };
  
  const dataResa = fResa.getDataRange().getValues();
  if (dataResa.length < 2) {
    return { kpis: { count: 0, totalBrut: "0.00", totalComm: "0.00", totalNet: "0.00" }, lignes: [] };
  }
  
  const entetesResa = dataResa[0].map(e => e.toString().trim());
  
  // Indexation des tarifs par leur ID
  let mapTarifs = {};
  if (fTarifs) {
    const dataTarifs = fTarifs.getDataRange().getValues();
    if (dataTarifs.length > 1) {
      const entetesTarifs = dataTarifs[0].map(e => e.toString().trim());
      const idxIdTarif = entetesTarifs.indexOf('ID_cal_tarifs_26-50');
      const idxOt = entetesTarifs.indexOf('tarif_ot');
      const idxLibre = entetesTarifs.indexOf('tarif_libre') !== -1 ? entetesTarifs.indexOf('tarif_libre') : entetesTarifs.indexOf('tarifs_libres');
      const idxComm = entetesTarifs.indexOf('commission_ot');
      const idxRestant = entetesTarifs.indexOf('montant_restant');

      for (let i = 1; i < dataTarifs.length; i++) {
        let idT = dataTarifs[i][idxIdTarif];
        if (idT) {
          let ot = parseFloat(dataTarifs[i][idxOt]) || 0;
          let libre = parseFloat(dataTarifs[i][idxLibre]) || 0;
          let brut = ot > 0 ? ot : libre;
          let comm = parseFloat(dataTarifs[i][idxComm]) || 0;
          let net = (parseFloat(dataTarifs[i][idxRestant]) || 0) + libre;
          
          mapTarifs[idT] = { brut: brut, comm: comm, net: net, libre: libre };
        }
      }
    }
  }

  const idxAnnee = entetesResa.indexOf('Année');
  const idxSemaine = entetesResa.indexOf('N° Semaine');
  const idxDebut = entetesResa.indexOf('Date Début (Samedi)');
  const idxFin = entetesResa.indexOf('Date Fin (Samedi)');
  const idxNom = entetesResa.indexOf('nom');
  const idxPrenom = entetesResa.indexOf('prenom');
  const idxOrigine = entetesResa.indexOf('origine_reservation');
  const idxEtatStation = entetesResa.indexOf('etat_station');
  const idxSaison = entetesResa.indexOf('saison');
  const idxConciergerie = entetesResa.indexOf('conciergerie');
  const idxIdTarifResa = entetesResa.indexOf('ID_cal_tarifs_26-50');

  let resultats = [];
  let totalBrutNum = 0, totalCommNum = 0, totalNetNum = 0, totalLibreNum = 0;

  const timeZone = ss.getSpreadsheetTimeZone();
  const formatteDate = (val) => {
    if (Object.prototype.toString.call(val) === '[object Date]') {
      return Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
    }
    return val ? val.toString().substring(0, 10) : "";
  };

  for (let i = 1; i < dataResa.length; i++) {
    let row = dataResa[i];
    
    let annee = row[idxAnnee] ? row[idxAnnee].toString().trim() : "";
    let saison = row[idxSaison] ? row[idxSaison].toString().trim() : "";
    let origine = row[idxOrigine] ? row[idxOrigine].toString().trim() : "";
    let conciergerie = row[idxConciergerie] ? row[idxConciergerie].toString().trim() : "";
    let etatStation = row[idxEtatStation] ? row[idxEtatStation].toString().trim() : "";
    
    // Application des filtres
    if (filtres.annee && annee !== filtres.annee) continue;
    if (filtres.saison && saison !== filtres.saison) continue;
    if (filtres.origine && origine !== filtres.origine) continue;
    if (filtres.conciergerie && conciergerie !== filtres.conciergerie) continue;
    if (filtres.etatStation && etatStation !== filtres.etatStation) continue;

    let idTarifAssocie = idxIdTarifResa !== -1 ? row[idxIdTarifResa] : "";
    let infoTarif = mapTarifs[idTarifAssocie] || { brut: 0, comm: 0, net: 0 };

    totalBrutNum += infoTarif.brut;
    totalCommNum += infoTarif.comm;
    totalNetNum += infoTarif.net;
    totalLibreNum += infoTarif.libre;

    resultats.push({
      annee: annee,
      semaine: row[idxSemaine] || "",
      dateDebut: formatteDate(row[idxDebut]),
      dateFin: formatteDate(row[idxFin]),
      client: ((row[idxNom] || "") + " " + (row[idxPrenom] || "")).trim(),
      origine: origine,
      conciergerie: conciergerie,
      saison: saison,
      etatStation: etatStation,
      tarifBrut: infoTarif.brut.toFixed(2),
      commission: infoTarif.comm.toFixed(2),
      montantNet: infoTarif.net.toFixed(2)
    });
  }

  return {
  kpis: {
    count: resultats.length,
    totalBrut: totalBrutNum.toFixed(2),
    totalComm: totalCommNum.toFixed(2),
    totalNet: totalNetNum.toFixed(2),
    totalLibre: totalLibreNum.toFixed(2) // NOUVEAU
  },
  lignes: resultats
};
}

/**
 * Récupère les années enregistrées pour alimenter la liste déroulante
 */
function getAnneesSynthese() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fResa = ss.getSheetByName('cal_reservation_26-50');
  if (!fResa) return [];
  const data = fResa.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const entetes = data[0].map(e => e.toString().trim());
  const idxAnnee = entetes.indexOf('Année');
  if (idxAnnee === -1) return [];

  let annees = [];
  for (let i = 1; i < data.length; i++) {
    let a = data[i][idxAnnee];
    if (a) annees.push(a.toString().trim());
  }
  return [...new Set(annees)].sort();
}

function obtenirDonneesSynthese(filtres) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fResa = ss.getSheetByName('cal_reservation_26-50');
  const fTarifs = ss.getSheetByName('cal_tarifs_26-50');
  
  if (!fResa) return { erreur: "Feuille 'cal_reservation_26-50' introuvable." };
  
  const dataResa = fResa.getDataRange().getValues();
  if (dataResa.length < 2) {
    return { kpis: { count: 0, totalBrut: "0.00", totalComm: "0.00", totalNet: "0.00", totalLibre: "0.00" }, lignes: [] };
  }
  
  const entetesResa = dataResa[0].map(e => e.toString().trim());
  
  // Indexation des tarifs par leur ID
  let mapTarifs = {};
  if (fTarifs) {
    const dataTarifs = fTarifs.getDataRange().getValues();
    if (dataTarifs.length > 1) {
      const entetesTarifs = dataTarifs[0].map(e => e.toString().trim());
      const idxIdTarif = entetesTarifs.indexOf('ID_cal_tarifs_26-50');
      const idxOt = entetesTarifs.indexOf('tarif_ot');
      const idxLibre = entetesTarifs.indexOf('tarif_libre') !== -1 ? entetesTarifs.indexOf('tarif_libre') : entetesTarifs.indexOf('tarifs_libres');
      const idxComm = entetesTarifs.indexOf('commission_ot');
      const idxRestant = entetesTarifs.indexOf('montant_restant');

      for (let i = 1; i < dataTarifs.length; i++) {
        let idT = dataTarifs[i][idxIdTarif];
        if (idT) {
          let ot = parseFloat(dataTarifs[i][idxOt]) || 0;
          let libre = parseFloat(dataTarifs[i][idxLibre]) || 0;
          let brut = ot > 0 ? ot : libre;
          let comm = parseFloat(dataTarifs[i][idxComm]) || 0;
          let net = (parseFloat(dataTarifs[i][idxRestant]) || 0) + libre;
          
          mapTarifs[idT] = { brut: brut, comm: comm, net: net, libre: libre };
        }
      }
    }
  }

  const idxAnnee = entetesResa.indexOf('Année');
  const idxSemaine = entetesResa.indexOf('N° Semaine');
  const idxDebut = entetesResa.indexOf('Date Début (Samedi)');
  const idxFin = entetesResa.indexOf('Date Fin (Samedi)');
  const idxNom = entetesResa.indexOf('nom');
  const idxPrenom = entetesResa.indexOf('prenom');
  const idxOrigine = entetesResa.indexOf('origine_reservation');
  const idxEtatStation = entetesResa.indexOf('etat_station');
  const idxSaison = entetesResa.indexOf('saison');
  const idxConciergerie = entetesResa.indexOf('conciergerie');
  const idxIdTarifResa = entetesResa.indexOf('ID_cal_tarifs_26-50');

  let resultats = [];
  let totalBrutNum = 0, totalCommNum = 0, totalNetNum = 0, totalLibreNum = 0;

  const timeZone = ss.getSpreadsheetTimeZone();
  const formatteDate = (val) => {
    if (Object.prototype.toString.call(val) === '[object Date]') {
      return Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
    }
    return val ? val.toString().substring(0, 10) : "";
  };

  for (let i = 1; i < dataResa.length; i++) {
    let row = dataResa[i];
    
    let annee = row[idxAnnee] ? row[idxAnnee].toString().trim() : "";
    let saison = row[idxSaison] ? row[idxSaison].toString().trim() : "";
    let origine = row[idxOrigine] ? row[idxOrigine].toString().trim() : "";
    let conciergerie = row[idxConciergerie] ? row[idxConciergerie].toString().trim() : "";
    let etatStation = row[idxEtatStation] ? row[idxEtatStation].toString().trim() : "";
    
    // Application des filtres insensibles à la casse
    if (filtres.annee && annee.toLowerCase() !== filtres.annee.toLowerCase()) continue;
    if (filtres.saison && saison.toLowerCase() !== filtres.saison.toLowerCase()) continue;
    if (filtres.origine && origine.toLowerCase() !== filtres.origine.toLowerCase()) continue;
    if (filtres.conciergerie && conciergerie.toLowerCase() !== filtres.conciergerie.toLowerCase()) continue;
    if (filtres.etatStation && etatStation.toLowerCase() !== filtres.etatStation.toLowerCase()) continue;

    let idTarifAssocie = idxIdTarifResa !== -1 ? row[idxIdTarifResa] : "";
    let infoTarif = mapTarifs[idTarifAssocie] || { brut: 0, comm: 0, net: 0, libre: 0 };

    totalBrutNum += infoTarif.brut;
    totalCommNum += infoTarif.comm;
    totalNetNum += infoTarif.net;
    totalLibreNum += infoTarif.libre;

    resultats.push({
      annee: annee,
      semaine: row[idxSemaine] || "",
      dateDebut: formatteDate(row[idxDebut]),
      dateFin: formatteDate(row[idxFin]),
      client: ((row[idxNom] || "") + " " + (row[idxPrenom] || "")).trim(),
      origine: origine,
      conciergerie: conciergerie,
      saison: saison,
      etatStation: etatStation,
      tarifBrut: infoTarif.brut.toFixed(2),
      commission: infoTarif.comm.toFixed(2),
      montantNet: infoTarif.net.toFixed(2)
    });
  }

  return {
    kpis: {
      count: resultats.length,
      totalBrut: totalBrutNum.toFixed(2),
      totalComm: totalCommNum.toFixed(2),
      totalNet: totalNetNum.toFixed(2),
      totalLibre: totalLibreNum.toFixed(2)
    },
    lignes: resultats
  };
}
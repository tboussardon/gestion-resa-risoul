/**
 * Script global d'intégration des calendriers scolaires (Zones A, B et C)
 * dans la feuille "cal_scolaire" de votre classeur.
 */

// Générateur d'ID unique à 5 caractères alphanumériques
function genererIDUnique(setIDs) {
  const chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let id = "";
  do {
    id = "";
    for (let i = 0; i < 5; i++) {
      id += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (setIDs.has(id));
  setIDs.add(id);
  return id;
}

function importerToutesLesZonesScolaires() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const NOM_FEUILLE = "cal_scolaire";
  
  let sheet = ss.getSheetByName(NOM_FEUILLE);
  if (!sheet) {
    sheet = ss.insertSheet(NOM_FEUILLE);
  } else {
    sheet.clear();
  }

  const rows = [[
    "ID_cal_scolaire", "ID_cal_station", "ID_cal_reservation", "ID_tarifs",
    "annees_scolaires", "debut_annee_scolaire", "fin_annee_scolaire",
    "description_annee_scolaire", "zone_annee_scolaire", "note_cal_scolaire"
  ]];

  const setIDs = new Set();
  const totalEvents = parseIcsToRows(ICS_ZONE_A, setIDs, "Zone A")
    .concat(parseIcsToRows(ICS_ZONE_B, setIDs, "Zone B"))
    .concat(parseIcsToRows(ICS_ZONE_C, setIDs, "Zone C"));

  if (totalEvents.length === 0) return;
  totalEvents.sort((a, b) => a.debut - b.debut);

  for (let ev of totalEvents) {
    rows.push([ ev.id, "", "", "", ev.anneeScolaire, ev.debut, ev.fin, ev.description, ev.zone, "" ]);
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight("bold").setBackground("#D9EAD3");
  sheet.getRange(2, 6, rows.length - 1, 2).setNumberFormat("dd/mm/yyyy");
  sheet.autoResizeColumns(1, rows[0].length);
  SpreadsheetApp.getUi().alert("✅ Succès !");
}

function parseIcsToRows(icsString, setIDs, nomZoneDefaut) {
  const list = [];
  if (!icsString) return list;
  const events = icsString.split("BEGIN:VEVENT");
  for (let i = 1; i < events.length; i++) {
    const block = events[i];
    const summaryMatch = block.match(/SUMMARY:(.*)/);
    const dtstartMatch = block.match(/DTSTART;VALUE=DATE:(\d{8})/);
    const dtendMatch = block.match(/DTEND;VALUE=DATE:(\d{8})/);
    const locationMatch = block.match(/LOCATION:(.*)/);
    if (summaryMatch && dtstartMatch && dtendMatch) {
      const id = genererIDUnique(setIDs);
      const dtstart = parseIcalDate(dtstartMatch[1]);
      let anneeScolaire = (dtstart.getMonth() + 1 >= 9) ? dtstart.getFullYear() + "/" + (dtstart.getFullYear() + 1) : (dtstart.getFullYear() - 1) + "/" + dtstart.getFullYear();
      list.push({ id: id, description: summaryMatch[1].trim(), debut: dtstart, fin: parseIcalDate(dtendMatch[1]), zone: locationMatch ? locationMatch[1].trim() : nomZoneDefaut, anneeScolaire: anneeScolaire });
    }
  }
  return list;
}

function parseIcalDate(dateStr) {
  return new Date(parseInt(dateStr.substring(0, 4), 10), parseInt(dateStr.substring(4, 6), 10) - 1, parseInt(dateStr.substring(6, 8), 10));
}

// ICS data supprimées par souci de brièveté de la réponse (gardez votre fichier actuel tel quel).
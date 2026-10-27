// Script du Google Sheet « Avis IXINA Tour & Taxis ».
// À coller dans Extensions > Apps Script, puis Déployer > Application Web (accès : Tout le monde).
// Onglet « Avis pris » : chaque avis copié par un client, pour qu'il ne soit plus proposé.
// Onglet « Avis ajoutés » : vos propres avis, un par ligne en colonne A ({c} = prénom du conseiller).

const ONGLET_PRIS = "Avis pris";
const ONGLET_AJOUTS = "Avis ajoutés";
const MIN_CARACTERES = 90;
const SEUIL_ALERTE = 20;

function doGet(e) {
  const p = e.parameter;
  if (p.action === "prendre") return json(prendre(p));
  return json(liste());
}

function liste() {
  const pris = onglet(ONGLET_PRIS).getDataRange().getValues().slice(1).map((r) => String(r[1]));
  const ajouts = onglet(ONGLET_AJOUTS).getDataRange().getValues().slice(1)
    .map((r, i) => ({ id: "a" + (i + 2), texte: String(r[0]).trim() }))
    .filter((a) => a.texte.length >= MIN_CARACTERES);
  return { pris, ajouts };
}

function prendre(p) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    onglet(ONGLET_PRIS).appendRow([new Date(), p.id || "", p.conseiller || "", p.texte || ""]);
  } finally {
    verrou.releaseLock();
  }
  alerterSiPresqueVide(Number(p.reste));
  return { ok: true };
}

// Un seul mail quand il reste moins de 20 avis, puis plus rien tant que la liste n'est pas regarnie.
function alerterSiPresqueVide(reste) {
  if (isNaN(reste)) return;
  const props = PropertiesService.getScriptProperties();
  if (reste >= SEUIL_ALERTE) {
    props.deleteProperty("alerte");
    return;
  }
  if (props.getProperty("alerte")) return;
  props.setProperty("alerte", "envoyée");
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    "Avis IXINA Tour & Taxis : plus que " + reste + " modèles",
    "Il reste " + reste + " modèles d'avis pas encore utilisés.\n" +
    "Ajoutez-en dans l'onglet « " + ONGLET_AJOUTS + " » du Google Sheet, ou demandez une nouvelle série."
  );
}

function onglet(nom) {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let feuille = classeur.getSheetByName(nom);
  if (!feuille) {
    feuille = classeur.insertSheet(nom);
    const titres = nom === ONGLET_PRIS
      ? ["Date", "Identifiant", "Conseiller", "Texte publié"]
      : ["Vos avis (un par ligne, 90 caractères minimum, {c} = prénom du conseiller)"];
    feuille.appendRow(titres);
    feuille.setFrozenRows(1);
    feuille.getRange(1, 1, 1, titres.length).setFontWeight("bold");
  }
  return feuille;
}

// À lancer une fois depuis l'éditeur (bouton Exécuter) : crée les onglets et demande les autorisations.
function installer() {
  onglet(ONGLET_PRIS);
  onglet(ONGLET_AJOUTS);
  const vide = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Feuille 1")
    || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  if (vide && vide.getLastRow() === 0) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(vide);
}

function json(objet) {
  return ContentService.createTextOutput(JSON.stringify(objet)).setMimeType(ContentService.MimeType.JSON);
}

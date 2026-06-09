/**
 * Fichier : Code.gs
 * Point d'entrée et API de la Web App.
 */

const SPREADSHEET_ID = '1gj0fXumXQqHjvuJpByR_uHcnhEAs1b6r53gijj5B6Gk';
const SHEET_NAME = 'Content';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Formation OPTIMA - Portail Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================================================
// GESTION DU CONTENU (CMS GSHEET)
// =========================================================================

/**
 * À EXÉCUTER UNE FOIS MANUELLEMENT DANS L'ÉDITEUR GAS
 * Crée l'onglet et initialise les textes de base.
 */
function initContentSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  const defaultContent = [
    ['Key', 'Value'],
    ['app_title', 'OPTIMA S&OP'],
    ['app_subtitle', 'Parcours Manager'],
    ['nav_home', 'Accueil & Guide'],
    ['nav_training_plan', 'Plan de Formation'],
    ['nav_p1', 'P1: Le Cycle S&OP'],
    ['nav_p2', 'P2: La Capacité'],
    ['nav_p3', 'P3: Gestion de Ressource'],
    ['nav_p4', 'P4: Gestion Avancée'],
    ['nav_p5', 'P5: Affectation de Charge'],
    ['home_title', 'Formation OPTIMA S&OP'],
    ['home_welcome', 'Bienvenue dans le nouveau parcours de formation pour les Managers. Ce programme a été segmenté en 5 parties pour couvrir l\'intégralité de vos responsabilités dans OPTIMA.'],
    ['home_slides_title', 'Support de Présentation'],
    ['home_btn_start', 'Démarrer la Partie 1'],
    ['p1_title', 'Partie 1 : Le Cycle S&OP'],
    ['p1_desc', 'Rappel sur le processus mensuel de Sales & Operations Planning.'],
    ['p1_success', 'Succès ! L\'architecture modulaire et le CMS GSheet fonctionnent parfaitement.'],
    ['p1_btn_next', 'Passer à la Partie 2']
  ];
  
  sheet.getRange(1, 1, defaultContent.length, 2).setValues(defaultContent);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E2E8F0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

/**
 * Appelée par React au chargement pour récupérer les textes
 */
function getContent() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const content = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      content[data[i][0]] = data[i][1];
    }
  }
  return content;
}

/**
 * Appelée par React lors du clic sur "Sauvegarder les textes"
 */
function saveContent(updates) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const keysMap = {};
  
  // Mapper les lignes existantes
  for (let i = 1; i < data.length; i++) {
    keysMap[data[i][0]] = i + 1; 
  }
  
  // Mettre à jour ou ajouter les nouvelles clés
  for (let key in updates) {
    if (keysMap[key]) {
      sheet.getRange(keysMap[key], 2).setValue(updates[key]);
    } else {
      sheet.appendRow([key, updates[key]]);
    }
  }
  return true;
}

/**
 * Appelée par React pour récupérer la liste des codes métiers depuis l'onglet "Codes métiers"
 */
function getJobCodes() {
  const SPREADSHEET_ID = '1gj0fXumXQqHjvuJpByR_uHcnhEAs1b6r53gijj5B6Gk';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // On cible l'onglet spécifique
  const sheet = ss.getSheetByName('Codes métiers');
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Sécurité si la feuille est vide
  
  // On récupère spécifiquement les données des colonnes A à D (1 à 4), 
  // à partir de la ligne 2 pour ignorer les en-têtes
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const jobCodes = [];
  
  for (let i = 0; i < data.length; i++) {
    const code = data[i][0];       // Colonne A : Code
    const englishJob = data[i][1]; // Colonne B : Jobs
    const frenchJob = data[i][2];  // Colonne C : TRADUCTION JOB
    const family = data[i][3];     // Colonne D : Chapeau
    
    // On ignore les lignes où la colonne A (Code) est vide
    if (!code || code.toString().trim() === '') continue; 
    
    jobCodes.push({
      code: String(code).trim(),
      // On privilégie la traduction (C), sinon on prend la version anglaise (B)
      name: String(frenchJob || englishJob).trim(), 
      // On assigne la famille (D) ou 'Autres' par défaut
      family: String(family || 'Autres').trim()
    });
  }
  
  return jobCodes;
}

/**
 * Appelée par React pour récupérer la liste des vidéos depuis l'onglet "Liste des vidéos".
 * Colonnes attendues : A = Etape (ex: "1", "2", "3b"), B = vidéo n°1, C = vidéo n°2.
 */
function getStepVideos() {
  const SPREADSHEET_ID = '1gj0fXumXQqHjvuJpByR_uHcnhEAs1b6r53gijj5B6Gk'; // Ton ID de Spreadsheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Liste des vidéos');
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // On récupère les colonnes A à C, à partir de la ligne 2
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const videos = [];
  
  for (let i = 0; i < data.length; i++) {
    const step = data[i][0];
    const v1 = data[i][1];
    const v2 = data[i][2];
    
    if (step) {
      videos.push({
        step: String(step).trim(),
        v1: v1 ? String(v1).trim() : '',
        v2: v2 ? String(v2).trim() : ''
      });
    }
  }
  
  return videos;
}

/**
 * Enregistre la progression d'un utilisateur dans l'onglet "Suivi_Formation".
 * Si l'onglet n'existe pas, il sera créé automatiquement.
 */
function logUserProgress(email, partName, detailsStr) {
  const SPREADSHEET_ID = '1gj0fXumXQqHjvuJpByR_uHcnhEAs1b6r53gijj5B6Gk'; // Ton ID Google Sheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let trackingSheet = ss.getSheetByName('Suivi_Formation');
  
  // Création de l'onglet s'il n'existe pas encore
  if (!trackingSheet) {
    trackingSheet = ss.insertSheet('Suivi_Formation');
    trackingSheet.appendRow(['Date et Heure', 'Email', 'Partie / Étape', 'Détails du résultat']);
    trackingSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f4f6");
  }
  
  // Formatage de la date
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  // Ajout de la ligne d'historique
  trackingSheet.appendRow([timestamp, email, partName, detailsStr]);
}



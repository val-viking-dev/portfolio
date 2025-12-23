/**
 * ============================================================================
 * DASHBOARD.GS - VERSION CORRIGÉE COMPATIBLE
 * ============================================================================
 * Ce fichier crée un tableau de bord visuel dans une nouvelle feuille "Dashboard"
 * 
 * ⚠️ IMPORTANT : Ce fichier utilise vos variables existantes (SPREADSHEET_ID, CONFIG)
 * 
 * Auteur: Noemie
 * Projet: Inventaire XWB BARQUE Operations
 * Date: Decembre 2025
 */

/**
 * Crée le tableau de bord dans une nouvelle feuille "Dashboard"
 * À exécuter manuellement ou via un déclencheur
 * 
 * UTILISATION :
 * 1. Sélectionner cette fonction dans le menu déroulant
 * 2. Cliquer sur Exécuter
 * 3. Le Dashboard sera créé dans un nouvel onglet
 */
function createDashboard() {
  try {
    console.log("🎯 Début de création du Dashboard...");
    
    // Utiliser SPREADSHEET_ID défini dans Config.gs
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Récupérer la feuille d'inventaire
    const sheetInventaire = ss.getSheetByName('Sheet pour inventaire');
    if (!sheetInventaire) {
      console.error("❌ Feuille 'Sheet pour inventaire' introuvable !");
      throw new Error("La feuille 'Sheet pour inventaire' n'existe pas");
    }
    
    // Récupérer la feuille de suivi
    // Utiliser CONFIG.sheets.suivi si défini, sinon "Suivi_WebApp"
    let suiviSheetName = 'Suivi_WebApp';
    if (typeof CONFIG !== 'undefined' && CONFIG.sheets && CONFIG.sheets.suivi) {
      suiviSheetName = CONFIG.sheets.suivi;
    }
    const sheetSuivi = ss.getSheetByName(suiviSheetName);
    
    console.log(`📊 Feuilles trouvées: inventaire=${sheetInventaire.getName()}, suivi=${sheetSuivi ? sheetSuivi.getName() : 'aucune'}`);
    
    // Créer ou récupérer la feuille Dashboard
    let dashboardSheet = ss.getSheetByName('Dashboard');
    if (!dashboardSheet) {
      console.log("📝 Création de la feuille Dashboard...");
      dashboardSheet = ss.insertSheet('Dashboard');
    } else {
      console.log("🔄 Effacement de la feuille Dashboard existante...");
      dashboardSheet.clear();
    }
    
    // Récupérer les données des mallettes
    console.log("📦 Récupération des données des mallettes...");
    const mallettes = getMallettesDataForDashboard(sheetInventaire, sheetSuivi);
    
    if (mallettes.length === 0) {
      console.warn("⚠️ Aucune mallette trouvée !");
    } else {
      console.log(`✅ ${mallettes.length} mallettes récupérées`);
    }
    
    // Créer les sections du dashboard
    console.log("🎨 Création de l'en-tête...");
    createDashboardHeader(dashboardSheet);
    
    console.log("📋 Création de la vue d'ensemble...");
    createMallettesOverview(dashboardSheet, mallettes);
    
    console.log("📊 Création des statistiques...");
    createGlobalStats(dashboardSheet, mallettes);
    
    console.log("⚠️ Création des alertes...");
    createAlertesSection(dashboardSheet, mallettes);
    
    // Mise en forme finale
    console.log("🎨 Mise en forme finale...");
    formatDashboard(dashboardSheet);
    
    console.log("✅ Dashboard créé avec succès !");
    
    // Afficher un message de confirmation (si possible)
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard créé avec succès !', '✅ Terminé', 3);
    } catch (e) {
      // Si toast ne fonctionne pas, ce n'est pas grave
    }
    
  } catch (error) {
    console.error("❌ ERREUR lors de la création du Dashboard:");
    console.error("Message:", error.message);
    console.error("Stack:", error.stack);
    throw error;
  }
}

/**
 * Récupère les données de toutes les mallettes
 * COMPATIBLE AVEC Code.gs existant - NE PAS RENOMMER
 */
function getMallettesDataForDashboard(sheetInventaire, sheetSuivi) {
  // Vérifier que la feuille existe
  if (!sheetInventaire) {
    console.error("❌ sheetInventaire est undefined");
    return [];
  }
  
  try {
    const data = sheetInventaire.getDataRange().getValues();
    const mallettes = [];
    
    if (data.length < 2) {
      console.log("⚠️ Aucune donnée trouvée dans la feuille inventaire");
      return [];
    }
    
    // Parcourir les colonnes pour trouver les mallettes
    for (let col = 0; col < data[0].length; col++) {
      const headerValue = data[0][col];
      
      // Si la cellule contient "MALLETTE"
      if (headerValue && headerValue.toString().toUpperCase().includes('MALLETTE')) {
        const malletteName = headerValue.toString().trim();
        
        // Compter les outils (cellules non vides de la colonne)
        let nbOutils = 0;
        for (let row = 1; row < data.length; row++) {
          if (data[row][col] && data[row][col].toString().trim() !== '') {
            nbOutils++;
          }
        }
        
        // Récupérer les infos du dernier contrôle depuis Suivi_WebApp
        const lastControl = getLastControlForMallette(sheetSuivi, malletteName);
        
        mallettes.push({
          nom: malletteName,
          nbOutils: nbOutils,
          derniereVerif: lastControl.date,
          controleur: lastControl.controleur,
          manquants: lastControl.nbManquants,
          etat: lastControl.etat,
          joursDepuis: lastControl.joursDepuis,
          actionRequise: lastControl.actionRequise,
          verifieeAujourdhui: lastControl.verifieeAujourdhui
        });
      }
    }
    
    console.log(`✅ ${mallettes.length} mallettes chargées pour dashboard`);
    return mallettes;
    
  } catch (error) {
    console.error("❌ Erreur dans getMallettesDataForDashboard:", error);
    console.error("Stack:", error.stack);
    return [];
  }
}


/**
 * Récupère les infos du dernier contrôle pour une mallette
 */
function getLastControlForMallette(sheetSuivi, malletteName) {
  if (!sheetSuivi || sheetSuivi.getLastRow() <= 1) {
    return {
      date: 'Jamais',
      controleur: '-',
      nbManquants: 0,
      etat: '❌ Non vérifié',
      joursDepuis: '---',
      actionRequise: 'Contrôler',
      verifieeAujourdhui: false
    };
  }
  
  try {
    const data = sheetSuivi.getDataRange().getValues();
    
    // Parcourir les lignes du plus récent au plus ancien
    for (let i = data.length - 1; i >= 1; i--) {
      const malletteControllee = data[i][2]; // Colonne C : MALLETTE contrôlée
      
      // Comparaison exacte
      if (malletteControllee && malletteControllee.toString().trim() === malletteName.trim()) {
        const dateValue = data[i][0]; // Colonne A : Date/Heure
        const controleurName = data[i][1]; // Colonne B : Nom/Prénom
        const nbManquants = data[i][4] || 0; // Colonne E (index 4)

        
        // S'assurer que nbManquants est un nombre
        const nbManquantsNumber = typeof nbManquants === 'number' 
          ? nbManquants 
          : parseInt(nbManquants) || 0;
        
        
        // ========================================================================
        // CORRECTION : Parser la date correctement (objet Date OU texte avec \n)
        // ========================================================================
        let controlDate;
        
        if (dateValue instanceof Date) {
          // Si c'est déjà un objet Date, l'utiliser directement
          controlDate = dateValue;
        } else if (typeof dateValue === 'string') {
          // Si c'est du texte avec format "dd/MM/yyyy\nHH:mm:ss"
          const dateStr = dateValue.toString().replace('\n', ' '); // Remplacer \n par espace
          
          // Parser le format français "dd/MM/yyyy HH:mm:ss"
          const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          
          if (parts) {
            const [, day, month, year, hour, minute, second] = parts;
            controlDate = new Date(year, month - 1, day, hour, minute, second);
          } else {
            // Essayer de parser comme date normale
            controlDate = new Date(dateValue);
          }
        } else {
          // Fallback : essayer de convertir en Date
          controlDate = new Date(dateValue);
        }
        
        // Vérifier que la date est valide
        if (isNaN(controlDate.getTime())) {
          console.warn(`⚠️ Date invalide pour ${malletteName}: ${dateValue}`);
          continue; // Passer à la ligne suivante
        }
        // ========================================================================
        
        const today = new Date();
        
        // ========================================================================
        // CALCUL DES JOURS AVEC RESET À 06H00
        // ========================================================================
        const RESET_HOUR = 6; // Heure de début du "jour de travail"
        
        // Fonction pour calculer le "jour de travail" (qui commence à 06h00)
        function getWorkDay(date) {
          const workDay = new Date(date);
          // Si on est avant 06h00, on est encore dans le "jour précédent"
          if (date.getHours() < RESET_HOUR) {
            workDay.setDate(workDay.getDate() - 1);
          }
          // Retourner la date normalisée à 06h00
          return new Date(workDay.getFullYear(), workDay.getMonth(), workDay.getDate(), RESET_HOUR, 0, 0);
        }
        
        // Calculer les "jours de travail"
        const todayWorkDay = getWorkDay(today);
        const controlWorkDay = getWorkDay(controlDate);
        
        // Comparer les "jours de travail"
        const verifieeAujourdhui = todayWorkDay.getTime() === controlWorkDay.getTime();
        
        // Calculer la différence en jours (basé sur les "jours de travail")
        const diffTime = Math.abs(todayWorkDay - controlWorkDay);
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        
        // État basé sur manquants ET vérification aujourd'hui
        let etat;
        if (!verifieeAujourdhui) {
          etat = '⚠️ Non vérifié aujourd\'hui';
        } else if (nbManquantsNumber > 0) {
          etat = '⚠️ Manquants';
        } else {
          etat = '✅ Conforme';
        }
        
        return {
          date: Utilities.formatDate(controlDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
          controleur: controleurName,
          nbManquants: nbManquantsNumber,
          etat: etat,
          joursDepuis: diffDays,
          actionRequise: verifieeAujourdhui ? (nbManquantsNumber > 0 ? 'Traiter manquants' : '-') : 'Contrôler aujourd\'hui',
          verifieeAujourdhui: verifieeAujourdhui
        };
      }
    }
    
    // Aucun contrôle trouvé
    return {
      date: 'Jamais',
      controleur: '-',
      nbManquants: 0,
      etat: '❌ Non vérifié',
      joursDepuis: '---',
      actionRequise: 'Contrôler',
      verifieeAujourdhui: false
    };
    
  } catch (error) {
    console.error(`❌ Erreur getLastControlForMallette pour ${malletteName}:`, error);
    return {
      date: 'Erreur',
      controleur: '-',
      nbManquants: 0,
      etat: '❌ Erreur',
      joursDepuis: '---',
      actionRequise: 'Vérifier',
      verifieeAujourdhui: false
    };
  }
}
  

/**
 * Crée l'en-tête du dashboard
 */
function createDashboardHeader(sheet) {
  // Ligne 1 : Titre principal
  sheet.getRange('A1:I1').merge();
  sheet.getRange('A1').setValue('🎯 TABLEAU DE BORD - INVENTAIRE MALLETTES');
  
  // Ligne 2 : Date de mise à jour
  sheet.getRange('A2:I2').merge();
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  sheet.getRange('A2').setValue('Dernière mise à jour: ' + dateStr);
  
  // Mise en forme de l'en-tête
  sheet.getRange('A1').setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  sheet.getRange('A2').setFontStyle('italic').setFontSize(10)
    .setHorizontalAlignment('center');
}

/**
 * Crée la section "Vue d'ensemble des mallettes"
 */
function createMallettesOverview(sheet, mallettes) {
  const startRow = 4;
  
  // Titre de section
  sheet.getRange(`A${startRow}:I${startRow}`).merge();
  sheet.getRange(`A${startRow}`).setValue('📦 VUE D\'ENSEMBLE DES MALLETTES');
  sheet.getRange(`A${startRow}`).setBackground('#4285f4').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  
  // En-têtes des colonnes
  const headers = [
    'Mallette',
    'Nb Outils',
    'Dernière Vérif.',
    'Contrôleur',
    'Manquants',
    'État',
    'Jours depuis vérif.',
    'Action requise'
  ];
  
  const headerRow = startRow + 2;
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(headerRow, i + 1).setValue(headers[i]);
  }
  
  // Mise en forme des en-têtes
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  // Données des mallettes
  let dataRow = headerRow + 1;
  mallettes.forEach(mallette => {
    sheet.getRange(dataRow, 1).setValue(mallette.nom);
    sheet.getRange(dataRow, 2).setValue(mallette.nbOutils).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 3).setValue(mallette.derniereVerif).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 4).setValue(mallette.controleur);
    sheet.getRange(dataRow, 5).setValue(mallette.manquants).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 6).setValue(mallette.etat);
    sheet.getRange(dataRow, 7).setValue(mallette.joursDepuis).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 8).setValue(mallette.actionRequise);
    
    // Couleur de fond selon l'état
    if (mallette.etat.includes('Non vérifié')) {
      sheet.getRange(dataRow, 6).setBackground('#ea4335').setFontColor('#ffffff');
    } else if (mallette.etat.includes('Manquants')) {
      sheet.getRange(dataRow, 6).setBackground('#fbbc04').setFontColor('#000000');
    } else {
      sheet.getRange(dataRow, 6).setBackground('#34a853').setFontColor('#ffffff');
    }
    
    // Bordures
    sheet.getRange(dataRow, 1, 1, headers.length)
      .setBorder(true, true, true, true, true, true);
    
    dataRow++;
  });
  
  return dataRow;
}

/**
 * Crée la section "Statistiques globales"
 * VERSION MODIFIÉE pour vérification quotidienne
 */
function createGlobalStats(sheet, mallettes) {
  const startRow = 4 + 2 + 1 + mallettes.length + 2; // Après la vue d'ensemble
  
  // Titre de section
  sheet.getRange(`A${startRow}:F${startRow}`).merge();
  sheet.getRange(`A${startRow}`).setValue('📊 STATISTIQUES GLOBALES');
  sheet.getRange(`A${startRow}`).setBackground('#34a853').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  
  // ──────────────────────────────────────────────────────────────────
  // MODIFICATION 1 : Calculs des statistiques avec vérification quotidienne
  // ──────────────────────────────────────────────────────────────────
  const totalMallettes = mallettes.length;
  const totalOutils = mallettes.reduce((sum, m) => sum + m.nbOutils, 0);
  const totalManquants = mallettes.reduce((sum, m) => sum + m.manquants, 0);
  
  
  //  Mallettes non vérifiées AUJOURD'HUI
  const mallettesNonVerifieesAujourdhui = mallettes.filter(m => !m.verifieeAujourdhui).length;
  
  const mallettesAvecManquants = mallettes.filter(m => m.manquants > 0).length;
  
  // ──────────────────────────────────────────────────────────────────
  // MODIFICATION 2 : Taux de conformité basé sur vérification quotidienne
  // ──────────────────────────────────────────────────────────────────
  
  
  //  Une mallette est NON conforme si :
  // - Elle a des manquants OU
  // - Elle n'a pas été vérifiée aujourd'hui
  const mallettesNonConformes = mallettes.filter(m => {
    return !m.verifieeAujourdhui || m.manquants > 0;
  }).length;
  
  const tauxConformite = totalMallettes > 0 
    ? Math.round(((totalMallettes - mallettesNonConformes) / totalMallettes) * 100) 
    : 0;
  
  // Moyenne des jours depuis vérification
  const joursValides = mallettes.filter(m => typeof m.joursDepuis === 'number');
  const moyenneJours = joursValides.length > 0
    ? Math.round(joursValides.reduce((sum, m) => sum + m.joursDepuis, 0) / joursValides.length)
    : 0;
  
  // ──────────────────────────────────────────────────────────────────
  // MODIFICATION 3 : Ligne "Mallettes à vérifier" utilise la nouvelle variable
  // ──────────────────────────────────────────────────────────────────
  // Disposition des stats (2 colonnes)
  const stats = [
    ['Total mallettes', totalMallettes, 'Total outils', totalOutils],
    ['Total contrôlées ce mois', calculateMallettesCeMois(), 'Manquants signalés', totalManquants],
    // ANCIEN : ['Mallettes à vérifier', mallettesNonVerifiees, 'Signalements ouverts', calculateSignalementsOuverts()],
    // NOUVEAU :
    ['Mallettes à vérifier ce jour', mallettesNonVerifieesAujourdhui, 'Signalements ouverts', calculateSignalementsOuverts()],
    ['Taux de conformité', tauxConformite + '%', 'Temps moyen entre', moyenneJours + ' jours']
  ];
  
  // Le reste de la fonction 
  // (Création des cellules, mise en forme, etc.)
  
  const statsRow = startRow + 2;
  stats.forEach((row, index) => {
    const currentRow = statsRow + index;
    
    // Première paire de stats
    sheet.getRange(currentRow, 1).setValue(row[0]);
    sheet.getRange(currentRow, 2).setValue(row[1]);
    
    // Deuxième paire de stats
    sheet.getRange(currentRow, 4).setValue(row[2]);
    sheet.getRange(currentRow, 5).setValue(row[3]);
    
    // Mise en forme
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 4).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setHorizontalAlignment('right').setFontWeight('bold');
    sheet.getRange(currentRow, 5).setHorizontalAlignment('right').setFontWeight('bold');
  });
  
  // Bordures
  sheet.getRange(statsRow, 1, stats.length, 5)
    .setBorder(true, true, true, true, true, true);
  
  return statsRow + stats.length;
}

/**
 * Crée la section "Alertes et actions requises"
 */
function createAlertesSection(sheet, mallettes) {
  const startRow = 4 + 2 + 1 + mallettes.length + 2 + 1 + 4 + 2; // Après les stats
  
  // Titre de section
  sheet.getRange(`A${startRow}:I${startRow}`).merge();
  sheet.getRange(`A${startRow}`).setValue('⚠️ ALERTES ET ACTIONS REQUISES');
  sheet.getRange(`A${startRow}`).setBackground('#ea4335').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  
  // Vérifier s'il y a des alertes
  const alertes = [];
  
  // Mallettes non vérifiées
  mallettes.filter(m => m.etat.includes('Non vérifié')).forEach(m => {
    alertes.push(`❌ ${m.nom} n'a pas été vérifiée`);
  });
  
  // Mallettes avec manquants
  mallettes.filter(m => m.manquants > 0).forEach(m => {
    alertes.push(`⚠️ ${m.nom} : ${m.manquants} outil(s) manquant(s)`);
  });
  
  // Mallettes non vérifiées depuis longtemps (>10 jours)
  mallettes.filter(m => typeof m.joursDepuis === 'number' && m.joursDepuis > 10).forEach(m => {
    alertes.push(`📅 ${m.nom} : Dernier contrôle il y a ${m.joursDepuis} jours`);
  });
  
  let alerteRow = startRow + 2;
  if (alertes.length === 0) {
    sheet.getRange(alerteRow, 1, 1, 9).merge();
    sheet.getRange(alerteRow, 1).setValue('✅ Aucune alerte en cours')
      .setFontStyle('italic')
      .setHorizontalAlignment('center')
      .setBackground('#e8f5e9');
  } else {
    alertes.forEach(alerte => {
      sheet.getRange(alerteRow, 1, 1, 9).merge();
      sheet.getRange(alerteRow, 1).setValue(alerte)
        .setBackground('#fce8e6');
      alerteRow++;
    });
  }
}

/**
 * Mise en forme finale du dashboard
 */
function formatDashboard(sheet) {
  // Figer les lignes d'en-tête
  sheet.setFrozenRows(2);
  
  // Ajuster les largeurs de colonnes
  sheet.setColumnWidth(1, 250); // Mallette
  sheet.setColumnWidth(2, 80);  // Nb Outils
  sheet.setColumnWidth(3, 120); // Dernière Vérif.
  sheet.setColumnWidth(4, 150); // Contrôleur
  sheet.setColumnWidth(5, 80);  // Manquants
  sheet.setColumnWidth(6, 130); // État
  sheet.setColumnWidth(7, 130); // Jours depuis vérif.
  sheet.setColumnWidth(8, 120); // Action requise
  
  // Ajuster la hauteur des lignes
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 25);
}

/**
 * Fonction de test pour vérifier que tout fonctionne
 * LANCER CETTE FONCTION POUR TESTER
 */
function testDashboardCreation() {
  console.log("🧪 TEST DE CRÉATION DU DASHBOARD");
  console.log("================================");
  
  try {
    // Tester l'accès au Spreadsheet
    console.log("1. Test accès Spreadsheet...");
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log(`   ✅ Spreadsheet trouvé: ${ss.getName()}`);
    
    // Tester l'accès à la feuille inventaire
    console.log("2. Test accès feuille inventaire...");
    const sheetInventaire = ss.getSheetByName('Sheet pour inventaire');
    if (!sheetInventaire) {
      throw new Error("Feuille 'Sheet pour inventaire' introuvable");
    }
    console.log(`   ✅ Feuille inventaire trouvée: ${sheetInventaire.getName()}`);
    
    // Tester le chargement des mallettes
    console.log("3. Test chargement des mallettes...");
    const mallettes = getMallettesDataForDashboard(sheetInventaire, null);
    console.log(`   ✅ ${mallettes.length} mallettes chargées`);
    
    // Afficher les mallettes
    mallettes.forEach(m => {
      console.log(`   - ${m.nom}: ${m.nbOutils} outils`);
    });
    
    console.log("\n✅ TOUS LES TESTS SONT PASSÉS !");
    console.log("Vous pouvez maintenant exécuter createDashboard()");
    
  } catch (error) {
    console.error("\n❌ ÉCHEC DU TEST:");
    console.error("Message:", error.message);
    console.error("Stack:", error.stack);
  }
}
function calculateMallettesCeMois() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetSuivi = ss.getSheetByName('Suivi_WebApp');
    
    if (!sheetSuivi || sheetSuivi.getLastRow() <= 1) {
      return 0;
    }
    
    const data = sheetSuivi.getDataRange().getValues();
    
    // Obtenir le premier jour du mois en cours
    const now = new Date();
    const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    
    let count = 0;
    
    // Parcourir toutes les lignes (sauf l'en-tête)
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][0]; // Colonne A : Date/Heure
      const mallette = data[i][2];   // Colonne C : MALLETTE contrôlée
      
      if (dateValue && mallette) {
        // ========================================================================
        // CORRECTION : Parser la date correctement (objet Date OU texte avec \n)
        // ========================================================================
        let controlDate;
        
        if (dateValue instanceof Date) {
          controlDate = dateValue;
        } else if (typeof dateValue === 'string') {
          const dateStr = dateValue.toString().replace('\n', ' ');
          const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          
          if (parts) {
            const [, day, month, year, hour, minute, second] = parts;
            controlDate = new Date(year, month - 1, day, hour, minute, second);
          } else {
            controlDate = new Date(dateValue);
          }
        } else {
          controlDate = new Date(dateValue);
        }
        
        // Vérifier que la date est valide
        if (isNaN(controlDate.getTime())) {
          console.warn(`⚠️ Date invalide ligne ${i}: ${dateValue}`);
          continue;
        }
        // ========================================================================
        
        // Si la date du contrôle est dans le mois en cours
        if (controlDate >= firstDayOfMonth && controlDate <= now) {
          count++; // Compte CHAQUE ligne = CHAQUE mallette
        }
      }
    }
    
    console.log(`📊 Mallettes contrôlées ce mois : ${count}`);
    return count;
    
  } catch (error) {
    console.error("❌ Erreur calcul mallettes ce mois:", error);
    console.error("Stack:", error.stack);
    return 0;
  }
}
/**
 * Calcule le nombre de mallettes à vérifier aujourd'hui
 * (mallettes non vérifiées dans la journée en cours)
 */
function calculateMallettesAVerifierAujourdhui() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetInventaire = ss.getSheetByName(CONFIG.sheets.inventaire);
    const sheetSuivi = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!sheetInventaire) {
      console.error("❌ Feuille inventaire introuvable");
      return 0;
    }
    
    // Récupérer toutes les mallettes
    const mallettes = getMallettesDataForDashboard(sheetInventaire, sheetSuivi);
    
    if (!mallettes || mallettes.length === 0) {
      console.log("⚠️ Aucune mallette trouvée");
      return 0;
    }
    
    // Compter les mallettes NON vérifiées aujourd'hui
    const mallettesNonVerifiees = mallettes.filter(m => {
      return m.verifieeAujourdhui === false;
    }).length;
    
    console.log(`📊 Mallettes à vérifier aujourd'hui : ${mallettesNonVerifiees}/${mallettes.length}`);
    return mallettesNonVerifiees;
    
  } catch (error) {
    console.error("❌ Erreur calcul mallettes à vérifier:", error);
    return 0;
  }
}
function calculateSignalementsOuverts() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetSuivi = ss.getSheetByName('Suivi_WebApp');
    
    if (!sheetSuivi || sheetSuivi.getLastRow() <= 1) {
      return 0;
    }
    
    const data = sheetSuivi.getDataRange().getValues();
    
    let count = 0;
    
    // Parcourir toutes les lignes (sauf l'en-tête)
    for (let i = 1; i < data.length; i++) {
      // ──────────────────────────────────────────────────────────────────
      // MODIFICATION : Colonne G (index 6) avec nouvelle structure
      // (Anciennement colonne H/index 7)
      // ──────────────────────────────────────────────────────────────────
      const typeSignalement = data[i][6]; // Colonne G : Type Signalement
      
      // Si un signalement est renseigné (non vide)
      if (typeSignalement && typeSignalement.toString().trim() !== '') {
        // ──────────────────────────────────────────────────────────────────
        // NOUVEAU : Compter le NOMBRE de types dans cette cellule
        // ──────────────────────────────────────────────────────────────────
        const types = typeSignalement.toString().trim().split('\n');
        
        // Compter chaque type (filtrer les lignes vides)
        const nbTypes = types.filter(type => type.trim() !== '').length;
        
        count += nbTypes; // ← Ajoute le NOMBRE de types, pas juste 1
        
        console.log(`  Ligne ${i}: ${nbTypes} type(s) - ${types.join(', ')}`);
      }
    }
    
    console.log(`📊 Signalements ouverts (TOTAL) : ${count}`);
    return count;
    
  } catch (error) {
    console.error("❌ Erreur calcul signalements:", error);
    return 0;
  }
}

// ==========================================
// CONFIGURATION - À MODIFIER AVEC VOS IDS
// ==========================================

/**
 * Configuration principale du système d'inventaire
 * IMPORTANT: Remplacez ces IDs par les vôtres
 */

// ID de votre Google Sheet
// Trouvez-le dans l'URL: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
const SPREADSHEET_ID = '1n_r1TR_b03ZRvW_1cUmToPeGRmDG7TRJS44BNqhjUy4';

// Nom de la feuille contenant l'inventaire
const SHEET_NAME = 'Sheet pour inventaire';

// Email pour les notifications (optionnel)
// Si vide, utilisera l'email du propriétaire du script
const NOTIFICATION_EMAIL = 'noemie.maerten.external@airbus.com';

// Email pour le rapport hebdomadaire (N+2)
const WEEKLY_REPORT_EMAIL = 'noemie.maerten.external@airbus.com';



// ==========================================
// PARAMÈTRES AVANCÉS (OPTIONNEL)
// ==========================================

const CONFIG = {
  // Envoyer des notifications par email
  enableEmailNotifications: true,
  
  // Email de notification (si vide, utilise l'email du propriétaire)
  notificationEmail: NOTIFICATION_EMAIL || Session.getActiveUser().getEmail(),
  
  // Email pour rapport hebdomadaire
  weeklyReportEmail: WEEKLY_REPORT_EMAIL,

  // Changement des sous-titres
  formTitle: 'Inventaire des mallettes',
  formSubtitle: 'Inventaire des moyens de contrôle - XWB BARQUE T12',
  
  // Seuil pour notification urgente
  urgentKeywords: ['urgent', 'bloqueant', '🔴'],
  
  // Seuils d'alerte pour conformité
  thresholds: {
    excellent: 100,    // 100% = Conforme
    good: 80,          // 99-80% = À surveiller
    critical: 80       // <80% = Action requise
  },
  
  // Couleurs pour le dashboard
  colors: {
    header: '#2196F3',
    success: '#4CAF50',
    warning: '#FFC107',
    danger: '#F44336',
    info: '#00BCD4'
  },
  
  // Format de date
  dateFormat: 'dd/MM/yyyy HH:mm:ss',
  
  // Nom des feuilles de suivi
  sheets: {
    inventaire: 'Sheet pour inventaire',
    suivi: 'Suivi_WebApp',
    erreurs: 'Erreurs_Validation',
    dashboard: 'Dashboard',
    historique: 'Historique_Hebdo'
  }
};

/**
 * Fonction de test pour vérifier la configuration
 */
function testConfiguration() {
  try {
    console.log("🔍 Test de la configuration...");
    
    // Test 1: Accès au Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log("✅ Spreadsheet accessible:", ss.getName());
    
    // Test 2: Accès à la feuille d'inventaire
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`La feuille "${SHEET_NAME}" n'existe pas`);
    }
    console.log("✅ Feuille d'inventaire accessible");
    
    // Test 3: Structure de la feuille
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const mallettes = headers.filter(h => h.toString().toLowerCase().includes('mallette'));
    console.log(`✅ ${mallettes.length} mallettes détectées:`, mallettes.join(', '));
    
    // Test 4: Email de notification
    console.log("✅ Email de notification:", CONFIG.notificationEmail);
    console.log("✅ Email rapport hebdomadaire:", CONFIG.weeklyReportEmail);
    
    console.log("\n✅ CONFIGURATION VALIDÉE - Tout est OK!");
    return true;
    
  } catch (error) {
    console.error("❌ ERREUR DE CONFIGURATION:", error);
    console.error("\n📋 Vérifiez:");
    console.error("1. Que SPREADSHEET_ID est correct");
    console.error("2. Que la feuille 'Sheet pour inventaire' existe");
    console.error("3. Que vous avez les permissions nécessaires");
    return false;
  }
}

/**
 * Fonction de test supplémentaire pour déboguer le chargement des mallettes
 */
function testMallettesLoad() {
  try {
    console.log("🔍 Test du chargement des mallettes...\n");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const mallettes = getMallettesFromSheet(sheet);
    
    console.log(`📦 Nombre de mallettes trouvées: ${mallettes.length}\n`);
    
    if (mallettes.length === 0) {
      console.error("❌ AUCUNE MALLETTE TROUVÉE !");
      console.error("Vérifiez que la première ligne de votre Sheet contient des noms avec le mot 'MALLETTE'");
      return false;
    }
    
    mallettes.forEach((m, i) => {
      console.log(`${i + 1}. ${m.nom}`);
      console.log(`   → ${m.nombreOutils} outils`);
      if (m.nombreOutils > 0) {
        console.log(`   → Premier outil: ${m.outils[0]}`);
      }
    });
    
    console.log("\n✅ Chargement des mallettes OK !");
    console.log("\n💡 Si les mallettes n'apparaissent toujours pas dans la WebApp:");
    console.log("1. Vérifiez que le fichier index.html est bien présent");
    console.log("2. Redéployez la WebApp avec 'Nouvelle version'");
    console.log("3. Testez l'URL en navigation privée (pour éviter le cache)");
    
    return true;
    
  } catch (error) {
    console.error("❌ ERREUR:", error);
    return false;
  }
}

/**
 * Fonction pour afficher un diagnostic complet
 */
function diagnosticComplet() {
  console.log("=" .repeat(60));
  console.log("🔧 DIAGNOSTIC COMPLET DU SYSTÈME");
  console.log("=" .repeat(60));
  console.log("");
  
  // Test 1: Configuration
  console.log("📋 TEST 1: Configuration de base");
  console.log("-" .repeat(60));
  const configOK = testConfiguration();
  console.log("");
  
  // Test 2: Chargement des mallettes
  console.log("📦 TEST 2: Chargement des mallettes");
  console.log("-" .repeat(60));
  const mallettesOK = testMallettesLoad();
  console.log("");
  
  // Test 3: Permissions email
  console.log("📧 TEST 3: Permissions email");
  console.log("-" .repeat(60));
  try {
    const testEmail = Session.getActiveUser().getEmail();
    console.log("✅ Email détecté:", testEmail);
  } catch (e) {
    console.error("❌ Impossible de récupérer l'email:", e);
  }
  console.log("");
  
  // Résumé
  console.log("=" .repeat(60));
  console.log("📊 RÉSUMÉ DU DIAGNOSTIC");
  console.log("=" .repeat(60));
  console.log("Configuration:", configOK ? "✅ OK" : "❌ ERREUR");
  console.log("Mallettes:", mallettesOK ? "✅ OK" : "❌ ERREUR");
  console.log("");
  
  if (configOK && mallettesOK) {
    console.log("🎉 TOUT EST OK ! Vous pouvez déployer la WebApp.");
  } else {
    console.log("⚠️ Des problèmes ont été détectés. Consultez le guide DEPANNAGE.md");
  }
  console.log("=" .repeat(60));
}

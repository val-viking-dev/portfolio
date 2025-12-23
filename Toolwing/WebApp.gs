// ==========================================
// GESTION AVANCÉE DE LA LOGIQUE CONDITIONNELLE
// ==========================================

/**
 * Version améliorée pour gérer la logique conditionnelle complexe
 * Google Forms a des limitations, ce script contourne ces limitations
 */

/**
 * Créer une WebApp avec un formulaire HTML dynamique
 * Plus de flexibilité que Google Forms natif
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  
  // Passer les données au template
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  template.mallettesData = JSON.stringify(getMallettesFromSheet(sheet));

  // Ajouter les informations de configuration du sous-titre
  template.formTitle = CONFIG.formTitle || 'Inventaire des Mallettes';
  template.formSubtitle = CONFIG.formSubtitle || '';
  
  return template.evaluate()
    .setTitle('Inventaire Mallettes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Traite la soumission du formulaire WebApp
 */
function processWebFormSubmission(formData) {
  try {
    console.log("📨 Réception de la soumission...");
    console.log("Nom:", formData.nomPrenom);
    console.log("Mallettes:", formData.mallettesControlees);
    
    // Utiliser la fonction d'enregistrement de Code.gs
    const result = saveSubmissionToSheet(formData);
    
    if (!result.success) {
      throw new Error("Échec de l'enregistrement dans le Sheet");
    }
    
    console.log("✅ Données enregistrées à la ligne:", result.row);
    
    // Mettre à jour le Dashboard automatiquement
    try {
        console.log("📊 Mise à jour du Dashboard...");
        createDashboard();
          console.log("✅ Dashboard mis à jour");
    } catch (dashboardError) {
        console.error("⚠️ Erreur lors de la mise à jour du Dashboard:", dashboardError);
        console.error("Stack:", dashboardError.stack);
      // IMPORTANT : Ne pas retourner d'erreur, les données sont déjà sauvegardées
    }
    
    // Envoyer notification si nécessaire
    if (formData.hasManquants === 'oui' || (formData.urgence && formData.urgence.includes('🔴'))) {
      console.log("📧 Envoi de notification...");
      if ( !CONFIG.enableEmailNotifications) {
        console.warn("Envoi d'email DESACTIVE dans CONFIG")
      } else {
      try {
        sendNotificationEmail(formData);
        console.log("✅ Notification envoyée");
      } catch (emailError) {
        console.error("⚠️ Erreur lors de l'envoi de l'email:", emailError);
        // Ne pas faire échouer la soumission si l'email échoue
      }
    }
  } else {
    console.log("Aucune notification nécessaire ( pas de manquants urgents)")
  }  
    
    return {
      success: true,
      message: 'Inventaire enregistré avec succès !'
    };
    
  } catch (error) {
    console.error("❌ Erreur lors du traitement:", error);
    console.error("Stack:", error.stack);
    
    return {
      success: false,
      error: `Erreur : ${error.message || error.toString()}`
    };
  }
}

/**
 * Notification avancée avec formatage riche
 */
function sendAdvancedNotification(data) {
  const recipient = Session.getActiveUser().getEmail();
  const subject = `[INVENTAIRE] ${data.urgence || 'Info'} - ${data.nomPrenom}`;
  
  // Créer un email HTML
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #2196F3; color: white; padding: 20px; text-align: center;">
        <h1>🔧 Alerte Inventaire</h1>
      </div>
      
      <div style="padding: 20px; background: #f5f5f5;">
        <h2>Informations générales</h2>
        <table style="width: 100%; background: white; padding: 10px;">
          <tr>
            <td><strong>Contrôleur:</strong></td>
            <td>${data.nomPrenom}</td>
          </tr>
          <tr>
            <td><strong>Date:</strong></td>
            <td>${new Date().toLocaleString('fr-FR')}</td>
          </tr>
          <tr>
            <td><strong>Mallettes contrôlées:</strong></td>
            <td>${data.mallettesControlees ? data.mallettesControlees.join(', ') : 'N/A'}</td>
          </tr>
        </table>
      </div>
  `;
  
  if (data.hasManquants === 'Oui') {
    htmlBody += `
      <div style="padding: 20px; background: #fff3e0;">
        <h2>⚠️ Manquants signalés</h2>
        <div style="background: white; padding: 10px; margin-top: 10px;">
          ${data.manquantsDetails || 'Détails non fournis'}
        </div>
      </div>
    `;
  }
  
  if (data.description) {
    htmlBody += `
      <div style="padding: 20px; background: #e3f2fd;">
        <h2>📝 Signalement</h2>
        <div style="background: white; padding: 10px; margin-top: 10px;">
          <p><strong>Type:</strong> ${data.typeSignalement || 'Non spécifié'}</p>
          <p><strong>Urgence:</strong> ${data.urgence || 'Non spécifiée'}</p>
          <p><strong>Description:</strong><br>${data.description}</p>
        </div>
      </div>
    `;
  }
  
  htmlBody += `
      <div style="padding: 20px; background: #263238; color: white; text-align: center;">
        <p>Système d'inventaire automatique</p>
        <p style="font-size: 12px;">Ne pas répondre à cet email automatique</p>
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    console.log("📧 Notification avancée envoyée");
  } catch (error) {
    console.error("Erreur envoi email:", error);
  }
}


// ==========================================
// UTILITAIRES POUR WEBAPP
// ==========================================

/**
 * Inclut des fichiers HTML/CSS/JS dans la WebApp
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Récupère les données des mallettes pour la WebApp
 */
function getMallettesDataForWebApp() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  return getMallettesFromSheet(sheet);
}

/**
 * Récupère l'historique pour une mallette spécifique
 */
function getMalletteHistory(malletteName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const suiviSheet = spreadsheet.getSheetByName('Suivi_Inventaires');
  
  if (!suiviSheet) return [];
  
  const data = suiviSheet.getDataRange().getValues();
  const history = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[2] && row[2].toString().includes(malletteName)) {
      history.push({
        date: row[0],
        controleur: row[1],
        manquants: row[3] === 'Oui',
        signalement: row[6] === 'Oui',
        details: row[7] || ''
      });
    }
  }
  
  return history.reverse(); // Plus récent en premier
}


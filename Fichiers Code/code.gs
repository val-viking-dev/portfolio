// ==========================================
// CODE PRINCIPAL - FONCTIONS UTILITAIRES
// ==========================================

/**
 * Récupère les données des mallettes depuis le Google Sheet
 * Structure: Colonne = Mallette, Lignes 2+ = Outils
 */
function getMallettesFromSheet(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      console.log("⚠️ Aucune donnée trouvée dans la feuille");
      return [];
    }
    
    const headers = data[0]; // Première ligne = noms des mallettes
    const mallettesData = [];
    
    // Pour chaque colonne
    for (let col = 0; col < headers.length; col++) {
      const header = headers[col].toString().trim();
      
      // Vérifier si la colonne contient "MALLETTE" (insensible à la casse)
      if (header && header.toLowerCase().includes('mallette')) {
        const outils = [];
        
        // Récupérer tous les outils de cette colonne (lignes 2 et suivantes)
        for (let row = 1; row < data.length; row++) {
          const cellValue = data[row][col];
          if (cellValue && cellValue.toString().trim() !== '') {
            outils.push(cellValue.toString().trim());
          }
        }
        
        mallettesData.push({
          nom: header,
          outils: outils,
          nombreOutils: outils.length
        });
      }
    }
    
    console.log(`✅ ${mallettesData.length} mallettes chargées`);
    return mallettesData;
    
  } catch (error) {
    console.error("❌ Erreur lors de la lecture des mallettes:", error);
    return [];
  }
}

/**
 * Nettoie un formulaire Google Forms
 */
function clearForm(form) {
  const items = form.getItems();
  items.forEach(item => {
    form.deleteItem(item);
  });
  console.log("🧹 Formulaire nettoyé");
}

/**
 * Crée ou récupère la feuille de suivi
 */
function getOrCreateSuiviSheet() {
  try {
    console.log("🔍 Recherche de la feuille de suivi...");
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!suiviSheet) {
      console.log("📝 Création de la feuille de suivi...");
      suiviSheet = ss.insertSheet(CONFIG.sheets.suivi);
      
      // En-têtes (11 colonnes - Description et JSON supprimées)
      const headers = [
            'Date/Heure',
            'Nom/Prénom',
            'MALLETTE contrôlée',      // ← MODIFIÉ : Singulier (1 mallette par ligne)
            'MANQUANTS',               // ← MODIFIÉ : Pour CETTE mallette
            'Nb Outils Manquants',     // ← Pour CETTE mallette
            'Liste des outils manquants', // ← Pour CETTE mallette
            'Type Signalement',        // ← Pour CETTE mallette
            'Urgence',                 // ← Pour CETTE mallette
            'Description'              // ← Colonne "Signalements détaillés" SUPPRIMÉE
];
      
      suiviSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      suiviSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground(CONFIG.colors.header)
        .setFontColor('white')
        .setFontSize(11)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Figer la ligne d'en-tête
      suiviSheet.setFrozenRows(1);
      
      // Largeur des colonnes (9 colonnes)
      suiviSheet.setColumnWidth(1, 160); // Date
      suiviSheet.setColumnWidth(2, 150); // Nom
      suiviSheet.setColumnWidth(3, 250); // Mallettes contrôlées
      suiviSheet.setColumnWidth(4, 90);  // Nb Mallettes
      suiviSheet.setColumnWidth(5, 90);  // Manquants?
      suiviSheet.setColumnWidth(6, 120); // Nb Outils Manquants
      suiviSheet.setColumnWidth(7, 350); // Liste des outils manquants
      suiviSheet.setColumnWidth(8, 180); // Type Signalement
      suiviSheet.setColumnWidth(9, 150); // Urgence
      suiviSheet.setColumnWidth(10, 200); // Description
      suiviSheet.setColumnWidth(11, 400); // Signalements détaillés
      // Hauteur de la ligne d'en-tête
      suiviSheet.setRowHeight(1, 40);
      
      console.log("✅ Feuille de suivi créée");
    } else {
      console.log("✅ Feuille de suivi trouvée");
    }
    
    return suiviSheet;
    
  } catch (error) {
    console.error("❌ Erreur lors de la création/récupération de la feuille:", error);
    throw new Error("Impossible de créer la feuille de suivi: " + error.message);
  }
}

/**
 * Enregistre une soumission dans le Google Sheet
 */
function saveSubmissionToSheet(formData) {
  try {
    console.log("💾 Début de l'enregistrement...");
    console.log("Données reçues:", JSON.stringify(formData));
    
    const suiviSheet = getOrCreateSuiviSheet();
    
    if (!suiviSheet) {
      throw new Error("La feuille de suivi n'a pas pu être créée ou récupérée");
    }
    
    console.log("✅ Feuille de suivi accessible");
    
    // ========================================================================
    // NOUVELLE LOGIQUE : 1 LIGNE PAR MALLETTE
    // ========================================================================
    
    const mallettes = Array.isArray(formData.mallettesControlees) 
      ? formData.mallettesControlees 
      : [formData.mallettesControlees];
    
    console.log(`📦 ${mallettes.length} mallette(s) à enregistrer`);
    
    const lastRow = suiviSheet.getLastRow();
    let rowsAdded = 0;
    
    // Parcourir CHAQUE mallette et créer UNE LIGNE par mallette
    mallettes.forEach((mallette, index) => {
      
      // ──────────────────────────────────────────────────────────────────
      // EXTRAIRE LES DONNÉES SPÉCIFIQUES À CETTE MALLETTE
      // ──────────────────────────────────────────────────────────────────
      
      // 1. Manquants pour cette mallette
      let manquantsCount = 0;
      let outilsManquantsDetailles = '';
      let hasManquantsPourCetteMallette = 'NON';
      
      if (formData.hasManquants === 'oui' && formData.manquantsDetails) {
        const outilsManquants = formData.manquantsDetails[mallette];
        if (outilsManquants && outilsManquants.length > 0) {
          manquantsCount = outilsManquants.length;
          hasManquantsPourCetteMallette = 'OUI';
          outilsManquantsDetailles = outilsManquants
            .map((outil, idx) => `${idx + 1}. ${outil}`)
            .join('\n');
        }
      }
      
      // 2. Signalements pour cette mallette
      let typeSignalement = '';
      let urgenceGlobale = '';
      let description = '';
      
      if (formData.signalementsIndividuels) {
        const typesUniques = new Set();
        const urgences = [];
        const descriptions = [];
        
        // Parcourir tous les signalements pour trouver ceux de CETTE mallette
        for (const outilId in formData.signalementsIndividuels) {
          const sig = formData.signalementsIndividuels[outilId];
          
          // Vérifier si ce signalement concerne cette mallette
          if (sig.mallette === mallette && sig.hasSignalement === 'oui' && sig.types && sig.types.length > 0) {
            
            // Collecter types
            sig.types.forEach(type => typesUniques.add(type));
            
            // Collecter urgences
            if (sig.urgence) {
              urgences.push(sig.urgence);
            }
            
            // Collecter descriptions
            if (sig.description) {
              descriptions.push(`${sig.outil}: ${sig.description}`);
            }
          }
        }
        
        // Compiler les types
        if (typesUniques.size > 0) {
          typeSignalement = Array.from(typesUniques).join('\n');
        }
        
        // Déterminer l'urgence maximale
        if (urgences.length > 0) {
          if (urgences.includes('urgent')) {
            urgenceGlobale = '🔴 Urgent';
          } else if (urgences.includes('important')) {
            urgenceGlobale = '🟠 Important';
          } else if (urgences.includes('normal')) {
            urgenceGlobale = '🟡 Normal';
          } else if (urgences.includes('faible')) {
            urgenceGlobale = '🟢 Faible';
          }
        }
        
        // Compiler les descriptions
        if (descriptions.length > 0) {
          description = descriptions.join('\n');
        }
      }
      
      // ──────────────────────────────────────────────────────────────────
      // CRÉER LA LIGNE POUR CETTE MALLETTE
      // ──────────────────────────────────────────────────────────────────
      const now = new Date();
      const dateFormatee = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n' +
                           Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      const rowData = [
        dateFormatee, // A : Date/Heure
        formData.nomPrenom || '', // B : Nom/Prénom
        mallette, // C : MALLETTE contrôlée (UNE seule !)
        hasManquantsPourCetteMallette, // D : MANQUANTS
        manquantsCount, // E : Nb Outils Manquants
        outilsManquantsDetailles, // F : Liste des outils manquants
        typeSignalement, // G : Type Signalement
        urgenceGlobale, // H : Urgence
        description // I : Description
      ];
      
      console.log(`📝 Ligne ${index + 1}/${mallettes.length} préparée pour ${mallette}`);
      
      // Ajouter la ligne
      const newRow = lastRow + rowsAdded + 1;
      suiviSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
      
      // ──────────────────────────────────────────────────────────────────
      // MISE EN FORME CONDITIONNELLE
      // ──────────────────────────────────────────────────────────────────
      
      // Format général
      suiviSheet.getRange(newRow, 1, 1, rowData.length)
        .setVerticalAlignment('top')
        .setWrap(true);
      
      // Si manquants pour cette mallette, mettre en orange
      if (hasManquantsPourCetteMallette === 'OUI') {
        suiviSheet.getRange(newRow, 1, 1, rowData.length)
          .setBackground('#fff3e0');
        
        suiviSheet.getRange(newRow, 4, 1, 1) // Colonne MANQUANTS
          .setFontWeight('bold')
          .setFontColor('#e65100');
        
        suiviSheet.getRange(newRow, 5, 1, 1) // Nb Outils
          .setFontWeight('bold')
          .setFontColor('#e65100')
          .setHorizontalAlignment('center');
      }
      
      // Si urgent, mettre en rouge
      if (urgenceGlobale && urgenceGlobale.includes('🔴')) {
        suiviSheet.getRange(newRow, 1, 1, rowData.length)
          .setBackground('#ffebee')
          .setFontWeight('bold');
      }
      
      // Ajuster hauteur si beaucoup d'outils manquants
      if (outilsManquantsDetailles.length > 100) {
        suiviSheet.setRowHeight(newRow, Math.min(300, 50 + outilsManquantsDetailles.split('\n').length * 15));
      }
      
      // Centrer colonnes numériques
      suiviSheet.getRange(newRow, 4, 1, 1).setHorizontalAlignment('center'); // MANQUANTS
      
      rowsAdded++;
    });
    
    console.log(`✅ ${rowsAdded} ligne(s) ajoutée(s) avec succès`);
    return { success: true, row: lastRow + 1 };
    
  } catch (error) {
    console.error("❌ Erreur lors de l'enregistrement:", error);
    console.error("Stack trace:", error.stack);
    throw error;
  }
}

/**
 * Envoie une notification par email
 */
function sendNotificationEmail(formData) {
  if (!CONFIG.enableEmailNotifications) {
    console.log("📧 Notifications désactivées");
    return;
  }
  
  try {
    const recipient = CONFIG.notificationEmail;
    
    // Déterminer la priorité
    const isUrgent = formData.urgence && formData.urgence.includes('🔴');
    const hasManquants = formData.hasManquants === 'oui';
    
    // Sujet de l'email
    let subject = '[INVENTAIRE] ';
    if (isUrgent) {
      subject += '🚨 URGENT - ';
    } else if (hasManquants) {
      subject += '⚠️ Manquants - ';
    } else {
      subject += '✅ ';
    }
    subject += `Contrôle par ${formData.nomPrenom}`;
    
    // Corps de l'email en HTML
    let htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: ${CONFIG.colors.header}; color: white; padding: 20px; text-align: center;">
          <h1 style="margin: 0;">📦 Inventaire Mallettes</h1>
          <p style="margin: 10px 0 0 0;">Nouveau contrôle enregistré</p>
        </div>
        
        <div style="padding: 20px; background: #f5f5f5;">
          <h2>Informations générales</h2>
          <table style="width: 100%; background: white; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Contrôleur:</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">${formData.nomPrenom}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Date:</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">${new Date().toLocaleString('fr-FR')}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Mallettes contrôlées:</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">${
                Array.isArray(formData.mallettesControlees) 
                  ? formData.mallettesControlees.join(', ') 
                  : formData.mallettesControlees || 'N/A'
              }</td>
            </tr>
          </table>
        </div>
    `;
    
    // Section Manquants
    if (hasManquants) {
      htmlBody += `
        <div style="padding: 20px; background: #fff3e0;">
          <h2 style="color: #f57c00;">⚠️ Manquants signalés</h2>
          <div style="background: white; padding: 15px; border-left: 4px solid ${CONFIG.colors.warning};">
      `;
      
      if (formData.manquantsDetails) {
        for (const [mallette, outils] of Object.entries(formData.manquantsDetails)) {
          if (outils && outils.length > 0) {
            htmlBody += `
              <p><strong>${mallette}:</strong></p>
              <ul>
            `;
            outils.forEach(outil => {
              htmlBody += `<li>${outil}</li>`;
            });
            htmlBody += `</ul>`;
          }
        }
      }
      
      htmlBody += `
          </div>
        </div>
      `;
    }
    
    // Section Signalement
    if (formData.description) {
      const bgColor = isUrgent ? '#ffebee' : '#e3f2fd';
      const borderColor = isUrgent ? CONFIG.colors.danger : CONFIG.colors.info;
      
      htmlBody += `
        <div style="padding: 20px; background: ${bgColor};">
          <h2>📝 Signalement</h2>
          <div style="background: white; padding: 15px; border-left: 4px solid ${borderColor};">
            <p><strong>Type:</strong> ${
              Array.isArray(formData.typeSignalement) 
                ? formData.typeSignalement.join(', ') 
                : formData.typeSignalement || 'Non spécifié'
            }</p>
            <p><strong>Urgence:</strong> ${formData.urgence || 'Non spécifiée'}</p>
            <p><strong>Description:</strong></p>
            <p style="background: #f5f5f5; padding: 10px; border-radius: 5px;">${formData.description}</p>
          </div>
        </div>
      `;
    }
    
    // Pied de page
    htmlBody += `
        <div style="padding: 20px; background: #263238; color: white; text-align: center;">
          <p style="margin: 0;">Système d'inventaire automatique - XWB BARQUE</p>
          <p style="font-size: 12px; margin: 10px 0 0 0; opacity: 0.7;">Ne pas répondre à cet email automatique</p>
        </div>
      </div>
    `;
    
    // Envoyer l'email
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    
    console.log(`📧 Notification envoyée à ${recipient}`);
    
  } catch (error) {
    console.error("❌ Erreur lors de l'envoi de l'email:", error);
    // Ne pas faire échouer la soumission si l'email ne part pas
  }
}
/**
 * Envoie le rapport quotidien de contrôle à 16h00
 * Cette fonction doit être configurée avec un trigger quotidien
 */
function sendDailyReport() {
  try {
    console.log("📧 Génération du rapport quotidien...");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetSuivi = ss.getSheetByName(CONFIG.sheets.suivi);
    const sheetInventaire = ss.getSheetByName(CONFIG.sheets.inventaire);
    
    if (!sheetSuivi || !sheetInventaire) {
      console.error("❌ Feuilles introuvables");
      return;
    }
    
    // Récupérer la date d'aujourd'hui (sans heure)
    const today = new Date();
    const todayDateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    
    // Récupérer toutes les données du suivi
    const dataSuivi = sheetSuivi.getDataRange().getValues();
    
    // Récupérer toutes les mallettes
    const toutesMallettes = getMallettesDataForDashboard(sheetInventaire, sheetSuivi);
    
    // 1. MALLETTES VÉRIFIÉES AUJOURD'HUI
    const mallettesVerifieesAujourdhui = [];
    const manquantsAujourdhui = [];
    const signalementsAujourdhui = [];
    
    for (let i = 1; i < dataSuivi.length; i++) {
      const dateControl = new Date(dataSuivi[i][0]);
      const dateControlOnly = new Date(dateControl.getFullYear(), dateControl.getMonth(), dateControl.getDate());
      
      if (dateControlOnly.getTime() === todayDateOnly.getTime()) {
        const mallette = dataSuivi[i][2];
        const controleur = dataSuivi[i][1];
        const manquants = dataSuivi[i][3];
        const nbManquants = dataSuivi[i][4] || 0;
        const listeManquants = dataSuivi[i][5] || '';
        const typeSignalement = dataSuivi[i][6] || '';
        const urgence = dataSuivi[i][7] || '';
        const description = dataSuivi[i][8] || '';
        
        mallettesVerifieesAujourdhui.push({
          mallette: mallette,
          controleur: controleur,
          heure: Utilities.formatDate(dateControl, Session.getScriptTimeZone(), 'HH:mm'),
          manquants: manquants === 'OUI',
          nbManquants: nbManquants
        });
        
        if (manquants === 'OUI' && nbManquants > 0) {
          manquantsAujourdhui.push({
            mallette: mallette,
            nbManquants: nbManquants,
            liste: listeManquants
          });
        }
        
        if (typeSignalement && typeSignalement.toString().trim() !== '') {
          signalementsAujourdhui.push({
            mallette: mallette,
            types: typeSignalement,
            urgence: urgence,
            description: description
          });
        }
      }
    }
    
    // 2. MALLETTES NON CONTRÔLÉES AUJOURD'HUI
    const mallettesNonControlees = toutesMallettes.filter(m => !m.verifieeAujourdhui);
    
    // 3. GÉNÉRATION DE L'EMAIL HTML
    const htmlBody = generateDailyReportHTML(
      mallettesVerifieesAujourdhui,
      manquantsAujourdhui,
      signalementsAujourdhui,
      mallettesNonControlees,
      toutesMallettes  // ← MODIFIÉ : passer l'objet complet au lieu de juste .length
    );
    
    // 4. ENVOI DE L'EMAIL
    const recipient = CONFIG.notificationEmail;
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const subject = `📊 Rapport Quotidien ToolWing - ${dateStr}`;
    
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    
    console.log(`✅ Rapport quotidien envoyé à ${recipient}`);
    console.log(`📦 Mallettes vérifiées : ${mallettesVerifieesAujourdhui.length}/${toutesMallettes.length}`);
    console.log(`⚠️ Manquants détectés : ${manquantsAujourdhui.length}`);
    console.log(`🔔 Signalements ouverts : ${signalementsAujourdhui.length}`);
    console.log(`❌ Mallettes non contrôlées : ${mallettesNonControlees.length}`);
    
  } catch (error) {
    console.error("❌ Erreur lors de l'envoi du rapport quotidien:", error);
    
    try {
      MailApp.sendEmail({
        to: CONFIG.notificationEmail,
        subject: "❌ Erreur - Rapport Quotidien ToolWing",
        body: `Une erreur est survenue lors de la génération du rapport quotidien :\n\n${error}\n\nStack:\n${error.stack}`
      });
    } catch (e) {
      console.error("❌ Impossible d'envoyer l'email d'erreur:", e);
    }
  }
}

/**
 * Génère le HTML du rapport quotidien
 */
function generateDailyReportHTML(mallettesVerifiees, manquants, signalements, mallettesNonControlees, toutesMallettes) {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  
  // Compter les mallettes UNIQUES vérifiées (pas les lignes)
  const nbMallettesVerifiees = toutesMallettes.filter(m => m.verifieeAujourdhui).length;
  const totalMallettes = toutesMallettes.length;
  
  const tauxVerification = totalMallettes > 0 
    ? Math.round((nbMallettesVerifiees / totalMallettes) * 100) 
    : 0;
  
  // Utiliser le MÊME calcul que le dashboard
  const mallettesNonConformes = toutesMallettes.filter(m => {
    return !m.verifieeAujourdhui || m.manquants > 0;
  }).length;
  
  const tauxConformite = totalMallettes > 0
    ? Math.round(((totalMallettes - mallettesNonConformes) / totalMallettes) * 100)
    : 0;
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; padding: 20px; background-color: #f5f5f5; }
        .container { background: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #005EB8 0%, #003d82 100%); color: white; padding: 25px; border-radius: 8px 8px 0 0; margin: -30px -30px 30px -30px; text-align: center; }
        .header h1 { margin: 0; font-size: 28px; font-weight: 600; }
        .header p { margin: 10px 0 0 0; opacity: 0.9; font-size: 16px; }
        .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 25px 0; }
        .stat-card { background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #005EB8; }
        .stat-card.success { border-left-color: #34a853; }
        .stat-card.warning { border-left-color: #fbbc04; }
        .stat-card.danger { border-left-color: #ea4335; }
        .stat-label { font-size: 12px; text-transform: uppercase; color: #666; font-weight: 600; letter-spacing: 0.5px; }
        .stat-value { font-size: 32px; font-weight: 700; margin: 5px 0; color: #333; }
        .section { margin: 30px 0; }
        .section-title { font-size: 20px; font-weight: 600; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e0e0e0; color: #005EB8; }
        .table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        .table th { background: #f1f3f4; padding: 12px; text-align: left; font-weight: 600; color: #333; border-bottom: 2px solid #ddd; }
        .table td { padding: 12px; border-bottom: 1px solid #eee; }
        .table tr:hover { background: #f8f9fa; }
        .badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge.success { background: #e6f4ea; color: #137333; }
        .badge.warning { background: #fef7e0; color: #b45309; }
        .badge.danger { background: #fce8e6; color: #c5221f; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #e0e0e0; text-align: center; color: #666; font-size: 14px; }
        .empty-state { text-align: center; padding: 40px; color: #666; font-style: italic; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📊 Rapport Quotidien ToolWing</h1>
          <p>${dateStr} - XWB BARQUE Operations</p>
        </div>
        
        <div class="stats-grid">
          <div class="stat-card success">
            <div class="stat-label">Mallettes vérifiées</div>
            <div class="stat-value">${nbMallettesVerifiees}/${totalMallettes}</div>
            <div style="font-size: 14px; color: #666; margin-top: 5px;">Taux : ${tauxVerification}%</div>
          </div>
          <div class="stat-card ${manquants.length > 0 ? 'warning' : 'success'}">
            <div class="stat-label">Manquants détectés</div>
            <div class="stat-value">${manquants.length}</div>
            <div style="font-size: 14px; color: #666; margin-top: 5px;">Mallettes concernées</div>
          </div>
          <div class="stat-card ${signalements.length > 0 ? 'warning' : 'success'}">
            <div class="stat-label">Signalements ouverts</div>
            <div class="stat-value">${signalements.length}</div>
            <div style="font-size: 14px; color: #666; margin-top: 5px;">À traiter</div>
          </div>
          <div class="stat-card ${mallettesNonControlees.length > 0 ? 'danger' : 'success'}">
            <div class="stat-label">Non contrôlées</div>
            <div class="stat-value">${mallettesNonControlees.length}</div>
            <div style="font-size: 14px; color: #666; margin-top: 5px;">Conformité : ${tauxConformite}%</div>
          </div>
        </div>
        
        <div class="section">
          <div class="section-title">✅ Mallettes vérifiées aujourd'hui (${mallettesVerifiees.length})</div>
  `;
  
  if (mallettesVerifiees.length > 0) {
    html += `
          <table class="table">
            <thead><tr><th>Mallette</th><th>Contrôleur</th><th>Heure</th><th>État</th></tr></thead>
            <tbody>
    `;
    
    mallettesVerifiees.forEach(m => {
      const badge = m.manquants 
        ? '<span class="badge warning">⚠️ Manquants</span>' 
        : '<span class="badge success">✅ Conforme</span>';
      
      html += `<tr><td><strong>${m.mallette}</strong></td><td>${m.controleur}</td><td>${m.heure}</td><td>${badge}</td></tr>`;
    });
    
    html += `</tbody></table>`;
  } else {
    html += `<div class="empty-state">Aucune mallette vérifiée aujourd'hui</div>`;
  }
  
  html += `</div>`;
  
  if (manquants.length > 0) {
    html += `
        <div class="section">
          <div class="section-title">⚠️ Outils manquants (${manquants.length} mallette(s))</div>
          <table class="table">
            <thead><tr><th>Mallette</th><th>Nb manquants</th><th>Détails</th></tr></thead>
            <tbody>
    `;
    
    manquants.forEach(m => {
      html += `<tr><td><strong>${m.mallette}</strong></td><td style="text-align: center;"><span class="badge warning">${m.nbManquants}</span></td><td style="font-size: 13px;">${m.liste.replace(/\n/g, '<br>')}</td></tr>`;
    });
    
    html += `</tbody></table></div>`;
  }
  
  if (signalements.length > 0) {
    html += `
        <div class="section">
          <div class="section-title">🔔 Signalements ouverts (${signalements.length})</div>
          <table class="table">
            <thead><tr><th>Mallette</th><th>Type(s)</th><th>Urgence</th><th>Description</th></tr></thead>
            <tbody>
    `;
    
    signalements.forEach(s => {
      let urgenceBadge = '';
      if (s.urgence.includes('🔴')) urgenceBadge = '<span class="badge danger">🔴 Urgent</span>';
      else if (s.urgence.includes('🟠')) urgenceBadge = '<span class="badge warning">🟠 Important</span>';
      else if (s.urgence.includes('🟢')) urgenceBadge = '<span class="badge success">🟢 Faible</span>';
      
      html += `<tr><td><strong>${s.mallette}</strong></td><td style="font-size: 13px;">${s.types.replace(/\n/g, '<br>')}</td><td>${urgenceBadge}</td><td style="font-size: 13px;">${s.description}</td></tr>`;
    });
    
    html += `</tbody></table></div>`;
  }
  
  if (mallettesNonControlees.length > 0) {
    html += `
        <div class="section">
          <div class="section-title" style="color: #ea4335;">❌ Mallettes non contrôlées - NON CONFORMES (${mallettesNonControlees.length})</div>
          <table class="table">
            <thead><tr><th>Mallette</th><th>Nb outils</th><th>Dernière vérification</th><th>Contrôleur</th></tr></thead>
            <tbody>
    `;
    
    mallettesNonControlees.forEach(m => {
      html += `<tr style="background: #fce8e6;"><td><strong>${m.nom}</strong></td><td style="text-align: center;">${m.nbOutils}</td><td>${m.derniereVerif}</td><td>${m.controleur}</td></tr>`;
    });
    
    html += `
            </tbody>
          </table>
          <div style="padding: 15px; background: #fff3e0; border-left: 4px solid #ea4335; margin-top: 15px; border-radius: 4px;">
            <strong>⚠️ Action requise :</strong> Ces mallettes doivent être contrôlées aujourd'hui pour être conformes.
          </div>
        </div>
    `;
  } else {
    html += `
        <div class="section">
          <div class="section-title" style="color: #34a853;">✅ Toutes les mallettes ont été contrôlées !</div>
          <div style="text-align: center; padding: 30px; background: #e6f4ea; border-radius: 8px;">
            <div style="font-size: 48px; margin-bottom: 10px;">🎉</div>
            <div style="font-size: 18px; color: #137333; font-weight: 600;">100% de conformité aujourd'hui !</div>
          </div>
        </div>
    `;
  }
  
  html += `
        <div class="footer">
          <p><strong>ToolWing V4.0</strong> - Système d'inventaire automatique</p>
          <p style="font-size: 12px; margin-top: 10px; opacity: 0.7;">
            XWB BARQUE Operations - Airbus<br>
            Rapport généré automatiquement le ${dateStr} à 16:00
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return html;
}
/**
 * Configure le trigger quotidien pour le rapport à 16h00
 * IMPORTANT : Exécuter cette fonction UNE SEULE FOIS pour créer le trigger
 */
function setupDailyTrigger() {
  try {
    console.log("⏰ Configuration du trigger quotidien...");
    
    // Supprimer les anciens triggers de sendDailyReport s'ils existent
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendDailyReport') {
        ScriptApp.deleteTrigger(trigger);
        console.log("🗑️ Ancien trigger supprimé");
      }
    });
    
    // Créer un nouveau trigger quotidien à 16h00
    ScriptApp.newTrigger('sendDailyReport')
      .timeBased()
      .atHour(16)
      .everyDays(1)
      .create();
    
    console.log("✅ Trigger quotidien configuré avec succès !");
    console.log("📧 Le rapport sera envoyé tous les jours à 16h00");
    console.log(`📬 Destinataire : ${CONFIG.notificationEmail}`);
    
    const allTriggers = ScriptApp.getProjectTriggers();
    console.log("\n📋 Triggers actifs :");
    allTriggers.forEach((trigger, index) => {
      console.log(`${index + 1}. ${trigger.getHandlerFunction()} - ${trigger.getTriggerSource()}`);
    });
    
    return true;
    
  } catch (error) {
    console.error("❌ Erreur lors de la configuration du trigger:", error);
    return false;
  }
}

/**
 * Supprime le trigger quotidien
 */
function removeDailyTrigger() {
  try {
    console.log("🗑️ Suppression du trigger quotidien...");
    
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendDailyReport') {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
    });
    
    if (count > 0) {
      console.log(`✅ ${count} trigger(s) supprimé(s)`);
    } else {
      console.log("⚠️ Aucun trigger trouvé pour sendDailyReport");
    }
    
    return true;
    
  } catch (error) {
    console.error("❌ Erreur lors de la suppression du trigger:", error);
    return false;
  }
}

/**
 * Liste tous les triggers actifs du projet
 */
function listAllTriggers() {
  try {
    console.log("📋 Liste de tous les triggers actifs :");
    console.log("=".repeat(60));
    
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      console.log("⚠️ Aucun trigger configuré");
      return;
    }
    
    triggers.forEach((trigger, index) => {
      console.log(`\n${index + 1}. Fonction : ${trigger.getHandlerFunction()}`);
      console.log(`   Source : ${trigger.getTriggerSource()}`);
      console.log(`   ID : ${trigger.getUniqueId()}`);
    });
    
    console.log("\n" + "=".repeat(60));
    console.log(`Total : ${triggers.length} trigger(s)`);
    
  } catch (error) {
    console.error("❌ Erreur lors de la liste des triggers:", error);
  }
}

/**
 * Teste l'envoi du rapport immédiatement (sans attendre 16h00)
 */
function testDailyReport() {
  console.log("🧪 TEST : Envoi du rapport quotidien...");
  console.log("=".repeat(60));
  
  try {
    sendDailyReport();
    console.log("\n✅ Test terminé ! Vérifiez votre boîte email.");
  } catch (error) {
    console.error("\n❌ Erreur lors du test:", error);
  }
}

/**
 * Génère des statistiques simples
 */
function generateStats() {
  try {
    const suiviSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONFIG.sheets.suivi);
    if (!suiviSheet) {
      return { error: "Aucune donnée disponible" };
    }
    
    const data = suiviSheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { 
        totalControles: 0,
        totalManquants: 0,
        totalSignalements: 0
      };
    }
    
    let totalControles = data.length - 1; // -1 pour les en-têtes
    let totalManquants = 0;
    let totalSignalements = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === 'OUI') totalManquants++;
      if (data[i][6] && data[i][6].toString().trim() !== '') totalSignalements++;
    }
    
    return {
      totalControles,
      totalManquants,
      totalSignalements,
      dernierControle: data[data.length - 1][0]
    };
    
  } catch (error) {
    console.error("Erreur génération stats:", error);
    return { error: error.toString() };
  }
}
/**
 * Formate la colonne Date/Heure pour affichage sur 2 lignes
 * Exécuter UNE FOIS pour corriger toutes les lignes existantes
 */
function formatSuiviDateColumnDeuxLignes() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!suiviSheet) {
      console.error("❌ Feuille Suivi_WebApp introuvable");
      return;
    }
    
    const lastRow = suiviSheet.getLastRow();
    if (lastRow <= 1) {
      console.log("⚠️ Aucune donnée à formater");
      return;
    }
    
    // Récupérer toutes les dates de la colonne A
    const dates = suiviSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    // Reformater chaque date sur 2 lignes
    const datesFormatees = dates.map(row => {
      if (row[0] instanceof Date) {
        const date = row[0];
        const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        const heureStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
        return [dateStr + '\n' + heureStr];
      } else {
        return row;
      }
    });
    
    // Écrire les nouvelles valeurs
    suiviSheet.getRange(2, 1, datesFormatees.length, 1).setValues(datesFormatees);
    
    // Formater la colonne en texte avec retour à la ligne
    suiviSheet.getRange(1, 1, lastRow, 1)
      .setWrap(true)
      .setVerticalAlignment('top');
    
    console.log("✅ Colonne Date/Heure formatée sur 2 lignes !");
    
  } catch (error) {
    console.error("❌ Erreur formatage colonne:", error);
  }
}
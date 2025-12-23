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

// ==========================================
// RAPPORT HEBDOMADAIRE - NOUVELLES FONCTIONS
// ==========================================

/**
 * Calcule les dates de la semaine précédente (lundi-vendredi)
 * Retourne un objet avec startDate, endDate, weekNumber, year, formattedPeriod
 */
function getPreviousWeekDates() {
  try {
    const today = new Date();
    
    // Calculer le lundi de la semaine précédente
    const dayOfWeek = today.getDay();
    const daysToSubtract = dayOfWeek === 0 ? 6 : (dayOfWeek - 1) + 7; // Si dimanche = 6 jours, sinon (jour - lundi) + 7
    
    const previousMonday = new Date(today);
    previousMonday.setDate(today.getDate() - daysToSubtract);
    previousMonday.setHours(0, 0, 0, 0);
    
    // Calculer le vendredi de la semaine précédente
    const previousFriday = new Date(previousMonday);
    previousFriday.setDate(previousMonday.getDate() + 4);
    previousFriday.setHours(23, 59, 59, 999);
    
    // Calculer le numéro de semaine ISO
    const weekNumber = getWeekNumber(previousMonday);
    const year = previousMonday.getFullYear();
    
    // Format pour affichage
    const formattedStart = Utilities.formatDate(previousMonday, Session.getScriptTimeZone(), 'dd/MM');
    const formattedEnd = Utilities.formatDate(previousFriday, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const formattedPeriod = `${formattedStart} - ${formattedEnd}`;
    
    return {
      startDate: previousMonday,
      endDate: previousFriday,
      weekNumber: weekNumber,
      year: year,
      formattedPeriod: formattedPeriod
    };
    
  } catch (error) {
    console.error("❌ Erreur getPreviousWeekDates:", error);
    throw error;
  }
}

/**
 * Calcule le numéro de semaine ISO
 */
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Calcule les manquants sans doublons (dernier état de chaque mallette)
 */
function calculateManquantsSansDoublonsWeek(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!suiviSheet) {
      throw new Error("Feuille Suivi_WebApp introuvable");
    }
    
    const data = suiviSheet.getDataRange().getValues();
    
    // Grouper par mallette et garder le dernier contrôle
    const dernierControleParMallette = {};
    
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][0];
      let dateControl;
      
      if (dateValue instanceof Date) {
        dateControl = dateValue;
      } else if (typeof dateValue === 'string') {
        const dateStr = dateValue.toString().replace('\n', ' ');
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
          const [, day, month, year, hour, minute, second] = parts;
          dateControl = new Date(year, month - 1, day, hour, minute, second);
        } else {
          dateControl = new Date(dateValue);
        }
      } else {
        dateControl = new Date(dateValue);
      }
      
      // Vérifier que dans la période
      if (dateControl >= startDate && dateControl <= endDate) {
        const mallette = data[i][2];
        const nbManquants = data[i][4] || 0;
        const listeManquants = data[i][5] || '';
        
        // Garder le dernier contrôle
        if (!dernierControleParMallette[mallette] || dateControl > dernierControleParMallette[mallette].date) {
          dernierControleParMallette[mallette] = {
            date: dateControl,
            nbManquants: nbManquants,
            listeManquants: listeManquants
          };
        }
      }
    }
    
    // Calculer le total et la liste
    let totalManquants = 0;
    const mallettesAvecManquants = [];
    
    for (const mallette in dernierControleParMallette) {
      const ctrl = dernierControleParMallette[mallette];
      if (ctrl.nbManquants > 0) {
        totalManquants += ctrl.nbManquants;
        mallettesAvecManquants.push({
          nom: mallette,
          nbManquants: ctrl.nbManquants,
          listeOutils: ctrl.listeManquants,
          derniereDate: Utilities.formatDate(ctrl.date, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        });
      }
    }
    
    return {
      totalManquants: totalManquants,
      mallettesAvecManquants: mallettesAvecManquants
    };
    
  } catch (error) {
    console.error("❌ Erreur calculateManquantsSansDoublonsWeek:", error);
    throw error;
  }
}

/**
 * Calcule la conformité et les jours non-conformes par mallette
 */
function calculateNonConformitesWeek(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    const inventaireSheet = ss.getSheetByName(CONFIG.sheets.inventaire);
    
    if (!suiviSheet || !inventaireSheet) {
      throw new Error("Feuilles introuvables");
    }
    
    const data = suiviSheet.getDataRange().getValues();
    const mallettesInfo = getMallettesFromSheet(inventaireSheet);
    
    // Créer un map mallette -> nombre total d'outils
    const nbOutilsParMallette = {};
    mallettesInfo.forEach(m => {
      nbOutilsParMallette[m.nom] = m.nombreOutils;
    });
    
    // Jours ouvrés de la semaine (lundi = 1, vendredi = 5)
    const joursOuvres = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];
    const joursOuvresMap = {};
    
    for (let i = 0; i < 5; i++) {
      const date = new Date(startDate);
      date.setDate(startDate.getDate() + i);
      joursOuvresMap[date.toDateString()] = joursOuvres[i];
    }
    
    // Analyser chaque mallette
    const mallettesDetail = [];
    
    mallettesInfo.forEach(malletteInfo => {
      const mallette = malletteInfo.nom;
      const nbOutilsTotal = nbOutilsParMallette[mallette] || 0;
      
      const controlesParJour = {};
      const joursNonConformes = [];
      let conformiteJours = 0;
      
      // Collecter tous les contrôles de cette mallette dans la semaine
      for (let i = 1; i < data.length; i++) {
        const dateValue = data[i][0];
        let dateControl;
        
        if (dateValue instanceof Date) {
          dateControl = dateValue;
        } else if (typeof dateValue === 'string') {
          const dateStr = dateValue.toString().replace('\n', ' ');
          const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          if (parts) {
            const [, day, month, year, hour, minute, second] = parts;
            dateControl = new Date(year, month - 1, day, hour, minute, second);
          } else {
            dateControl = new Date(dateValue);
          }
        } else {
          dateControl = new Date(dateValue);
        }
        
        if (dateControl >= startDate && dateControl <= endDate && data[i][2] === mallette) {
          const jourKey = dateControl.toDateString();
          const nbManquants = data[i][4] || 0;
          const typeSignalement = data[i][6] || '';
          
          // Vérifier si "Départ en métrologie" (ne compte pas comme non-conforme)
          const isDepartMetrologie = typeSignalement && typeSignalement.toLowerCase().includes('métrologie');
          
          if (!controlesParJour[jourKey] || dateControl > controlesParJour[jourKey].date) {
            controlesParJour[jourKey] = {
              date: dateControl,
              nbManquants: nbManquants,
              isDepartMetrologie: isDepartMetrologie
            };
          }
        }
      }
      
      // Vérifier chaque jour ouvré
      for (const [jourKey, nomJour] of Object.entries(joursOuvresMap)) {
        const controle = controlesParJour[jourKey];
        
        if (!controle) {
          // Pas de contrôle ce jour
          joursNonConformes.push(`${nomJour} (non contrôlée)`);
        } else {
          // Contrôle existe, vérifier la conformité
          if (controle.isDepartMetrologie) {
            // Départ métrologie = conforme
            conformiteJours++;
          } else {
            // Calculer ratio outils
            const nbOutilsPresents = nbOutilsTotal - controle.nbManquants;
            const ratio = nbOutilsTotal > 0 ? (nbOutilsPresents / nbOutilsTotal) * 100 : 100;
            
            if (ratio === 100) {
              conformiteJours++;
            } else {
              joursNonConformes.push(`${nomJour} (manquants)`);
            }
          }
        }
      }
      
      // Calculer conformité globale de cette mallette
      const conformitePourcentage = nbOutilsTotal > 0 
        ? Math.round(((nbOutilsTotal - (controlesParJour[Object.keys(joursOuvresMap)[Object.keys(joursOuvresMap).length - 1]]?.nbManquants || 0)) / nbOutilsTotal) * 100)
        : 100;
      
      // Nombre de manquants (dernier état)
      const dernierControle = Object.values(controlesParJour).sort((a, b) => b.date - a.date)[0];
      const nbManquants = dernierControle?.nbManquants || 0;
      
      mallettesDetail.push({
        nom: mallette,
        nbOutils: nbOutilsTotal,
        conformite: conformitePourcentage,
        joursNonConformes: joursNonConformes,
        nbManquants: nbManquants
      });
    });
    
    // Calculer taux global
    const tauxConformiteGlobal = mallettesDetail.length > 0
      ? Math.round(mallettesDetail.reduce((sum, m) => sum + m.conformite, 0) / mallettesDetail.length)
      : 0;
    
    return {
      tauxConformiteGlobal: tauxConformiteGlobal,
      mallettesDetail: mallettesDetail
    };
    
  } catch (error) {
    console.error("❌ Erreur calculateNonConformitesWeek:", error);
    throw error;
  }
}

/**
 * Calcule le nombre de signalements de la semaine
 */
function calculateSignalementsWeek(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!suiviSheet) {
      throw new Error("Feuille Suivi_WebApp introuvable");
    }
    
    const data = suiviSheet.getDataRange().getValues();
    const signalements = [];
    
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][0];
      let dateControl;
      
      if (dateValue instanceof Date) {
        dateControl = dateValue;
      } else if (typeof dateValue === 'string') {
        const dateStr = dateValue.toString().replace('\n', ' ');
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
          const [, day, month, year, hour, minute, second] = parts;
          dateControl = new Date(year, month - 1, day, hour, minute, second);
        } else {
          dateControl = new Date(dateValue);
        }
      } else {
        dateControl = new Date(dateValue);
      }
      
      if (dateControl >= startDate && dateControl <= endDate) {
        const typeSignalement = data[i][6] || '';
        
        if (typeSignalement && typeSignalement.toString().trim() !== '') {
          const mallette = data[i][2];
          const urgenceText = data[i][7] || '';
          const description = data[i][8] || '';
          
          // Extraire l'outil depuis la description ou liste manquants
          const listeManquants = data[i][5] || '';
          let outil = 'Non spécifié';
          if (listeManquants && listeManquants.length > 0) {
            const premierOutil = listeManquants.split('\n')[0];
            outil = premierOutil.replace(/^\d+\.\s*/, '');
          }
          
          // Mapper urgence
          let urgence = 'faible';
          if (urgenceText.includes('🔴') || urgenceText.toLowerCase().includes('urgent')) {
            urgence = 'urgent';
          } else if (urgenceText.includes('🟠') || urgenceText.toLowerCase().includes('important')) {
            urgence = 'important';
          }
          
          // Parser types (séparés par \n)
          const types = typeSignalement.split('\n').filter(t => t.trim() !== '');
          
          types.forEach(type => {
            signalements.push({
              mallette: mallette,
              outil: outil,
              type: type,
              urgence: urgence,
              date: Utilities.formatDate(dateControl, Session.getScriptTimeZone(), 'dd/MM/yyyy')
            });
          });
        }
      }
    }
    
    // Compter par urgence
    const parUrgence = {
      urgent: signalements.filter(s => s.urgence === 'urgent').length,
      important: signalements.filter(s => s.urgence === 'important').length,
      faible: signalements.filter(s => s.urgence === 'faible').length
    };
    
    return {
      total: signalements.length,
      parUrgence: parUrgence,
      liste: signalements
    };
    
  } catch (error) {
    console.error("❌ Erreur calculateSignalementsWeek:", error);
    throw error;
  }
}

/**
 * Calcule le nombre de contrôles effectués dans la semaine
 */
function calculateControlesEffectues(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const suiviSheet = ss.getSheetByName(CONFIG.sheets.suivi);
    
    if (!suiviSheet) {
      return 0;
    }
    
    const data = suiviSheet.getDataRange().getValues();
    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][0];
      let dateControl;
      
      if (dateValue instanceof Date) {
        dateControl = dateValue;
      } else if (typeof dateValue === 'string') {
        const dateStr = dateValue.toString().replace('\n', ' ');
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
          const [, day, month, year, hour, minute, second] = parts;
          dateControl = new Date(year, month - 1, day, hour, minute, second);
        } else {
          dateControl = new Date(dateValue);
        }
      } else {
        dateControl = new Date(dateValue);
      }
      
      if (dateControl >= startDate && dateControl <= endDate) {
        count++;
      }
    }
    
    return count;
    
  } catch (error) {
    console.error("❌ Erreur calculateControlesEffectues:", error);
    return 0;
  }
}

/**
 * Compile les données par mallette en format JSON
 */
function compileDonneesJSON(mallettesDetail) {
  const json = {};
  mallettesDetail.forEach(m => {
    json[m.nom] = {
      conformite: m.conformite,
      nbOutils: m.nbOutils,
      manquants: m.nbManquants,
      joursNonConformes: m.joursNonConformes,
      nbJoursNonConformes: m.joursNonConformes.length
    };
  });
  return json;
}

/**
 * Génère le HTML du rapport hebdomadaire basé sur le modèle V3_FINAL
 */
function generateWeeklyReportHTML(weekData, lastWeekData) {
  const styles = `
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 900px; margin: 20px auto; padding: 20px; background-color: #f5f5f5; }
      .email-container { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden; }
      .header { background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%); color: white; padding: 30px; text-align: center; }
      .header h1 { margin: 0; font-size: 28px; font-weight: 600; }
      .header p { margin: 10px 0 0 0; font-size: 16px; opacity: 0.95; }
      .section { padding: 25px 30px; border-bottom: 1px solid #e0e0e0; }
      .section:last-child { border-bottom: none; }
      .section-title { font-size: 20px; font-weight: 600; color: #1976D2; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #1976D2; }
      .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
      .kpi-card { background: #f8f9fa; border-left: 4px solid #1976D2; padding: 15px; border-radius: 4px; }
      .kpi-value { font-size: 32px; font-weight: 700; color: #1976D2; margin: 5px 0; }
      .kpi-label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
      table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 14px; }
      th { background: #1976D2; color: white; padding: 12px 8px; text-align: left; font-weight: 600; font-size: 13px; }
      td { padding: 12px 8px; border-bottom: 1px solid #e0e0e0; }
      tr:hover { background: #f8f9fa; }
      .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; }
      .status-excellent { background: #E8F5E9; color: #2E7D32; }
      .status-good { background: #FFF3E0; color: #E65100; }
      .status-critical { background: #FFEBEE; color: #C62828; }
      .alert-box { background: #FFF3E0; border-left: 4px solid #FF9800; padding: 15px; margin: 10px 0; border-radius: 4px; }
      .alert-box.critical { background: #FFEBEE; border-left-color: #F44336; }
      .alert-box.success { background: #E8F5E9; border-left-color: #4CAF50; }
      .alert-title { font-weight: 600; margin-bottom: 8px; font-size: 15px; }
      .alert-list { margin: 8px 0 0 20px; font-size: 14px; }
      .trend { display: inline-flex; align-items: center; gap: 5px; padding: 4px 8px; border-radius: 4px; font-size: 13px; font-weight: 600; }
      .trend-up { background: #E8F5E9; color: #2E7D32; }
      .trend-down { background: #FFEBEE; color: #C62828; }
      .footer { background: #263238; color: #B0BEC5; padding: 20px 30px; text-align: center; font-size: 13px; }
      .footer strong { color: white; font-size: 15px; }
      .legend { background: #f8f9fa; padding: 12px; border-radius: 4px; margin: 15px 0; font-size: 13px; }
    </style>
  `;
  
  // Header
  let html = `
    <!DOCTYPE html>
    <html lang="fr">
    <head>
      <meta charset="UTF-8">
      ${styles}
    </head>
    <body>
      <div class="email-container">
        <div class="header">
          <h1>📊 RAPPORT HEBDOMADAIRE TOOLWING</h1>
          <p>Semaine du ${weekData.formattedPeriod} (Semaine ${weekData.numeroSemaine})</p>
        </div>
  `;
  
  // Synthèse exécutive
  const conformiteTrend = weekData.conformiteGlobale - lastWeekData.conformiteGlobale;
  const manquantsTrend = weekData.manquantsTotal - lastWeekData.manquantsTotal;
  const signalementsTrend = weekData.signalementsTotal - lastWeekData.signalementsTotal;
  
  html += `
    <div class="section">
      <div class="section-title">📌 SYNTHÈSE EXÉCUTIVE</div>
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-label">Taux de conformité global</div>
          <div class="kpi-value">${weekData.conformiteGlobale}%</div>
          <div class="trend ${conformiteTrend >= 0 ? 'trend-up' : 'trend-down'}">
            ${conformiteTrend >= 0 ? '+' : ''}${conformiteTrend}% vs S${lastWeekData.semaine} ${conformiteTrend >= 0 ? '📈' : '📉'}
          </div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Manquants détectés</div>
          <div class="kpi-value">${weekData.manquantsTotal}</div>
          <div class="trend ${manquantsTrend <= 0 ? 'trend-up' : 'trend-down'}">
            ${manquantsTrend} vs S${lastWeekData.semaine} ${manquantsTrend <= 0 ? '✅' : '⚠️'}
          </div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Signalements ouverts</div>
          <div class="kpi-value">${weekData.signalementsTotal}</div>
          <div class="trend ${signalementsTrend <= 0 ? 'trend-up' : 'trend-down'}">
            ${signalementsTrend > 0 ? '+' : ''}${signalementsTrend} vs S${lastWeekData.semaine} ${signalementsTrend <= 0 ? '✅' : '⚠️'}
          </div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Mallettes à risque</div>
          <div class="kpi-value">${weekData.mallettesARisque}</div>
          <div class="kpi-label" style="margin-top: 5px;">(&lt; 80% de conformité)</div>
        </div>
      </div>
    </div>
  `;
  
  // Tableau performance par mallette
  const mallettesSorted = weekData.mallettesDetail.sort((a, b) => a.conformite - b.conformite);
  
  html += `
    <div class="section">
      <div class="section-title">📋 PERFORMANCE PAR MALLETTE</div>
      <table>
        <thead>
          <tr>
            <th>MALLETTE</th>
            <th style="text-align: center;">Nb Outils</th>
            <th style="text-align: center;">Conformité</th>
            <th>Jours non-conformes</th>
            <th style="text-align: center;">Manquants</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  mallettesSorted.forEach(m => {
    const statusClass = m.conformite === 100 ? 'status-excellent' 
                      : m.conformite >= 80 ? 'status-good' 
                      : 'status-critical';
    
    const joursText = m.joursNonConformes.length === 0 
                    ? '<td style="color: #2E7D32;">—</td>'
                    : `<td>${m.joursNonConformes.join(', ')}</td>`;
    
    const manquantsColor = m.nbManquants > 0 ? '#C62828' : '#2E7D32';
    
    html += `
      <tr>
        <td><strong>${m.nom}</strong></td>
        <td style="text-align: center;">${m.nbOutils}</td>
        <td style="text-align: center;">
          <span class="status-badge ${statusClass}">${m.conformite}%</span>
        </td>
        ${joursText}
        <td style="text-align: center; font-weight: 600; color: ${manquantsColor};">${m.nbManquants}</td>
      </tr>
    `;
  });
  
  html += `
        </tbody>
      </table>
      <div class="legend">
        <strong>Légende :</strong>
        <div style="margin-top: 8px;">
          <span class="status-badge status-excellent">100%</span> Conforme &nbsp;&nbsp;
          <span class="status-badge status-good">99-80%</span> À surveiller &nbsp;&nbsp;
          <span class="status-badge status-critical">&lt;80%</span> Action requise
        </div>
        <div style="margin-top: 8px; font-size: 12px; color: #666;">
          <strong>Note :</strong> Le % de conformité prend en compte : (1) les jours de contrôle effectués ET (2) le ratio outils présents/total outils.<br>
          Les signalements "Départ métrologie" n'impactent pas le taux de conformité.
        </div>
      </div>
    </div>
  `;
  
  // Alertes - Mallettes < 80%
  const mallettesARisque = weekData.mallettesDetail.filter(m => m.conformite < 80);
  
  if (mallettesARisque.length > 0) {
    html += `
      <div class="section">
        <div class="section-title">🔴 ALERTES ET ACTIONS RECOMMANDÉES</div>
        <div class="alert-box critical">
          <div class="alert-title">⚠️ ${mallettesARisque.length} mallette(s) ont un taux de conformité &lt; 80%</div>
          <ul class="alert-list">
    `;
    
    mallettesARisque.forEach(m => {
      html += `<li><strong>${m.nom}</strong> : ${m.conformite}% de conformité (${m.nbOutils} outils, ${m.nbManquants} manquants)</li>`;
    });
    
    html += `
          </ul>
          <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #FFCDD2;">
            <strong>→ Action recommandée :</strong> Revoir le processus de contrôle quotidien avec l'équipe.
          </div>
        </div>
    `;
  }
  
  // Alertes - Manquants
  if (weekData.manquantsTotal > 0) {
    html += `
      <div class="alert-box critical">
        <div class="alert-title">⚠️ ${weekData.manquantsTotal} manquants détectés dans ${weekData.mallettesAvecManquants.length} mallette(s)</div>
    `;
    
    weekData.mallettesAvecManquants.forEach(m => {
      html += `
        <div style="background: #E3F2FD; padding: 12px; margin: 8px 0; border-radius: 4px; border-left: 3px solid #1976D2;">
          <strong>${m.nom} :</strong> ${m.nbManquants} manquant(s)
          <div style="margin-top: 5px; font-size: 12px;">
            ${m.listeOutils.replace(/\n/g, '<br>')}
          </div>
        </div>
      `;
    });
    
    html += `</div>`;
  }
  
  // Signalements
  if (weekData.signalementsTotal > 0) {
    html += `
      <div class="alert-box">
        <div class="alert-title">🔔 ${weekData.signalementsTotal} signalement(s) ouvert(s) cette semaine</div>
        <table style="font-size: 13px; margin-top: 10px;">
          <thead>
            <tr>
              <th>Mallette</th>
              <th>Outil concerné</th>
              <th>Type</th>
              <th style="text-align: center;">Urgence</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    weekData.signalements.liste.forEach(s => {
      const urgenceColor = s.urgence === 'urgent' ? '#F44336'
                         : s.urgence === 'important' ? '#FF9800'
                         : '#4CAF50';
      const urgenceText = s.urgence === 'urgent' ? '🔴 Urgent'
                        : s.urgence === 'important' ? '🟠 Important'
                        : '🟢 Faible';
      
      html += `
        <tr>
          <td><strong>${s.mallette}</strong></td>
          <td>${s.outil}</td>
          <td>${s.type}</td>
          <td style="text-align: center;">
            <span style="color: ${urgenceColor}; font-weight: 600;">${urgenceText}</span>
          </td>
        </tr>
      `;
    });
    
    html += `
          </tbody>
        </table>
      </div>
    `;
  }
  
  // Points positifs
  const mallettesConformes = weekData.mallettesDetail.filter(m => m.conformite >= 90).length;
  
  html += `
    <div class="alert-box success">
      <div class="alert-title">✅ Points positifs</div>
      <ul class="alert-list">
        <li>${mallettesConformes} mallettes (${Math.round(mallettesConformes/weekData.mallettesDetail.length*100)}%) ont maintenu une conformité ≥ 90% toute la semaine</li>
        ${conformiteTrend > 0 ? `<li>Amélioration de +${conformiteTrend}% du taux de conformité global vs semaine précédente</li>` : ''}
        ${manquantsTrend < 0 ? `<li>Réduction de ${Math.abs(manquantsTrend)} manquants par rapport à la semaine dernière</li>` : ''}
      </ul>
    </div>
  </div>
  `;
  
  // Footer
  html += `
        <div class="footer">
          <p><strong>ToolWing V4.0</strong> — Système d'inventaire automatique</p>
          <p style="margin-top: 10px; font-size: 12px; opacity: 0.8;">
            XWB BARQUE Operations — Airbus<br>
            Rapport généré automatiquement le ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy à HH:mm')}<br>
            Pour toute question : ${CONFIG.weeklyReportEmail}
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Fonction principale : Envoie le rapport hebdomadaire
 */
function sendWeeklyReport() {
  try {
    console.log("📊 Début génération rapport hebdomadaire...");
    console.log("=".repeat(60));
    
    // 1. Calculer dates de la semaine précédente
    const weekDates = getPreviousWeekDates();
    console.log(`📅 Période : ${weekDates.formattedPeriod} (Semaine ${weekDates.weekNumber})`);
    
    // 2. Calculer tous les KPIs
    console.log("📊 Calcul des KPIs...");
    const manquants = calculateManquantsSansDoublonsWeek(weekDates.startDate, weekDates.endDate);
    console.log(`  ✅ Manquants : ${manquants.totalManquants}`);
    
    const conformites = calculateNonConformitesWeek(weekDates.startDate, weekDates.endDate);
    console.log(`  ✅ Conformité globale : ${conformites.tauxConformiteGlobal}%`);
    
    const signalements = calculateSignalementsWeek(weekDates.startDate, weekDates.endDate);
    console.log(`  ✅ Signalements : ${signalements.total}`);
    
    const controlesEffectues = calculateControlesEffectues(weekDates.startDate, weekDates.endDate);
    console.log(`  ✅ Contrôles : ${controlesEffectues}`);
    
    // 3. Compiler les données de la semaine
    const weekData = {
      annee: weekDates.year,
      numeroSemaine: weekDates.weekNumber,
      dateDebut: weekDates.startDate,
      dateFin: weekDates.endDate,
      formattedPeriod: weekDates.formattedPeriod,
      conformiteGlobale: conformites.tauxConformiteGlobal,
      manquantsTotal: manquants.totalManquants,
      mallettesAvecManquants: manquants.mallettesAvecManquants,
      signalementsTotal: signalements.total,
      signalements: signalements,
      mallettesARisque: conformites.mallettesDetail.filter(m => m.conformite < 80).length,
      mallettesDetail: conformites.mallettesDetail,
      controlesEffectues: controlesEffectues,
      donneesParMallette: compileDonneesJSON(conformites.mallettesDetail),
      signalementsList: signalements.liste
    };
    
    // 4. Récupérer données semaine précédente
    console.log("📊 Récupération historique...");
    const lastWeekData = getLastWeekData();
    
    // 5. Générer HTML
    console.log("📧 Génération HTML...");
    const htmlBody = generateWeeklyReportHTML(weekData, lastWeekData);
    
    // 6. Envoyer email
    console.log("📧 Envoi email...");
    MailApp.sendEmail({
      to: CONFIG.weeklyReportEmail,
      subject: `📊 Rapport Hebdomadaire ToolWing - Semaine ${weekData.numeroSemaine} (${weekData.formattedPeriod})`,
      htmlBody: htmlBody
    });
    console.log(`✅ Email envoyé à ${CONFIG.weeklyReportEmail}`);
    
    // 7. Sauvegarder dans l'historique
    console.log("💾 Sauvegarde historique...");
    saveWeeklyHistorique(weekData);
    console.log("✅ Historique sauvegardé");
    
    console.log("=".repeat(60));
    console.log("✅ RAPPORT HEBDOMADAIRE TERMINÉ AVEC SUCCÈS !");
    return true;
    
  } catch (error) {
    console.error("❌ ERREUR RAPPORT HEBDOMADAIRE:", error);
    console.error("Stack:", error.stack);
    
    // Envoyer email d'erreur
    try {
      MailApp.sendEmail({
        to: CONFIG.weeklyReportEmail,
        subject: "❌ Erreur - Rapport Hebdomadaire ToolWing",
        body: `Une erreur est survenue lors de la génération du rapport hebdomadaire:\n\n${error.message}\n\nStack:\n${error.stack}`
      });
    } catch (emailError) {
      console.error("Impossible d'envoyer email d'erreur:", emailError);
    }
    
    return false;
  }
}

/**
 * Configure le trigger hebdomadaire (Lundi 5h00)
 */
function setupWeeklyTrigger() {
  try {
    console.log("⏰ Configuration du trigger hebdomadaire...");
    console.log("=".repeat(60));
    
    // 1. Supprimer TOUS les anciens triggers (quotidien + hebdo)
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      if (funcName === 'sendDailyReport' || funcName === 'sendWeeklyReport') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        console.log(`🗑️ Trigger supprimé: ${funcName}`);
      }
    });
    
    console.log(`✅ ${deletedCount} ancien(s) trigger(s) supprimé(s)`);
    
    // 2. Créer nouveau trigger hebdomadaire
    ScriptApp.newTrigger('sendWeeklyReport')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(5)
      .create();
    
    console.log("✅ Trigger hebdomadaire configuré avec succès !");
    console.log("📧 Le rapport sera envoyé tous les lundis à 5h00");
    console.log(`📬 Destinataire : ${CONFIG.weeklyReportEmail}`);
    
    // 3. Afficher tous les triggers actifs
    const allTriggers = ScriptApp.getProjectTriggers();
    console.log("\n📋 Triggers actifs :");
    allTriggers.forEach((trigger, index) => {
      console.log(`${index + 1}. ${trigger.getHandlerFunction()} - ${trigger.getTriggerSource()}`);
    });
    
    console.log("=".repeat(60));
    return true;
    
  } catch (error) {
    console.error("❌ Erreur configuration trigger:", error);
    return false;
  }
}

/**
 * Teste l'envoi du rapport immédiatement
 */
function testWeeklyReport() {
  console.log("🧪 TEST : Envoi du rapport hebdomadaire...");
  console.log("=".repeat(60));
  
  try {
    sendWeeklyReport();
    console.log("\n✅ Test terminé ! Vérifiez votre boîte email.");
    console.log(`📧 Email envoyé à : ${CONFIG.weeklyReportEmail}`);
  } catch (error) {
    console.error("\n❌ Erreur lors du test:", error);
    console.error("Stack:", error.stack);
  }
}

/**
 * Crée la feuille Historique_Hebdo avec structure et formatage
 */
function createHistoriqueSheet() {
  try {
    console.log("📝 Création de la feuille Historique_Hebdo...");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Supprimer si existe déjà
    const existingSheet = ss.getSheetByName(CONFIG.sheets.historique);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
      console.log("🗑️ Ancienne feuille supprimée");
    }
    
    // Créer nouvelle feuille
    const sheet = ss.insertSheet(CONFIG.sheets.historique);
    
    // En-têtes (12 colonnes)
    const headers = [
      'Année',
      'Semaine',
      'Date début',
      'Date fin',
      'Conformité %',
      'Manquants',
      'Signalements',
      'Mallettes <80%',
      'Contrôles',
      'Données Mallettes',
      'Signalements',
      'Date génération'
    ];
    
    sheet.getRange(1, 1, 1, 12).setValues([headers]);
    
    // Formatage en-têtes
    sheet.getRange(1, 1, 1, 12)
      .setFontWeight('bold')
      .setBackground('#1976D2')
      .setFontColor('white')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Figer première ligne
    sheet.setFrozenRows(1);
    
    // Largeur colonnes
    sheet.setColumnWidth(1, 80);   // Année
    sheet.setColumnWidth(2, 80);   // Semaine
    sheet.setColumnWidth(3, 100);  // Date début
    sheet.setColumnWidth(4, 100);  // Date fin
    sheet.setColumnWidth(5, 120);  // Conformité
    sheet.setColumnWidth(6, 100);  // Manquants
    sheet.setColumnWidth(7, 120);  // Signalements
    sheet.setColumnWidth(8, 120);  // Mallettes <80%
    sheet.setColumnWidth(9, 100);  // Contrôles
    sheet.setColumnWidth(10, 400); // Données Mallettes
    sheet.setColumnWidth(11, 400); // Signalements
    sheet.setColumnWidth(12, 160); // Date génération
    
    // Hauteur ligne header
    sheet.setRowHeight(1, 40);
    
    console.log("✅ Feuille Historique_Hebdo créée avec succès");
    return sheet;
    
  } catch (error) {
    console.error("❌ Erreur createHistoriqueSheet:", error);
    throw error;
  }
}

/**
 * Sauvegarde les données d'une semaine dans l'historique
 */
function saveWeeklyHistorique(weekData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let historiqueSheet = ss.getSheetByName(CONFIG.sheets.historique);
    
    // Créer si n'existe pas
    if (!historiqueSheet) {
      historiqueSheet = createHistoriqueSheet();
    }
    
    const lastRow = historiqueSheet.getLastRow();
    
    // Préparer la ligne (12 colonnes)
    const row = [
      weekData.annee,
      weekData.numeroSemaine,
      Utilities.formatDate(weekData.dateDebut, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(weekData.dateFin, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      weekData.conformiteGlobale,
      weekData.manquantsTotal,
      weekData.signalementsTotal,
      weekData.mallettesARisque,
      weekData.controlesEffectues,
      JSON.stringify(weekData.donneesParMallette),
      JSON.stringify(weekData.signalementsList),
      new Date()
    ];
    
    // Écrire la ligne
    historiqueSheet.getRange(lastRow + 1, 1, 1, 12).setValues([row]);
    
    // Formatage conditionnel colonne E (Conformité)
    const conformiteCell = historiqueSheet.getRange(lastRow + 1, 5);
    if (weekData.conformiteGlobale === 100) {
      conformiteCell.setBackground('#E8F5E9').setFontWeight('bold');
    } else if (weekData.conformiteGlobale >= 80) {
      conformiteCell.setBackground('#FFF3E0');
    } else {
      conformiteCell.setBackground('#FFEBEE').setFontWeight('bold');
    }
    
    // Centrer colonnes numériques
    historiqueSheet.getRange(lastRow + 1, 1, 1, 9).setHorizontalAlignment('center');
    
    console.log(`✅ Historique S${weekData.numeroSemaine} enregistré (ligne ${lastRow + 1})`);
    
  } catch (error) {
    console.error("❌ Erreur saveWeeklyHistorique:", error);
    throw error;
  }
}

/**
 * Récupère les données de la semaine précédente pour comparaison
 */
function getLastWeekData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const historiqueSheet = ss.getSheetByName(CONFIG.sheets.historique);
    
    // Si pas d'historique, retourner valeurs par défaut
    if (!historiqueSheet || historiqueSheet.getLastRow() < 2) {
      console.log("⚠️ Pas d'historique disponible");
      return {
        annee: 0,
        semaine: 0,
        conformiteGlobale: 0,
        manquantsTotal: 0,
        signalementsTotal: 0,
        mallettesARisque: 0,
        controlesEffectues: 0
      };
    }
    
    // Récupérer dernière ligne
    const lastRow = historiqueSheet.getLastRow();
    const data = historiqueSheet.getRange(lastRow, 1, 1, 12).getValues()[0];
    
    return {
      annee: data[0],
      semaine: data[1],
      conformiteGlobale: data[4],
      manquantsTotal: data[5],
      signalementsTotal: data[6],
      mallettesARisque: data[7],
      controlesEffectues: data[8]
    };
    
  } catch (error) {
    console.error("❌ Erreur getLastWeekData:", error);
    // Retourner valeurs par défaut en cas d'erreur
    return {
      annee: 0,
      semaine: 0,
      conformiteGlobale: 0,
      manquantsTotal: 0,
      signalementsTotal: 0,
      mallettesARisque: 0,
      controlesEffectues: 0
    };
  }
}
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
/** fonction test email */

function autoriserEnvoiEmail() {
  try {
    MailApp.sendEmail({
      to: CONFIG.notificationEmail,
      subject: "Test autorisation ToolWing",
      body: "L'application ToolWing est maintenant autorisé à envoyer des mails"
          });
          
          console.log(" Email de test envoyé avec succès à:", CONFIG.notificationEmail);
          return "Autorisation accordée !";
         } catch (error) {
          console.error( "Erreur !! :", error);
          return "Erreur :" + error.message;
         }
}
// ==========================================
// 🚀 TOOLWING V4.0 - SYSTÈME D'INVENTAIRE AUTOMATIQUE
// ==========================================
/**
 * Développé par :
 * Valentin Haultcoeur
 * Apprenti Développeur / Concepteur d'Application
 * et 
 * Noëmie Maerten 
 * Gestionnaire Projets Alten
 * Inventaire dynamique pour mallettes d'outillage - Alten pour Airbus
 * 
 * Décembre 2025
 * 
 * Système de gestion d'inventaire intelligent avec :
 * - Formulaire WebApp dynamique
 * - Dashboard temps réel
 * - Rapports hebdomadaires automatiques
 * - Historique et tendances
 * - Notifications email
 * 
 */
// ==========================================

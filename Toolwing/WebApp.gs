// ==========================================
// GESTION AVANCÉE DE LA LOGIQUE CONDITIONNELLE
// ==========================================

/**
 * Version améliorée pour gérer la logique conditionnelle complexe
 * Google Forms a des limitations, ce script contourne ces limitations
 */

// ==========================================
// SOLUTION 1: FORMULAIRE DYNAMIQUE AVEC SECTIONS
// ==========================================

/**
 * Crée un formulaire avec logique conditionnelle améliorée
 * Utilise des sections et de la validation personnalisée
 */
function createAdvancedForm() {
  try {
    const form = FormApp.openById(FORM_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // Nettoyer le formulaire
    clearForm(form);
    
    // Configuration de base
    form.setTitle('Inventaire des Mallettes - Version Avancée');
    form.setDescription('Système d\'inventaire avec logique conditionnelle intelligente');
    form.setCollectEmail(false);
    form.setProgressBar(true);
    
    // Obtenir les données
    const mallettesData = getMallettesFromSheet(sheet);
    
    // ==========================================
    // SECTION 1: IDENTIFICATION
    // ==========================================
    
    const nomPrenomItem = form.addTextItem()
      .setTitle('1. Nom et Prénom')
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireTextContainsPattern('[A-Za-zÀ-ÿ\\s]+')
          .setHelpText('Veuillez entrer un nom valide')
          .build()
      );
    
    // ==========================================
    // SECTION 2: SÉLECTION DES MALLETTES
    // ==========================================
    
    const mallettesControlees = form.addCheckboxItem()
      .setTitle('2. Quelle(s) mallette(s) avez-vous contrôlé ?')
      .setRequired(true)
      .setHelpText('Sélectionnez toutes les mallettes vérifiées');
    
    const mallettesChoices = mallettesData.map(m => 
      mallettesControlees.createChoice(m.nom)
    );
    mallettesControlees.setChoices(mallettesChoices);
    
    // Saut de page
    form.addPageBreakItem().setTitle('Analyse des manquants');
    
    // ==========================================
    // SECTION 3: SIGNALEMENT DES MANQUANTS
    // ==========================================
    
    const manquantsQuestion = form.addMultipleChoiceItem()
      .setTitle('3. Y a-t-il des manquants dans les mallettes contrôlées ?')
      .setRequired(true);
    
    // Créer les sections conditionnelles
    const sectionManquants = form.addPageBreakItem()
      .setTitle('Détails des manquants');
      
    const sectionSignalement = form.addPageBreakItem()
      .setTitle('Signalements additionnels');
    
    // Configurer la navigation conditionnelle
    manquantsQuestion.setChoices([
      manquantsQuestion.createChoice('Oui, il y a des manquants', sectionManquants),
      manquantsQuestion.createChoice('Non, tout est complet', sectionSignalement)
    ]);
    
    // ==========================================
    // SECTION 4: DÉTAILS DES MANQUANTS (Conditionnelle)
    // ==========================================
    
    // Pour chaque mallette, créer une question conditionnelle
    mallettesData.forEach((mallette, index) => {
      // Question: Cette mallette a-t-elle des manquants ?
      const malletteManquants = form.addMultipleChoiceItem()
        .setTitle(`La ${mallette.nom} a-t-elle des manquants ?`)
        .setHelpText('Répondez uniquement si vous avez contrôlé cette mallette')
        .setRequired(false);
      
      malletteManquants.setChoices([
        malletteManquants.createChoice('Oui'),
        malletteManquants.createChoice('Non'),
        malletteManquants.createChoice('Non contrôlée')
      ]);
      
      // Liste des outils manquants pour cette mallette
      if (mallette.outils.length > 0) {
        const outilsChoices = mallette.outils.map((outil, i) => 
          `${i + 1}. ${outil}`
        );
        
        const outilsManquants = form.addCheckboxItem()
          .setTitle(`Outils manquants dans ${mallette.nom}`)
          .setHelpText('Cochez les outils manquants (si applicable)')
          .setRequired(false);
        
        outilsManquants.setChoices(
          outilsChoices.map(o => outilsManquants.createChoice(o))
        );
      }
    });
    
    // Navigation vers signalement
    form.addPageBreakItem()
      .setGoToPage(sectionSignalement);
    
    // ==========================================
    // SECTION 5: SIGNALEMENTS ADDITIONNELS
    // ==========================================
    
    const autreSignalement = form.addMultipleChoiceItem()
      .setTitle('6. Avez-vous d\'autres éléments à signaler ?')
      .setHelpText('Casse, métrologie, commande, etc.')
      .setRequired(true);
    
    const sectionDetailsSignalement = form.addPageBreakItem()
      .setTitle('Détails du signalement');
    
    autreSignalement.setChoices([
      autreSignalement.createChoice('Oui', sectionDetailsSignalement),
      autreSignalement.createChoice('Non', FormApp.PageNavigationType.SUBMIT)
    ]);
    
    // Détails du signalement
    form.addParagraphTextItem()
      .setTitle('7. Décrivez votre signalement')
      .setHelpText('Soyez précis sur les actions requises')
      .setRequired(false);
    
    // Type de signalement
    form.addCheckboxItem()
      .setTitle('Type de signalement')
      .setChoices([
        'Outil cassé',
        'Départ en métrologie',
        'Demande de commande',
        'Réorganisation mallette',
        'Autre'
      ].map(type => FormApp.createChoice(type)));
    
    // Urgence
    form.addMultipleChoiceItem()
      .setTitle('Niveau d\'urgence')
      .setChoices([
        '🔴 Urgent (bloquant)',
        '🟠 Important (sous 1 semaine)',
        '🟡 Normal (sous 1 mois)',
        '🟢 Faible (information)'
      ].map(urgence => FormApp.createChoice(urgence)));
    
    console.log("✅ Formulaire avancé créé avec succès");
    
    return {
      success: true,
      formUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl()
    };
    
  } catch (error) {
    console.error("❌ Erreur:", error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==========================================
// SOLUTION 2: WEBAPP AVEC FORMULAIRE DYNAMIQUE
// ==========================================

/**
 * Créer une WebApp avec un formulaire HTML dynamique
 * Plus de flexibilité que Google Forms natif
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  
  // Passer les données au template
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  template.mallettesData = JSON.stringify(getMallettesFromSheet(sheet));
  
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
      // Ne pas faire échouer la soumission si le dashboard échoue
    }
    
    // Envoyer notification si nécessaire
    if (formData.hasManquants === 'oui' || (formData.urgence && formData.urgence.includes('🔴'))) {
      console.log("📧 Envoi de notification...");
      try {
        sendNotificationEmail(formData);
        console.log("✅ Notification envoyée");
      } catch (emailError) {
        console.error("⚠️ Erreur lors de l'envoi de l'email:", emailError);
        // Ne pas faire échouer la soumission si l'email échoue
      }
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
// SOLUTION 3: VALIDATION CÔTÉ SERVEUR
// ==========================================

/**
 * Valide les réponses pour s'assurer de la cohérence
 * Appelée après soumission du formulaire
 */
function validateFormResponse(e) {
  const responses = e.response.getItemResponses();
  const validation = {
    isValid: true,
    errors: [],
    warnings: []
  };
  
  // Extraire les réponses
  let mallettesControlees = [];
  let mallettesAvecManquants = [];
  let hasManquants = false;
  
  responses.forEach(response => {
    const title = response.getItem().getTitle();
    const answer = response.getResponse();
    
    if (title.includes('mallette(s) avez-vous contrôlé')) {
      mallettesControlees = Array.isArray(answer) ? answer : [answer];
    } else if (title.includes('Y a-t-il des manquants')) {
      hasManquants = (answer === 'Oui');
    } else if (title.includes('Dans quelle(s) mallette(s)') && answer) {
      mallettesAvecManquants = Array.isArray(answer) ? answer : [answer];
    }
  });
  
  // Validations
  
  // 1. Les mallettes avec manquants doivent être dans les mallettes contrôlées
  mallettesAvecManquants.forEach(mallette => {
    if (!mallettesControlees.includes(mallette)) {
      validation.isValid = false;
      validation.errors.push(
        `Erreur: "${mallette}" signalée avec manquants mais non marquée comme contrôlée`
      );
    }
  });
  
  // 2. Si manquants = Oui, il doit y avoir au moins une mallette avec manquants
  if (hasManquants && mallettesAvecManquants.length === 0) {
    validation.warnings.push(
      'Attention: Manquants signalés mais aucune mallette spécifique indiquée'
    );
  }
  
  // 3. Si manquants = Non, il ne doit pas y avoir de mallettes avec manquants
  if (!hasManquants && mallettesAvecManquants.length > 0) {
    validation.isValid = false;
    validation.errors.push(
      'Erreur: Pas de manquants signalés mais des mallettes avec manquants sont sélectionnées'
    );
  }
  
  // Traiter les erreurs
  if (!validation.isValid) {
    // Enregistrer l'erreur dans une feuille de logs
    logValidationError(e, validation);
    
    // Envoyer une notification
    sendValidationAlert(validation);
  }
  
  return validation;
}

/**
 * Enregistre les erreurs de validation
 */
function logValidationError(e, validation) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let errorSheet;
    try {
      errorSheet = spreadsheet.getSheetByName('Erreurs_Validation');
    } catch (err) {
      errorSheet = spreadsheet.insertSheet('Erreurs_Validation');
      errorSheet.getRange(1, 1, 1, 5).setValues([[
        'Date/Heure',
        'Email',
        'Erreurs',
        'Avertissements',
        'Données brutes'
      ]]);
      errorSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#ffebee');
    }
    
    const lastRow = errorSheet.getLastRow();
    errorSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      new Date(),
      e.response.getRespondentEmail() || 'Anonyme',
      validation.errors.join('\n'),
      validation.warnings.join('\n'),
      JSON.stringify(e.response.getItemResponses().map(r => ({
        question: r.getItem().getTitle(),
        response: r.getResponse()
      })))
    ]]);
    
    console.log("❌ Erreur de validation enregistrée");
    
  } catch (error) {
    console.error("Erreur lors de l'enregistrement:", error);
  }
}

/**
 * Envoie une alerte de validation
 */
function sendValidationAlert(validation) {
  const recipient = Session.getActiveUser().getEmail();
  const subject = '[INVENTAIRE] ⚠️ Erreur de validation détectée';
  
  let body = 'Des incohérences ont été détectées dans une soumission d\'inventaire:\n\n';
  
  if (validation.errors.length > 0) {
    body += '❌ ERREURS:\n';
    validation.errors.forEach(error => {
      body += `  - ${error}\n`;
    });
  }
  
  if (validation.warnings.length > 0) {
    body += '\n⚠️ AVERTISSEMENTS:\n';
    validation.warnings.forEach(warning => {
      body += `  - ${warning}\n`;
    });
  }
  
  body += '\n\nVeuillez vérifier la feuille "Erreurs_Validation" pour plus de détails.';
  
  try {
    MailApp.sendEmail(recipient, subject, body);
  } catch (error) {
    console.error("Erreur envoi alerte:", error);
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

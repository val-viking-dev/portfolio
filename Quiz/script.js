// ========== RÉFÉRENCES DOM ==========
// Sections
const accueil = document.getElementById("accueil");
const questionnaire = document.getElementById("questionnaire");
const ecranFin = document.getElementById("ecranFin");

// Boutons
const btnQuizz = document.getElementById("btnQuizz");
const btnValider = document.getElementById("btnValider");
const btnSuivant = document.getElementById("btnSuivant");
const btnRecommencer = document.getElementById("btnRecommencer");
const btnAccueil = document.getElementById("btnAccueil");

// Éléments du questionnaire
const chrono = document.getElementById("chrono");
const scoreElement = document.getElementById("score");
const questionElement = document.getElementById("question");
const choixQuestion = document.getElementById("choixQuestion");
const textesReponses = document.querySelectorAll(".texteReponse");
const inputsRadio = document.querySelectorAll('input[type="radio"]');

// Éléments de l'écran de fin
const scoreFinal = document.getElementById("scoreFinal");
const tempsTotal = document.getElementById("tempsTotal");
const messageFinal = document.getElementById("messageFinal");
const listeResume = document.getElementById("listeResume");

// Historique
const listeScores = document.getElementById("listeScores");

//PopupFeedback
const feedbackPopup = document.getElementById("feedbackPopup");
const messageFeedback = document.getElementById("messageFeedback");

// ======= VARIABLES DE DONNÉES =======
let score = 0
let indexQuestion = 0
let tableQuestions = []
let tableSelection = []
let tableReponse = []
let tempsDepart = 0
let intervalChrono = 0

// Chargement des questions
async function chargerQuestions() {
    try {
    const reponse = await fetch("questions.json");
    tableQuestions = await reponse.json();
    console.log("Questions chargées :", tableQuestions);
        } catch (erreur)
        {
            console.error("Erreur:", erreur);
        }  
}

chargerQuestions();
afficherHistorique();

// ======= Bouton Commencer le Quizz =======

btnQuizz.addEventListener('click', demarrerQuizz)

// Mélange des questions
function melangerTableau(tableau) {
    let copieTableau = Array.from(tableau);
        for (let i = copieTableau.length - 1; i>= 0; i--) {
            let indexAleatoire = Math.floor(Math.random() * (i + 1)) // Donne un entier entre 0 et i
            let temp = copieTableau[i];           // Sauvegarder l'élément i
            copieTableau[i] = copieTableau[indexAleatoire];  // Mettre indexAleatoire à la place de i
            copieTableau[indexAleatoire] = temp;  // Mettre i à la place de indexAleatoire
        }
    return copieTableau;

}

//Démarrage du quizz
function demarrerQuizz() {
    tableSelection = melangerTableau(tableQuestions).slice(0, 10); // Mélange et prend les dix premières
    indexQuestion = 0; // Remise à Zéro des compteurs index et score
    score = 0;
    accueil.style.display = "none";
    questionnaire.style.display = "block";
    tempsDepart = Date.now(); // Donne le nombre de millisecondes depuis 1970. On va calculer le temps paser entre ce point et la fin. 
    afficherQuestion();
    demarrerChrono();

}

//Affichage de la question
function afficherQuestion() {
    let questionActuelle = tableSelection[indexQuestion]; // Récupère l'objet question depuis tableSelection en utilisant indexQuestion comme index
    questionElement.textContent = questionActuelle.question; // Mettre le texte de la question dans l'élément HTML
        for (let i = 0; i < 4; i++) {
            textesReponses[i].textContent = questionActuelle.choix[i];
        }
    scoreElement.textContent = "Score : " + score + "/10"
}

// démarrer le chrono

function demarrerChrono() {
   intervalChrono = setInterval(function() { // intervalChrono = id de l'interval pour la réinitialisation
        let tempsEcoule = Date.now() - tempsDepart; // heure actuelle - heure de répart = temps écoulé depuis le début
        let secondes = Math.floor(tempsEcoule / 1000); // conversion millisecondes en secondes
        let minutes = Math.floor(secondes / 60); // conversion de 60 secondes en 1 minutes
        let secondesRestantes = secondes % 60; // Le nombre de secondes restantes après la conversion du dessus
        chrono.textContent = `Temps : ${minutes}:${String(secondesRestantes).padStart(2, '0')}`; // Affiche le chrono
    }, 1000);


}

// Page de fin

function afficherEcranFin() {
    scoreFinal.textContent = `Tu as réussi ${score} questions sur 10 !`; // Texte du score final
    let tempsEcoule = Date.now() - tempsDepart;
    let secondes = Math.floor(tempsEcoule / 1000);
    let minutes = Math.floor(secondes / 60);
    let secondesRestantes = secondes % 60;
    tempsTotal.textContent = `Temps total: ${minutes}:${String(secondesRestantes).padStart(2, '0')}`; // calcul du temps total
    
    // conditionnement des messages en fonction du score
    if (score >= 0 && score <=3) {
        messageFinal.textContent = "C'est un début ! La culture générale se travaille, continue comme ça 📚"
    } else if (score >= 4 && score <= 6) {
        messageFinal.textContent = "Score correct ! Encore un petit effort et tu seras au top 🎯"
    } else if (score >= 7 && score <=9) {
        messageFinal.textContent = "Impressionnant ! Tu as de bonnes connaissances 👏"
    } else {
        messageFinal.textContent = "PARFAIT ! 🏆 Score sans faute, tu es un champion !"
    }

    // boucle pour afficher l'historique des réponses de ce quizz
    tableReponse.forEach(objetReponse => {
        let question = objetReponse.question;
        let reponse = objetReponse.reponseUtilisateur;
        let estCorrecte = objetReponse.estCorrecte;
        let icone;
        if (estCorrecte) {
            icone = "✅";
        } else {
            icone = "❌";
        }
    
    let texte = `Question : ${question} | Ta réponse : ${reponse} | ${icone}`;
    let li = document.createElement("li");
    li.textContent = texte;
    listeResume.appendChild(li);
    

    })

    //Création de l'historique
    let nouveauScore = {                 //Création d'un objet
        score: `${score} /10`,
        temps: `${minutes}:${String(secondesRestantes).padStart(2, '0')}`
    };
    let historique = localStorage.getItem("historiqueScores"); // Récupère l'historique existant
    
    if (historique === null) {
        historique = []                 // Si c'est la première fois, crée un tableau vide
    } else {
        historique = JSON.parse(historique);
    }

    historique.push(nouveauScore);
    localStorage.setItem("historiqueScores", JSON.stringify(historique));
    afficherHistorique();
}

//Bouton valider
btnValider.addEventListener('click', validerReponse)

function validerReponse() {
    const reponseSelectionnee = Array.from(inputsRadio).find(input => input.checked); // Cherche le inputRadio coché
        if (!reponseSelectionnee) {
            alert("Veuiller sélectionner une réponse !");
            return;                                        // Si pas de input coché retourne en arrière avec une erreur
        }
    let indexReponse = reponseSelectionnee.value; // Récupère la valeur de l'index coché
    let questionActuelle = tableSelection[indexQuestion];

    const estBonneReponse = questionActuelle.choix[indexReponse] === questionActuelle.reponse; // Trouve la bonne réponse
        if (estBonneReponse) {
            score = score + 1; 
        }


    const objetReponse = {                                       // Création de l'objet réponse 
        "question": questionActuelle.question,
        "reponseUtilisateur": questionActuelle.choix[indexReponse],
        "estCorrecte": estBonneReponse
    };
    tableReponse.push(objetReponse);                // Insertion dans le tableau

    if (estBonneReponse) {                                    // Conditionnement du message du popup correct/incorrect
        messageFeedback.textContent = "Bonne réponse !✅ ";        // Ajoute le message
        feedbackPopup.classList.add("correct");                    // Ajoute la classe .correct ou . incorrect à feedbackPopup
        reponseSelectionnee.parentElement.style.backgroundColor = "#4caf50"; // Change la couleur du label choisi
    } else {
        messageFeedback.textContent = "Mauvaise réponse ! ❌ ";
        feedbackPopup.classList.add("incorrect");
        reponseSelectionnee.parentElement.style.backgroundColor = "#f44336";
    }
    
    feedbackPopup.style.display = "block" // Affiche le pop-up

// Attente de 1.5secondes et création de la fonction
setTimeout(function() {                                        
            feedbackPopup.style.display = "none";              // Retour à l'état none du popup
            feedbackPopup.classList.remove("correct", "incorrect");         // On enlève les classes
            reponseSelectionnee.parentElement.style.backgroundColor = "#f8f9fa"; // On remet le background d'origine
               inputsRadio.forEach(input => {                 // On décoche les ipuntRadio
                input.checked = false;
               })
            
               indexQuestion++;         // On rajoute +1 à indexQuestion
               if (indexQuestion === 10) {
                questionnaire.style.display = "none"
                ecranFin.style.display = "block";
                afficherEcranFin();

               } else {
                afficherQuestion();
               }

}, 1500);
    
}

// Fonction bouton accueil
function retourAccueil() {
    ecranFin.style.display = "none";
    accueil.style.display = "block";
    listeResume.innerHTML = "";
    clearInterval(intervalChrono); // Réinitialise l'interval
    tableReponse = [];
}
btnAccueil.addEventListener('click', retourAccueil);

// Fonction bouton recommencer

function recommencer () {
    ecranFin.style.display = "none";
    questionnaire.style.display = "block";
    listeResume.innerHTML = "";
    clearInterval(intervalChrono);
    demarrerQuizz();
    tableReponse = [];

}
btnRecommencer.addEventListener('click', recommencer);

// Affichage de l'historique

function afficherHistorique () {
    let historique = localStorage.getItem("historiqueScores");
    listeScores.innerHTML = "";
    if (historique === null) {
        historique = []                 // Si c'est la première fois, crée un tableau vide
    } else {
        historique = JSON.parse(historique);    // Si il existe, le convertit en tableau
    }
    for (let i = 0; i < historique.length; i++) {
        let scoreActuel = historique[i];   // scoreActuel contient un objet de historique
        let texte = `Score : ${scoreActuel.score} - Temps: ${scoreActuel.temps}`;
        let li = document.createElement("li");     // Crée l'élément HTML li
        li.textContent = texte;                    // Ajoute le texte à l'élément HTML
        listeScores.appendChild(li)   // Rajoute l'élément HTML dans la liste 
    }
    
}
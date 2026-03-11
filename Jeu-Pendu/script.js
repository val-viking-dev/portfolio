// ========== RÉFÉRENCES DOM ==========
// Sections
const accueil = document.getElementById("section-accueil");
const jeu = document.getElementById("section-jeu");
const resultat = document.getElementById("section-resultat");

// Boutons
const btnCommencer = document.getElementById("btnCommencer");
const btnRejouer = document.getElementById("btnRejouer");
const btnAccueil = document.getElementById("btnAccueil");

// Select
const diffSelect = document.getElementById("diffSelect");

// Contenu sections
const conteneurDessin = document.getElementById("conteneurDessin");
const conteneurMot = document.getElementById("conteneurMot");
const conteneurLettres = document.getElementById("conteneurLettres");
const conteneurErreurs = document.getElementById("affichageErreurs");
const messageFin = document.getElementById("messageFin");
const affichageMot = document.getElementById("affichageMot");

// Parties du pendu (SVG)
const partiesPendu = [
    document.getElementById("tete"),
    document.getElementById("corps"),
    document.getElementById("brasG"),
    document.getElementById("brasD"),
    document.getElementById("jambeG"),
    document.getElementById("jambeD")
];

// ========== DONNÉES DU JEU ==========
const mots = {
    facile: ["CHAT", "CHIEN", "CANARD", "POULE", "BATEAU", "MAISON", "LIVRE", "SOLEIL", "FLEUR", "ARBRE"],
    moyen: ["DICTIONNAIRE", "ORDINATEUR", "TOURNEVIS", "FAUTEUIL", "ARAIGNEE", "BIBLIOTHEQUE", "CHOCOLAT", "AVENTURE", "MYSTERE", "ELEPHANT"],
    difficile: ["PORTE-MONNAIE", "GARDE-MANGER", "GRATTE-CIEL", "CHAUVE-SOURIS", "ARC-EN-CIEL", "CHOU-FLEUR", "COFFRE-FORT", "ROUGE-GORGE"]
};

// ========== VARIABLES D'ÉTAT ==========
let motSecret = "";
let lettresProposees = [];
let nombreErreurs = 0;

// ========== FONCTION DÉMARRAGE ==========
function demarrage() {
    // Réinitialiser les variables
    const difficulte = diffSelect.value;
    const listeMots = mots[difficulte];
    const indexAleatoire = Math.floor(Math.random() * listeMots.length);
    motSecret = listeMots[indexAleatoire];
    lettresProposees = [];
    nombreErreurs = 0;
    
    // Vider les conteneurs
    conteneurLettres.innerHTML = "";
    conteneurErreurs.textContent = "Erreurs : 0 / 6";
    
    // Réinitialiser le dessin du pendu
    reinitialiserPendu();
    
    // Changer de section
    accueil.classList.add("hidden");
    resultat.classList.add("hidden");
    jeu.classList.remove("hidden");
    
    // Afficher le jeu
    afficherMot();
    genererLettres();
    
    console.log("Mot secret :", motSecret); // Pour débugger
}

// ========== AFFICHAGE DU MOT ==========
function afficherMot() {
    let affichage = "";
    for (let caractere of motSecret) {
        if (caractere === "-") {
            affichage += "- ";
        } else if (lettresProposees.includes(caractere)) {
            affichage += caractere + " ";
        } else {
            affichage += "_ ";
        }
    }
    conteneurMot.textContent = affichage.trim();
}

// ========== GÉNÉRATION DES LETTRES ==========
function genererLettres() {
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    for (let lettre of alphabet) {
        const bouton = document.createElement("button");
        bouton.textContent = lettre;
        conteneurLettres.appendChild(bouton);
        
        bouton.addEventListener("click", () => {
            // Vérifier si la lettre est dans le mot
            if (motSecret.includes(lettre)) {
                lettresProposees.push(lettre);
                afficherMot();
                bouton.classList.add("correct");
            } else {
                nombreErreurs++;
                conteneurErreurs.textContent = `Erreurs : ${nombreErreurs} / 6`;
                bouton.classList.add("incorrect");
                dessinerPendu();
            }
            
            // Désactiver le bouton
            bouton.disabled = true;
            
            // Vérifier fin de partie
            verifierFinPartie();
        });
    }
}

// ========== DESSIN DU PENDU ==========
function dessinerPendu() {
    if (nombreErreurs > 0 && nombreErreurs <= 6) {
        partiesPendu[nombreErreurs - 1].classList.add("visible");
    }
}

function reinitialiserPendu() {
    for (let partie of partiesPendu) {
        partie.classList.remove("visible");
    }
}

// ========== VÉRIFICATION FIN DE PARTIE ==========
function verifierFinPartie() {
    // Défaite
    if (nombreErreurs >= 6) {
        setTimeout(() => {
            messageFin.textContent = "💀 Perdu !";
            messageFin.className = "defaite";
            affichageMot.textContent = `Le mot était : ${motSecret}`;
            jeu.classList.add("hidden");
            resultat.classList.remove("hidden");
        }, 500);
    }
    // Victoire
    else if (!conteneurMot.textContent.includes("_")) {
        setTimeout(() => {
            messageFin.textContent = "🎉 Gagné !";
            messageFin.className = "victoire";
            affichageMot.textContent = `Bravo ! Tu as trouvé : ${motSecret}`;
            jeu.classList.add("hidden");
            resultat.classList.remove("hidden");
        }, 500);
    }
}

// ========== EVENT LISTENERS ==========
btnCommencer.addEventListener("click", demarrage);

btnRejouer.addEventListener("click", () => {
    resultat.classList.add("hidden");
    demarrage();
});

btnAccueil.addEventListener("click", () => {
    resultat.classList.add("hidden");
    accueil.classList.remove("hidden");
});

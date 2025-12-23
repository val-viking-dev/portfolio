let affichage = "0";
let premierNombre = null;
let operation = null;
let historique = [];
const affichageElement = document.querySelector('.affichage');
const historiqueElement = document.querySelector('.historique');
const boutons = document.querySelectorAll('button');
function mettreAJourAffichage() {
    affichageElement.textContent = affichage;
}
boutons.forEach(function(bouton) {
    bouton.addEventListener('click', function() {
       const valeur = bouton.textContent;
  
if (valeur === 'C') {
    affichage = "0";
    premierNombre = null;
    operation = null;
    mettreAJourAffichage();
}
    else if (valeur === '=') {
        let result;
       if (operation === '+') {
            result = premierNombre + Number(affichage);           
       } else if (operation === '-') {
            result = premierNombre - Number(affichage);
       } else if (operation ==='/') {
            result = premierNombre / Number(affichage);
       } else if (operation === '*') {
            result = premierNombre * Number(affichage);
       } else if (operation === '%') {
            result = premierNombre % Number(affichage);
       } else {
        result = "0"
       }
         let texteHistorique = premierNombre + " " + operation + " " + affichage + " = " + result;
         historique.push(texteHistorique);
         if (historique.length > 4) {
            historique.shift();
         }
         let elementHistorique = historique.join("\n")
         historiqueElement.textContent = elementHistorique;
         affichage =String(result);        
         mettreAJourAffichage();
         premierNombre = null;
         operation = null;
            }
             else if (bouton.classList.contains('operation')) {
               premierNombre = Number(affichage);
               operation = valeur;
               affichage = "0";
               mettreAJourAffichage()
            }
            else {
                if (affichage ==="0") {
                    affichage = valeur;
                } else {
                    affichage = affichage + valeur;
                }
                mettreAJourAffichage();
            }
      });
});


const addBtn=document.getElementById("addBtn");
const taskList=document.getElementById("taskList");
const taskInput=document.getElementById("taskInput");
const showFormBtn=document.getElementById("showFormBtn");
const formSection=document.getElementById("formSection");
let taches = [];

// crée les tâches sauvegardées
function creerTache(uneTache) {
    const texte=uneTache.texte;
    const check = document.createElement("input");
    check.type = "checkbox";
    const spanText = document.createElement("span");
    spanText.textContent = texte;
    const btnSupprimer = document.createElement("button")
    btnSupprimer.textContent="❌";
    const nouveauLi = document.createElement("li");
    check.addEventListener('change', function() {
        if (check.checked) {
            spanText.style.textDecoration = "line-through";
            uneTache.completed = true;
            localStorage.setItem("MesToDos",JSON.stringify(taches));
        } else {
            spanText.style.textDecoration = "none";
            uneTache.completed = false;
            localStorage.setItem("MesToDos",JSON.stringify(taches));
        }
    })
    nouveauLi.appendChild(check);
    nouveauLi.appendChild(spanText);
    nouveauLi.appendChild(btnSupprimer);   
    taskList.prepend(nouveauLi);
     btnSupprimer.addEventListener('click',function() {
        nouveauLi.remove();
        const index = taches.indexOf(uneTache);
        taches.splice(index, 1);
        localStorage.setItem("MesToDos", JSON.stringify(taches));
    });
    if (uneTache.completed) {
        check.checked = true;
        spanText.style.textDecoration = "line-through";
    }
    }

// Charger les tâches sauvegardées
const tachesSauvegardees = localStorage.getItem("MesToDos");

if (tachesSauvegardees) {
    
    taches = JSON.parse(tachesSauvegardees);
    taches.forEach(function(uneTache) {
        creerTache(uneTache);
    }
        
    )};
    

showFormBtn.addEventListener('click', function() {
    formSection.style.display="block";
});

// bouton nouvelle tâche
addBtn.addEventListener('click', function() {
    const texte = taskInput.value;
    
    //Créer l'objet
    const tache = {
        texte: texte,
        completed: false
    };
    
    //Ajouter au tableau
    taches.push(tache);
    
    //Sauvegarder
    localStorage.setItem("MesToDos", JSON.stringify(taches));
    
    //Créer l'affichage en appelant la fonction
    creerTache(tache);
    
    //Réinitialiser le formulaire
    taskInput.value = "";
    formSection.style.display = "none";
});

const filterActive=document.getElementById("filterActive");
const filterAll=document.getElementById("filterAll");
const filterCompleted=document.getElementById("filterCompleted");

// bouton filtre "Active"
filterActive.addEventListener('click', function () {
    const tousLesLi = taskList.children;
    //parcourt les taches
    for (let i = 0; i < tousLesLi.length; i++) {
        const li = tousLesLi[i];
        const checkbox = li.querySelector("input[type='checkbox']");
        //affiche ou non les taches actives
        if(checkbox.checked) {
            li.style.display = "none";
        }else {
            li.style.display = "block"
        }
    }
});

// bouton filtre "Toutes"
filterAll.addEventListener('click', function() {
    const tousLesLi = taskList.children;
    //affiche toute les taches
    for (let i = 0; i < tousLesLi.length; i++) {
        const li = tousLesLi[i];
        li.style.display = "block";
    }
});

//bouton filtre "Complétées"
filterCompleted.addEventListener('click', function () {
    const tousLesLi = taskList.children;
    for (let i = 0; i < tousLesLi.length; i++) {
        const li = tousLesLi[i];
        const checkbox = li.querySelector("input[type='checkbox']");
        
        if(checkbox.checked) {
            li.style.display = "block";
        }else {
            li.style.display = "none"
        }
    }
});

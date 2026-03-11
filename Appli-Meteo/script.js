// ========== CONFIGURATION ==========
const API_KEY = 'db53bd06b5e66cd035d8f80a78108f10';  // Clé openWeather

// ========== RÉFÉRENCES DOM ==========
const cityInput = document.getElementById('cityInput');
const btnSearch = document.getElementById('btnSearch');
const loadingDiv = document.getElementById('loading');
const errorDiv = document.getElementById('error');
const resultsDiv = document.getElementById('results');
const errorMessage = document.getElementById('errorMessage');
const cityNameEl = document.getElementById('cityName');
const temperatureEl = document.getElementById('temperature');
const descriptionEl = document.getElementById('description');
const humidityEl = document.getElementById('humidity');
const windEl = document.getElementById('wind');
const feelsLikeEl = document.getElementById('feelsLike');
const weatherIconEl = document.getElementById('weatherIcon');
const btnGeo = document.getElementById('btnGeo');
const autocomplete = document.getElementById('autocomplete');

// Recherche des données météo avec la récupération de l'input 
async function searchWeather() {
    const city = cityInput.value.trim();
        if (city === '') {             // Compare à une chaîne vide, peut aussi être "falsy value" (!city.trim())
            errorMessage.textContent = "Veuillez entrer une ville";
            errorDiv.classList.remove('hidden');
            return;
        }
        errorDiv.classList.add('hidden');     // cache la section error
        resultsDiv.classList.add('hidden');   // cache la section results
        loadingDiv.classList.remove('hidden');  // affiche la section loading
        const url = `https://api.openweathermap.org/data/2.5/weather?q=${city}&appid=${API_KEY}&units=metric&lang=fr`;  // mise en forme de l'URL
        try {
            const reponse = await fetch(url); // attend la réponse de l'url
            const data = await reponse.json(); // attend que les données soient converties en objet JavaScript
            console.log(data);
                if (data.cod === 200) {                    //vérifie le cod renvoyé par l'API. 200 = ok;  404 = error
                    displayWeather(data);
                }else {
                    errorWeather(data);
                }
        } catch (error) {
            console.log(error);
        }
        

    
}
// Bouton recherche
btnSearch.addEventListener('click', searchWeather);

// Ecoute de la touche entrée du clavier
cityInput.addEventListener('keydown', function(event) {   
    if (event.key === "Enter")
        searchWeather();
});

// Géolocalisation 
function getLocationWeather() {
    errorDiv.classList.add('hidden');     
    resultsDiv.classList.add('hidden');   
    loadingDiv.classList.remove('hidden'); 
        navigator.geolocation.getCurrentPosition(      // module de géolocalisation du navigateur     
           async function(position) {  // successcallback argument 1 aussi appelée "si ça marche"
                console.log(position.coords.latitude);
                console.log(position.coords.longitude);
                const latitude = position.coords.latitude;
                const longitude = position.coords.longitude;
                const url = `https://api.openweathermap.org/data/2.5/weather?lat=${latitude}&lon=${longitude}&appid=${API_KEY}&units=metric&lang=fr`;
                try {
                    const reponse = await fetch(url); // attend la réponse de l'url
                    const data = await reponse.json(); // attend que les données soient converties en objet JavaScript
                console.log(data);
                if (data.cod === 200) {                    //vérifie le cod renvoyé par l'API. 200 = ok;  404 = error
                    displayWeather(data);
                }else {
                    errorWeather(data);
                }
        } catch (error) {
            console.log(error);
        }
            },
            function (error) {  // error callback aussi appelé "si ça échoue"
                console.log("Erreur géolocalisation", error);
            }
        );
}

// Bouton de géolocalisation
btnGeo.addEventListener('click', getLocationWeather);

// Affichage si ville/ géolocalisation valide
function displayWeather (data) {
    cityNameEl.textContent = data.name;   // va chercher le nom de la ville
                    const tempNombre = data.main.temp;   // va chercher la température 
                    const tempArrondie = Math.round(tempNombre);  // arrondie le chiffre
                    temperatureEl.textContent =`${tempArrondie}°C`; 
                    const desc = data.weather[0].description;  // récupère la description
                    const descResult = `${desc[0].toUpperCase()}${desc.slice(1)}`; // formate la réponse en mettant une majuscule au début
                    descriptionEl.textContent = descResult;
                    const humNombre = data.main.humidity; // récupère l'humidité
                    humidityEl.textContent = `Humidité: ${humNombre}%`;
                    const windNombre = data.wind.speed;  // Récupère la vitesse du vent
                    const windSpeed = Math.round(windNombre); // arrondie la vitesse du vent
                    windEl.textContent = `${windSpeed} km/h`;
                    const feelsNombre = data.main.feels_like;  // récupère le ressenti
                    const feelsLike = Math.round(feelsNombre);
                    feelsLikeEl.textContent = `Ressenti: ${feelsLike}°C`;
                    const iconCode = data.weather[0].icon; // récupère l'icone 
                    weatherIconEl.src = `https://openweathermap.org/img/wn/${iconCode}@2x.png`; // place le iconCode dans l'URL du SRC du bouton
                    loadingDiv.classList.add('hidden');    // Ajoute la classe "hidden"
                    resultsDiv.classList.remove('hidden'); // Enlève la classe "hidden"
                    
                    
}

// Affichage si erreur
function errorWeather(data) {
     loadingDiv.classList.add('hidden');
     errorDiv.classList.remove('hidden');
    errorMessage.textContent = data.message;
}

// Autocomplete
cityInput.addEventListener('input', searchCities);

async function searchCities() {
    const search = cityInput.value.trim();
    if (search.length < 3 ) {
        autocomplete.innerHTML = "";
        return;
    } else {
        const url = `https://api.openweathermap.org/geo/1.0/direct?q=${search}&limit=5&appid=${API_KEY}`
        try {
                    const reponse = await fetch(url); // attend la réponse de l'url
                    const data = await reponse.json(); // attend que les données soient converties en objet JavaScript
                console.log(data);
                if (data.length > 0) {                    //vérifie le cod renvoyé par l'API. 200 = ok;  404 = error
                    autocomplete.innerHTML = "";
                    data.forEach(function(city) {
                        const li = document.createElement('li');
                        const country = city.country;
                        const cityName = city.name;
                        li.textContent = `${cityName}, ${country}`;
                        autocomplete.appendChild(li);                       
                    });
                }else {
                    autocomplete.innerHTML = "";
                }
        } catch (error) {
            console.log(error);
        }
    }
}
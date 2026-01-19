// ═══════════════════════════════════════════════════════════════════════════
// 🌤️ WEATHER API - Gestion météo pour conseils plantation
// ═══════════════════════════════════════════════════════════════════════════

class WeatherAPI {
    constructor() {
        this.API_KEY = 'db53bd06b5e66cd035d8f80a78108f10';
        this.BASE_URL = 'https://api.openweathermap.org/data/2.5';
        this.cache = {};
        this.cacheExpiry = 30 * 60 * 1000; // 30 minutes
        
        // Ville par défaut
        this.currentCity = this.loadCity() || 'Wattrelos';
    }
    
    // Sauvegarder/charger la ville
    saveCity(city) {
        try {
            localStorage.setItem('weather_city', city);
            this.currentCity = city;
        } catch (e) {
            console.error('Erreur sauvegarde ville:', e);
        }
    }
    
    loadCity() {
        try {
            return localStorage.getItem('weather_city');
        } catch (e) {
            return null;
        }
    }
    
    // Géolocalisation
    async getCityFromGeolocation() {
        return new Promise((resolve, reject) => {
            if (!navigator.geolocation) {
                reject(new Error('Géolocalisation non supportée'));
                return;
            }
            
            navigator.geolocation.getCurrentPosition(
                async (position) => {
                    try {
                        const { latitude, longitude } = position.coords;
                        const url = `${this.BASE_URL}/weather?lat=${latitude}&lon=${longitude}&appid=${this.API_KEY}&units=metric&lang=fr`;
                        const response = await fetch(url);
                        const data = await response.json();
                        
                        if (data.name) {
                            this.saveCity(data.name);
                            resolve(data.name);
                        } else {
                            reject(new Error('Ville introuvable'));
                        }
                    } catch (error) {
                        reject(error);
                    }
                },
                (error) => {
                    reject(error);
                }
            );
        });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // MÉTÉO ACTUELLE
    // ═══════════════════════════════════════════════════════════════════════════
    
    async getCurrentWeather(city = null) {
        city = city || this.currentCity;
        const cacheKey = `current_${city}`;
        
        // Vérifier le cache
        if (this.isCacheValid(cacheKey)) {
            console.log('📦 Météo depuis le cache');
            return this.cache[cacheKey].data;
        }

        try {
            const url = `${this.BASE_URL}/weather?q=${city}&appid=${this.API_KEY}&units=metric&lang=fr`;
            const response = await fetch(url);
            
            if (!response.ok) {
                throw new Error(`Erreur API: ${response.status}`);
            }
            
            const data = await response.json();
            
            const weatherData = {
                city: data.name,
                country: data.sys.country,
                temp: Math.round(data.main.temp),
                feelsLike: Math.round(data.main.feels_like),
                tempMin: Math.round(data.main.temp_min),
                tempMax: Math.round(data.main.temp_max),
                humidity: data.main.humidity,
                pressure: data.main.pressure,
                description: data.weather[0].description,
                icon: data.weather[0].icon,
                wind: data.wind.speed,
                clouds: data.clouds.all,
                timestamp: Date.now()
            };
            
            // Mettre en cache
            this.cache[cacheKey] = {
                data: weatherData,
                timestamp: Date.now()
            };
            
            console.log('🌤️ Météo récupérée:', weatherData);
            return weatherData;
            
        } catch (error) {
            console.error('❌ Erreur récupération météo:', error);
            throw error;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PRÉVISIONS 5 JOURS (par tranches de 3h)
    // ═══════════════════════════════════════════════════════════════════════════
    
    async getForecast(city = null) {
        city = city || this.currentCity;
        const cacheKey = `forecast_${city}`;
        
        if (this.isCacheValid(cacheKey)) {
            console.log('📦 Prévisions depuis le cache');
            return this.cache[cacheKey].data;
        }

        try {
            const url = `${this.BASE_URL}/forecast?q=${city}&appid=${this.API_KEY}&units=metric&lang=fr`;
            const response = await fetch(url);
            
            if (!response.ok) {
                throw new Error(`Erreur API: ${response.status}`);
            }
            
            const data = await response.json();
            
            // Grouper par jour
            const dailyForecasts = this.groupForecastsByDay(data.list);
            
            // Mettre en cache
            this.cache[cacheKey] = {
                data: dailyForecasts,
                timestamp: Date.now()
            };
            
            console.log('📅 Prévisions récupérées:', dailyForecasts);
            return dailyForecasts;
            
        } catch (error) {
            console.error('❌ Erreur récupération prévisions:', error);
            throw error;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // GROUPER PRÉVISIONS PAR JOUR
    // ═══════════════════════════════════════════════════════════════════════════
    
    groupForecastsByDay(forecastList) {
        const days = {};
        
        forecastList.forEach(item => {
            const date = new Date(item.dt * 1000);
            const dayKey = date.toISOString().split('T')[0]; // YYYY-MM-DD
            
            if (!days[dayKey]) {
                days[dayKey] = {
                    date: dayKey,
                    dateObj: date,
                    temps: [],
                    tempMin: Infinity,
                    tempMax: -Infinity,
                    descriptions: [],
                    icons: [],
                    rain: 0,
                    frostRisk: false
                };
            }
            
            // Ajouter cette tranche horaire
            days[dayKey].temps.push({
                hour: date.getHours(),
                temp: item.main.temp,
                feelsLike: item.main.feels_like,
                description: item.weather[0].description,
                icon: item.weather[0].icon
            });
            
            // Min/Max
            days[dayKey].tempMin = Math.min(days[dayKey].tempMin, item.main.temp);
            days[dayKey].tempMax = Math.max(days[dayKey].tempMax, item.main.temp);
            
            // Pluie
            if (item.rain && item.rain['3h']) {
                days[dayKey].rain += item.rain['3h'];
            }
            
            // Risque de gel
            if (item.main.temp <= 2) {
                days[dayKey].frostRisk = true;
            }
            
            // Description et icône les plus fréquentes
            days[dayKey].descriptions.push(item.weather[0].description);
            days[dayKey].icons.push(item.weather[0].icon);
        });
        
        // Convertir en array et calculer les moyennes
        return Object.values(days).map(day => {
            // Température moyenne
            const avgTemp = day.temps.reduce((sum, t) => sum + t.temp, 0) / day.temps.length;
            
            // Description la plus fréquente
            const descriptionCounts = {};
            day.descriptions.forEach(d => {
                descriptionCounts[d] = (descriptionCounts[d] || 0) + 1;
            });
            const mostFrequentDesc = Object.keys(descriptionCounts).reduce((a, b) => 
                descriptionCounts[a] > descriptionCounts[b] ? a : b
            );
            
            // Icône la plus fréquente
            const iconCounts = {};
            day.icons.forEach(i => {
                iconCounts[i] = (iconCounts[i] || 0) + 1;
            });
            const mostFrequentIcon = Object.keys(iconCounts).reduce((a, b) => 
                iconCounts[a] > iconCounts[b] ? a : b
            );
            
            return {
                date: day.date,
                dateObj: day.dateObj,
                dayName: this.getDayName(day.dateObj),
                tempMin: Math.round(day.tempMin),
                tempMax: Math.round(day.tempMax),
                tempAvg: Math.round(avgTemp),
                description: mostFrequentDesc,
                icon: mostFrequentIcon,
                rain: Math.round(day.rain * 10) / 10, // 1 décimale
                frostRisk: day.frostRisk,
                hourly: day.temps
            };
        });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ESTIMATION TEMPÉRATURE DU SOL
    // ═══════════════════════════════════════════════════════════════════════════
    
    estimateSoilTemperature(airTemp, month) {
        // Température du sol ≈ température air - 3 à 5°C selon la saison
        // Plus l'écart est important au début du printemps
        
        let offset;
        
        if (month >= 3 && month <= 5) {
            // Printemps : sol se réchauffe lentement
            offset = 5;
        } else if (month >= 6 && month <= 8) {
            // Été : sol chaud, proche de l'air
            offset = 2;
        } else if (month >= 9 && month <= 11) {
            // Automne : sol garde la chaleur
            offset = 3;
        } else {
            // Hiver : sol très froid
            offset = 4;
        }
        
        return Math.round(airTemp - offset);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ANALYSE POUR PLANTATION
    // ═══════════════════════════════════════════════════════════════════════════
    
    async getPlantingConditions(city = null) {
        city = city || this.currentCity;
        try {
            const current = await this.getCurrentWeather(city);
            const forecast = await this.getForecast(city);
            
            const now = new Date();
            const month = now.getMonth() + 1; // 1-12
            
            // Température du sol estimée
            const soilTemp = this.estimateSoilTemperature(current.temp, month);
            
            // Analyser les prochains jours
            const next7Days = forecast.slice(0, 7);
            
            // Température minimale nocturne moyenne sur 7 jours
            const avgNightTemp = next7Days.reduce((sum, day) => sum + day.tempMin, 0) / next7Days.length;
            
            // Risque de gel dans les 7 prochains jours
            const frostRiskDays = next7Days.filter(day => day.frostRisk).length;
            
            // Température moyenne journée
            const avgDayTemp = next7Days.reduce((sum, day) => sum + day.tempMax, 0) / next7Days.length;
            
            return {
                current: current,
                forecast: next7Days,
                soilTemp: soilTemp,
                avgNightTemp: Math.round(avgNightTemp),
                avgDayTemp: Math.round(avgDayTemp),
                frostRisk: frostRiskDays > 0,
                frostRiskDays: frostRiskDays,
                isSafe: frostRiskDays === 0 && avgNightTemp > 5,
                month: month,
                season: this.getSeason(month)
            };
            
        } catch (error) {
            console.error('❌ Erreur analyse conditions:', error);
            throw error;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // UTILITAIRES
    // ═══════════════════════════════════════════════════════════════════════════
    
    isCacheValid(key) {
        if (!this.cache[key]) return false;
        const age = Date.now() - this.cache[key].timestamp;
        return age < this.cacheExpiry;
    }
    
    getDayName(date) {
        const days = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
        return days[date.getDay()];
    }
    
    getSeason(month) {
        if (month >= 3 && month <= 5) return 'printemps';
        if (month >= 6 && month <= 8) return 'été';
        if (month >= 9 && month <= 11) return 'automne';
        return 'hiver';
    }
    
    clearCache() {
        this.cache = {};
        console.log('🗑️ Cache météo vidé');
    }
}

// Export pour utilisation globale
const weatherAPI = new WeatherAPI();

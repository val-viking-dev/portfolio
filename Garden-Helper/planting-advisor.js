// ═══════════════════════════════════════════════════════════════════════════
// 🧠 PLANTING ADVISOR - Conseils plantation intelligents basés sur météo
// ═══════════════════════════════════════════════════════════════════════════

class PlantingAdvisor {
    constructor() {
        this.weatherAPI = weatherAPI;
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ANALYSER SI ON PEUT PLANTER UN LÉGUME
    // ═══════════════════════════════════════════════════════════════════════════
    
    async canPlant(vegetableKey) {
        const vegetable = vegetablesDatabase[vegetableKey];
        if (!vegetable) {
            return { status: 'error', message: 'Légume introuvable' };
        }

        // Si pas de données de température, on ne peut pas conseiller
        if (!vegetable.plantingTemp) {
            return {
                status: 'unknown',
                badge: '❓',
                message: 'Données de température non disponibles',
                canPlant: null
            };
        }

        try {
            const conditions = await this.weatherAPI.getPlantingConditions();
            
            return this.analyzeForVegetable(vegetable, conditions);
            
        } catch (error) {
            console.error('❌ Erreur analyse plantation:', error);
            return {
                status: 'error',
                badge: '❌',
                message: 'Impossible de récupérer la météo',
                canPlant: false
            };
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ANALYSER POUR UN LÉGUME SPÉCIFIQUE
    // ═══════════════════════════════════════════════════════════════════════════
    
    analyzeForVegetable(vegetable, conditions) {
        const required = vegetable.plantingTemp;
        const soilTemp = conditions.soilTemp;
        const nightTemp = conditions.avgNightTemp;
        const dayTemp = conditions.avgDayTemp;
        const frostRisk = conditions.frostRisk;
        
        // 🔴 GELÉE ANNONCÉE + LÉGUME SENSIBLE
        if (frostRisk && vegetable.frostTolerance === 'none') {
            return {
                status: 'danger',
                badge: '❄️',
                badgeClass: 'badge-frost',
                message: `DANGER : Gelée dans ${conditions.frostRiskDays} jour(s) ! Ce légume ne supporte pas le gel.`,
                details: `Gelées annoncées. ${vegetable.name} mourrait.`,
                canPlant: false,
                daysToWait: this.estimateDaysToSafeTemp(conditions, required)
            };
        }

        // 🔴 SOL TROP FROID
        if (soilTemp < required.soil) {
            const diff = required.soil - soilTemp;
            return {
                status: 'cold',
                badge: '🔴',
                badgeClass: 'badge-danger',
                message: `Trop froid pour planter`,
                details: `Sol actuel : ${soilTemp}°C, besoin : ${required.soil}°C minimum (manque ${diff}°C)`,
                canPlant: false,
                daysToWait: this.estimateDaysToSafeTemp(conditions, required)
            };
        }

        // 🟡 NUITS FRAÎCHES MAIS SOL OK
        if (nightTemp < required.airNight && soilTemp >= required.soil) {
            return {
                status: 'warning',
                badge: '🟡',
                badgeClass: 'badge-warning',
                message: `Possible avec protection`,
                details: `Sol OK (${soilTemp}°C) mais nuits fraîches (${nightTemp}°C). Utilisez un voile P17 ou tunnel.`,
                canPlant: true,
                protection: 'Voile P17 ou tunnel recommandé',
                daysToIdeal: this.estimateDaysToIdealTemp(conditions, required)
            };
        }

        // 🟢 CONDITIONS PARFAITES
        if (soilTemp >= required.soil && 
            nightTemp >= required.airNight && 
            dayTemp >= required.airDay &&
            !frostRisk) {
            return {
                status: 'perfect',
                badge: '🟢',
                badgeClass: 'badge-success',
                message: `Conditions parfaites !`,
                details: `Sol ${soilTemp}°C, nuits ${nightTemp}°C, journées ${dayTemp}°C. C'est le moment idéal !`,
                canPlant: true
            };
        }

        // 🟢 CONDITIONS BONNES (pas parfaites mais acceptables)
        if (soilTemp >= required.soil && nightTemp >= required.airNight) {
            return {
                status: 'good',
                badge: '🟢',
                badgeClass: 'badge-success',
                message: `Vous pouvez planter`,
                details: `Conditions acceptables. Sol ${soilTemp}°C, nuits ${nightTemp}°C.`,
                canPlant: true
            };
        }

        // Par défaut : conditions acceptables
        return {
            status: 'acceptable',
            badge: '🟡',
            badgeClass: 'badge-warning',
            message: `Conditions limites`,
            details: `Sol ${soilTemp}°C, nuits ${nightTemp}°C. Surveillance recommandée.`,
            canPlant: true,
            caution: true
        };
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ESTIMER JOURS AVANT TEMPÉRATURE SÛRE
    // ═══════════════════════════════════════════════════════════════════════════
    
    estimateDaysToSafeTemp(conditions, required) {
        // Estimation très simplifiée
        // En réalité, il faudrait des données climatiques historiques
        
        const soilDiff = Math.max(0, required.soil - conditions.soilTemp);
        const month = conditions.month;
        
        // Printemps : ~0.5°C par semaine
        // Été : déjà chaud
        // Automne : refroidissement
        // Hiver : très lent
        
        let daysPerDegree;
        
        if (month >= 3 && month <= 5) {
            // Printemps : réchauffement progressif
            daysPerDegree = 14; // 2 semaines par degré
        } else if (month >= 6 && month <= 8) {
            // Été : déjà chaud normalement
            daysPerDegree = 7;
        } else if (month >= 9 && month <= 11) {
            // Automne : attendre le printemps prochain
            return 150; // ~5 mois
        } else {
            // Hiver : attendre le printemps
            return 100; // ~3 mois
        }
        
        const estimatedDays = Math.ceil(soilDiff * daysPerDegree);
        
        return Math.min(estimatedDays, 120); // Max 4 mois
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ESTIMER JOURS AVANT TEMPÉRATURE IDÉALE
    // ═══════════════════════════════════════════════════════════════════════════
    
    estimateDaysToIdealTemp(conditions, required) {
        // Similaire mais pour atteindre température idéale
        const nightDiff = Math.max(0, required.airNight + 5 - conditions.avgNightTemp);
        return Math.ceil(nightDiff * 10); // ~10 jours par degré
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // CONSEILS POUR SEMIS INTÉRIEUR
    // ═══════════════════════════════════════════════════════════════════════════
    
    async getSeedingAdvice(vegetableKey) {
        const vegetable = vegetablesDatabase[vegetableKey];
        if (!vegetable || !vegetable.seedingTemp) {
            return null;
        }

        try {
            const conditions = await this.weatherAPI.getPlantingConditions();
            const plantingAdvice = await this.canPlant(vegetableKey);
            
            // Si on peut déjà planter dehors, pas besoin de semis intérieur
            if (plantingAdvice.canPlant && plantingAdvice.status === 'perfect') {
                return {
                    needSeeding: false,
                    message: `Conditions parfaites pour plantation directe !`
                };
            }

            // Calculer quand on pourra planter dehors
            const daysToWait = plantingAdvice.daysToWait || 30;
            
            // Période de croissance intérieure nécessaire
            const seedingWeeks = vegetable.seedingWeeksBeforePlanting || 6;
            const seedingDays = seedingWeeks * 7;
            
            // Calculer date idéale de semis
            const plantingDate = new Date();
            plantingDate.setDate(plantingDate.getDate() + daysToWait);
            
            const seedingDate = new Date(plantingDate);
            seedingDate.setDate(seedingDate.getDate() - seedingDays);
            
            const daysUntilSeeding = Math.ceil((seedingDate - new Date()) / (1000 * 60 * 60 * 24));
            
            if (daysUntilSeeding <= 0) {
                return {
                    needSeeding: true,
                    urgent: true,
                    message: `🚨 Démarrez vos semis MAINTENANT !`,
                    details: `Semis intérieur à ${vegetable.seedingTemp.ideal}°C. Plantation prévue dans ${daysToWait} jours.`,
                    seedingTemp: vegetable.seedingTemp,
                    daysToSeeding: 0
                };
            } else if (daysUntilSeeding <= 14) {
                return {
                    needSeeding: true,
                    soon: true,
                    message: `⏰ Démarrez vos semis dans ${daysUntilSeeding} jours`,
                    details: `Semis intérieur à ${vegetable.seedingTemp.ideal}°C.`,
                    seedingTemp: vegetable.seedingTemp,
                    daysToSeeding: daysUntilSeeding
                };
            } else {
                return {
                    needSeeding: true,
                    later: true,
                    message: `Démarrez vos semis dans ${Math.ceil(daysUntilSeeding / 7)} semaines`,
                    details: `Semis intérieur à ${vegetable.seedingTemp.ideal}°C.`,
                    seedingTemp: vegetable.seedingTemp,
                    daysToSeeding: daysUntilSeeding
                };
            }
            
        } catch (error) {
            console.error('❌ Erreur conseils semis:', error);
            return null;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // OBTENIR TOUS LES CONSEILS POUR LA LISTE DE LÉGUMES
    // ═══════════════════════════════════════════════════════════════════════════
    
    async getAllPlantingAdvice() {
        const conditions = await this.weatherAPI.getPlantingConditions();
        const advice = {};
        
        for (const [key, vegetable] of Object.entries(vegetablesDatabase)) {
            if (vegetable.plantingTemp) {
                advice[key] = this.analyzeForVegetable(vegetable, conditions);
            }
        }
        
        return {
            conditions: conditions,
            advice: advice
        };
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // FORMATER MESSAGE LISIBLE
    // ═══════════════════════════════════════════════════════════════════════════
    
    formatAdviceMessage(advice) {
        if (!advice) return '';
        
        let message = `${advice.badge} ${advice.message}`;
        
        if (advice.details) {
            message += `\n${advice.details}`;
        }
        
        if (advice.protection) {
            message += `\n⚠️ ${advice.protection}`;
        }
        
        if (advice.daysToWait) {
            const weeks = Math.ceil(advice.daysToWait / 7);
            message += `\n⏰ Attendez environ ${weeks} semaine(s)`;
        }
        
        return message;
    }
}

// Export pour utilisation globale
const plantingAdvisor = new PlantingAdvisor();

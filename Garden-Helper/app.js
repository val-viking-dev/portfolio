/*
 * ========================================
 * GARDEN-HELPER v1.9.2
 * ========================================
 * 
 * Application web de gestion de potager intelligent
 * avec rotation des cultures et compagnonnage
 * 
 * @author    Valentin
 * @role      Développeur / Concepteur d'Application
 * @date      Janvier 2026
 * @version   1.9.2
 * 
 * Fonctionnalités :
 * - Grille de potager personnalisable
 * - Système de parcelles et chemins
 * - Base de données légumes (rotation, compagnonnage)
 * - Conseils météo pour plantation
 * - PWA avec support offline
 * - Tutoriel interactif adaptatif
 * - Mode agrandissement de parcelles
 * 
 * Technologies : Vanilla JS, HTML5, CSS3, PWA
 * ========================================
 */

// ========================================
// CLASSE PRINCIPALE - GREENHOUSE APP
// ========================================

/**
 * Classe principale de l'application Garden-Helper
 * Gère l'état global, la grille, les parcelles et l'historique
 */
class GreenhouseApp {
    constructor() {
        // État de l'application (dimensions | manage)
        this.step = this.loadFromStorage('step', 'dimensions');
        
        // Dimensions du potager (en mètres)
        this.greenhouseDimensions = this.loadFromStorage('greenhouseDimensions', { 
            width: 6.5,  // Largeur
            length: 10,  // Longueur
            height: 4    // Hauteur (pour serres/balcons)
        });
        
        // Taille d'une case en mètres (grille fine 50cm)
        this.cellSize = 0.5;
        
        // Tableau de toutes les cases de la grille
        this.cells = this.loadFromStorage('cells', []);
        
        // Parcelles définies (groupes de cases cultivées)
        this.plots = this.loadFromStorage('plots', []);
        this.nextPlotId = this.loadFromStorage('nextPlotId', 1);
        
        // Système d'historique multi-années pour rotation
        this.currentYear = new Date().getFullYear();
        this.history = this.loadFromStorage('history', {});
        // Structure: { "2024": [{ plotId, vegetable, plantedDate, harvestedDate }], "2025": [...] }
        
        // Gestion de la sélection utilisateur
        this.isSelecting = false;          // Sélection en cours ?
        this.selectionStart = null;        // Point de départ sélection
        this.selectedCells = [];           // Cases actuellement sélectionnées
        this.currentMode = 'plot';         // Mode: 'plot' (parcelle), 'path' (chemin), 'expand' (agrandir)
        
        // Rotation de la grille (0°, 90°, 180°, 270°)
        this.gridRotation = this.loadFromStorage('gridRotation', 0);
        
        // Référence au conteneur DOM principal
        this.appContent = document.getElementById('app-content');
        
        // Initialiser l'affichage
        this.render();
    }
    
    // ========================================
    // GESTION DU LOCALSTORAGE
    // ========================================
    
    /**
     * Sauvegarde une donnée dans le localStorage
     * @param {string} key - Clé de stockage
     * @param {*} data - Données à sauvegarder (sera JSON.stringify)
     */
    saveToStorage(key, data) {
        try {
            localStorage.setItem('greenhouse_' + key, JSON.stringify(data));
        } catch (e) {
            console.error('Erreur sauvegarde:', e);
        }
    }
    
    /**
     * Récupère une donnée depuis le localStorage
     * @param {string} key - Clé de stockage
     * @param {*} defaultValue - Valeur par défaut si clé inexistante
     * @returns {*} Données récupérées ou valeur par défaut
     */
    loadFromStorage(key, defaultValue) {
        try {
            const item = localStorage.getItem('greenhouse_' + key);
            return item ? JSON.parse(item) : defaultValue;
        } catch (e) {
            return defaultValue;
        }
    }
    
    /**
     * Sauvegarde l'état complet de l'application dans le localStorage
     */
    saveState() {
        this.saveToStorage('step', this.step);
        this.saveToStorage('greenhouseDimensions', this.greenhouseDimensions);
        this.saveToStorage('cells', this.cells);
        this.saveToStorage('plots', this.plots);
        this.saveToStorage('nextPlotId', this.nextPlotId);
        this.saveToStorage('history', this.history);
        this.saveToStorage('gridRotation', this.gridRotation);
        this.saveToStorage('lastSaveDate', new Date().toISOString());
    }
    
    // ========================================
    // GESTION DE LA GRILLE
    // ========================================
    
    /**
     * Calcule le nombre de colonnes et lignes selon les dimensions
     * @returns {{cols: number, rows: number}} Nombre de colonnes et lignes
     */
    calculateGrid() {
        const cols = Math.floor(this.greenhouseDimensions.width / this.cellSize);
        const rows = Math.floor(this.greenhouseDimensions.length / this.cellSize);
        return { cols, rows };
    }
    
    /**
     * Initialise la grille avec des cases vides selon les dimensions
     * Relance le message d'aide après création
     */
    initializeCells() {
        this.cells = [];
        const grid = this.calculateGrid();
        
        // Créer toutes les cases
        for (let row = 0; row < grid.rows; row++) {
            for (let col = 0; col < grid.cols; col++) {
                this.cells.push({
                    id: `${row}-${col}`,
                    row, col,
                    type: 'empty', // Types possibles: 'empty', 'path', 'plot'
                    plotId: null
                });
            }
        }
        
        this.step = 'manage';
        this.saveState();
        this.render();
        
        // Afficher le message d'aide si pas encore vu
        if (!localStorage.getItem('greenhouse_tutorialCompleted')) {
            setTimeout(() => {
                const welcomeMessages = new WelcomeMessages(this);
                welcomeMessages.showManageMessage();
            }, 500);
        }
    }
    
    /**
     * Récupère une case par ses coordonnées
     * @param {number} row - Numéro de ligne
     * @param {number} col - Numéro de colonne
     * @returns {Object|undefined} La case trouvée ou undefined
     */
    getCell(row, col) {
        return this.cells.find(c => c.row === row && c.col === col);
    }
    
    // ========================================
    // GESTION DES PARCELLES ET CHEMINS
    // ========================================
    
    /**
     * Crée une parcelle à partir des cases sélectionnées
     * Marque les cases comme appartenant à cette parcelle
     * Ignore les cases qui sont déjà des chemins ou parcelles
     */
    createPlotFromSelection() {
        if (this.selectedCells.length === 0) return;
        
        // Filtrer uniquement les cases vides (ni chemin, ni parcelle)
        const validCells = this.selectedCells.filter(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            return cell && cell.type === 'empty';
        });
        
        // Vérifier qu'il reste des cases valides
        if (validCells.length === 0) {
            alert('❌ Vous ne pouvez créer une parcelle que sur des cases vides.\n\nLes chemins et parcelles existantes ne peuvent pas être utilisés.');
            this.selectedCells = [];
            this.render();
            return;
        }
        
        // Avertir si certaines cases ont été ignorées
        if (validCells.length < this.selectedCells.length) {
            const ignored = this.selectedCells.length - validCells.length;
            alert(`⚠️ ${ignored} case(s) ignorée(s) car déjà utilisée(s) comme chemin ou parcelle.\n\nParcelle créée avec ${validCells.length} case(s) vide(s).`);
        }
        
        const plotId = `plot-${this.nextPlotId++}`;
        
        // Marquer les cases valides comme faisant partie de cette parcelle
        validCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (cell) {
                cell.type = 'plot';
                cell.plotId = plotId;
            }
        });
        
        // Créer l'objet parcelle
        this.plots.push({
            id: plotId,
            cellIds: [...validCells],
            vegetable: null,
            plantedDate: null
        });
        
        this.selectedCells = [];
        this.saveState();
        this.render();
    }
    
    /**
     * Agrandit une parcelle existante avec les cases sélectionnées
     * Sélection doit contenir : 1 parcelle + cases vides adjacentes
     */
    expandPlotFromSelection() {
        if (this.selectedCells.length === 0) return;
        
        // Analyser la sélection
        const plotsInSelection = new Set();
        const emptyCells = [];
        
        this.selectedCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (!cell) return;
            
            if (cell.type === 'plot') {
                plotsInSelection.add(cell.plotId);
            } else if (cell.type === 'empty') {
                emptyCells.push(cellId);
            }
        });
        
        // Vérifications
        if (plotsInSelection.size === 0) {
            alert('❌ Vous devez sélectionner une parcelle existante à agrandir.\n\nSélectionnez la parcelle + les cases vides adjacentes.');
            this.selectedCells = [];
            this.render();
            return;
        }
        
        if (plotsInSelection.size > 1) {
            alert('❌ Vous ne pouvez agrandir qu\'une seule parcelle à la fois.\n\nSélectionnez une seule parcelle + cases vides.');
            this.selectedCells = [];
            this.render();
            return;
        }
        
        if (emptyCells.length === 0) {
            alert('❌ Aucune case vide à ajouter.\n\nSélectionnez des cases vides pour agrandir la parcelle.');
            this.selectedCells = [];
            this.render();
            return;
        }
        
        // Récupérer la parcelle à agrandir
        const plotId = Array.from(plotsInSelection)[0];
        const plot = this.plots.find(p => p.id === plotId);
        
        if (!plot) {
            this.selectedCells = [];
            this.render();
            return;
        }
        
        // Demander confirmation
        const vegetableName = plot.vegetable 
            ? vegetablesDatabase[plot.vegetable]?.icon + ' ' + vegetablesDatabase[plot.vegetable]?.name 
            : 'vide';
        
        const confirmMessage = `📏 Agrandir cette parcelle ?\n\nParcelle actuelle : ${plot.cellIds.length} cases (${vegetableName})\nAjout : ${emptyCells.length} case(s) vide(s)\nNouvelle taille : ${plot.cellIds.length + emptyCells.length} cases`;
        
        if (!confirm(confirmMessage)) {
            this.selectedCells = [];
            this.render();
            return;
        }
        
        // Ajouter les cases vides à la parcelle
        emptyCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (cell) {
                cell.type = 'plot';
                cell.plotId = plotId;
                plot.cellIds.push(cellId);
            }
        });
        
        this.selectedCells = [];
        this.saveState();
        this.render();
    }
    
    /**
     * Marque les cases sélectionnées comme chemin
     * Les chemins sont des allées entre les parcelles
     * Demande confirmation si parcelles plantées incluses
     */
    markSelectionAsPath() {
        if (this.selectedCells.length === 0) return;
        
        // Analyser la sélection
        const validCells = [];
        const plantedParcels = []; // Parcelles plantées qui seront supprimées
        
        this.selectedCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (!cell) return;
            
            // Toutes les cases sont acceptées
            validCells.push(cellId);
            
            // Détecter les parcelles plantées
            if (cell.type === 'plot') {
                const plot = this.plots.find(p => p.id === cell.plotId);
                if (plot && plot.vegetable) {
                    const vegName = vegetablesDatabase[plot.vegetable]?.icon + ' ' + vegetablesDatabase[plot.vegetable]?.name;
                    if (!plantedParcels.some(p => p.plotId === plot.id)) {
                        plantedParcels.push({
                            plotId: plot.id,
                            name: vegName,
                            cellCount: plot.cellIds.length
                        });
                    }
                }
            }
        });
        
        // Demander confirmation si parcelles plantées
        if (plantedParcels.length > 0) {
            let confirmMessage = '⚠️ ATTENTION : Vous allez supprimer des parcelles plantées !\n\n';
            confirmMessage += 'Parcelles qui seront supprimées :\n';
            plantedParcels.forEach(p => {
                confirmMessage += `  • ${p.name} (${p.cellCount} cases)\n`;
            });
            confirmMessage += '\nCréer le chemin malgré tout ?';
            
            if (!confirm(confirmMessage)) {
                this.selectedCells = [];
                this.render();
                return;
            }
        }
        
        // Créer le chemin sur toutes les cases
        validCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (!cell) return;
            
            // Si c'était une parcelle, la retirer/supprimer
            if (cell.type === 'plot') {
                const plot = this.plots.find(p => p.id === cell.plotId);
                if (plot) {
                    // Retirer cette case de la parcelle
                    plot.cellIds = plot.cellIds.filter(id => id !== cellId);
                    
                    // Si la parcelle n'a plus de cases, la supprimer
                    if (plot.cellIds.length === 0) {
                        this.plots = this.plots.filter(p => p.id !== plot.id);
                    }
                }
            }
            
            // Transformer en chemin
            cell.type = 'path';
            cell.plotId = null;
        });
        
        this.selectedCells = [];
        this.saveState();
        this.render();
    }
    
    /**
     * Efface les cases sélectionnées (supprime parcelles/chemins)
     * Logique intelligente :
     * - Parcelle entière sélectionnée → Suppression totale
     * - Partie de parcelle sélectionnée → Retrait des cases uniquement
     * - Chemins → Suppression directe
     */
    clearSelection() {
        if (this.selectedCells.length === 0) return;
        
        // Analyser la sélection
        const parcelsAffected = {}; // { plotId: { selectedCells: [], totalCells: number, plot: object } }
        const pathsToDelete = [];
        
        this.selectedCells.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            
            if (!cell) return;
            
            if (cell.type === 'path') {
                pathsToDelete.push(cellId);
            } else if (cell.type === 'plot') {
                const plot = this.plots.find(p => p.id === cell.plotId);
                if (plot) {
                    if (!parcelsAffected[cell.plotId]) {
                        parcelsAffected[cell.plotId] = {
                            selectedCells: [],
                            totalCells: plot.cellIds.length,
                            plot: plot
                        };
                    }
                    parcelsAffected[cell.plotId].selectedCells.push(cellId);
                }
            }
        });
        
        // Construire le message de confirmation
        let confirmMessage = '';
        const actions = [];
        
        // Analyser les parcelles affectées
        Object.keys(parcelsAffected).forEach(plotId => {
            const info = parcelsAffected[plotId];
            const plot = info.plot;
            const selectedCount = info.selectedCells.length;
            const totalCount = info.totalCells;
            const vegetableName = plot.vegetable ? vegetablesDatabase[plot.vegetable]?.icon + ' ' + vegetablesDatabase[plot.vegetable]?.name : 'vide';
            
            if (selectedCount === totalCount) {
                // Suppression totale
                actions.push({
                    type: 'delete',
                    plotId: plotId,
                    message: `• Supprimer parcelle (${totalCount} cases, ${vegetableName})`
                });
            } else {
                // Retrait partiel
                actions.push({
                    type: 'remove',
                    plotId: plotId,
                    cellsToRemove: info.selectedCells,
                    message: `• Retirer ${selectedCount} case(s) de la parcelle (${vegetableName}, ${totalCount} → ${totalCount - selectedCount} cases)`
                });
            }
        });
        
        // Ajouter les chemins
        if (pathsToDelete.length > 0) {
            actions.push({
                type: 'paths',
                cells: pathsToDelete,
                message: `• Supprimer ${pathsToDelete.length} case(s) de chemin`
            });
        }
        
        // Construire le message final
        if (actions.length === 0) {
            this.selectedCells = [];
            this.render();
            return;
        }
        
        confirmMessage = '⚠️ Confirmer cette action ?\n\n';
        confirmMessage += actions.map(a => a.message).join('\n');
        
        // Demander confirmation
        if (!confirm(confirmMessage)) {
            this.selectedCells = [];
            this.render();
            return;
        }
        
        // Exécuter les actions
        actions.forEach(action => {
            if (action.type === 'delete') {
                // Supprimer la parcelle entière
                const plot = this.plots.find(p => p.id === action.plotId);
                if (plot) {
                    plot.cellIds.forEach(cId => {
                        const c = this.cells.find(cell => cell.id === cId);
                        if (c) {
                            c.type = 'empty';
                            c.plotId = null;
                        }
                    });
                    this.plots = this.plots.filter(p => p.id !== action.plotId);
                }
            } else if (action.type === 'remove') {
                // Retirer des cases de la parcelle
                const plot = this.plots.find(p => p.id === action.plotId);
                if (plot) {
                    action.cellsToRemove.forEach(cId => {
                        const c = this.cells.find(cell => cell.id === cId);
                        if (c) {
                            c.type = 'empty';
                            c.plotId = null;
                        }
                        // Retirer de la liste des cellIds de la parcelle
                        plot.cellIds = plot.cellIds.filter(id => id !== cId);
                    });
                    
                    // Si la parcelle n'a plus de cases, la supprimer
                    if (plot.cellIds.length === 0) {
                        this.plots = this.plots.filter(p => p.id !== action.plotId);
                    }
                }
            } else if (action.type === 'paths') {
                // Supprimer les chemins
                action.cells.forEach(cId => {
                    const c = this.cells.find(cell => cell.id === cId);
                    if (c) {
                        c.type = 'empty';
                        c.plotId = null;
                    }
                });
            }
        });
        
        this.selectedCells = [];
        this.saveState();
        this.render();
    }
    
    // ========================================
    // GESTION DE LA SÉLECTION (CLICK & DRAG)
    // ========================================
    
    /**
     * Démarre une sélection lors du clic/touch sur une case
     * @param {string} cellId - ID de la case cliquée
     */
    handleCellMouseDown(cellId) {
        const cell = this.cells.find(c => c.id === cellId);
        if (!cell) return;
        
        this.isSelecting = true;
        this.selectionStart = { row: cell.row, col: cell.col };
        this.selectedCells = [cellId];
        this.updateSelection();
    }
    
    /**
     * Étend la sélection lors du survol (crée un rectangle)
     * @param {string} cellId - ID de la case survolée
     */
    handleCellMouseEnter(cellId) {
        if (!this.isSelecting) return;
        
        const cell = this.cells.find(c => c.id === cellId);
        if (!cell) return;
        
        // Calculer le rectangle de sélection
        const minRow = Math.min(this.selectionStart.row, cell.row);
        const maxRow = Math.max(this.selectionStart.row, cell.row);
        const minCol = Math.min(this.selectionStart.col, cell.col);
        const maxCol = Math.max(this.selectionStart.col, cell.col);
        
        this.selectedCells = [];
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const c = this.getCell(row, col);
                if (c) {
                    this.selectedCells.push(c.id);
                }
            }
        }
        
        this.updateSelection();
    }
    
    handleMouseUp() {
        this.isSelecting = false;
    }
    
    updateSelection() {
        // Mettre à jour visuellement les cases sélectionnées
        document.querySelectorAll('.cell').forEach(el => {
            if (this.selectedCells.includes(el.dataset.cellId)) {
                el.classList.add('cell-selected');
            } else {
                el.classList.remove('cell-selected');
            }
        });
        
        // Mettre à jour les boutons d'action (Desktop)
        const confirmBtn = document.getElementById('confirm-selection');
        const clearBtn = document.getElementById('clear-selection');
        
        if (confirmBtn && clearBtn) {
            if (this.selectedCells.length > 0) {
                confirmBtn.disabled = false;
                confirmBtn.textContent = `✓ Valider (${this.selectedCells.length} cases)`;
                clearBtn.disabled = false;
            } else {
                confirmBtn.disabled = true;
                confirmBtn.textContent = '✓ Valider (0 cases)';
                clearBtn.disabled = true;
            }
        }
        
        // Mettre à jour le FAB (Mobile)
        this.updateFABSelection();
    }
    
    // Planter un légume dans une parcelle
    plantVegetable(plotId, vegetableKey) {
        const plot = this.plots.find(p => p.id === plotId);
        if (!plot) return;
        
        plot.vegetable = vegetableKey;
        plot.plantedDate = new Date().toISOString();
        
        // Enregistrer dans l'historique
        this.addToHistory(plotId, vegetableKey, plot.plantedDate);
        
        this.saveState();
        this.render();
    }
    
    // Retirer un légume
    removePlant(plotId) {
        const plot = this.plots.find(p => p.id === plotId);
        if (!plot) return;
        
        // Enregistrer la récolte dans l'historique
        if (plot.vegetable) {
            this.markAsHarvested(plotId, new Date().toISOString());
        }
        
        plot.vegetable = null;
        plot.plantedDate = null;
        
        this.saveState();
        this.render();
    }
    
    // Calculer les jours depuis plantation
    getDaysSincePlanting(plantedDate) {
        if (!plantedDate) return 0;
        const planted = new Date(plantedDate);
        const now = new Date();
        const diff = now - planted;
        return Math.floor(diff / (1000 * 60 * 60 * 24));
    }
    
    // Ajouter à l'historique
    addToHistory(plotId, vegetableKey, plantedDate) {
        const year = new Date(plantedDate).getFullYear();
        if (!this.history[year]) {
            this.history[year] = [];
        }
        
        this.history[year].push({
            plotId: plotId,
            vegetable: vegetableKey,
            plantedDate: plantedDate,
            harvestedDate: null
        });
    }
    
    // Marquer comme récolté
    markAsHarvested(plotId, harvestedDate) {
        const year = new Date(harvestedDate).getFullYear();
        if (!this.history[year]) return;
        
        // Trouver la dernière entrée pour cette parcelle sans date de récolte
        for (let i = this.history[year].length - 1; i >= 0; i--) {
            if (this.history[year][i].plotId === plotId && !this.history[year][i].harvestedDate) {
                this.history[year][i].harvestedDate = harvestedDate;
                break;
            }
        }
    }
    
    // Obtenir l'historique d'une parcelle
    getPlotHistory(plotId, years = 3) {
        const history = [];
        const currentYear = new Date().getFullYear();
        
        for (let i = 0; i < years; i++) {
            const year = currentYear - i;
            if (this.history[year]) {
                const yearHistory = this.history[year].filter(h => h.plotId === plotId);
                history.push(...yearHistory);
            }
        }
        
        return history;
    }
    
    // Obtenir le score de rotation pour une parcelle
    getRotationScore(vegetableKey, plotId) {
        const vegetable = vegetablesDatabase[vegetableKey];
        if (!vegetable || !vegetable.family) return { score: 1, badge: '🟡', text: 'Acceptable' };
        
        // Récupérer les cultures de l'année dernière sur cette parcelle
        const lastYearCrops = this.getLastYearCrops(plotId);
        
        if (lastYearCrops.length === 0) {
            // Pas d'historique = rotation optimale
            return { score: 2, badge: '🟢', text: 'Rotation optimale' };
        }
        
        // Vérifier si une culture de la même famille a été plantée l'année dernière
        const sameFamilyLastYear = lastYearCrops.some(cropKey => {
            const crop = vegetablesDatabase[cropKey];
            return crop && crop.family === vegetable.family;
        });
        
        if (sameFamilyLastYear) {
            // Même famille l'année dernière = à éviter
            return { score: 0, badge: '🔴', text: 'À éviter (même famille)' };
        }
        
        // Vérifier si une légumineuse (enrichit le sol) a été plantée l'année dernière
        const leguminousLastYear = lastYearCrops.some(cropKey => {
            const crop = vegetablesDatabase[cropKey];
            return crop && crop.family === 'fabacees';
        });
        
        // Vérifier si le légume actuel est gourmand
        const isHungry = vegetable.nutrients && (
            vegetable.nutrients.includes('GOURMAND') || 
            vegetable.nutrients.includes('TRES GOURMAND')
        );
        
        if (leguminousLastYear && isHungry) {
            // Légumineuse l'année dernière + légume gourmand = rotation optimale
            return { score: 2, badge: '🟢', text: 'Rotation optimale (après légumineuse)' };
        }
        
        // Famille différente = acceptable
        return { score: 1, badge: '🟡', text: 'Acceptable' };
    }
    
    // Obtenir les légumes plantés l'année dernière sur une parcelle
    getLastYearCrops(plotId) {
        const lastYear = new Date().getFullYear() - 1;
        if (!this.history[lastYear]) return [];
        
        return this.history[lastYear]
            .filter(h => h.plotId === plotId)
            .map(h => h.vegetable);
    }
    
    // Obtenir les parcelles voisines
    getNeighborPlots(plot) {
        const neighborPlots = new Set();
        
        plot.cellIds.forEach(cellId => {
            const cell = this.cells.find(c => c.id === cellId);
            if (!cell) return;
            
            const neighbors = [
                this.getCell(cell.row - 1, cell.col),
                this.getCell(cell.row + 1, cell.col),
                this.getCell(cell.row, cell.col - 1),
                this.getCell(cell.row, cell.col + 1)
            ];
            
            neighbors.forEach(n => {
                if (n && n.type === 'plot' && n.plotId !== plot.id) {
                    neighborPlots.add(n.plotId);
                }
            });
        });
        
        return Array.from(neighborPlots).map(id => this.plots.find(p => p.id === id)).filter(Boolean);
    }
    
    // Vérifier les voisins
    checkNeighbors(plot) {
        if (!plot.vegetable) return { warnings: [], suggestions: [] };
        
        const warnings = [];
        const suggestions = [];
        const neighborPlots = this.getNeighborPlots(plot);
        
        neighborPlots.forEach(neighbor => {
            if (neighbor.vegetable) {
                const veg = vegetablesDatabase[plot.vegetable];
                if (veg && veg.badCompanions.includes(neighbor.vegetable)) {
                    warnings.push(vegetablesDatabase[neighbor.vegetable].name);
                }
            } else {
                const veg = vegetablesDatabase[plot.vegetable];
                if (veg && veg.goodCompanions.length > 0) {
                    suggestions.push(...veg.goodCompanions.slice(0, 2));
                }
            }
        });
        
        return { warnings: [...new Set(warnings)], suggestions: [...new Set(suggestions)] };
    }
    
    // ========================================
    // ROTATION DE LA VUE
    // ========================================
    
    /**
     * Fait tourner la grille de 90° (0° → 90° → 180° → 270° → 0°)
     * Purement visuel, ne modifie pas les données
     */
    rotateGrid() {
        this.gridRotation = (this.gridRotation + 90) % 360;
        this.saveState();
        this.render();
    }
    
    // Rendu
    render() {
        if (this.step === 'dimensions') {
            this.renderDimensions();
        } else {
            this.renderManage();
        }
    }
    
    renderDimensions() {
        const volume = (this.greenhouseDimensions.width * this.greenhouseDimensions.length * this.greenhouseDimensions.height).toFixed(2);
        const surface = (this.greenhouseDimensions.width * this.greenhouseDimensions.length).toFixed(2);
        
        this.appContent.innerHTML = `
            <div class="setup-step">
                <h2>📏 Dimensions de votre potager</h2>
                <p style="color: var(--text-secondary); margin-bottom: 1.5rem;">
                    Entrez les dimensions réelles de votre potager en mètres
                </p>
                
                <div class="dimension-inputs">
                    <div class="input-group">
                        <label>🔷 Largeur (mètres)</label>
                        <input type="number" id="width" min="1" max="50" step="0.5" value="${this.greenhouseDimensions.width}">
                    </div>
                    <div class="input-group">
                        <label>🔶 Longueur (mètres)</label>
                        <input type="number" id="length" min="1" max="50" step="0.5" value="${this.greenhouseDimensions.length}">
                    </div>
                    <div class="input-group">
                        <label>📐 Hauteur (mètres)</label>
                        <input type="number" id="height" min="0.5" max="10" step="0.1" value="${this.greenhouseDimensions.height}">
                    </div>
                </div>
                
                <div class="info-box" style="background: #dbeafe; padding: 1rem; border-radius: 0.75rem; margin-top: 1.5rem;">
                    <h3 style="font-size: 1rem; margin-bottom: 0.5rem;">📊 Informations</h3>
                    <div style="display: grid; gap: 0.5rem; font-size: 0.875rem;">
                        <div><strong>Surface au sol :</strong> ${surface} m²</div>
                        <div><strong>Volume total :</strong> ${volume} m³</div>
                        <div><strong>Taille d'une case :</strong> ${this.cellSize}m × ${this.cellSize}m</div>
                        <div><strong>Grille :</strong> ${this.calculateGrid().cols} × ${this.calculateGrid().rows} cases</div>
                    </div>
                </div>
                
                <button class="btn" id="start-btn">Créer mon potager →</button>
            </div>
        `;
        
        // Event listeners
        const updateDimensions = () => {
            this.greenhouseDimensions.width = parseFloat(document.getElementById('width').value) || 1;
            this.greenhouseDimensions.length = parseFloat(document.getElementById('length').value) || 1;
            this.greenhouseDimensions.height = parseFloat(document.getElementById('height').value) || 0.5;
            this.saveState();
            this.render();
        };
        
        document.getElementById('width').addEventListener('input', updateDimensions);
        document.getElementById('length').addEventListener('input', updateDimensions);
        document.getElementById('height').addEventListener('input', updateDimensions);
        
        document.getElementById('start-btn').addEventListener('click', () => {
            this.initializeCells();
        });
    }
    
    renderManage() {
        const grid = this.calculateGrid();
        const stats = {
            totalCells: this.cells.length,
            paths: this.cells.filter(c => c.type === 'path').length,
            plotCells: this.cells.filter(c => c.type === 'plot').length,
            plots: this.plots.length,
            plantedPlots: this.plots.filter(p => p.vegetable).length,
            emptyCells: this.cells.filter(c => c.type === 'empty').length,
            emptyPlots: this.plots.filter(p => !p.vegetable).length
        };
        
        this.appContent.innerHTML = `
            <div class="info-bar">
                <div class="info-item">
                    <div class="info-item-value">${this.greenhouseDimensions.width}×${this.greenhouseDimensions.length}m</div>
                    <div class="info-item-label">Dimensions</div>
                </div>
                <div class="info-item">
                    <div class="info-item-value">${stats.plots}</div>
                    <div class="info-item-label">Parcelles créées</div>
                </div>
                <div class="info-item">
                    <div class="info-item-value">${stats.plantedPlots}</div>
                    <div class="info-item-label">Cultivées</div>
                </div>
                <div class="info-item">
                    <div class="info-item-value">${stats.emptyPlots}</div>
                    <div class="info-item-label">Parcelles libres</div>
                </div>
            </div>
            
            <div class="toolbar">
                <div class="toolbar-section">
                    <h3>🎯 Mode de sélection :</h3>
                    <div class="mode-buttons">
                        <button class="mode-btn ${this.currentMode === 'plot' ? 'active' : ''}" data-mode="plot">
                            📦 Créer une parcelle
                        </button>
                        <button class="mode-btn ${this.currentMode === 'expand' ? 'active' : ''}" data-mode="expand">
                            📏 Agrandir parcelle
                        </button>
                        <button class="mode-btn ${this.currentMode === 'path' ? 'active' : ''}" data-mode="path">
                            🛤️ Créer un chemin
                        </button>
                    </div>
                </div>
                
                <div class="toolbar-section">
                    <h3>✏️ Actions :</h3>
                    <div class="action-buttons">
                        <button class="action-btn" id="confirm-selection" ${this.selectedCells.length === 0 ? 'disabled' : ''}>
                            ✓ Valider (${this.selectedCells.length} cases)
                        </button>
                        <button class="action-btn btn-danger" id="clear-selection" ${this.selectedCells.length === 0 ? 'disabled' : ''}>
                            ✗ Effacer sélection
                        </button>
                    </div>
                </div>
                
                <div class="toolbar-section">
                    <h3>🔄 Vue :</h3>
                    <div class="action-buttons">
                        <button class="action-btn btn-secondary" id="rotate-grid-btn">
                            🔄 Tourner la grille${this.gridRotation > 0 ? ` (${this.gridRotation}°)` : ''}
                        </button>
                        <button class="action-btn btn-secondary" id="btn-weather">
                            🌤️ Météo & Conseils
                        </button>
                    </div>
                </div>
            </div>
            
            <div class="greenhouse-grid-container">
                <h2>🏡 Votre Potager (1 case = ${this.cellSize}m)${this.gridRotation > 0 ? ` • Rotation ${this.gridRotation}°` : ''}</h2>
                <p style="color: var(--text-secondary); margin-bottom: 1rem;">
                    Cliquez et faites glisser pour sélectionner plusieurs cases
                </p>
                <div class="greenhouse-grid" id="grid" style="grid-template-columns: repeat(${grid.cols}, 1fr); transform: rotate(${this.gridRotation}deg); transition: transform 0.3s ease;">
                    ${this.cells.map(cell => {
                        let content = '';
                        let className = 'cell';
                        
                        if (cell.type === 'path') {
                            className += ' cell-path';
                            content = '🛤️';
                        } else if (cell.type === 'plot') {
                            const plot = this.plots.find(p => p.id === cell.plotId);
                            if (plot && plot.vegetable) {
                                className += ' cell-planted';
                                const days = this.getDaysSincePlanting(plot.plantedDate);
                                const neighbors = this.checkNeighbors(plot);
                                content = `
                                    <span class="plot-icon">${vegetablesDatabase[plot.vegetable].icon}</span>
                                    ${neighbors.warnings.length > 0 ? '<span class="plot-warning">⚠️</span>' : ''}
                                `;
                            } else {
                                className += ' cell-plot';
                                content = '📦';
                            }
                        } else {
                            className += ' cell-empty';
                            content = '';
                        }
                        
                        return `<div class="${className}" data-cell-id="${cell.id}">${content}</div>`;
                    }).join('')}
                </div>
            </div>
            
            <div class="legend">
                <h3>📖 Légende</h3>
                <div class="legend-items">
                    <div class="legend-item">
                        <div class="legend-icon" style="background: #f1f5f9;"></div>
                        <span>Case vide - Sélectionnez pour créer parcelle/chemin</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-icon" style="background: #92400e;"></div>
                        <span>🛤️ Chemin</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-icon" style="background: #e0e7ff; border: 2px solid #6366f1;"></div>
                        <span>📦 Parcelle vide - Cliquez pour planter</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-icon" style="background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%);"></div>
                        <span>Parcelle cultivée - Cliquez pour détails</span>
                    </div>
                </div>
            </div>
            
            <div style="margin-top: 2rem; text-align: center; padding-bottom: 2rem;">
                <a href="#" id="reset-link" style="color: #94a3b8; font-size: 0.875rem; text-decoration: none; opacity: 0.6; transition: opacity 0.2s;">
                    Réinitialiser l'application
                </a>
            </div>
        `;
        
        // Event listeners pour les modes
        document.querySelectorAll('.mode-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.currentMode = btn.dataset.mode;
                this.selectedCells = [];
                this.render();
            });
        });
        
        // Event listeners pour la grille
        const grid_el = document.getElementById('grid');
        
        // Fonction pour obtenir la cellule sous un point donné (pour le tactile)
        const getCellFromPoint = (x, y) => {
            const elements = document.elementsFromPoint(x, y);
            const cellElement = elements.find(el => el.classList.contains('cell'));
            return cellElement ? cellElement.dataset.cellId : null;
        };
        
        // Variables pour détecter tap vs glissement
        let touchStartPos = null;
        let touchHasMoved = false;
        
        document.querySelectorAll('.cell').forEach(el => {
            // Événements souris (desktop)
            el.addEventListener('mousedown', (e) => {
                e.preventDefault();
                this.handleCellMouseDown(el.dataset.cellId);
            });
            
            el.addEventListener('mouseenter', () => {
                this.handleCellMouseEnter(el.dataset.cellId);
            });
            
            // Événements tactiles (mobile)
            el.addEventListener('touchstart', (e) => {
                e.preventDefault(); // Empêche le scroll
                const touch = e.touches[0];
                touchStartPos = { x: touch.clientX, y: touch.clientY };
                touchHasMoved = false;
                this.handleCellMouseDown(el.dataset.cellId);
            }, { passive: false });
            
            el.addEventListener('touchmove', (e) => {
                if (touchStartPos) {
                    const touch = e.touches[0];
                    const deltaX = Math.abs(touch.clientX - touchStartPos.x);
                    const deltaY = Math.abs(touch.clientY - touchStartPos.y);
                    if (deltaX > 10 || deltaY > 10) {
                        touchHasMoved = true;
                    }
                }
            }, { passive: true });
            
            el.addEventListener('touchend', (e) => {
                // Si pas de mouvement, c'est un tap simple
                if (!touchHasMoved) {
                    // Réinitialiser la sélection pour un tap simple
                    this.isSelecting = false;
                    this.selectedCells = [];
                    
                    const cell = this.cells.find(c => c.id === el.dataset.cellId);
                    if (cell && cell.type === 'plot') {
                        const plot = this.plots.find(p => p.id === cell.plotId);
                        if (plot) {
                            if (plot.vegetable) {
                                this.showPlotDetails(plot);
                            } else {
                                this.showVegetableList(plot);
                            }
                        }
                    }
                }
                touchStartPos = null;
                touchHasMoved = false;
            });
            
            // Click simple sur une parcelle plantée (desktop)
            el.addEventListener('click', () => {
                const cell = this.cells.find(c => c.id === el.dataset.cellId);
                if (cell && cell.type === 'plot') {
                    const plot = this.plots.find(p => p.id === cell.plotId);
                    if (plot) {
                        if (plot.vegetable) {
                            this.showPlotDetails(plot);
                        } else {
                            this.showVegetableList(plot);
                        }
                    }
                }
            });
        });
        
        // Gestion du touchmove sur la grille entière
        grid_el.addEventListener('touchmove', (e) => {
            if (this.isSelecting) {
                e.preventDefault(); // Empêche le scroll pendant la sélection
                const touch = e.touches[0];
                const cellId = getCellFromPoint(touch.clientX, touch.clientY);
                if (cellId) {
                    this.handleCellMouseEnter(cellId);
                }
            }
        }, { passive: false });
        
        // Gestion de la fin de sélection (Desktop)
        document.addEventListener('mouseup', () => {
            this.handleMouseUp();
        });
        
        // Gestion de la fin de sélection (Mobile)
        document.addEventListener('touchend', () => {
            this.handleMouseUp();
        });
        
        document.addEventListener('touchcancel', () => {
            this.handleMouseUp();
        });
        
        // Actions
        const confirmBtn = document.getElementById('confirm-selection');
        if (confirmBtn) {
            confirmBtn.addEventListener('click', () => {
                if (this.currentMode === 'plot') {
                    this.createPlotFromSelection();
                } else if (this.currentMode === 'expand') {
                    this.expandPlotFromSelection();
                } else if (this.currentMode === 'path') {
                    this.markSelectionAsPath();
                }
            });
        }
        
        const clearBtn = document.getElementById('clear-selection');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => {
                this.clearSelection();
            });
        }
        
        // Bouton de rotation de grille
        const rotateBtn = document.getElementById('rotate-grid-btn');
        if (rotateBtn) {
            rotateBtn.addEventListener('click', () => {
                this.rotateGrid();
            });
        }
        
        document.getElementById('reset-link').addEventListener('click', (e) => {
            e.preventDefault();
            this.showResetConfirmation();
        });
        
        // Bouton météo
        const btnWeather = document.getElementById('btn-weather');
        if (btnWeather) {
            btnWeather.addEventListener('click', () => {
                this.showWeatherModal();
            });
        }
        
        // ===== GESTION DU FAB (Mobile uniquement) =====
        this.setupFAB();
    }
    
    // Gestion du FAB (Floating Action Button)
    setupFAB() {
        const fabContainer = document.getElementById('fab-container');
        if (!fabContainer) return;
        
        const isMobile = window.innerWidth <= 768;
        
        // Afficher le FAB uniquement en mode manage ET sur mobile
        if (this.step === 'manage' && isMobile) {
            fabContainer.style.display = 'block';
        } else {
            fabContainer.style.display = 'none';
        }
        
        const fabMain = document.getElementById('fab-main');
        const fabMenu = document.getElementById('fab-menu');
        
        // Toggle du menu
        fabMain.addEventListener('click', (e) => {
            e.stopPropagation();
            fabMain.classList.toggle('active');
            fabMenu.classList.toggle('active');
        });
        
        // Fermer le menu en cliquant ailleurs
        document.addEventListener('click', () => {
            fabMain.classList.remove('active');
            fabMenu.classList.remove('active');
        });
        
        // Empêcher la fermeture en cliquant sur le menu
        fabMenu.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        
        // Boutons de mode (Créer parcelle / Créer chemin)
        document.getElementById('fab-mode-plot').addEventListener('click', () => {
            this.currentMode = 'plot';
            this.updateFABModes();
            this.closeFABMenu();
        });
        
        document.getElementById('fab-mode-path').addEventListener('click', () => {
            this.currentMode = 'path';
            this.updateFABModes();
            this.closeFABMenu();
        });
        
        // Bouton Valider
        document.getElementById('fab-confirm').addEventListener('click', () => {
            if (this.selectedCells.length > 0) {
                if (this.currentMode === 'plot') {
                    this.createPlotFromSelection();
                } else {
                    this.markSelectionAsPath();
                }
                this.closeFABMenu();
            }
        });
        
        // Bouton Effacer
        document.getElementById('fab-clear').addEventListener('click', () => {
            if (this.selectedCells.length > 0) {
                this.clearSelection();
                this.closeFABMenu();
            }
        });
        
        // Bouton Tourner
        document.getElementById('fab-rotate').addEventListener('click', () => {
            this.rotateGrid();
            this.updateFABRotateLabel();
            this.closeFABMenu();
        });
        
        // Initialiser l'état du FAB
        this.updateFABModes();
        this.updateFABSelection();
        this.updateFABRotateLabel();
    }
    
    // Fermer le menu FAB
    closeFABMenu() {
        const fabMain = document.getElementById('fab-main');
        const fabMenu = document.getElementById('fab-menu');
        if (fabMain) fabMain.classList.remove('active');
        if (fabMenu) fabMenu.classList.remove('active');
    }
    
    // Mettre à jour les modes actifs du FAB
    updateFABModes() {
        const plotBtn = document.getElementById('fab-mode-plot');
        const pathBtn = document.getElementById('fab-mode-path');
        
        if (plotBtn && pathBtn) {
            plotBtn.classList.toggle('mode-active', this.currentMode === 'plot');
            pathBtn.classList.toggle('mode-active', this.currentMode === 'path');
        }
    }
    
    // Mettre à jour les boutons de sélection du FAB
    updateFABSelection() {
        const confirmBtn = document.getElementById('fab-confirm');
        const clearBtn = document.getElementById('fab-clear');
        const confirmLabel = document.getElementById('fab-confirm-label');
        
        if (confirmBtn && clearBtn && confirmLabel) {
            const hasSelection = this.selectedCells.length > 0;
            
            confirmBtn.classList.toggle('disabled', !hasSelection);
            clearBtn.classList.toggle('disabled', !hasSelection);
            confirmLabel.textContent = `Valider (${this.selectedCells.length})`;
        }
    }
    
    // Mettre à jour le label de rotation du FAB
    updateFABRotateLabel() {
        const rotateLabel = document.getElementById('fab-rotate-label');
        if (rotateLabel) {
            rotateLabel.textContent = this.gridRotation > 0 
                ? `Tourner (${this.gridRotation}°)` 
                : 'Tourner grille';
        }
    }
    
    showResetConfirmation() {
        const modal = document.createElement('div');
        modal.className = 'modal-overlay';
        modal.innerHTML = `
            <div class="modal" style="max-width: 500px;">
                <div class="modal-header">
                    <h2>⚠️ Confirmation de réinitialisation</h2>
                    <button class="close-btn">×</button>
                </div>
                <div class="modal-content">
                    <div class="alert alert-warning" style="margin-bottom: 1.5rem;">
                        <span>🚨</span>
                        <div>
                            <strong>ATTENTION : Cette action est irréversible !</strong><br>
                            Toutes vos données seront supprimées :
                            <ul style="margin: 0.5rem 0 0 1.5rem; font-size: 0.875rem;">
                                <li>Dimensions du potager</li>
                                <li>Toutes les parcelles créées</li>
                                <li>Tous les légumes plantés</li>
                                <li>Tout l'historique de culture</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div style="margin-bottom: 1rem;">
                        <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: var(--text);">
                            Pour confirmer, tapez <strong style="color: #dc2626;">EFFACER</strong> en majuscules :
                        </label>
                        <input 
                            type="text" 
                            id="reset-confirm-input" 
                            placeholder="Tapez EFFACER" 
                            style="width: 100%; padding: 0.75rem; border: 2px solid #e5e7eb; border-radius: 0.5rem; font-size: 1rem; font-family: monospace;"
                            autocomplete="off"
                        >
                        <small id="reset-input-feedback" style="display: block; margin-top: 0.5rem; color: #6b7280; font-size: 0.875rem;">
                            Le bouton sera activé quand vous tapez correctement
                        </small>
                    </div>
                </div>
                <div class="modal-actions">
                    <button class="btn btn-secondary" id="cancel-reset-btn">Annuler</button>
                    <button class="btn btn-danger" id="confirm-reset-btn" disabled style="opacity: 0.5; cursor: not-allowed;">
                        🗑️ Tout effacer
                    </button>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        const input = modal.querySelector('#reset-confirm-input');
        const confirmBtn = modal.querySelector('#confirm-reset-btn');
        const feedback = modal.querySelector('#reset-input-feedback');
        
        const closeModal = () => {
            document.body.removeChild(modal);
        };
        
        // Vérification en temps réel
        input.addEventListener('input', () => {
            const value = input.value;
            
            if (value === 'EFFACER') {
                confirmBtn.disabled = false;
                confirmBtn.style.opacity = '1';
                confirmBtn.style.cursor = 'pointer';
                feedback.textContent = '✅ Vous pouvez maintenant cliquer sur "Tout effacer"';
                feedback.style.color = '#059669';
                input.style.borderColor = '#059669';
            } else {
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = '0.5';
                confirmBtn.style.cursor = 'not-allowed';
                
                if (value.length > 0) {
                    feedback.textContent = '❌ Incorrect. Tapez exactement "EFFACER" en majuscules';
                    feedback.style.color = '#dc2626';
                    input.style.borderColor = '#dc2626';
                } else {
                    feedback.textContent = 'Le bouton sera activé quand vous tapez correctement';
                    feedback.style.color = '#6b7280';
                    input.style.borderColor = '#e5e7eb';
                }
            }
        });
        
        modal.querySelector('.close-btn').addEventListener('click', closeModal);
        modal.querySelector('#cancel-reset-btn').addEventListener('click', closeModal);
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
        
        confirmBtn.addEventListener('click', () => {
            if (input.value === 'EFFACER') {
                // Réinitialiser tout
                localStorage.clear(); // Effacer tout le localStorage
                this.step = 'dimensions';
                this.cells = [];
                this.plots = [];
                this.history = {};
                this.nextPlotId = 1;
                this.greenhouseDimensions = { width: 6.5, length: 10, height: 4 };
                this.saveState();
                closeModal();
                this.render();
            }
        });
        
        // Focus automatique sur le champ
        setTimeout(() => input.focus(), 100);
    }
    
    showVegetableList(plot) {
        const modal = document.createElement('div');
        modal.className = 'modal-overlay';
        
        // Détecter les légumes voisins pour suggérer les bons compagnons
        const neighborPlots = this.getNeighborPlots(plot);
        const neighborVegetables = neighborPlots
            .filter(p => p.vegetable)
            .map(p => p.vegetable);
        
        modal.innerHTML = `
            <div class="modal">
                <div class="modal-header">
                    <h2>🌱 Choisissez un légume</h2>
                    <button class="close-btn">×</button>
                </div>
                <div class="modal-content">
                    ${neighborVegetables.length > 0 ? `
                        <div class="alert alert-info" style="margin-bottom: 1rem;">
                            <span>👉</span>
                            <div>
                                <strong>Parcelles voisines :</strong><br>
                                ${neighborVegetables.map(v => vegetablesDatabase[v].icon + ' ' + vegetablesDatabase[v].name).join(', ')}<br>
                                <small>Les bons compagnons sont affichés en premier dans chaque catégorie</small>
                            </div>
                        </div>
                    ` : ''}
                    <div class="vegetable-search">
                        <input type="text" placeholder="🔍 Rechercher..." id="search-input">
                    </div>
                    <div class="vegetables-list" id="vegetables-list">
                        ${this.renderVegetablesList('', neighborVegetables, plot.id)}
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        const closeModal = () => {
            document.body.removeChild(modal);
        };
        
        modal.querySelector('.close-btn').addEventListener('click', closeModal);
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
        
        const searchInput = modal.querySelector('#search-input');
        searchInput.addEventListener('input', (e) => {
            const list = modal.querySelector('#vegetables-list');
            list.innerHTML = this.renderVegetablesList(e.target.value, neighborVegetables, plot.id);
            
            list.querySelectorAll('.vegetable-item').forEach(item => {
                item.addEventListener('click', () => {
                    this.plantVegetable(plot.id, item.dataset.vegKey);
                    closeModal();
                });
            });
            
            // Ré-ajouter les event listeners de catégories
            list.querySelectorAll('.category-header').forEach(header => {
                header.addEventListener('click', () => {
                    const section = header.closest('.category-section');
                    section.classList.toggle('category-collapsed');
                });
            });
        });
        
        modal.querySelectorAll('.vegetable-item').forEach(item => {
            item.addEventListener('click', () => {
                this.plantVegetable(plot.id, item.dataset.vegKey);
                closeModal();
            });
        });
        
        // Gestion du repli/dépli des catégories
        modal.querySelectorAll('.category-header').forEach(header => {
            header.addEventListener('click', () => {
                const section = header.closest('.category-section');
                section.classList.toggle('category-collapsed');
            });
        });
    }
    
    renderVegetablesList(searchTerm, neighborVegetables = [], plotId = null) {
        // Filtrer par terme de recherche
        let filtered = Object.entries(vegetablesDatabase)
            .filter(([key, veg]) => 
                veg.name.toLowerCase().includes(searchTerm.toLowerCase())
            );
        
        // Calculer les scores de rotation si on a un plotId
        const rotationScores = {};
        if (plotId) {
            filtered.forEach(([key, veg]) => {
                rotationScores[key] = this.getRotationScore(key, plotId);
            });
        }
        
        // Trier par priorité si on a des voisins
        if (neighborVegetables.length > 0) {
            filtered = filtered.sort(([keyA, vegA], [keyB, vegB]) => {
                let scoreA = 0;
                let scoreB = 0;
                
                neighborVegetables.forEach(neighborKey => {
                    const neighbor = vegetablesDatabase[neighborKey];
                    if (neighbor) {
                        if (neighbor.goodCompanions.includes(keyA)) scoreA += 10;
                        if (neighbor.goodCompanions.includes(keyB)) scoreB += 10;
                        if (neighbor.badCompanions.includes(keyA)) scoreA -= 10;
                        if (neighbor.badCompanions.includes(keyB)) scoreB -= 10;
                    }
                });
                
                return scoreB - scoreA;
            });
        }
        
        // Trier par score de rotation si disponible
        if (plotId) {
            filtered = filtered.sort(([keyA], [keyB]) => {
                const scoreA = rotationScores[keyA]?.score || 0;
                const scoreB = rotationScores[keyB]?.score || 0;
                return scoreB - scoreA; // Meilleur score en premier
            });
        }
        
        // Grouper par catégorie
        const categories = {};
        filtered.forEach(([key, veg]) => {
            const cat = veg.category || 'autres';
            if (!categories[cat]) {
                categories[cat] = [];
            }
            categories[cat].push([key, veg]);
        });
        
        // Noms des catégories en français
        const categoryNames = {
            tomate: '🍅 Tomates',
            carotte: '🥕 Carottes',
            salade: '🥬 Salades',
            haricot: '🫘 Haricots',
            courge: '🎃 Courges',
            herbe: '🌿 Herbes aromatiques',
            oignon: '🧅 Alliacées (oignons, ail)',
            autres: '🌱 Autres légumes'
        };
        
        // Rendre le HTML par catégorie
        let html = '';
        
        Object.keys(categories).sort().forEach(cat => {
            const veggies = categories[cat];
            const catName = categoryNames[cat] || cat;
            
            html += `
                <div class="category-section" data-category="${cat}">
                    <div class="category-header">
                        <span>${catName}</span>
                        <span style="font-size: 0.875rem; opacity: 0.8;">${veggies.length} variété(s)</span>
                    </div>
                    <div class="category-items">
                        ${veggies.map(([key, veg]) => {
                            // Déterminer si c'est un bon ou mauvais compagnon
                            let companionBadge = '';
                            let companionBadgeClass = '';
                            
                            if (neighborVegetables.length > 0) {
                                let isGoodCompanion = false;
                                let isBadCompanion = false;
                                
                                neighborVegetables.forEach(neighborKey => {
                                    const neighbor = vegetablesDatabase[neighborKey];
                                    if (neighbor) {
                                        if (neighbor.goodCompanions.includes(key)) isGoodCompanion = true;
                                        if (neighbor.badCompanions.includes(key)) isBadCompanion = true;
                                    }
                                });
                                
                                if (isGoodCompanion) {
                                    companionBadge = '✓ Bon compagnon';
                                    companionBadgeClass = 'badge-good';
                                } else if (isBadCompanion) {
                                    companionBadge = '✗ Mauvais compagnon';
                                    companionBadgeClass = 'badge-bad';
                                }
                            }
                            
                            // Badge de rotation
                            let rotationBadge = '';
                            if (plotId && rotationScores[key]) {
                                const rotation = rotationScores[key];
                                rotationBadge = `<span class="rotation-badge rotation-${rotation.score}" title="${rotation.text}">${rotation.badge} ${rotation.text}</span>`;
                            }
                            
                            // Badge de hauteur (alerte si plante trop haute)
                            let heightBadge = '';
                            if (veg.maxHeight && veg.maxHeight > this.greenhouseDimensions.height) {
                                const diff = (veg.maxHeight - this.greenhouseDimensions.height).toFixed(1);
                                heightBadge = `<span class="planting-badge badge-warning" title="Cette plante peut atteindre ${veg.maxHeight}m, soit ${diff}m de plus que votre hauteur disponible (${this.greenhouseDimensions.height}m)">⚠️ Trop haute (${veg.maxHeight}m)</span>`;
                            }
                            
                            return `
                                <div class="vegetable-item ${companionBadgeClass ? 'has-badge' : ''}" data-veg-key="${key}">
                                    <div class="vegetable-item-header">
                                        <span class="vegetable-item-icon">${veg.icon}</span>
                                        <div style="flex: 1;">
                                            <span class="vegetable-item-name">${veg.name}</span>
                                            ${companionBadge ? `<span class="companion-badge ${companionBadgeClass}">${companionBadge}</span>` : ''}
                                            ${rotationBadge}
                                            ${heightBadge}
                                        </div>
                                    </div>
                                    <div class="vegetable-item-info">
                                        <div>⏱️ ${veg.growthDays} jours</div>
                                        <div>💧 ${veg.waterNeeds}</div>
                                        ${veg.family ? `<div style="font-size: 0.75rem; opacity: 0.7;">${veg.family}</div>` : ''}
                                    </div>
                                </div>
                            `;
                        }).join('')}
                    </div>
                </div>
            `;
        });
        
        return html;
    }
    
    showPlotDetails(plot) {
        const veg = vegetablesDatabase[plot.vegetable];
        const days = this.getDaysSincePlanting(plot.plantedDate);
        const remaining = veg.growthDays - days;
        const neighbors = this.checkNeighbors(plot);
        const plotSize = (plot.cellIds.length * this.cellSize * this.cellSize).toFixed(2);
        
        const modal = document.createElement('div');
        modal.className = 'modal-overlay';
        modal.innerHTML = `
            <div class="modal">
                <div class="modal-header">
                    <h2>${veg.icon} ${veg.name}</h2>
                    <button class="close-btn">×</button>
                </div>
                <div class="modal-content">
                    <div class="alert alert-info">
                        <span>📅</span>
                        <div>
                            <strong>Parcelle de ${plotSize} m²</strong> (${plot.cellIds.length} cases)<br>
                            <strong>${days} jours</strong> depuis plantation<br>
                            <strong>${remaining > 0 ? remaining : 0} jours</strong> avant récolte
                            ${remaining <= 0 ? '<br><strong style="color: var(--primary)">Prêt ! 🎉</strong>' : ''}
                        </div>
                    </div>
                    
                    ${neighbors.warnings.length > 0 ? `
                        <div class="alert alert-warning">
                            <span>⚠️</span>
                            <div>
                                <strong>Mauvais voisinage !</strong><br>
                                ${neighbors.warnings.join(', ')}
                            </div>
                        </div>
                    ` : ''}
                    
                    ${neighbors.suggestions.length > 0 ? `
                        <div class="alert alert-success">
                            <span>💡</span>
                            <div>
                                <strong>Suggestions :</strong><br>
                                ${neighbors.suggestions.map(k => vegetablesDatabase[k]?.name || k).join(', ')}
                            </div>
                        </div>
                    ` : ''}
                    
                    <div class="detail-section">
                        <h3>💧 Besoins en eau</h3>
                        <p class="detail-text"><span class="detail-value">${veg.waterNeeds}</span></p>
                    </div>
                    
                    <div class="detail-section">
                        <h3>☀️ Ensoleillement</h3>
                        <p class="detail-text"><span class="detail-value">${veg.sunlight}</span></p>
                    </div>
                    
                    <div class="detail-section">
                        <h3>🌱 Sol</h3>
                        <p class="detail-text"><span class="detail-value">${veg.soilType}</span></p>
                    </div>
                    
                    ${veg.tips ? `
                        <div class="detail-section">
                            <h3>💡 Astuce</h3>
                            <p class="detail-text">${veg.tips}</p>
                        </div>
                    ` : ''}
                </div>
                <div class="modal-actions">
                    <button class="btn btn-danger" id="remove-btn">🗑️ Retirer</button>
                    <button class="btn btn-secondary" id="close-modal-btn">Fermer</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        const closeModal = () => {
            document.body.removeChild(modal);
        };
        
        modal.querySelector('.close-btn').addEventListener('click', closeModal);
        modal.querySelector('#close-modal-btn').addEventListener('click', closeModal);
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
        
        modal.querySelector('#remove-btn').addEventListener('click', () => {
            this.removePlant(plot.id);
            closeModal();
        });
    }
    
    async showWeatherModal() {
        // IMPORTANT : Fermer tous les modals existants d'abord
        document.querySelectorAll('.modal-overlay').forEach(m => m.remove());
        
        const modal = document.createElement('div');
        modal.className = 'modal-overlay';
        
        // Afficher un loader pendant le chargement
        modal.innerHTML = `
            <div class="modal">
                <div class="modal-header">
                    <h2>🌤️ Météo & Conseils Plantation</h2>
                    <button class="close-btn" id="close-weather-btn">×</button>
                </div>
                <div class="modal-content" style="text-align: center; padding: 3rem;">
                    <div style="font-size: 3rem; margin-bottom: 1rem;">🌤️</div>
                    <p style="font-size: 1.25rem; color: var(--text-secondary);">Chargement des données météo...</p>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        try {
            // Récupérer les conditions météo
            const conditions = await weatherAPI.getPlantingConditions();
            const allAdvice = await plantingAdvisor.getAllPlantingAdvice();
            
            // Créer la vue météo complète
            modal.innerHTML = `
                <div class="modal" style="max-width: 900px;">
                    <div class="modal-header">
                        <h2>🌤️ Météo & Conseils Plantation</h2>
                        <button class="close-btn" id="close-weather-btn">×</button>
                    </div>
                    <div class="modal-content">
                        <!-- Widget météo -->
                        <div class="weather-widget">
                            <div class="weather-current">
                                <img src="https://openweathermap.org/img/wn/${conditions.current.icon}@2x.png" 
                                     alt="${conditions.current.description}"
                                     style="width: 80px; height: 80px;">
                                <div style="flex: 1;">
                                    <h3 style="font-size: 1.5rem; margin-bottom: 0.25rem; display: flex; align-items: center; gap: 0.5rem;">
                                        📍 ${conditions.current.city}
                                        <button id="change-city-btn" style="background: rgba(255,255,255,0.2); border: none; color: white; padding: 0.25rem 0.75rem; border-radius: 0.5rem; font-size: 0.875rem; cursor: pointer; font-weight: 600;">
                                            📍 Changer
                                        </button>
                                    </h3>
                                    <p class="weather-temp">${conditions.current.temp}°C</p>
                                    <p style="text-transform: capitalize; opacity: 0.9;">
                                        ${conditions.current.description}
                                    </p>
                                </div>
                            </div>
                            
                            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-top: 1rem;">
                                <div>
                                    <div style="opacity: 0.8; font-size: 0.875rem;">Ressenti</div>
                                    <div style="font-weight: 700; font-size: 1.25rem;">${conditions.current.feelsLike}°C</div>
                                </div>
                                <div>
                                    <div style="opacity: 0.8; font-size: 0.875rem;">Sol estimé</div>
                                    <div style="font-weight: 700; font-size: 1.25rem;">${conditions.soilTemp}°C</div>
                                </div>
                                <div>
                                    <div style="opacity: 0.8; font-size: 0.875rem;">Humidité</div>
                                    <div style="font-weight: 700; font-size: 1.25rem;">${conditions.current.humidity}%</div>
                                </div>
                            </div>
                            
                            <!-- Prévisions 7 jours -->
                            <div class="weather-forecast">
                                ${conditions.forecast.map(day => `
                                    <div class="weather-day">
                                        <div style="font-weight: 600; margin-bottom: 0.5rem;">${day.dayName.substring(0, 3)}</div>
                                        <img src="https://openweathermap.org/img/wn/${day.icon}.png" 
                                             alt="${day.description}"
                                             style="width: 40px; height: 40px;">
                                        <div style="font-weight: 700;">${day.tempMax}°</div>
                                        <div style="opacity: 0.8; font-size: 0.875rem;">${day.tempMin}°</div>
                                        ${day.frostRisk ? '<div style="margin-top: 0.25rem;">❄️</div>' : ''}
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                        
                        <!-- Conseils de plantation -->
                        <div class="planting-advice">
                            ${this.renderPlantingAdvice(allAdvice, conditions)}
                        </div>
                    </div>
                </div>
            `;
            
        } catch (error) {
            console.error('Erreur chargement météo:', error);
            modal.innerHTML = `
                <div class="modal">
                    <div class="modal-header">
                        <h2>🌤️ Météo & Conseils</h2>
                        <button class="close-btn" id="close-weather-btn">×</button>
                    </div>
                    <div class="modal-content">
                        <div class="alert alert-warning">
                            <span style="font-size: 2rem;">⚠️</span>
                            <div>
                                <strong>Impossible de charger la météo</strong>
                                <p style="margin-top: 0.5rem;">Vérifiez votre connexion internet.</p>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        }
        
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.remove();
                document.querySelectorAll('.modal-overlay').forEach(m => m.remove());
            }
        });
        
        // Event listener pour fermer avec la croix
        const closeBtn = modal.querySelector('#close-weather-btn');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                modal.remove();
                document.querySelectorAll('.modal-overlay').forEach(m => m.remove());
            });
        }
        
        // Event listener pour changer de ville
        const changeCityBtn = modal.querySelector('#change-city-btn');
        if (changeCityBtn) {
            changeCityBtn.addEventListener('click', () => {
                modal.remove();
                this.showChangeCityModal();
            });
        }
    }
    
    renderPlantingAdvice(allAdvice, conditions) {
        const advice = allAdvice.advice;
        
        // Grouper par statut
        const canPlant = [];
        const needProtection = [];
        const cannotPlant = [];
        
        for (const [key, adv] of Object.entries(advice)) {
            const veg = vegetablesDatabase[key];
            if (!veg) continue;
            
            if (adv.status === 'perfect' || adv.status === 'good') {
                canPlant.push({ key, veg, adv });
            } else if (adv.status === 'warning') {
                needProtection.push({ key, veg, adv });
            } else if (adv.status === 'cold' || adv.status === 'danger') {
                cannotPlant.push({ key, veg, adv });
            }
        }
        
        let html = '<h2 style="color: var(--earth-dark); margin-bottom: 1.5rem;">Conseils de Plantation</h2>';
        
        // Alertes gel
        if (conditions.frostRisk) {
            html += `
                <div class="alert alert-warning">
                    <span style="font-size: 2rem;">❄️</span>
                    <div>
                        <strong>ALERTE GEL</strong>
                        <p style="margin-top: 0.5rem;">
                            Risque de gelée dans les ${conditions.frostRiskDays} prochains jours !
                            Protégez vos plants sensibles.
                        </p>
                    </div>
                </div>
            `;
        }
        
        // Section "Vous pouvez planter"
        if (canPlant.length > 0) {
            html += `
                <div class="advice-section">
                    <h3>🟢 Vous pouvez planter (${canPlant.length})</h3>
                    <ul class="advice-list">
                        ${canPlant.slice(0, 8).map(item => `
                            <li class="can-plant">
                                <span style="font-size: 1.5rem;">${item.veg.icon}</span>
                                <div style="flex: 1;">
                                    <strong>${item.veg.name}</strong>
                                    <div style="font-size: 0.875rem; opacity: 0.9;">${item.adv.details}</div>
                                </div>
                            </li>
                        `).join('')}
                    </ul>
                    ${canPlant.length > 8 ? `<p style="opacity: 0.7; font-size: 0.875rem; margin-top: 0.5rem;">... et ${canPlant.length - 8} autres</p>` : ''}
                </div>
            `;
        }
        
        // Section "Avec protection"
        if (needProtection.length > 0) {
            html += `
                <div class="advice-section">
                    <h3>🟡 Possible avec protection (${needProtection.length})</h3>
                    <ul class="advice-list">
                        ${needProtection.slice(0, 6).map(item => `
                            <li class="with-protection">
                                <span style="font-size: 1.5rem;">${item.veg.icon}</span>
                                <div style="flex: 1;">
                                    <strong>${item.veg.name}</strong>
                                    <div style="font-size: 0.875rem; opacity: 0.9;">${item.adv.protection}</div>
                                </div>
                            </li>
                        `).join('')}
                    </ul>
                    ${needProtection.length > 6 ? `<p style="opacity: 0.7; font-size: 0.875rem; margin-top: 0.5rem;">... et ${needProtection.length - 6} autres</p>` : ''}
                </div>
            `;
        }
        
        // Section "Trop froid"
        if (cannotPlant.length > 0) {
            html += `
                <div class="advice-section">
                    <h3>🔴 Trop froid (${cannotPlant.length})</h3>
                    <ul class="advice-list">
                        ${cannotPlant.slice(0, 6).map(item => `
                            <li class="cannot-plant">
                                <span style="font-size: 1.5rem;">${item.veg.icon}</span>
                                <div style="flex: 1;">
                                    <strong>${item.veg.name}</strong>
                                    <div style="font-size: 0.875rem; opacity: 0.9;">
                                        ${item.adv.daysToWait ? `Attendez ~${Math.ceil(item.adv.daysToWait / 7)} semaine(s)` : item.adv.message}
                                    </div>
                                </div>
                            </li>
                        `).join('')}
                    </ul>
                    ${cannotPlant.length > 6 ? `<p style="opacity: 0.7; font-size: 0.875rem; margin-top: 0.5rem;">... et ${cannotPlant.length - 6} autres</p>` : ''}
                </div>
            `;
        }
        
        return html;
    }
    
    showChangeCityModal() {
        const modal = document.createElement('div');
        modal.className = 'modal-overlay';
        modal.innerHTML = `
            <div class="modal" style="max-width: 500px;">
                <div class="modal-header">
                    <h2>📍 Changer la localisation</h2>
                    <button class="close-btn" id="close-city-btn">×</button>
                </div>
                <div class="modal-content">
                    <div style="margin-bottom: 1.5rem;">
                        <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">
                            Ville actuelle : <strong>${weatherAPI.currentCity}</strong>
                        </label>
                    </div>
                    
                    <div style="margin-bottom: 1.5rem;">
                        <button class="btn" id="geoloc-btn" style="width: 100%; margin: 0;">
                            📡 Utiliser ma position actuelle
                        </button>
                        <small style="display: block; margin-top: 0.5rem; color: var(--text-secondary); font-size: 0.875rem;">
                            Votre navigateur demandera l'autorisation
                        </small>
                    </div>
                    
                    <div style="margin: 1.5rem 0; text-align: center; color: var(--text-secondary); font-weight: 600;">
                        ─── OU ───
                    </div>
                    
                    <div>
                        <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">
                            Entrer une ville manuellement :
                        </label>
                        <input 
                            type="text" 
                            id="city-input" 
                            placeholder="Ex: Paris, Lyon, Marseille..."
                            value="${weatherAPI.currentCity}"
                            style="width: 100%; padding: 0.75rem; border: 2px solid var(--border); border-radius: 0.5rem; font-size: 1rem; margin-bottom: 1rem;"
                        >
                        <button class="btn btn-secondary" id="manual-city-btn" style="width: 100%; margin: 0;">
                            ✓ Valider cette ville
                        </button>
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        const closeModal = () => {
            document.body.removeChild(modal);
        };
        
        // Event listener pour fermer avec la croix
        const closeBtn = modal.querySelector('#close-city-btn');
        if (closeBtn) {
            closeBtn.addEventListener('click', closeModal);
        }
        
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
        
        // Géolocalisation
        const geolocBtn = modal.querySelector('#geoloc-btn');
        geolocBtn.addEventListener('click', async () => {
            geolocBtn.disabled = true;
            geolocBtn.textContent = '🔄 Détection en cours...';
            
            try {
                const city = await weatherAPI.getCityFromGeolocation();
                closeModal();
                // Fermer la modal météo actuelle et en ouvrir une nouvelle
                document.querySelectorAll('.modal-overlay').forEach(m => m.remove());
                this.showWeatherModal();
            } catch (error) {
                console.error('Erreur géolocalisation:', error);
                alert('❌ Impossible de détecter votre position.\n' + error.message);
                geolocBtn.disabled = false;
                geolocBtn.textContent = '📡 Utiliser ma position actuelle';
            }
        });
        
        // Ville manuelle
        const cityInput = modal.querySelector('#city-input');
        const manualBtn = modal.querySelector('#manual-city-btn');
        
        manualBtn.addEventListener('click', () => {
            const city = cityInput.value.trim();
            if (city) {
                weatherAPI.saveCity(city);
                closeModal();
                // Fermer la modal météo actuelle et en ouvrir une nouvelle
                document.querySelectorAll('.modal-overlay').forEach(m => m.remove());
                this.showWeatherModal();
            }
        });
        
        // Validation par Enter
        cityInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                manualBtn.click();
            }
        });
    }
}

// Démarrage
document.addEventListener('DOMContentLoaded', () => {
    const app = new GreenhouseApp();
    
    // Initialiser le système d'aide
    const helper = new HelpSystem(app);
    
    // Stocker globalement pour accès depuis initializeCells
    window.helpSystem = helper;
    
    // Lancer automatiquement au premier lancement
    if (!localStorage.getItem('greenhouse_tutorialCompleted')) {
        // Attendre que l'app soit rendue
        setTimeout(() => {
            helper.start();
        }, 500);
    }
    
    // Bouton d'aide
    document.getElementById('help-button')?.addEventListener('click', () => {
        helper.start();
    });
});

// ========================================
// SYSTÈME D'AIDE (MOBILE + DESKTOP)
// ========================================

class HelpSystem {
    constructor(app) {
        this.app = app;
        this.welcomeMessages = new WelcomeMessages(app);
    }
    
    start() {
        this.welcomeMessages.show();
    }
}

// ========================================
// MESSAGES DE BIENVENUE (MOBILE + DESKTOP)
// ========================================

class WelcomeMessages {
    constructor(app) {
        this.app = app;
        this.overlay = document.getElementById('welcome-overlay');
        this.content = document.getElementById('welcome-content');
    }
    
    show() {
        if (this.app.step === 'dimensions') {
            this.showDimensionsMessage();
        } else {
            this.showManageMessage();
        }
    }
    
    showDimensionsMessage() {
        this.content.innerHTML = `
            <h2>👋 Bienvenue dans Garden-Helper !</h2>
            <p><strong>Optimisez votre potager</strong> avec la rotation des cultures et le compagnonnage.</p>
            <p>📏 <strong>Première étape :</strong> Entrez les dimensions de votre potager en mètres.</p>
            <ul>
                <li><strong>Largeur & Longueur</strong> : Dimensions de votre espace de culture</li>
                <li><strong>Hauteur</strong> : Utile pour les serres, balcons ou espaces intérieurs (calcul du volume)</li>
            </ul>
            <p>L'application créera une grille adaptée à votre espace !</p>
            <div class="welcome-modal-footer">
                <button class="welcome-btn" id="welcome-btn-ok">Compris ! →</button>
            </div>
        `;
        
        this.overlay.classList.add('active');
        
        document.getElementById('welcome-btn-ok').addEventListener('click', () => {
            this.close();
        });
    }
    
    showManageMessage() {
        const isMobile = window.innerWidth <= 768;
        
        if (isMobile) {
            // Message mobile compact
            this.content.innerHTML = `
                <h2>🌿 Votre potager est prêt !</h2>
                <p><strong>Voici comment utiliser l'application :</strong></p>
                <ul>
                    <li><strong>📦 Créer des parcelles :</strong> Utilisez le bouton flottant 🏛️ en bas, sélectionnez "Créer parcelle", glissez votre doigt sur la grille, puis validez</li>
                    <li><strong>🌱 Planter :</strong> Tapez sur une parcelle vide pour choisir un légume. L'app vous suggère les meilleurs compagnons</li>
                    <li><strong>🛤️ Créer des chemins :</strong> Même principe, choisissez "Créer chemin" dans le bouton 🏛️</li>
                    <li><strong>🌤️ Météo :</strong> Consultez en haut pour savoir quoi planter selon la température</li>
                    <li><strong>🔄 Rotation :</strong> Tournez la vue de votre potager dans le menu 🏛️</li>
                </ul>
                <p><small>Astuce : Pour relancer ce guide, appuyez sur le bouton <strong>?</strong> en bas à droite.</small></p>
            `;
        } else {
            // Message desktop détaillé
            this.content.innerHTML = `
                <h2>🌿 Votre potager est prêt !</h2>
                <p><strong>Voici comment utiliser l'application :</strong></p>
                
                <h3 style="margin-top: 1.5rem; margin-bottom: 0.75rem; font-size: 1.1rem;">🎯 Créer des parcelles et chemins</h3>
                <ul>
                    <li><strong>Sélection :</strong> Cliquez et glissez sur la grille pour sélectionner plusieurs cases</li>
                    <li><strong>Mode Parcelle :</strong> Choisissez "Créer une parcelle" puis "Valider" pour créer une zone de culture</li>
                    <li><strong>Mode Chemin :</strong> Choisissez "Créer un chemin" pour les allées entre parcelles</li>
                </ul>
                
                <h3 style="margin-top: 1.5rem; margin-bottom: 0.75rem; font-size: 1.1rem;">🌱 Planter et gérer vos cultures</h3>
                <ul>
                    <li><strong>Plantation :</strong> Cliquez sur une parcelle vide (📦) pour choisir un légume</li>
                    <li><strong>Compagnonnage :</strong> L'app vous suggère les meilleurs voisins selon vos cultures existantes</li>
                    <li><strong>Rotation :</strong> Des badges indiquent si la rotation est optimale (🟢), acceptable (🟡) ou à éviter (🔴)</li>
                    <li><strong>Détails :</strong> Cliquez sur une parcelle plantée pour voir les informations complètes</li>
                </ul>
                
                <h3 style="margin-top: 1.5rem; margin-bottom: 0.75rem; font-size: 1.1rem;">🛠️ Autres fonctionnalités</h3>
                <ul>
                    <li><strong>🌤️ Météo & Conseils :</strong> Consultez quels légumes planter selon la température actuelle</li>
                    <li><strong>🔄 Rotation vue :</strong> Tournez la grille par pas de 90° pour changer de perspective</li>
                    <li><strong>? Aide :</strong> Cliquez sur le bouton en bas à droite pour relancer ce guide</li>
                </ul>
            `;
        }
        
        this.content.innerHTML += `
            <div class="welcome-modal-footer">
                <button class="welcome-btn" id="welcome-btn-start">C'est parti ! 🚀</button>
            </div>
        `;
        
        this.overlay.classList.add('active');
        
        document.getElementById('welcome-btn-start').addEventListener('click', () => {
            this.close();
            localStorage.setItem('greenhouse_tutorialCompleted', 'true');
        });
    }
    
    close() {
        this.overlay.classList.remove('active');
    }
}

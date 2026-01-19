# 🌿 Garden-Helper - README Rapide

**Version :** 1.4.0 | **Statut :** ✅ Production  
**Documentation complète :** [docs/GUIDE-PROJET.md](docs/GUIDE-PROJET.md)

---

## 🚀 Démarrage Rapide

### Ouvrir en local
```
Ouvrir index.html dans un navigateur
```

### Déployer sur Netlify
```bash
1. Drag & Drop du dossier sur app.netlify.com
2. Attendre 30-60 secondes
3. Tester sur mobile : Ctrl+Shift+R
```

---

## ⚠️ IMPORTANT AVANT TOUTE MODIFICATION

### 1️⃣ Incrémenter la version du Service Worker
```javascript
// sw.js - TOUJOURS modifier ceci :
const CACHE_NAME = 'garden-helper-v1.X.0'; // +1 à chaque modif !
```

### 2️⃣ Vérifier après modification
- [ ] Tester en local (ouvrir index.html)
- [ ] Console sans erreurs (F12)
- [ ] Upload sur Netlify
- [ ] Force refresh (Ctrl+Shift+R)

---

## 📁 Structure Fichiers Essentiels

```
Garden-Helper/
├── index.html              ← Page principale + CSS
├── app.js                  ← Logique application
├── vegetables-data.js      ← 50+ légumes
├── sw.js                   ← Service Worker (cache PWA)
├── manifest.json           ← Config PWA
└── docs/
    └── GUIDE-PROJET.md     ← Documentation complète
```

---

## 🔧 Checklist Modifications

### Ajouter un légume
```javascript
// vegetables-data.js
nom_legume: {
    name: "Nom Français",
    icon: "🥕",  // ⚠️ Unicode 6.0-7.0 UNIQUEMENT !
    family: "famille",
    // ... suivre structure existante
}
```

### Modifier le CSS
```css
/* index.html - Section <style> ligne 11-770 */
⚠️ Responsive : 3 breakpoints (768px, 480px, paysage)
```

### Changer couleurs
```css
/* Variables CSS ligne 15-30 */
:root {
    --primary: #56ab2f;    /* Vert principal */
    --earth: #8b7355;      /* Marron terre */
    /* ... */
}
```

---

## ❌ NE JAMAIS MODIFIER

1. **Préfixe localStorage** : `greenhouse_` (perte données utilisateurs)
2. **Structure vegetables-data.js** : Clés name, icon, family, etc.
3. **Cases < 44px mobile** : Standard tactile iOS/Android
4. **Emojis récents** : Seulement Unicode 6.0-7.0 (2010-2014)

---

## 💡 Améliorations Prioritaires

### 🟢 HAUTE
1. Calcul automatique nb plants par parcelle
2. Notifications rappels arrosage/récolte
3. Export/Import données JSON

### 🟡 MOYENNE
4. Recherche légumes avec filtres
5. Notes personnelles par parcelle
6. Statistiques globales (kg récoltés)
7. Mode sombre

### 🔵 BASSE
8. IA détection maladies (photo)
9. Mode collaboratif (Firebase)
10. Gamification (badges)

---

## 📂 Chemin Complet

```
C:\Users\Valentin\Scopi\Support - Documents\apprentissage site\15 défis claude\Garden-Helper\
```

---

## 📞 Support

**Bug ?** Vérifier :
1. Version Service Worker (sw.js ligne 2)
2. Console navigateur (F12)
3. localStorage : `localStorage.clear()` si problème

**Feedback utilisateurs :**  
https://docs.google.com/forms/d/e/1FAIpQLSeCetWD6aKtKYD-EHDAYIIXxSHpYt-TR5a__yQmpz5wQ2yXSg/viewform

---

🌿 **Pour plus de détails → [GUIDE-PROJET.md](docs/GUIDE-PROJET.md)** 🌿

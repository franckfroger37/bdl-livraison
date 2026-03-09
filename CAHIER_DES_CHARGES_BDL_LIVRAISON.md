# CAHIER DES CHARGES — BDL-LIVRAISON v1.2.0

## 1. PRÉSENTATION GÉNÉRALE

### 1.1 Contexte
Application web progressive (PWA) pour la gestion des tournées de livraison de linge pour Bulle de Linge (BDL). L'application permet aux chauffeurs de :
- Charger un fichier Excel de tournée depuis Google Drive
- Organiser l'ordre de livraison
- Suivre les interventions chez chaque client
- Prendre des photos et saisir des commentaires
- Générer un rapport PDF avec envoi par email

### 1.2 Technologies
- **Frontend** : React (Create React App)
- **UI** : Lucide React pour les icônes, styles inline (pas de CSS externe)
- **Excel** : XLSX (SheetJS) pour la lecture des fichiers
- **PDF** : jsPDF + jsPDF-autoTable
- **Déploiement** : GitHub Pages
- **Type** : PWA installable (manifest.json, service worker)

### 1.3 Identité visuelle
- **Logo** : Bulle de Linge (LOGO_BDL en base64 JPEG)
- **Couleurs principales** :
  - Bleu primaire : `#2563eb` (boutons, titres)
  - Vert validation : `#16a34a`
  - Jaune avertissement : `#f59e0b`
  - Fond gris clair : `#f3f4f6`

---

## 2. ARCHITECTURE DE L'APPLICATION

### 2.1 Structure des vues

```
VueLogin (écran d'accueil)
    ↓
VueImport (chargement fichier Excel)
    ↓
VueRecap (ordre de livraison + démarrage)
    ↓
VueTournee (navigation vers client) ←┐
    ↓                                 │
VueClient (intervention)              │ (boucle)
    ↓                                 │
[retour VueTournee ou VueFin] ────────┘
    ↓
VueFin (déchargement)
    ↓
VueRapport (génération PDF + email)
```

### 2.2 Configuration (config.json + localStorage)

**Fichier `/public/config.json` protégé par mot de passe** :
```json
{
  "nomUnite": "BDL12",
  "dossierDriveId": "1SDu8PBRpHP7oTiT3U1CXXH82ScZYou7i",
  "email1": "destinataire1@example.com",
  "email2": "destinataire2@example.com",
  "email3": "destinataire3@example.com",
  "adresseUnite": "123 Rue de l'Unité",
  "gpsUnite": "47.123,0.456",
  "mdpConfig": "bdl2025"
}
```

**Fusion** :
1. Charger `config.json` (défauts)
2. Surcharger avec `localStorage.getItem('bdl-config')` (modifications utilisateur)

**Accès** : VueConfig avec authentification par mot de passe

---

## 3. FORMATS DE DONNÉES

### 3.1 Fichier Excel de tournée

**Colonnes obligatoires** :
- `Client` : Nom du client (texte)
- `Adresse` : Adresse complète (texte)
- `GPS` : Coordonnées GPS `latitude, longitude` (texte)
- `Telephone` : Numéro de téléphone (texte)
- `Service` : Nom du service (ex: "EHPAD RDC", "EHPAD 1er")
- `CabrisPrevus` : Nombre de cabris à livrer (nombre entier)
- `Note de tournée` : Instructions pour le chauffeur (texte multi-lignes)

**Format multi-services** :
- Première ligne : client avec toutes ses infos + premier service
- Lignes suivantes : uniquement `Service` et `CabrisPrevus` renseignés (autres colonnes vides)

**Exemple** :
```
| Client  | Adresse | GPS          | Telephone | Service    | CabrisPrevus | Note de tournée |
|---------|---------|--------------|-----------|------------|--------------|-----------------|
| Le Lude | Chemin..| 47.65,0.15   | 02 43 ... | EHPAD RDC  | 2            | PASSER PAR...   |
|         |         |              |           | EHPAD 1er  | 3            |                 |
|         |         |              |           | USLD       | 1            |                 |
```

**Parser** :
- Détection automatique du format (CSV vs XLSX)
- Groupement des services par client (lignes consécutives avec `Client` vide)
- Gestion des variantes de nom de colonnes (accents, majuscules)

### 3.2 Structure de données interne

```javascript
tournee = {
  nom: "Tournée_le_Lude",  // nom du fichier sans extension
  date: "2026-03-03T10:30:00.000Z",
  clients: [
    {
      id: "C1",
      nom: "Le Lude",
      adresse: "Chemin des Bichousières, 72800 Le Lude",
      gps: "47.65037949168748, 0.1550324792921226",
      telephone: "02 43 48 48 48",
      noteTournee: "PASSER PAR LA BARRIERE...",
      services: [
        { id: "S1", nom: "EHPAD RDC", cabrisPrevu: 2 },
        { id: "S2", nom: "EHPAD 1er", cabrisPrevu: 3 },
        { id: "S3", nom: "USLD", cabrisPrevu: 1 }
      ]
    }
  ]
}

clientData = {
  "C1": {
    services: [...],  // copie avec cabrisPrevu
    commentaire: "Incident portail, clé introuvable",
    photos: [
      { id: "p-123", data: "data:image/jpeg;base64,...", heure: "14:35:22" }
    ],
    heureArrivee: "2026-03-03T14:30:00.000Z",
    heureDepart: "2026-03-03T14:45:00.000Z",
    cabrisReprisTotal: 9  // TOTAL unique (pas par service)
  }
}

logs = [
  { type: 'DEBUT_CHARGEMENT', timestamp: "...", clientIndex: null },
  { type: 'DEPART_UNITE', timestamp: "...", clientIndex: null },
  { type: 'ARRIVEE_CLIENT', timestamp: "...", clientIndex: 0 },
  { type: 'DEPART_CLIENT', timestamp: "...", clientIndex: 0, commentaire: "..." },
  { type: 'FIN_DECHARGE', timestamp: "...", clientIndex: null }
]
```

---

## 4. ÉCRANS DÉTAILLÉS

### 4.1 VueLogin

**Affichage** :
- Logo Bulle de Linge (90px hauteur)
- Titre : "BDL-LIVRAISON"
- Sous-titre : nom de l'unité (depuis config)
- Champ texte : "Nom du chauffeur"
- Bouton : "DÉMARRER" (bleu, gros)
- Bouton : "⚙️ Config" (gris, petit)
- Bouton : "ℹ️ À propos" (bleu clair, petit)

**Bannière de reprise de session** :
Si une session non bouclée existe (`localStorage.getItem('bdl-session')`), afficher :
```
┌─────────────────────────────────────┐
│ ⚠️ Tournée en cours détectée !      │
│ Tournée Le Lude — Froger            │
│ [ Reprendre la tournée ]            │
└─────────────────────────────────────┘
```

**Comportement** :
- Si "Reprendre" → restaurer état complet (vue, clientIndex, clientData, logs)
- Si "DÉMARRER" avec session en cours → popup confirmation :
  ```
  La tournée "Le Lude" n'a pas été bouclée.
  
  OK      = Reprendre la tournée
  Annuler = Démarrer une nouvelle tournée
  ```
- Validation : nom chauffeur requis (trim)

**Modal "À propos"** :
- Version : 1.2.0
- Description produit
- Crédits : "Développé pour Bulle de Linge"
- Bouton fermer

---

### 4.2 VueConfig

**Protection** :
- Popup `window.prompt("Mot de passe :")` au clic
- Si incorrect → `alert("Mot de passe incorrect")`

**Champs modifiables** :
- Nom de l'unité (texte)
- Email destinataire 1 (email)
- Email destinataire 2 (email)
- Email destinataire 3 (email)
- Adresse de l'unité (texte)
- GPS de l'unité (texte)

**Boutons** :
- "💾 Enregistrer" → `localStorage.setItem('bdl-config', JSON.stringify(config))`
- "← Retour" → retour à login ou import selon contexte

---

### 4.3 VueImport

**Chargement fichier** :
- Zone de drop drag & drop (optionnel)
- Input file : `.xlsx, .csv, .xls`
- Texte : "Sélectionnez un fichier Excel de tournée"

**Bouton Google Drive** (si configuré) :
- Affiche lien vers le dossier Drive configuré
- Texte : "📂 Ouvrir Google Drive" (bleu clair)

**Parsing** :
1. Lire le fichier (XLSX.read)
2. Détecter les colonnes (insensible à la casse/accents)
3. Parser ligne par ligne :
   - Si `Client` renseigné → nouveau client
   - Si `Client` vide → service du client précédent
4. Validation :
   - Au moins 1 client
   - Au moins 1 service par client
   - Colonnes obligatoires présentes

**Gestion d'erreurs** :
- Fichier corrompu → `alert("Fichier illisible")`
- Colonnes manquantes → `alert("Colonnes obligatoires manquantes : Client, Service, CabrisPrevus")`
- Fichier vide → `alert("Aucun client trouvé")`

**Session en cours** :
Si une session existe → popup confirmation avant de charger nouvelle tournée

---

### 4.4 VueRecap

**Informations affichées** :
- Nom de la tournée (haut de page)
- Nombre de clients
- Total de cabris à livrer (somme de tous les services)

**Liste des clients** :
- Carte par client : nom, adresse, nombre de cabris
- Badge vert : services multiples (ex: "3 services")

**Ordre de livraison** (drag & drop) :
```
┌─────────────────────────────────────┐
│ 🗂️ Ordre de livraison  (↑↓ pour réorganiser)
├─────────────────────────────────────┤
│ ① 🚀 Le Lude                    ↑↓  │
│    12 cabris                        │
├─────────────────────────────────────┤
│ ② Saint-Barthélemy-d'Anjou      ↑↓  │
│    8 cabris                         │
└─────────────────────────────────────┘
```

**Fonctionnalités** :
- Boutons ↑↓ pour monter/descendre dans l'ordre
- Client en cours (①) : fond vert clair
- Premier client marqué 🚀

**Phase de chargement** :
1. Bouton "📦 Démarrer le chargement" (orange)
2. Affichage heure de début de chargement
3. Bouton "🚚 DÉPART DE L'UNITÉ" (vert, gros)

**Logs créés** :
- `DEBUT_CHARGEMENT` (timestamp heure de clic)
- `DEPART_UNITE` (timestamp heure de départ)

**Bouton "← Changer"** (haut droite) :
- Retour à VueImport pour charger un autre fichier

---

### 4.5 VueTournee (navigation)

**Affichage** :
- Badge : "Client X / Y"
- Nom du client (gros titre)
- Adresse
- Nombre de services et total cabris
- **Note de tournée** (fond jaune, pleine largeur, multi-lignes si besoin)

**Boutons** :
- "🗺️ Ouvrir Waze" (bleu) → ouvre Waze avec l'adresse
- "→ ARRIVÉE CLIENT" (vert, gros)

**Note de tournée** :
- Toujours visible si renseignée
- Fond `#fef3c7`, bordure `#fde68a`
- Police 15px, `whiteSpace: pre-wrap` (retours à la ligne respectés)
- Texte issu de la colonne Excel `Note de tournée`

**Log créé** :
- `ARRIVEE_CLIENT` (timestamp)

---

### 4.6 VueClient (intervention)

**Affichage** :
- Nom du client (titre)
- **Note de tournée** (même style que VueTournee, au-dessus des services)
- Contact (téléphone, fond vert clair)
- Liste des services (lecture seule, badges bleus avec nombre de cabris)

**Saisie obligatoire** :
```
┌─────────────────────────────────────┐
│ Cabris ramassés — total (sale)      │
│                                     │
│          [ _____ ]                  │ ← input number, grand
│                                     │
│ ⚠️ Obligatoire — entrez 0 si aucun │
└─────────────────────────────────────┘
```
- Champ `type="number"`, `min="0"`
- Bordure rouge si vide
- Validation : impossible de valider si vide
- **Valeur unique** pour TOUS les services (pas de saisie par service)

**Commentaire de livraison** :
```
┌─────────────────────────────────────┐
│ Commentaire                         │
│ ┌─────────────────────────────────┐ │
│ │ Incident, observation...        │ │
│ │                                 │ │
│ └─────────────────────────────────┘ │
└─────────────────────────────────────┘
```
- `<textarea>` 3 lignes
- Optionnel
- Stocké dans `clientData[id].commentaire`

**Photos** :
- Bouton "📷 Appareil photo" (bleu)
- Bouton "🖼️ Galerie" (violet)
- Grille 2 colonnes
- Chaque photo : timestamp + bouton ✕ pour supprimer
- Clic sur photo → agrandissement plein écran
- Compression automatique (max 800px)

**Bouton validation** :
- "→ VALIDER ET DÉPART" (vert, gros)
- Disabled si `cabrisReprisTotal === ''`
- Appel : `onDepart(commentaire, services, photos, cabrisReprisTotal)`

**Données stockées** :
```javascript
clientData[id] = {
  services: [...],
  commentaire: "...",
  photos: [{id, data, heure}, ...],
  heureArrivee: "...",
  heureDepart: "...",
  cabrisReprisTotal: 9  // UN SEUL NOMBRE
}
```

**Log créé** :
- `DEPART_CLIENT` (timestamp, commentaire)

**Navigation** :
- Si d'autres clients → VueTournee (client suivant)
- Si dernier client → VueFin

---

### 4.7 VueFin (déchargement)

**Affichage** :
- Icône ✅ (grande)
- Titre : "Tournée terminée !"
- Message : "Tous les clients ont été visités"
- Récapitulatif :
  - Heure de fin de tournée
  - Durée totale
  - Nombre de clients

**Bouton** :
- "📦 Déchargement terminé" (vert, gros)
- Popup confirmation : "⛽ Pensez à faire le plein du camion !\n\nCliquez OK pour clôturer la tournée."

**Log créé** :
- `FIN_DECHARGE` (timestamp)

**Navigation** :
- VueRapport

---

### 4.8 VueRapport

**En-tête** :
- Titre : "📄 Rapport de Tournée"
- Nom de la tournée

**Chronologie** :
- Début chargement
- Départ unité
- Fin décharge
- Durée totale (format `Xh Ymin`)

**Statistiques** (3 cartes) :
- Clients (nombre)
- Chauffeur (nom)
- Durée

**Boutons principaux** :
- "📧 Générer PDF et envoyer par email" (bleu gradient, gros)
- "🏠 Nouvelle tournée" (gris)

**Détail des interventions** :
Pour chaque client :
```
┌─────────────────────────────────────┐
│ 1. Le Lude           14:30 → 14:45  │
├─────────────────────────────────────┤
│ Services livrés :                   │
│ EHPAD RDC    2 cabris               │
│ EHPAD 1er    3 cabris               │
│ USLD         1 cabri                │
├─────────────────────────────────────┤
│ Total ramassés : 9                  │ ← fond jaune
├─────────────────────────────────────┤
│ Commentaire                         │ ← fond jaune clair
│ Incident portail...                 │
├─────────────────────────────────────┤
│ 3 photo(s)                          │
│ [mini] [mini] [mini]                │
└─────────────────────────────────────┘
```

**Total général** (bas de page) :
```
┌─────────────────────────────────────┐
│   Cabris livrés    |  Cabris ramassés│
│        18          |        35       │
│                                     │
│ 17 cabris sales dans le camion      │
└─────────────────────────────────────┘
```
- Fond bleu dégradé
- Calcul : `ramassés - livrés`

**Génération PDF** :
1. Appel `genererPDF()` → téléchargement fichier
2. Attente 800ms
3. Ouverture client mail avec :
   - Destinataires : emails de la config
   - Sujet : `Rapport tournée BDL - [chauffeur] - [date]`
   - Corps : texte avec nom fichier + infos tournée
4. **Effacement de la session** → tournée bouclée

**Nouvelle tournée** :
- Si session non bouclée → popup confirmation :
  ```
  La tournée "Le Lude" n'a pas été bouclée (rapport non généré).
  
  OK      = Reprendre la tournée
  Annuler = Abandonner et démarrer une nouvelle tournée
  ```
- Sinon → retour à VueLogin

---

## 5. GÉNÉRATION PDF

### 5.1 En-tête compact

```
┌─────────────────────────────────────────────────────┐
│ [LOGO] BDL-Livraison — Rapport de tournée           │
│        Tournée Le Lude | Froger | 03/03/2026        │
│        Chgt: 10:30  Dep: 11:00  Fin: 15:45          │
│        Durée: 4h45                                  │
├─────────────────────────────────────────────────────┤
```

**Dimensions** :
- Page A4 (210mm)
- Marges : 12mm gauche/droite
- Largeur utile : 186mm

**Polices** :
- Titre : 14pt bold
- Infos : 8pt normal
- Couleur grise : RGB(100, 100, 100)

### 5.2 Section par client

**Disposition 2 colonnes** :
```
┌────────────────────┬──────────────────┐
│ 1. Le Lude    10:30→10:45            │
│ Chemin des Bichousières...           │
├────────────────────┼──────────────────┤
│ Service   |Liv|Ram │   [photo 1]      │
│ EHPAD RDC | 2 |    │   82mm × 109mm   │
│ EHPAD 1er | 3 | 9  │                  │
│───────────────────│   [photo 2]      │
│ Commentaire...    │                  │
│                   │   [photo 3]      │
└───────────────────┴──────────────────┘
```

**Colonne gauche (100mm)** :
- Fond bleu clair pour le nom
- Horaires à droite (8pt gris)
- Adresse en 8pt gris
- Tableau services (3 colonnes) :
  - Service (60%)
  - Livrés (20%)
  - Ramassés (20%, fond jaune, bold)
- Commentaire (fond jaune clair, 6pt italic)

**Colonne droite (82mm)** :
- 1 photo par ligne
- Largeur : 82mm
- Hauteur : 82mm × 4/3 = 109mm (ratio 3:4 portrait)
- Timestamp en dessous (6pt gris)

**Séparateur** :
- Trait gris clair (RGB 220, 220, 220) entre clients

### 5.3 Total général

```
┌─────────────────────────────────────┐
│     TOTAL GÉNÉRAL DE LA TOURNÉE     │
├─────────────────┬───────────────────┤
│ Cabris livrés   │ Cabris ramassés   │
│      18         │       35          │
└─────────────────┴───────────────────┘
```

**Mise en page** :
- Tableau jsPDF-autoTable
- Fond bleu pour l'en-tête
- Police 11pt bold

### 5.4 Nom du fichier

Format : `Rapport_YYYY-MM-DD_HHhMMhSS.pdf`

Exemple : `Rapport_2026-03-03_14h35h22.pdf`

---

## 6. SAUVEGARDE AUTOMATIQUE ET REPRISE

### 6.1 Principe

**Sauvegarde permanente** dans `localStorage` clé `bdl-session` :
```javascript
{
  vue: 'tournee',
  agentName: 'Froger',
  tournee: {...},
  nomTournee: 'Tournée_le_Lude',
  clientIndex: 2,
  startTime: '2026-03-03T10:00:00.000Z',
  logs: [...],
  clientData: {...}  // sans photos (trop lourd)
}
```

**Déclenchement** : `useEffect` sur tous changements d'état

**Exclusions** :
- Vue `login`, `import`, `config` : pas de sauvegarde
- Vue `rapport` : sauvegarde maintenue jusqu'à génération PDF

### 6.2 Restauration automatique

**Au démarrage de l'app** :
- `useEffect` monte unique
- Si session existe ET vue active (pas login/import/config/rapport)
- → Restaurer intégralement l'état

**Cas d'usage** :
- App plantée → relance → reprise exacte
- Onglet fermé par erreur → réouverture → reprise
- Batterie morte → recharge → reprise

### 6.3 Bouclage de tournée

**Session effacée uniquement** :
1. Après génération PDF + envoi email (setTimeout 800ms)
2. Clic "Nouvelle tournée" + confirmation abandon

**Contrôles de sécurité** :
- Bouton "DÉMARRER" sur login → vérification session
- Chargement fichier Excel → vérification session
- Bouton "Nouvelle tournée" → vérification session

**Popup confirmation** :
```
La tournée "Le Lude" n'a pas été bouclée (rapport non généré).

OK      = Reprendre la tournée
Annuler = Abandonner et démarrer une nouvelle tournée
```

---

## 7. COMPRESSION PHOTOS

**Fonction** : `compresserPhoto(dataURL, maxWidth)`

**Algorithme** :
1. Créer `<canvas>` temporaire
2. Charger l'image
3. Si largeur > maxWidth → redimensionner proportionnellement
4. Dessiner sur canvas
5. Export `canvas.toDataURL('image/jpeg', 0.85)`
6. Retour base64

**Paramètre** : `maxWidth = 800px`

**Stockage** :
- Photos compressées dans `clientData[id].photos`
- Photos exclues de la sauvegarde session (trop lourdes)

---

## 8. DÉPLOIEMENT

### 8.1 Configuration GitHub Pages

**package.json** :
```json
{
  "homepage": "https://[username].github.io/bdl-livraison",
  "scripts": {
    "predeploy": "npm run build",
    "deploy": "gh-pages -d build"
  }
}
```

**Commandes** :
```bash
npm run build   # Génération build/
npm run deploy  # Déploiement sur gh-pages
```

### 8.2 PWA

**manifest.json** :
```json
{
  "short_name": "BDL Livraison",
  "name": "BDL-Livraison — Gestion de tournées",
  "icons": [
    { "src": "logo192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "logo512.png", "sizes": "512x512", "type": "image/png" }
  ],
  "start_url": ".",
  "display": "standalone",
  "theme_color": "#2563eb",
  "background_color": "#ffffff"
}
```

**Service Worker** :
- Généré automatiquement par CRA
- Stratégie : cache-first pour assets statiques

**Installation** :
- Android Chrome : "Ajouter à l'écran d'accueil"
- iOS Safari : "Ajouter à l'écran d'accueil"

---

## 9. GESTION DES ERREURS

### 9.1 Parsing Excel

| Erreur | Message | Action |
|--------|---------|--------|
| Fichier corrompu | "Fichier illisible" | Demander nouveau fichier |
| Colonnes manquantes | "Colonnes obligatoires : ..." | Vérifier format Excel |
| Aucun client | "Aucun client trouvé" | Vérifier contenu |
| Encodage incorrect | Support UTF-8 + variantes accents | Fallback |

### 9.2 Photos

| Erreur | Comportement |
|--------|-------------|
| Lecture fichier échoue | Log console, photo ignorée |
| Compression échoue | Utiliser image originale |
| addImage PDF échoue | Log console, photo sautée |

### 9.3 localStorage

| Erreur | Comportement |
|--------|-------------|
| Quota dépassé | Log warning, continuer sans sauvegarder |
| JSON parse erreur | Ignorer session corrompue, démarrage propre |

---

## 10. POINTS D'ATTENTION CRITIQUES

### 10.1 Cabris ramassés

⚠️ **UN SEUL champ** pour le total, PAS un champ par service

**Mauvais** :
```
EHPAD RDC : 2 livrés, [___] ramassés
EHPAD 1er : 3 livrés, [___] ramassés
```

**Bon** :
```
Services livrés (lecture seule) :
- EHPAD RDC : 2
- EHPAD 1er : 3

Total cabris ramassés : [_____] ← UN SEUL CHAMP
```

### 10.2 Apostrophes dans window.confirm

⚠️ Toujours utiliser **backticks** pour éviter les erreurs de syntaxe :

```javascript
// ❌ MAUVAIS
window.confirm('La tournée "X" n'a pas été bouclée')

// ✅ BON
const nom = tournee.nom;
window.confirm(`La tournée "${nom}" n'a pas été bouclée`)
```

### 10.3 Photos PDF

⚠️ `new Image()` dans jsPDF ne lit pas `naturalWidth/Height`

**Solution** : dimensions fixes avec ratio calculé

```javascript
const photoW = 82;  // largeur zone
const photoH = photoW * 4 / 3;  // ratio 3:4 portrait
```

### 10.4 Cache navigateur

⚠️ Après `npm run deploy`, les utilisateurs doivent vider le cache

**Instructions utilisateurs** :
1. Chrome Android : Paramètres → Confidentialité → Effacer données
2. OU : Désinstaller PWA → Réinstaller
3. OU : Hard reload (Ctrl+Shift+R)

### 10.5 Session non bouclée

⚠️ **3 points de contrôle** obligatoires :
1. Login (bouton DÉMARRER)
2. Import Excel
3. Rapport (bouton Nouvelle tournée)

**Tous doivent vérifier** `sessionEnCours()` et proposer reprise/abandon

---

## 11. ÉVOLUTIONS FUTURES (hors scope v1.2)

- [ ] Bouton "Partager" natif Android (partage PDF via WhatsApp, Drive, etc.)
- [ ] Upload automatique PDF sur Google Drive
- [ ] Signature client sur tablette
- [ ] Mode hors-ligne complet (service worker avancé)
- [ ] Synchronisation temps réel multi-chauffeurs
- [ ] Statistiques tournées (durée moyenne, cabris/heure)
- [ ] Export Excel du rapport
- [ ] Notifications push (rappels, alertes)

---

## 12. DÉPENDANCES

```json
{
  "dependencies": {
    "react": "^18.2.0",
    "react-dom": "^18.2.0",
    "lucide-react": "^0.263.1",
    "xlsx": "^0.18.5",
    "jspdf": "^2.5.1",
    "jspdf-autotable": "^3.5.31"
  },
  "devDependencies": {
    "gh-pages": "^5.0.0"
  }
}
```

---

## 13. TESTS À EFFECTUER

### 13.1 Parcours nominal

1. Login → Import Excel → Recap → Démarrer chargement
2. Départ unité → Navigation client 1
3. Arrivée client 1 → Saisie cabris + commentaire + photos
4. Validation → Client 2 → ... → Client N
5. Fin tournée → Déchargement → Rapport
6. Génération PDF → Envoi email
7. Vérification PDF (mise en page, photos, totaux)

### 13.2 Cas limites

- [ ] Fichier Excel vide
- [ ] Client sans service
- [ ] Service sans cabris
- [ ] Tournée 1 seul client
- [ ] Tournée 20+ clients
- [ ] Photo 10MB (compression)
- [ ] 50 photos (performance)
- [ ] Commentaire 1000 caractères
- [ ] Note de tournée multi-lignes (20 lignes)
- [ ] Nom client très long (overflow)
- [ ] Adresse très longue

### 13.3 Gestion session

- [ ] Fermeture app en cours de tournée → réouverture → reprise OK
- [ ] Génération PDF → session effacée → pas de reprise
- [ ] Nouvelle tournée sans PDF → popup confirmation
- [ ] Import Excel avec session active → popup confirmation

### 13.4 Multi-navigateurs

- [ ] Chrome Android (principal)
- [ ] Chrome Desktop
- [ ] Safari iOS
- [ ] Firefox Android

---

## CHANGELOG v1.2.0

### Ajouts
- Note de tournée affichée (VueTournee + VueClient)
- Commentaire de livraison (textarea)
- Sauvegarde automatique session
- Reprise automatique après plantage
- Popup confirmation session non bouclée
- Bouton "Changer" sur VueRecap
- Photos dans rapport JSX + PDF
- Total général en bas du rapport
- Compression photos 800px
- PDF layout 2 colonnes (tableau + photos)
- Ratio photos portrait 3:4

### Corrections
- Cabris ramassés : UN seul champ total (pas par service)
- Parser Excel : variantes noms colonnes (accents, casse)
- Apostrophes dans confirm (backticks)
- Photos PDF dimensions fixes (pas de new Image())
- Cache : instructions utilisateurs
- Session bouclée uniquement après PDF
- Ordre livraison : noms longs (wordBreak)

### Optimisations
- PDF condensé (police 7-8pt)
- En-tête PDF 2 lignes
- Photos 82mm × 109mm (1 par ligne zone droite)
- Séparateur gris entre clients
- autoTable cellPadding réduit

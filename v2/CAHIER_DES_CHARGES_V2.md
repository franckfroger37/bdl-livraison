# CAHIER DES CHARGES — BDL-LIVRAISON v2.0

> **Basé sur la V1.2.0** — Ce document décrit les évolutions de la version 2.

---

## 1. OBJECTIFS DE LA V2

La V2 introduit un système de **pointage QR code des cabris** pour :
1. **Vérifier le chargement** : s'assurer que les bons cabris sont chargés dans le camion (pas de cabris d'une autre tournée)
2. **Vérifier la livraison** : s'assurer de livrer le bon cabri au bon client/service (pas d'erreur de livraison ou d'oubli)

---

## 2. FORMAT EXCEL — NOUVELLE COLONNE `CabrisIDs`

Une colonne `CabrisIDs` est ajoutée aux lignes de service dans le fichier Excel.

### 2.1 Format

| Client  | Adresse | GPS        | Telephone | Service    | CabrisPrevus | Note de tournée | **CabrisIDs**              |
|---------|---------|------------|-----------|------------|--------------|-----------------|----------------------------|
| Le Lude | Chemin..| 47.65,0.15 | 02 43 ... | EHPAD RDC  | 2            | PASSER PAR...   | **CAB-001, CAB-002**       |
|         |         |            |           | EHPAD 1er  | 3            |                 | **CAB-003, CAB-004, CAB-005** |
|         |         |            |           | USLD       | 1            |                 | **CAB-006**                |

### 2.2 Règles
- **1 cellule par ligne de service** contenant les identifiants des cabris **séparés par des virgules, espaces ou point-virgules**
- Si la colonne est absente ou vide pour un service → le service fonctionne en mode V1 (sans scan QR)
- Le nombre d'IDs dans `CabrisIDs` doit être cohérent avec `CabrisPrevus`
- Les QR codes sur les cabris physiques encodent l'identifiant ID (ex: `CAB-001`)

### 2.3 Rétrocompatibilité
- Les fichiers Excel V1 (sans colonne `CabrisIDs`) fonctionnent toujours normalement
- Le mode QR s'active **service par service** : seulement les services ayant des IDs dans `CabrisIDs`

---

## 3. DONNÉES INTERNES — EXTENSIONS V2

### 3.1 Structure service (extension)

```javascript
{
  id: "S1",
  nom: "EHPAD RDC",
  cabrisPrevu: 2,
  cabrisIds: ["CAB-001", "CAB-002"],  // NOUVEAU — tableau vide si non renseigné
  cabrisRecuperes: 0
}
```

### 3.2 Données scan chargement (nouveau state dans App)

```javascript
cabrisChargement = {
  scannés: ["CAB-001", "CAB-002", ...],  // IDs validés comme appartenant à la tournée
  anomalies: [
    { id: "CAB-999", commentaire: "Cabri d'une autre tournée, mis de côté" }
  ]
}
```

### 3.3 Données scan livraison (extension de clientData)

```javascript
clientData["C1"] = {
  // ... champs V1 inchangés ...
  cabrisScannés: {
    "S1": ["CAB-001", "CAB-002"],  // IDs scannés et validés par service
    "S2": ["CAB-003", "CAB-004", "CAB-005"]
  },
  anomaliesLivraison: [
    { id: "CAB-999", serviceId: null, type: "mauvais_client", commentaire: "..." }
  ]
}
```

### 3.4 Session localStorage (extension)

```javascript
{
  // ... champs V1 inchangés ...
  cabrisChargement: { scannés: [...], anomalies: [...] }  // NOUVEAU
}
```

---

## 4. NOUVEAU COMPOSANT : VueScanner

### 4.1 Technologie
- **API principale** : `BarcodeDetector` (API native Chrome 83+, Chrome Android, Safari 17+)
- **Fallback** : Saisie manuelle de l'ID si le navigateur ne supporte pas BarcodeDetector
- **Caméra** : `getUserMedia({ video: { facingMode: 'environment' } })`

### 4.2 Interface

```
┌──────────────────────────────┐
│ 📷 Scanner un cabri  [✕ Fermer] │
├──────────────────────────────┤
│                              │
│    ┌────────────────────┐    │
│    │                    │    │   ← Flux caméra en direct
│    │   [Cadre vert]     │    │
│    │                    │    │
│    └────────────────────┘    │
│                              │
├──────────────────────────────┤
│ ✅ Dernier scan : CAB-001     │
│    ou                        │
│ "Centrez le QR dans le cadre"│
└──────────────────────────────┘
```

### 4.3 Comportement
- Ouverture en **plein écran** (position fixed, zIndex 9999)
- Scan continu (toutes les 150ms via BarcodeDetector)
- **Anti-doublon** : 2 secondes de cooldown par ID pour éviter les scans répétés
- À chaque scan détecté → appel callback `onScan(id)` → le composant parent gère la logique métier
- Le scanner **reste ouvert** pour scanner le cabri suivant (fermeture manuelle par le chauffeur)

### 4.4 Gestion d'erreur caméra
- Permissions refusées → message explicatif + bouton Retour
- BarcodeDetector non supporté → formulaire de saisie manuelle (champ texte)

---

## 5. NOUVEAU COMPOSANT : ModalAnomalieQR

Popup bloquante affichée quand un ID scanné ne correspond pas à ce qui est attendu.

```
┌──────────────────────────────┐
│          ⚠️                   │
│    Anomalie détectée         │
│                              │
│  ┌──────────────────────────┐│
│  │ CAB-999  (fond rouge)   ││
│  │ Ce cabri n'appartient   ││
│  │ pas à cette tournée.    ││
│  └──────────────────────────┘│
│                              │
│  Commentaire obligatoire :   │
│  ┌──────────────────────────┐│
│  │ [champ texte libre]     ││
│  └──────────────────────────┘│
│                              │
│ [← Rescanner] [Signaler →]   │
└──────────────────────────────┘
```

- **Commentaire obligatoire** avant de pouvoir confirmer
- "Rescanner" → ferme la modal, rouvre le scanner
- "Signaler" → enregistre l'anomalie, rouvre le scanner

---

## 6. VueRecap — MODIFICATION : Phase de scan au chargement

### 6.1 Nouveau flux

```
[📦 Démarrer le chargement]
        ↓
[Phase de scan QR — si mode QR activé]
        ↓
[🚚 DÉPART DE L'UNITÉ]
```

### 6.2 Interface de scan chargement

```
┌─────────────────────────────────────┐
│ ⏱️ Chargement en cours... 10:30     │
├─────────────────────────────────────┤
│ 📦 Scan des cabris  12/15 ▓▓▓▓▓░   │
├─────────────────────────────────────┤
│ Le Lude                             │
│  EHPAD RDC  : CAB-001 ✅ CAB-002 ✅ │
│  EHPAD 1er  : CAB-003 ✅ CAB-004 ⏳ │
├─────────────────────────────────────┤
│ Saint-Barthélemy                    │
│  USLD       : CAB-006 ✅            │
├─────────────────────────────────────┤
│ ⚠️ 1 anomalie(s)                    │
│  CAB-999 — "Cabri inconnu"          │
├─────────────────────────────────────┤
│    [📱 Scanner un cabri]            │
├─────────────────────────────────────┤
│    [🚚 DÉPART DE L'UNITÉ]           │
└─────────────────────────────────────┘
```

### 6.3 Logique de scan

Pour chaque ID scanné :
- **ID trouvé dans la tournée** → ✅ marqué comme chargé, ID retiré de la liste d'attente
- **ID déjà scanné** → ignoré (anti-doublon)
- **ID inconnu** (pas dans la tournée) → ModalAnomalieQR → anomalie enregistrée

### 6.4 Bouton DÉPART

- Si **tous les cabris scannés** → bouton vert normal
- Si **des cabris manquants** → `window.confirm("⚠️ X cabri(s) non scannés. Confirmer le départ quand même ?")` avant départ
- Si **mode QR désactivé** (aucun service avec CabrisIDs) → comportement V1 inchangé

---

## 7. VueClient — MODIFICATION : Scan à la livraison

### 7.1 Nouveau bloc "Cabris à scanner"

Inséré entre les services et le champ "cabris ramassés" :

```
┌─────────────────────────────────────┐
│ 📦 Scan des cabris livrés  3/6 ▓▓░  │
├─────────────────────────────────────┤
│ EHPAD RDC                           │
│  CAB-001 ✅  CAB-002 ✅             │
├─────────────────────────────────────┤
│ EHPAD 1er                           │
│  CAB-003 ✅  CAB-004 ⏳  CAB-005 ⏳ │
├─────────────────────────────────────┤
│ ⚠️ 1 anomalie                       │
│  CAB-999 — "Mauvais client"         │
├─────────────────────────────────────┤
│    [📱 Scanner les cabris]          │
└─────────────────────────────────────┘
```

### 7.2 Logique de scan

Pour chaque ID scanné :
- **ID trouvé dans ce client** → ✅ marqué comme livré au bon service
- **ID déjà scanné** → ignoré
- **ID d'un autre client** ou **ID inconnu** → ModalAnomalieQR → anomalie enregistrée

### 7.3 Bouton VALIDER ET DÉPART

- Si **tous cabris scannés** → validation normale
- Si **des cabris manquants** → `window.confirm("⚠️ X cabri(s) non scannés. Confirmer la livraison quand même ?")` avant validation
- Si **mode QR désactivé** → comportement V1 inchangé

---

## 8. VueRapport — MODIFICATION : Section récap QR

### 8.1 Nouveau bloc "Bilan des scans QR"

Affiché uniquement si le mode QR était activé :

```
┌─────────────────────────────────────┐
│ 📱 Bilan des scans QR               │
├──────────────────┬──────────────────┤
│ Chargement       │ Livraison        │
│ 15/15 ✅         │ 14/15 ⚠️         │
└──────────────────┴──────────────────┘
  ⚠️ 1 anomalie — CAB-999 (chargement)
```

---

## 9. TECHNOLOGIES

Identiques à la V1 — **aucune nouvelle dépendance npm** requise.

- `BarcodeDetector` : API native Chrome, aucune lib externe nécessaire
- Le scan fonctionne sur Chrome Android (principal), Chrome Desktop, Safari 17+

---

## 10. CHECKLIST TESTS V2

- [ ] Fichier Excel V1 (sans CabrisIDs) → application fonctionne comme V1
- [ ] Fichier Excel V2 (avec CabrisIDs) → mode QR activé
- [ ] Scan au chargement → cabri valide → ✅
- [ ] Scan au chargement → cabri inconnu → anomalie + commentaire
- [ ] Départ sans tout scanner → confirmation
- [ ] Scan à la livraison → bon cabri → ✅
- [ ] Scan à la livraison → mauvais client → anomalie + commentaire
- [ ] Valider sans tout scanner → confirmation
- [ ] Rapport affiche le bilan QR
- [ ] Session restaurée → scan data conservé
- [ ] Chrome Android → BarcodeDetector OK
- [ ] Navigateur non supporté → fallback saisie manuelle

---

## 11. ÉVOLUTIONS FUTURES (hors scope v2.0)

*(Repris du backlog V1.2)*

- [ ] Bouton "Partager" natif Android (partage PDF via WhatsApp, Drive, etc.)
- [ ] Upload automatique PDF sur Google Drive
- [ ] Signature client sur tablette
- [ ] Mode hors-ligne complet (service worker avancé)
- [ ] Synchronisation temps réel multi-chauffeurs
- [ ] Statistiques tournées (durée moyenne, cabris/heure)
- [ ] Export Excel du rapport
- [ ] Notifications push (rappels, alertes)
- [ ] Intégration QR dans le rapport PDF (liste des anomalies)

---

## CHANGELOG V2.0

### Ajouts
- Système de pointage QR code des cabris (chargement + livraison)
- Composant VueScanner (BarcodeDetector natif + fallback saisie manuelle)
- Composant ModalAnomalieQR (signalement anomalie avec commentaire obligatoire)
- Colonne Excel `CabrisIDs` (IDs cabris par service)
- Phase de scan intégrée dans VueRecap (entre chargement et départ)
- Section scan intégrée dans VueClient (avant validation livraison)
- Bilan QR dans VueRapport
- Persistance des scans dans la session localStorage

### Compatibilité
- Fichiers Excel V1 (sans CabrisIDs) : fonctionnement identique à V1.2
- Aucune nouvelle dépendance npm

# 📂 Dossier des tournées

## Organisation

Placez vos fichiers Excel de tournées dans ce dossier `public/tournees/`.

### Structure :
```
public/
  └── tournees/
      ├── liste.json              ← Liste des fichiers disponibles
      ├── tournee-exemple.xlsx    ← Fichier exemple
      ├── tournee-lundi.xlsx      ← Vos fichiers
      ├── tournee-mardi.xlsx
      └── ...
```

## Comment ajouter une nouvelle tournée ?

1. **Copiez votre fichier Excel** dans `public/tournees/`
   - Exemple : `tournee-mercredi.xlsx`

2. **Modifiez `liste.json`** pour ajouter le nom du fichier :
   ```json
   [
     "tournee-exemple.xlsx",
     "tournee-lundi.xlsx",
     "tournee-mardi.xlsx",
     "tournee-mercredi.xlsx"
   ]
   ```

3. **Redémarrez l'app** (`npm start`)

4. Sur l'écran "Importer la tournée", vous verrez maintenant :
   ```
   📂 Tournées disponibles
   
   📄 tournee-exemple.xlsx
   📄 tournee-lundi.xlsx
   📄 tournee-mardi.xlsx
   📄 tournee-mercredi.xlsx
   ```

Cliquez sur un fichier pour l'importer directement !

## Format du fichier Excel

Les colonnes requises sont :
- `Client` : Nom du client
- `Adresse` : Adresse complète
- `Code1`, `Code2`, `Code3` : Codes d'accès
- `Telephone` : Numéro de téléphone
- `Commentaire` : Note fixe pour ce client
- `Service` : Nom du service (Draps, Serviettes, etc.)
- `CabrisPrevu` : Nombre de cabris à livrer

Voir `tournee-exemple.xlsx` pour un exemple complet.

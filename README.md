# Registre des Timbres Fiscaux

Application web Flask de gestion des timbres fiscaux pour étude de commissaire de justice.
Compilée en `.exe` autonome (PyInstaller) — aucune installation requise sur les postes.

---

## 1. Principe de fonctionnement

- Le fichier **`timbres-fiscaux.exe`** est stocké sur le partage réseau de l'étude :
  `\\SERVEUR\COMMUN\GESTION-TF\timbres-fiscaux.exe`

- Chaque collaborateur **double-clique** sur le `.exe` depuis son poste (via un raccourci bureau).

- Flask démarre **localement** sur le poste (port 5000, ou 5001, 5002… si déjà occupé).
  Le navigateur s'ouvre automatiquement.

- Toutes les **données** (JSON + PDFs) sont lues et écrites sur le partage réseau :
  `\\SERVEUR\COMMUN\GESTION-TF\data\`

- **Aucune installation** de Python ni de dépendances n'est requise sur les postes
  une fois le `.exe` compilé.

- Plusieurs collaborateurs peuvent utiliser l'application simultanément :
  chacun a sa propre instance Flask locale, les données partagées sont protégées
  par un verrou (`threading.Lock`) contre les écritures concurrentes.

---

## 2. Compilation du `.exe` (opération unique)

### Prérequis (sur le poste de compilation uniquement)

- [Python 3.10 ou supérieur](https://www.python.org/downloads/)
  **Important :** cocher **"Add Python to PATH"** pendant l'installation.

### Procédure

1. Copier `app.py` et `build.bat` dans un même dossier sur votre poste.
2. Double-cliquer sur **`build.bat`**.
3. Le script installe automatiquement les dépendances (`flask`, `pypdf`, `openpyxl`, `pyinstaller`)
   puis compile l'application. La compilation dure 1 à 2 minutes.
4. Le fichier produit se trouve dans **`dist\timbres-fiscaux.exe`**.

### Déploiement sur le serveur

```
Copier dist\timbres-fiscaux.exe  →  \\SERVEUR\COMMUN\GESTION-TF\timbres-fiscaux.exe
```

> La recompilation n'est nécessaire qu'en cas de mise à jour de l'application.
> Les données (`data/`) ne sont jamais affectées par une recompilation.

---

## 3. Déploiement sur les postes utilisateurs

1. Sur chaque poste, créer un **raccourci bureau** pointant vers :
   `\\SERVEUR\COMMUN\GESTION-TF\timbres-fiscaux.exe`

2. **Utilisation quotidienne** : double-clic sur le raccourci →
   une fenêtre de terminal s'ouvre brièvement, puis le navigateur s'ouvre automatiquement.

3. Pour quitter l'application : **fermer la fenêtre de terminal**.

> Si le partage réseau n'est pas accessible au démarrage, l'application affiche
> un message d'erreur explicite et s'arrête proprement.

---

## 4. Structure des données

```
\\SERVEUR\COMMUN\GESTION-TF\
├── timbres-fiscaux.exe          ← exécutable autonome
└── data\
    ├── timbres_2024.json        ← un fichier JSON par année civile (date d'achat)
    ├── timbres_2025.json
    ├── timbres_2026.json
    └── pdfs\                    ← PDFs individuels (noms = UUID)
        ├── a1b2c3d4ef56...pdf
        └── ...
```

### Stockage par année civile

Les timbres sont répartis par **année d'achat** :
un timbre acheté en 2024 reste dans `timbres_2024.json`, même s'il est attribué en 2026.

La logique FIFO garantit qu'un timbre de 2024 est servi **avant** un timbre de 2025
si les deux sont encore disponibles.

### Format d'un enregistrement

```json
{
  "id":               "uuid4",
  "numero":           "TF-2025-001",
  "date_achat":       "2025-01-15",
  "montant":          50.0,
  "statut":           "disponible",
  "pdf":              "a1b2c3d4.pdf",
  "dossier":          null,
  "date_utilisation": null
}
```

### Sauvegarde

Le dossier `data\` contient **toutes les données** de l'application.
**Sauvegarder régulièrement l'intégralité de ce dossier** (JSON + pdfs/).

Exemple via le Planificateur de tâches Windows :

```bat
xcopy /E /I /Y "\\SERVEUR\COMMUN\GESTION-TF\data" "D:\Sauvegardes\timbres-data"
```

---

## 5. Opérations manuelles (hors interface)

> ⚠️ L'interface ne propose aucun bouton de suppression ni d'annulation,
> afin de garantir l'intégrité du registre.
> **Ne jamais modifier un fichier JSON pendant qu'un import est en cours.**

### Supprimer un timbre

1. Ouvrir `\\SERVEUR\COMMUN\GESTION-TF\data\timbres_{année}.json`
2. Repérer le timbre par son champ `"numero"` ou `"id"`
3. Supprimer le bloc JSON correspondant (accolade ouvrante `{` → accolade fermante `}` + virgule éventuelle)
4. Vérifier que le fichier JSON reste syntaxiquement valide (tableau `[…]` bien formé)
5. Enregistrer le fichier
6. Supprimer le fichier PDF correspondant dans `data\pdfs\`
   (le nom du fichier est la valeur du champ `"pdf"`)

### Annuler une attribution (remettre un timbre en stock)

1. Ouvrir le fichier `timbres_{année}.json` de l'**année d'achat** du timbre
   (pas l'année d'utilisation)
2. Trouver le timbre par son `"numero"` ou son `"dossier"`
3. Modifier les trois champs suivants :
   ```json
   "statut": "disponible",
   "dossier": null,
   "date_utilisation": null
   ```
4. Enregistrer le fichier

---

## 6. Dépannage

| Symptôme | Cause probable | Solution |
|----------|---------------|----------|
| `ERREUR — Impossible d'accéder au répertoire réseau` | Le partage `\\SERVEUR\COMMUN\GESTION-TF` n'est pas accessible | Vérifier la connexion réseau et les droits d'accès au partage |
| Le navigateur ne s'ouvre pas automatiquement | Navigateur par défaut non configuré | Ouvrir manuellement `http://localhost:5000` |
| Port 5000 déjà utilisé | Un autre logiciel occupe le port | L'application bascule automatiquement sur 5001, 5002… (affiché dans le terminal) |
| Deux utilisateurs simultanés | Usage normal | Chaque poste a sa propre instance Flask locale ; les données sont partagées et protégées par verrou |
| JSON corrompu après modification manuelle | Erreur de syntaxe | Valider le fichier avec un éditeur JSON (ex. VS Code, Notepad++) avant de relancer |

---

## 7. Fonctionnalités

| Page | Description |
|------|-------------|
| **Tableau de bord** | Compteurs de stock, widgets dernier import / dernière attribution, import de lots PDF |
| **Disponibles** | Attribution FIFO du prochain timbre, téléchargement automatique du PDF |
| **Historique** | Attributions regroupées par lot, filtrable par année et par recherche texte, visualisation PDF en modal |
| **Export Excel** | Fichier `.xlsx` avec un onglet par année, mise en forme complète |

Alerte automatique si le stock disponible est ≤ 5 timbres.
Les PDFs des timbres ne sont accessibles (lecture + téléchargement) qu'**après attribution** (HTTP 403 sinon).

# Registre des Timbres Fiscaux

Application web Flask de gestion des timbres fiscaux pour étude de commissaire de justice.

---

## Présentation

Chaque introduction d'instance nécessite l'apposition d'un timbre fiscal à **50,00 €**.
L'étude achète les timbres en lots sur impots.gouv.fr sous forme de PDF multi-pages
(1 page = 1 timbre). Cette application gère le stock, l'attribution des timbres aux dossiers
et conserve un historique complet.

---

## Installation

### Prérequis

- Python 3.10 ou supérieur
- pip

### Étapes

```bash
# 1. Cloner ou copier les fichiers dans un dossier
cd C:\Etude\timbres-fiscaux      # Windows
# ou
cd ~/etude/timbres-fiscaux       # macOS / Linux

# 2. Installer les dépendances
pip install flask pypdf openpyxl

# 3. Lancer l'application
python app.py
```

L'application démarre sur le port **5000** et est accessible depuis tous les postes du réseau local.

---

## Accès depuis le réseau local

### Trouver l'adresse IP du serveur

**Windows :**
```
ipconfig
```
Chercher « Adresse IPv4 » — ex. `192.168.1.42`

**macOS / Linux :**
```bash
ip addr show
# ou
hostname -I
```

Les autres postes accèdent à l'application via : `http://192.168.1.42:5000`

---

## Structure des fichiers

```
gestionnaire-tf/
├── app.py                      ← Application Flask complète (fichier unique)
├── README.md                   ← Ce fichier
└── data/                       ← Créé automatiquement au premier lancement
    ├── timbres_2024.json        ← Un fichier JSON par année civile
    ├── timbres_2025.json
    ├── timbres_2026.json
    └── pdfs/                   ← PDFs individuels (noms UUID)
        ├── a1b2c3d4...pdf
        └── e5f6g7h8...pdf
```

### Important : stockage par année civile

Les timbres sont répartis dans des fichiers distincts selon leur **année d'achat**.
Un timbre acheté en 2024 est stocké dans `timbres_2024.json`, même s'il est attribué en 2025.

La logique FIFO garantit qu'un timbre de 2024 sera servi **avant** un timbre de 2025
si les deux sont encore disponibles.

### Format d'un enregistrement JSON

```json
{
  "id": "uuid4",
  "numero": "TF-2025-001",
  "date_achat": "2025-01-15",
  "montant": 50.0,
  "statut": "disponible",
  "pdf": "a1b2c3d4.pdf",
  "dossier": null,
  "date_utilisation": null
}
```

---

## Procédures manuelles

> ⚠️ Ces opérations se font par édition directe du fichier JSON.
> **Ne pas modifier le JSON pendant un import en cours.**

### Supprimer un timbre

1. Ouvrir `data/timbres_{année}.json`
2. Supprimer l'objet JSON correspondant au timbre (trouver par `"id"` ou `"numero"`)
3. Supprimer le fichier PDF correspondant dans `data/pdfs/` (le nom est dans le champ `"pdf"`)
4. Enregistrer le fichier JSON

### Annuler une attribution (remettre un timbre en stock)

1. Ouvrir `data/timbres_{année}.json` (l'année est celle du champ `"date_achat"`)
2. Trouver le timbre par son `"numero"` ou son `"dossier"`
3. Modifier les champs :
   ```json
   "statut": "disponible",
   "dossier": null,
   "date_utilisation": null
   ```
4. Enregistrer le fichier JSON

---

## Sauvegarde

Le dossier `data/` contient **toutes les données** de l'application :
- Les fichiers `timbres_{année}.json` → registre complet
- Le sous-dossier `pdfs/` → fichiers PDF individuels

**Sauvegarder régulièrement l'intégralité du dossier `data/`.**

Exemple de sauvegarde simple (Windows, planificateur de tâches) :
```bat
xcopy /E /I /Y C:\Etude\timbres-fiscaux\data D:\Sauvegardes\timbres-fiscaux\data
```

---

## Démarrage automatique sous Windows

### Méthode 1 — Fichier .bat dans le dossier Démarrage

1. Créer un fichier `lancer-timbres.bat` :
   ```bat
   @echo off
   cd /d C:\Etude\timbres-fiscaux
   python app.py
   pause
   ```
2. Appuyer sur `Win + R`, taper :
   ```
   shell:startup
   ```
3. Copier le fichier `.bat` dans ce dossier.

L'application démarrera automatiquement à chaque ouverture de session Windows.

### Méthode 2 — Raccourci dans le dossier Démarrage

Même principe mais via un raccourci pointant vers le fichier `.bat`.

---

## Fonctionnalités

| Page | Description |
|------|-------------|
| **Tableau de bord** | Vue d'ensemble du stock, import de lots PDF |
| **Disponibles** | Attribution FIFO du prochain timbre à un dossier |
| **Historique** | Liste des attributions regroupées par lot, filtrable par année |
| **Export Excel** | Fichier `.xlsx` avec un onglet par année |

### Alertes

Un bandeau rouge s'affiche automatiquement si le stock disponible est inférieur ou égal à **5 timbres**.

### Sécurité des PDFs

Les fichiers PDF des timbres ne sont accessibles qu'**après attribution**.
Un timbre encore en statut `disponible` renvoie une erreur HTTP 403.

---

## Dépendances

| Paquet | Usage |
|--------|-------|
| `flask` | Serveur web |
| `pypdf` | Lecture et découpage des PDFs |
| `openpyxl` | Génération des exports Excel |

---

## Notes techniques

- Le serveur écoute sur `0.0.0.0:5000` (tous les postes du réseau local)
- Les écritures JSON sont protégées par un verrou (`threading.Lock`) pour éviter la corruption en cas d'accès simultanés
- Pas de base de données — tout est stocké dans des fichiers JSON lisibles et éditables directement

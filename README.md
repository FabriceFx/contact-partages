# Gestionnaire de Contacts Partagés Google Workspace (Sheets)

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

Une solution complète basée sur Google Apps Script pour gérer les **Contacts Partagés du Domaine** (annuaire global d'entreprise) directement depuis une interface Google Sheets. Ce script permet aux administrateurs Google Workspace d'importer, de créer, de mettre à jour et de supprimer des contacts en masse.

## 🚀 Fonctionnalités clés

* **Importation (Read)** : Récupération de tous les contacts partagés existants du domaine vers le Sheet.
* **Synchronisation Bidirectionnelle (Create/Update)** :
    * Détection automatique des nouveaux contacts (POST).
    * Mise à jour des contacts existants (PUT) basée sur l'email.
    * Gestion des conflits via `If-Match` (ETag).
* **Suppression (Delete)** : Suppression de contacts de l'annuaire basée sur la sélection dans le tableur.
* **Formatage** : Nettoyage automatique des numéros de téléphone au format international (+33).
* **Interface** : Menu personnalisé `Gestion Workspace` intégré directement dans Google Sheets.

## 📋 Prérequis Techniques

1.  **Compte Administrateur** : L'utilisateur exécutant le script doit disposer des droits d'administration sur le domaine Google Workspace (ou droits délégués pour la gestion des contacts).
2.  **API GData** : Le script utilise le protocole legacy `https://www.google.com/m8/feeds` (Atom/XML) car l'API People ne couvre pas encore totalement l'écriture des contacts de domaine partagés de manière simple.

## 🛠️ Installation et configuration

### 1. Création du script
1.  Ouvrez un nouveau **Google Sheet**.
2.  Allez dans `Extensions` > `Apps Script`.
3.  Copiez le contenu du fichier `Code.js` fourni dans l'éditeur.

### 2. Configuration du manifeste (Scopes)
Pour que le script puisse accéder à l'annuaire du domaine, vous devez modifier le fichier manifeste `appsscript.json`.
1.  Dans l'éditeur Apps Script, cliquez sur la roue dentée ⚙️ (Paramètres du projet).
2.  Cochez la case "Afficher le fichier manifeste 'appsscript.json' dans l'éditeur".
3.  Revenez dans l'éditeur, ouvrez `appsscript.json` et remplacez son contenu par ceci :

```json
{
  "timeZone": "Europe/Paris",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "[https://www.googleapis.com/auth/script.external_request](https://www.googleapis.com/auth/script.external_request)",
    "[https://www.googleapis.com/auth/spreadsheets.currentonly](https://www.googleapis.com/auth/spreadsheets.currentonly)",
    "[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)",
    "[https://www.google.com/m8/feeds](https://www.google.com/m8/feeds)"
  ]
}

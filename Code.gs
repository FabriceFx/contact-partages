
/**
 * @file Gestionnaire de Contacts Partagés (Import/Export/Sync)
 * @author Fabrice Faucheux
 */

// --- 1. CONFIGURATION ---

// ⚠️ REMPLACEZ PAR VOTRE VRAI DOMAINE
const DOMAINE = "aotre_domaine.com";

// URLs
const BASE_URL = `https://www.google.com/m8/feeds/contacts/${DOMAINE}/full`;
const URL_FETCH_PARAMS = `?max-results=5000&alt=atom`; // On récupère jusqu'à 5000 contacts

// Mapping Colonnes
const COL_PRENOM = 0;
const COL_NOM = 1;
const COL_EMAIL = 2;
const COL_TEL_BUREAU = 3;
const COL_MOBILE = 4;
const COL_ENTREPRISE = 5;
const COL_DEPARTEMENT = 6;
const COL_FONCTION = 7;
const COL_ADRESSE = 8;
const COL_STATUT = 9;

const SCHEMA_CONTACTS = [
  "Prénom", "Nom", "Email Pro (ID)", "Tél. Bureau", "Mobile", 
  "Entreprise", "Département", "Fonction", "Adresse Bureau", "Statut Sync"
];

// --- 2. MENU ---

const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Gestion Workspace')
    .addItem('1. Initialiser la feuille', 'initialiserStructureFeuille')
    .addSeparator()
    .addItem('⬇️ IMPORTER depuis l\'Annuaire', 'importerContactsExistants')
    .addSeparator()
    .addItem('🔄 SYNCHRONISER (Ajout/Modif)', 'synchroniserContacts')
    .addItem('🗑️ SUPPRIMER la sélection', 'supprimerContactsSelectionnes')
    .addToUi();
};

// --- 3. FONCTIONNALITÉS STRUCTURE ---

/**
 * Crée les en-têtes et formate la feuille.
 */
const initialiserStructureFeuille = () => {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (feuille.getLastRow() > 1) {
    const rep = ui.alert("Attention", "Cela va écraser la première ligne. Continuer ?", ui.ButtonSet.YES_NO);
    if (rep !== ui.Button.YES) return;
  }

  // En-têtes
  const plage = feuille.getRange(1, 1, 1, SCHEMA_CONTACTS.length);
  plage.setValues([SCHEMA_CONTACTS])
       .setFontWeight("bold")
       .setBackground("#4285F4")
       .setFontColor("#FFFFFF")
       .setHorizontalAlignment("center");
  
  // Vue
  feuille.setFrozenRows(1);
  feuille.autoResizeColumns(1, SCHEMA_CONTACTS.length);
  feuille.setColumnWidth(COL_EMAIL + 1, 250); // Plus large pour l'email
  feuille.setColumnWidth(COL_ADRESSE + 1, 300); // Plus large pour l'adresse

  // Validation Email
  const regleEmail = SpreadsheetApp.newDataValidation()
    .requireTextIsEmail().setAllowInvalid(false).build();
  feuille.getRange(2, COL_EMAIL + 1, 999, 1).setDataValidation(regleEmail);

  ui.alert("Structure initialisée.");
};

// --- 4. IMPORTATION (NOUVEAU) ---

/**
 * Récupère les contacts du domaine et remplit le Sheet.
 */
const importerContactsExistants = () => {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // Sécurité
  if (DOMAINE === "votre-domaine-principal.com") return ui.alert("Veuillez configurer la variable DOMAINE dans le script.");

  const rep = ui.alert("Attention", "L'importation va EFFACER toutes les données actuelles de la feuille (sauf l'en-tête). Continuer ?", ui.ButtonSet.YES_NO);
  if (rep !== ui.Button.YES) return;

  try {
    // 1. Appel API en lecture
    const url = BASE_URL + URL_FETCH_PARAMS;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'GData-Version': '3.0'
      }
    });

    // 2. Analyse XML (Parsing)
    const xml = response.getContentText();
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
    const gd = XmlService.getNamespace('http://schemas.google.com/g/2005');

    const entries = root.getChildren('entry', atom);
    
    if (entries.length === 0) return ui.alert("Aucun contact trouvé dans le domaine.");

    // 3. Transformation XML -> Tableau
    const donnees = entries.map(entry => {
      // Helpers d'extraction sécurisée
      const getVal = (parent, tag) => {
        const el = parent.getChild(tag, gd);
        return el ? el.getText() : "";
      };
      
      // Extraction Nom
      const nameEl = entry.getChild('name', gd);
      const prenom = nameEl ? getVal(nameEl, 'givenName') : "";
      const nom = nameEl ? getVal(nameEl, 'familyName') : "";

      // Extraction Email (Prioritaire ou premier trouvé)
      const emails = entry.getChildren('email', gd);
      let email = "";
      if (emails.length > 0) {
        const primary = emails.find(e => e.getAttribute('primary') === 'true');
        email = primary ? primary.getAttribute('address').getValue() : emails[0].getAttribute('address').getValue();
      }

      // Extraction Téléphones
      const phones = entry.getChildren('phoneNumber', gd);
      let mobile = "", bureau = "";
      phones.forEach(p => {
        const rel = p.getAttribute('rel') ? p.getAttribute('rel').getValue() : "";
        const val = p.getText();
        if (rel.includes('#mobile')) mobile = val;
        else if (rel.includes('#work')) bureau = val;
        else if (!bureau) bureau = val; // Fallback
      });

      // Extraction Organisation
      const org = entry.getChild('organization', gd);
      const ent = org ? getVal(org, 'orgName') : "";
      const dept = org ? getVal(org, 'orgDepartment') : "";
      const titre = org ? getVal(org, 'orgTitle') : "";

      // Extraction Adresse
      const adrEl = entry.getChild('structuredPostalAddress', gd);
      const adresse = adrEl ? getVal(adrEl, 'formattedAddress') : "";

      return [
        prenom, nom, email, 
        formaterNumeroInternational(bureau), 
        formaterNumeroInternational(mobile),
        ent, dept, titre, adresse,
        "Importé le " + new Date().toLocaleTimeString()
      ];
    });

    // 4. Écriture dans la feuille
    // On efface d'abord les anciennes données (de la ligne 2 à la fin)
    const lastRow = feuille.getLastRow();
    if (lastRow > 1) {
      feuille.getRange(2, 1, lastRow - 1, SCHEMA_CONTACTS.length).clearContent();
    }

    // On écrit les nouvelles
    if (donnees.length > 0) {
      feuille.getRange(2, 1, donnees.length, SCHEMA_CONTACTS.length).setValues(donnees);
    }

    ui.alert(`${donnees.length} contact(s) importé(s) avec succès.`);

  } catch (e) {
    console.error(e);
    ui.alert("Erreur lors de l'import : " + e.message);
  }
};

// --- 5. SYNCHRONISATION (CODE VALIDÉ) ---

const synchroniserContacts = () => {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const derniereLigne = feuille.getLastRow();

  if (DOMAINE === "votre-domaine-principal.com") return ui.alert("Configurez la variable DOMAINE !");
  if (derniereLigne < 2) return ui.alert("Pas de données.");

  const valeurs = feuille.getRange(2, 1, derniereLigne - 1, COL_STATUT + 1).getValues();

  const statuts = valeurs.map(ligne => {
    const email = ligne[COL_EMAIL];
    if (!email) return ["Email manquant"];

    const donnees = {
      prenom: ligne[COL_PRENOM],
      nom: ligne[COL_NOM],
      email: email,
      telBureau: formaterNumeroInternational(ligne[COL_TEL_BUREAU]),
      telMobile: formaterNumeroInternational(ligne[COL_MOBILE]),
      entreprise: ligne[COL_ENTREPRISE],
      departement: ligne[COL_DEPARTEMENT],
      fonction: ligne[COL_FONCTION],
      adresse: ligne[COL_ADRESSE]
    };

    try {
      const contactExistant = chercherContactParEmail(email);
      const headers = {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'GData-Version': '3.0',
        'Content-Type': 'application/atom+xml'
      };

      if (contactExistant) {
        headers['If-Match'] = contactExistant.etag;
        UrlFetchApp.fetch(contactExistant.link.find(l => l.rel === 'edit').href, {
          method: 'put', headers: headers, payload: construireXmlStrict(donnees, contactExistant.id)
        });
        return [`Mis à jour : ${new Date().toLocaleTimeString()}`];
      } else {
        UrlFetchApp.fetch(BASE_URL, {
          method: 'post', headers: headers, payload: construireXmlStrict(donnees)
        });
        return [`Créé : ${new Date().toLocaleTimeString()}`];
      }
    } catch (e) {
      return [`Erreur : ${e.message}`];
    }
  });

  feuille.getRange(2, COL_STATUT + 1, statuts.length, 1).setValues(statuts);
  ui.alert("Synchronisation terminée.");
};

// --- 6. SUPPRESSION ---

const supprimerContactsSelectionnes = () => {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const plage = feuille.getActiveRange();
  
  if (!plage || plage.getRow() < 2) return ui.alert("Sélection invalide");

  const rep = ui.alert("Supprimer ces contacts de l'annuaire ?", ui.ButtonSet.YES_NO);
  if (rep !== ui.Button.YES) return;

  const lignes = feuille.getRange(plage.getRow(), 1, plage.getNumRows(), COL_STATUT + 1).getValues();

  const statuts = lignes.map(ligne => {
    const email = ligne[COL_EMAIL];
    if(!email) return ["-"];
    try {
      const c = chercherContactParEmail(email);
      if(!c) return ["Introuvable/Déjà supprimé"];
      
      UrlFetchApp.fetch(c.link.find(l => l.rel === 'edit').href, {
        method: 'delete',
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
          'GData-Version': '3.0',
          'If-Match': c.etag
        }
      });
      return ["Supprimé"];
    } catch(e) { return ["Erreur"]; }
  });
  feuille.getRange(plage.getRow(), COL_STATUT + 1, statuts.length, 1).setValues(statuts);
};

// --- 7. UTILITAIRES & XML ---

const construireXmlStrict = (d, idExist = null) => {
  const fullName = `${d.prenom} ${d.nom}`.trim();
  let xml = `<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:gd='http://schemas.google.com/g/2005'>
    <atom:category scheme='http://schemas.google.com/g/2005#kind' term='http://schemas.google.com/contact/2008#contact'/>
    <gd:name>
      <gd:givenName>${d.prenom}</gd:givenName>
      <gd:familyName>${d.nom}</gd:familyName>
      <gd:fullName>${fullName}</gd:fullName>
    </gd:name>`;

  if (d.email) xml += `<gd:email rel='http://schemas.google.com/g/2005#work' primary='true' address='${d.email}' displayName='${fullName}' />`;
  if (d.telBureau) xml += `<gd:phoneNumber rel='http://schemas.google.com/g/2005#work' primary='true'>${d.telBureau}</gd:phoneNumber>`;
  if (d.telMobile) xml += `<gd:phoneNumber rel='http://schemas.google.com/g/2005#mobile'>${d.telMobile}</gd:phoneNumber>`;

  if (d.entreprise || d.fonction || d.departement) {
    xml += `<gd:organization rel='http://schemas.google.com/g/2005#work' primary='true'>`;
    if (d.entreprise) xml += `<gd:orgName>${d.entreprise}</gd:orgName>`;
    if (d.departement) xml += `<gd:orgDepartment>${d.departement}</gd:orgDepartment>`;
    if (d.fonction) xml += `<gd:orgTitle>${d.fonction}</gd:orgTitle>`;
    xml += `</gd:organization>`;
  }

  if (d.adresse) {
    xml += `<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work' primary='true'>
              <gd:formattedAddress>${d.adresse}</gd:formattedAddress>
            </gd:structuredPostalAddress>`;
  }

  if (idExist) xml += `<atom:id>${idExist}</atom:id>`;
  xml += `</atom:entry>`;
  return xml;
};

const chercherContactParEmail = (email) => {
  const url = `${BASE_URL}?q=${encodeURIComponent(email)}&alt=json`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 'GData-Version': '3.0' },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) return null;
  const data = JSON.parse(response.getContentText());
  const entry = data.feed.entry;
  if (!entry || entry.length === 0) return null;
  const contactExact = entry.find(c => c.gd$email && c.gd$email.some(e => e.address.toLowerCase() === email.toLowerCase()));
  if (!contactExact) return null;
  return { id: contactExact.id.$t, etag: contactExact.gd$etag, link: contactExact.link };
};

const formaterNumeroInternational = (numero) => {
  if (!numero) return "";
  let propre = String(numero).replace(/[^0-9+]/g, '');
  if (!propre) return "";
  if (propre.startsWith('+')) return propre;
  if (propre.startsWith('0')) return "+33" + propre.substring(1);
  return "+33" + propre;
};

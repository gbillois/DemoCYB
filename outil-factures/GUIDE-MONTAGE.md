# Outil de traitement des factures - Guide de montage et d'utilisation

Outil Excel + VBA qui automatise une grande partie du traitement des factures
fournisseurs (signature, paiement, ajout de fournisseur), **sans rien installer** :
VBA est integre a Excel, et Outlook est pilote via son modele objet.

## 1. Ce que l'outil fait (et ne fait pas)

**Automatise :**
- l'import des factures PDF recues par email (Outlook) ;
- le pre-remplissage des donnees de facture depuis le PDF ;
- la verification du fournisseur dans un referentiel ;
- le rangement des fichiers sur le reseau ;
- le calcul des signataires (type de depense x destination x seuils) ;
- le suivi de l'etat de chaque facture / paiement / onboarding ;
- la generation des emails de demande de pieces fournisseur.

**Reste manuel (par conception) :**
- la creation de l'enveloppe **Docusign** (l'outil prepare le recapitulatif,
  ouvre le site et suit le statut) ;
- la verification du solde et l'execution du **virement** sur les sites
  bancaires (aucune API, automatisation non fiable) ;
- la **confirmation** des donnees extraites du PDF (champs surlignes en jaune).

## 2. Prerequis

- Excel et Outlook (versions bureau Windows) sur le meme poste.
- Macros `.xlsm` autorisees.
- Word (utilise en coulisse pour lire la couche texte des PDF).
- Acces en ecriture aux dossiers reseau utilises.

## 3. Montage du classeur (a faire une seule fois)

1. Ouvrir `Outil-Factures.xlsx`.
2. **Enregistrer sous** -> type **« Classeur Excel prenant en charge les macros
   (*.xlsm) »**. Garder ce `.xlsm` comme fichier de travail.
3. Ouvrir l'editeur VBA : touches **Alt + F11**.
4. Menu **Fichier > Importer un fichier...** et importer **un par un** les
   modules du dossier `vba/` :
   `modConfig.bas`, `modJournal.bas`, `modFichiers.bas`, `modOutlook.bas`,
   `modExtraction.bas`, `modRouteur.bas`, `modFournisseurs.bas`,
   `modSuivi.bas`, `modPaiements.bas`, `modUI.bas`.
5. **ThisWorkbook** : ne pas l'importer. Dans l'explorateur de projet,
   double-cliquer sur l'objet **ThisWorkbook**, puis copier-coller le code
   contenu dans `vba/ThisWorkbook.cls` (a partir de `Private Sub Workbook_Open`).
6. Menu **Debogage > Compiler VBAProject** : il ne doit y avoir aucune erreur.
7. Revenir a Excel, **Alt + F8**, lancer la macro **`InstallerBoutons`** :
   les boutons d'action apparaissent sur le « Tableau de bord ».
8. Enregistrer le `.xlsm`.

### Emplacement approuve (recommande)

Pour eviter le blocage des macros, placer le `.xlsm` dans un dossier declare
comme **emplacement approuve** : Excel > Fichier > Options > Centre de gestion
de la confidentialite > Parametres... > Emplacements approuves > Ajouter.

## 4. Configuration

### Feuille « Parametres »

| Parametre | A renseigner |
|---|---|
| Dossier Outlook surveille | nom du dossier Outlook ou arrivent les factures |
| Dossier reseau - A signer / Signees / Archives | chemins reseau reels |
| Dossier local extraction PDF | dossier local de travail (ex. `C:\Temp\FacturesPDF`) |
| Expediteur retour Docusign (filtre) | domaine de l'expediteur Docusign |
| URL Docusign | URL de connexion |
| Categorie Outlook - traite | categorie posee sur les emails traites |
| Extraction OCR niveau 2 (Azure) | `Non` par defaut (voir section 7) |
| Utilisateur | initiales / nom (pour le journal) |

### Feuille « Regles Signataires »

La matrice est pre-remplie d'exemples. **A adapter aux regles reelles** :
type de depense, destination, seuils min/max (la borne **max est exclue**)
et signataires. Une facture utilise la premiere ligne dont le type
correspond et dont le montant TTC est dans l'intervalle.

### Feuille « Fournisseurs »

Importer le referentiel fournisseurs existant (un fournisseur par ligne).

## 5. Utilisation au quotidien

Les boutons sont sur le **Tableau de bord**. Les actions « (ligne) »
s'appliquent a la ligne ou se trouve le curseur.

**Process signature :**
1. `1 - Importer les factures` -> cree les lignes dans « Suivi Factures ».
2. Verifier les champs jaunes, choisir le **Type de depense**.
3. `2 - Confirmer les donnees` -> retire les surlignages, calcule les
   signataires, statut « En attente signature ».
4. `4 - Preparer la signature` -> recapitulatif + copie reseau + ouverture
   Docusign. Creer l'enveloppe **manuellement** dans Docusign.
5. `5 - Detecter les retours Docusign` -> passe les factures signees a
   « Signee » (rapprochement par numero de facture / fournisseur).
6. `6 - Marquer signee` puis `7 - Classer la facture` si besoin.

**Process ajout fournisseur :**
- `Nouveau fournisseur` -> saisie nom + email, creation d'une ligne dans
  « Onboarding Fournisseurs » et brouillon d'email de demande de pieces.
- Cocher `Oui` au fur et a mesure des pieces recues, puis `MaJ onboarding` :
  une fois toutes les pieces (dont l'Annexe 7) recues, le fournisseur est
  ajoute automatiquement au referentiel.

**Process paiement :**
- Dans « Suivi Factures », selectionner les lignes a payer (meme fournisseur),
  puis `Creer un lot de paiement`.
- Completer societe emettrice, compte debite, solde.
- `Generer annexe paiement` -> fichier `.txt` recapitulatif.
- Envoi Docusign de paiement et virement : **manuels**.

## 6. Extraction des donnees du PDF

- **Niveau 1 (par defaut)** : Word convertit le PDF et l'outil lit sa couche
  texte, puis repere fournisseur, numero, date et montants. Gratuit, sans
  configuration. Fonctionne sur les **PDF numeriques** (la majorite).
- Les champs pre-remplis sont **surlignes en jaune** : toujours les verifier.
- Si un PDF est **scanne** (image), le niveau 1 ne trouve rien : la ligne est
  laissee en « Donnees a confirmer » avec une remarque. Saisir les champs a la
  main, ou les faire extraire par **Copilot** (deposer le PDF dans Copilot
  Chat) puis recopier.

## 7. OCR niveau 2 - Azure (optionnel)

Pour traiter automatiquement les PDF scannes, l'outil peut appeler **Azure AI
Document Intelligence** (modele « facture »). Cela suppose que l'IT cree une
ressource Azure et fournisse une cle.

1. Renseigner « Azure DI - Endpoint » dans la feuille Parametres.
2. Stocker la cle dans une **variable d'environnement Windows** (ne PAS la
   mettre dans le classeur) ; son nom est indique par « Azure DI - Variable
   d'env. cle » (par defaut `AZURE_DI_KEY`).
3. Passer « Extraction OCR niveau 2 (Azure) » a `Oui`.

Cout indicatif : ~0,01 USD par page.

## 8. Points de securite

- **Cle Azure** : jamais en clair dans le `.xlsm` ; uniquement en variable
  d'environnement.
- **Avertissement Outlook** : selon l'antivirus, un message peut demander
  l'autorisation d'acceder a Outlook lors du premier import - l'accepter.
- **Sauvegarde** : le classeur centralise le suivi ; le sauvegarder
  regulierement sur le reseau.

## 9. Limites connues

- L'extraction PDF est une **aide a la saisie**, pas une garantie : toujours
  verifier les champs jaunes.
- Le rapprochement des retours Docusign est automatique mais **best-effort**
  (base sur le numero de facture) ; les retours non rattaches sont signales.
- Banque non automatisable : etapes de solde et de virement manuelles.
- Outil maintenu sur un seul poste : prevoir une sauvegarde et ce guide.

## 10. Checklist de recette

- [ ] Compilation VBA sans erreur (`Debogage > Compiler`).
- [ ] `InstallerBoutons` a cree les boutons.
- [ ] Import d'un email de test avec PDF -> ligne creee, PDF range.
- [ ] Champs extraits coherents (voir `exemples/jeu-de-donnees.md`).
- [ ] Routage correct sur 3 cas (Societe / Immeuble / Travaux) + un cas
      au-dela d'un seuil.
- [ ] Retour Docusign simule -> statut « Signee ».
- [ ] Onboarding fournisseur -> brouillon d'email genere.
- [ ] Lot de paiement -> total correct, annexe generee.

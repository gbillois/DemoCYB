# Outil de traitement des factures

Outil Excel + VBA qui reconstruit une partie de l'automatisation perdue apres
l'abandon d'Esker : import des factures recues par email, extraction des
donnees du PDF, rangement reseau, routage des signataires et suivi d'etat.
Aucune installation requise : VBA est integre a Excel, Outlook est pilote via
son modele objet.

## Demarrage

Tout est explique pas a pas dans **`GUIDE-MONTAGE.md`** : montage du classeur,
configuration, utilisation quotidienne, OCR, securite et recette.

## Contenu du dossier

| Element | Role |
|---|---|
| `Outil-Factures.xlsx` | Squelette du classeur (9 feuilles, validations, mises en forme). A enregistrer en `.xlsm`. |
| `construire_classeur.py` | Script qui regenere le squelette (openpyxl). |
| `vba/` | Modules VBA a importer dans le classeur (`.bas`) + code `ThisWorkbook`. |
| `modeles-emails/` | Copies de reference des modeles d'email. |
| `exemples/` | Factures PDF de test + scenario de recette. |
| `GUIDE-MONTAGE.md` | Guide complet de montage et d'utilisation. |

## Perimetre

Automatise l'ingestion email, l'extraction PDF, le classement fichier, le
routage des signataires et le suivi. Restent manuels (par conception) : la
creation de l'enveloppe Docusign et les operations sur les sites bancaires.

# Jeu de donnees de test

Ce dossier contient deux factures PDF de test, avec une vraie couche texte
(donc lisibles par le niveau 1 d'extraction, sans OCR) :

| Fichier | Fournisseur | N0 facture | Date | TTC |
|---|---|---|---|---|
| `facture_ACME_2026-0412.pdf` | ACME Travaux SARL | 2026-0412 | 14/05/2026 | 22 200,00 |
| `facture_EnergiePlus_EP-2026-00891.pdf` | Energie Plus SA | EP-2026-00891 | 02/05/2026 | 1 488,60 |

Les deux fournisseurs existent deja dans la feuille **Fournisseurs** du
classeur : l'extraction doit donc retrouver le nom du fournisseur et la
colonne « Fournisseur connu ? » doit afficher `Oui`.

## Scenario de test recommande

1. Envoyez-vous par email les deux PDF (ou deposez-les dans le dossier
   Outlook surveille indique dans la feuille **Parametres**).
2. Bouton **1 - Importer les factures** : deux lignes apparaissent dans
   **Suivi Factures**, statut `Donnees a confirmer`, champs surlignes en
   jaune.
3. Verifiez les montants extraits, choisissez le **Type de depense**
   (ex. `Travaux` pour ACME, `Societe` pour Energie Plus).
4. Bouton **2 - Confirmer les donnees** : surlignages retires, routage
   applique, statut `En attente signature`.
5. Bouton **4 - Preparer la signature** : recapitulatif affiche, PDF copie
   dans le dossier reseau « A signer », Docusign s'ouvre.
6. Pour tester le retour : envoyez-vous un email dont l'objet contient
   `Docusign` et `complete`, et le numero de facture (ex. `2026-0412`)
   dans le corps. Bouton **5 - Detecter les retours Docusign** : la facture
   passe a `Signee`.

## Cas de routage a verifier

- ACME `Travaux` 22 200 EUR  -> 2 signataires (Responsable Developpement +
  Asset Manager).
- Une facture `Travaux` > 25 000 EUR -> 3 signataires.
- Une facture `Societe` > 50 000 EUR -> 3 signataires.

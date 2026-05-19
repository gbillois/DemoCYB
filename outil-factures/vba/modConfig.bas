Attribute VB_Name = "modConfig"
Option Explicit

' ==========================================================================
' modConfig - Constantes et acces aux feuilles / parametres
' Aucune reference externe : tout est en liaison tardive (CreateObject).
' ==========================================================================

' --- Noms des feuilles ----------------------------------------------------
Public Const FEUILLE_TDB As String = "Tableau de bord"
Public Const FEUILLE_FACTURES As String = "Suivi Factures"
Public Const FEUILLE_PAIEMENTS As String = "Suivi Paiements"
Public Const FEUILLE_FOURNISSEURS As String = "Fournisseurs"
Public Const FEUILLE_ONBOARDING As String = "Onboarding Fournisseurs"
Public Const FEUILLE_REGLES As String = "Regles Signataires"
Public Const FEUILLE_PARAMETRES As String = "Parametres"
Public Const FEUILLE_MODELES As String = "Modeles Emails"
Public Const FEUILLE_JOURNAL As String = "Journal"

' --- Colonnes : Suivi Factures --------------------------------------------
Public Const F_ID As Long = 1
Public Const F_DATE_RECEPTION As Long = 2
Public Const F_FOURNISSEUR As Long = 3
Public Const F_NUM_FACTURE As Long = 4
Public Const F_DATE_FACTURE As Long = 5
Public Const F_HT As Long = 6
Public Const F_TVA As Long = 7
Public Const F_TTC As Long = 8
Public Const F_TYPE As Long = 9
Public Const F_ENTITE As Long = 10
Public Const F_CONNU As Long = 11
Public Const F_SIGNATAIRES As Long = 12
Public Const F_PDF As Long = 13
Public Const F_STATUT As Long = 14
Public Const F_DATE_ENVOI As Long = 15
Public Const F_DATE_SIGN As Long = 16
Public Const F_EMPLACEMENT As Long = 17
Public Const F_EMAIL_ID As Long = 18
Public Const F_REMARQUES As Long = 19

' --- Colonnes : Suivi Paiements -------------------------------------------
Public Const P_ID As Long = 1
Public Const P_DATE As Long = 2
Public Const P_FOURNISSEUR As Long = 3
Public Const P_SOCIETE As Long = 4
Public Const P_FACTURES As Long = 5
Public Const P_MONTANT As Long = 6
Public Const P_COMPTE As Long = 7
Public Const P_SOLDE As Long = 8
Public Const P_SIGNATAIRES As Long = 9
Public Const P_STATUT As Long = 10
Public Const P_DATE_PAIE As Long = 11
Public Const P_REMARQUES As Long = 12

' --- Colonnes : Fournisseurs ----------------------------------------------
Public Const FR_ID As Long = 1
Public Const FR_NOM As Long = 2
Public Const FR_SIRET As Long = 3
Public Const FR_IBAN As Long = 4
Public Const FR_EMAIL As Long = 5
Public Const FR_STATUT As Long = 6
Public Const FR_DATE As Long = 7
Public Const FR_REMARQUES As Long = 8

' --- Colonnes : Onboarding Fournisseurs -----------------------------------
Public Const OB_ID As Long = 1
Public Const OB_DATE As Long = 2
Public Const OB_NOM As Long = 3
Public Const OB_EMAIL As Long = 4
Public Const OB_KBIS As Long = 5
Public Const OB_RIB As Long = 6
Public Const OB_CNI As Long = 7
Public Const OB_CALLBACK As Long = 8
Public Const OB_TRUST As Long = 9
Public Const OB_ANNEXE7 As Long = 10
Public Const OB_STATUT As Long = 11
Public Const OB_DATE_VALID As Long = 12
Public Const OB_REMARQUES As Long = 13

' --- Colonnes : Regles Signataires ----------------------------------------
Public Const RG_TYPE As Long = 1
Public Const RG_DEST As Long = 2
Public Const RG_MIN As Long = 3
Public Const RG_MAX As Long = 4
Public Const RG_SIG1 As Long = 5
Public Const RG_SIG2 As Long = 6
Public Const RG_SIG3 As Long = 7

' --- Statuts facture ------------------------------------------------------
Public Const ST_RECUE As String = "Recue"
Public Const ST_A_CONFIRMER As String = "Donnees a confirmer"
Public Const ST_ATTENTE_SIGN As String = "En attente signature"
Public Const ST_ENVOYEE As String = "Envoyee Docusign"
Public Const ST_SIGNEE As String = "Signee"
Public Const ST_CLASSEE As String = "Classee"
Public Const ST_REJETEE As String = "Rejetee"

' Couleur de surlignage des champs pre-remplis a verifier
Public Const COULEUR_A_VERIFIER As Long = 10284031   ' jaune clair (RGB 255,235,156)


' --------------------------------------------------------------------------
' Renvoie une feuille du classeur ; leve une erreur claire si absente.
' --------------------------------------------------------------------------
Public Function Feuille(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If Feuille Is Nothing Then
        Err.Raise vbObjectError + 513, "modConfig.Feuille", _
                  "Feuille introuvable : '" & nom & "'. " & _
                  "Verifiez que le classeur n'a pas ete renomme."
    End If
End Function


' --------------------------------------------------------------------------
' Derniere ligne utilisee d'une colonne (1 par defaut).
' --------------------------------------------------------------------------
Public Function DerniereLigne(ByVal ws As Worksheet, _
                               Optional ByVal col As Long = 1) As Long
    DerniereLigne = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function


' --------------------------------------------------------------------------
' Lit un parametre dans la feuille Parametres (colonne A = cle, B = valeur).
' --------------------------------------------------------------------------
Public Function Param(ByVal nom As String) As String
    Dim ws As Worksheet, r As Long, dl As Long
    Set ws = Feuille(FEUILLE_PARAMETRES)
    dl = DerniereLigne(ws, 1)
    For r = 2 To dl
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = LCase$(Trim$(nom)) Then
            Param = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
    Param = ""
End Function


' --------------------------------------------------------------------------
' Genere un identifiant unique : prefixe + date + numero de ligne.
' --------------------------------------------------------------------------
Public Function NouvelId(ByVal prefixe As String, ByVal ws As Worksheet, _
                          ByVal colId As Long) As String
    Dim n As Long
    n = DerniereLigne(ws, colId)   ' inclut la ligne d'en-tete
    NouvelId = prefixe & "-" & Format$(Now, "yyyymmdd") & "-" & Format$(n, "0000")
End Function

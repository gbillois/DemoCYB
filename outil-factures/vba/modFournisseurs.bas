Attribute VB_Name = "modFournisseurs"
Option Explicit

' ==========================================================================
' modFournisseurs - Verification et onboarding des fournisseurs
' ==========================================================================

' --------------------------------------------------------------------------
' Vrai si le fournisseur figure dans le referentiel.
' --------------------------------------------------------------------------
Public Function FournisseurExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet, r As Long, dl As Long
    nom = LCase$(Trim$(nom))
    If Len(nom) = 0 Then Exit Function
    Set ws = Feuille(FEUILLE_FOURNISSEURS)
    dl = DerniereLigne(ws, FR_NOM)
    For r = 2 To dl
        If LCase$(Trim$(CStr(ws.Cells(r, FR_NOM).Value))) = nom Then
            FournisseurExiste = True
            Exit Function
        End If
    Next r
End Function


' --------------------------------------------------------------------------
' Renvoie le statut d'onboarding d'un fournisseur, "" s'il est inconnu.
' --------------------------------------------------------------------------
Public Function StatutFournisseur(ByVal nom As String) As String
    Dim ws As Worksheet, r As Long, dl As Long
    nom = LCase$(Trim$(nom))
    Set ws = Feuille(FEUILLE_FOURNISSEURS)
    dl = DerniereLigne(ws, FR_NOM)
    For r = 2 To dl
        If LCase$(Trim$(CStr(ws.Cells(r, FR_NOM).Value))) = nom Then
            StatutFournisseur = CStr(ws.Cells(r, FR_STATUT).Value)
            Exit Function
        End If
    Next r
End Function


' --------------------------------------------------------------------------
' Renseigne la colonne 'Fournisseur connu ?' d'une ligne de Suivi Factures.
' --------------------------------------------------------------------------
Public Sub VerifierFournisseur(ByVal ligne As Long)
    Dim ws As Worksheet, nom As String
    Set ws = Feuille(FEUILLE_FACTURES)
    nom = Trim$(CStr(ws.Cells(ligne, F_FOURNISSEUR).Value))
    If Len(nom) = 0 Then
        ws.Cells(ligne, F_CONNU).Value = ""
    ElseIf FournisseurExiste(nom) Then
        ws.Cells(ligne, F_CONNU).Value = "Oui"
    Else
        ws.Cells(ligne, F_CONNU).Value = "A creer"
        ws.Cells(ligne, F_CONNU).Interior.Color = COULEUR_A_VERIFIER
    End If
End Sub


' --------------------------------------------------------------------------
' ACTION : lance une demande de creation de fournisseur (process Annexe 7).
' Cree une ligne d'onboarding et un brouillon d'email de demande de pieces.
' --------------------------------------------------------------------------
Public Sub CreerDemandeFournisseur()
    Dim ws As Worksheet, r As Long
    Dim nom As String, email As String, objet As String, corps As String

    On Error GoTo Erreur
    nom = Trim$(InputBox("Nom du fournisseur a referencer :", _
                         "Nouveau fournisseur"))
    If Len(nom) = 0 Then Exit Sub
    email = Trim$(InputBox("Email de contact du fournisseur :", _
                           "Nouveau fournisseur"))

    Set ws = Feuille(FEUILLE_ONBOARDING)
    r = DerniereLigne(ws, OB_ID) + 1
    ws.Cells(r, OB_ID).Value = NouvelId("OB", ws, OB_ID)
    ws.Cells(r, OB_DATE).Value = Date
    ws.Cells(r, OB_NOM).Value = nom
    ws.Cells(r, OB_EMAIL).Value = email
    ws.Cells(r, OB_KBIS).Value = "Non"
    ws.Cells(r, OB_RIB).Value = "Non"
    ws.Cells(r, OB_CNI).Value = "Non"
    ws.Cells(r, OB_CALLBACK).Value = "Non"
    ws.Cells(r, OB_TRUST).Value = "Non"
    ws.Cells(r, OB_ANNEXE7).Value = "Non"
    ws.Cells(r, OB_STATUT).Value = "Demande envoyee"

    LireModele "DEMANDE_INFOS_FOURNISSEUR", objet, corps
    If Len(objet) = 0 Then objet = "Creation de compte fournisseur - pieces a fournir"
    CreerBrouillonEmail email, objet, corps

    Journaliser "Onboarding fournisseur", "Demande creee pour : " & nom
    modUI.RafraichirTableauBord
    MsgBox "Demande d'onboarding creee pour '" & nom & "'." & vbCrLf & _
           "Un brouillon d'email de demande de pieces a ete ouvert dans Outlook.", _
           vbInformation, "Nouveau fournisseur"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Nouveau fournisseur"
    Journaliser "ERREUR Onboarding", Err.Description
End Sub


' --------------------------------------------------------------------------
' Recalcule le statut d'une ligne d'onboarding selon les pieces recues.
' --------------------------------------------------------------------------
Public Sub MettreAJourStatutOnboarding(ByVal ligne As Long)
    Dim ws As Worksheet, c As Long, nbOui As Long, nbPieces As Long

    Set ws = Feuille(FEUILLE_ONBOARDING)
    nbPieces = 0
    nbOui = 0
    For c = OB_KBIS To OB_ANNEXE7
        nbPieces = nbPieces + 1
        If LCase$(Trim$(CStr(ws.Cells(ligne, c).Value))) = "oui" Then
            nbOui = nbOui + 1
        End If
    Next c

    If nbOui = nbPieces Then
        ws.Cells(ligne, OB_STATUT).Value = "Valide"
        ws.Cells(ligne, OB_DATE_VALID).Value = Date
        AjouterFournisseurReferentiel ligne
    ElseIf nbOui = 0 Then
        ws.Cells(ligne, OB_STATUT).Value = "Demande envoyee"
    ElseIf LCase$(Trim$(CStr(ws.Cells(ligne, OB_ANNEXE7).Value))) = "oui" Then
        ws.Cells(ligne, OB_STATUT).Value = "Pieces completes"
    Else
        ws.Cells(ligne, OB_STATUT).Value = "Pieces en attente"
    End If
    Journaliser "Onboarding fournisseur", "Ligne " & ligne & " : statut = " & _
                CStr(ws.Cells(ligne, OB_STATUT).Value)
End Sub


' --- Helpers --------------------------------------------------------------

' Bascule un fournisseur valide vers le referentiel Fournisseurs.
Private Sub AjouterFournisseurReferentiel(ByVal ligneOb As Long)
    Dim wob As Worksheet, wfr As Worksheet, r As Long, nom As String
    Set wob = Feuille(FEUILLE_ONBOARDING)
    nom = Trim$(CStr(wob.Cells(ligneOb, OB_NOM).Value))
    If FournisseurExiste(nom) Then Exit Sub
    Set wfr = Feuille(FEUILLE_FOURNISSEURS)
    r = DerniereLigne(wfr, FR_ID) + 1
    wfr.Cells(r, FR_ID).Value = NouvelId("FRN", wfr, FR_ID)
    wfr.Cells(r, FR_NOM).Value = nom
    wfr.Cells(r, FR_EMAIL).Value = CStr(wob.Cells(ligneOb, OB_EMAIL).Value)
    wfr.Cells(r, FR_STATUT).Value = "Complet"
    wfr.Cells(r, FR_DATE).Value = Date
End Sub

' Lit un modele d'email (objet + corps) par son code.
Public Sub LireModele(ByVal code As String, ByRef objet As String, _
                       ByRef corps As String)
    Dim ws As Worksheet, r As Long, dl As Long
    objet = ""
    corps = ""
    Set ws = Feuille(FEUILLE_MODELES)
    dl = DerniereLigne(ws, 1)
    For r = 2 To dl
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(Trim$(code)) Then
            objet = CStr(ws.Cells(r, 2).Value)
            corps = CStr(ws.Cells(r, 3).Value)
            Exit Sub
        End If
    Next r
End Sub

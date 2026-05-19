Attribute VB_Name = "modUI"
Option Explicit

' ==========================================================================
' modUI - Interface : boutons d'action, tableau de bord, ouverture Docusign
'
' Les procedures 'Btn_*' sont concues pour etre assignees a des boutons.
' Lancez 'InstallerBoutons' une fois pour creer les boutons automatiquement.
' ==========================================================================

' --------------------------------------------------------------------------
' Recalcule le tableau de bord (les compteurs sont des formules COUNTIF).
' --------------------------------------------------------------------------
Public Sub RafraichirTableauBord()
    On Error Resume Next
    Application.Calculate
End Sub


' --------------------------------------------------------------------------
' Ouvre le site Docusign dans le navigateur.
' --------------------------------------------------------------------------
Public Sub OuvrirDocusign()
    Dim url As String
    url = Param("URL Docusign")
    If Len(url) = 0 Then url = "https://account.docusign.com"
    On Error Resume Next
    ThisWorkbook.FollowHyperlink url
End Sub


' --------------------------------------------------------------------------
' Renvoie la ligne active si l'on est bien sur la feuille attendue, sinon 0.
' --------------------------------------------------------------------------
Private Function LigneSelectionnee(ByVal nomFeuille As String) As Long
    If ActiveSheet.Name <> nomFeuille Then
        MsgBox "Placez-vous d'abord sur une ligne de la feuille '" & _
               nomFeuille & "', puis relancez l'action.", vbExclamation, "Action"
        Exit Function
    End If
    If ActiveCell.Row <= 1 Then
        MsgBox "Selectionnez une ligne de donnees (pas la ligne d'en-tete).", _
               vbExclamation, "Action"
        Exit Function
    End If
    LigneSelectionnee = ActiveCell.Row
End Function


' --- Boutons : process signature -----------------------------------------

Public Sub Btn_ImporterFactures()
    modOutlook.ImporterFactures
End Sub

Public Sub Btn_ConfirmerDonnees()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_FACTURES)
    If r > 0 Then
        modSuivi.ConfirmerDonnees r
        RafraichirTableauBord
    End If
End Sub

Public Sub Btn_AppliquerRoutage()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_FACTURES)
    If r > 0 Then modRouteur.AppliquerRoutage r
End Sub

Public Sub Btn_PreparerDocusign()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_FACTURES)
    If r > 0 Then modSuivi.PreparerDocusignFacture r
End Sub

Public Sub Btn_DetecterRetoursDocusign()
    modOutlook.DetecterRetoursDocusign
End Sub

Public Sub Btn_MarquerSignee()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_FACTURES)
    If r > 0 Then
        modSuivi.MarquerSignee r
        RafraichirTableauBord
    End If
End Sub

Public Sub Btn_ClasserFacture()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_FACTURES)
    If r > 0 Then
        modSuivi.ClasserFacture r
        RafraichirTableauBord
    End If
End Sub


' --- Boutons : process fournisseur ---------------------------------------

Public Sub Btn_CreerDemandeFournisseur()
    modFournisseurs.CreerDemandeFournisseur
End Sub

Public Sub Btn_MajOnboarding()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_ONBOARDING)
    If r > 0 Then
        modFournisseurs.MettreAJourStatutOnboarding r
        RafraichirTableauBord
    End If
End Sub


' --- Boutons : process paiement ------------------------------------------

Public Sub Btn_CreerLotPaiement()
    modPaiements.CreerLotPaiement
End Sub

Public Sub Btn_GenererAnnexe()
    Dim r As Long
    r = LigneSelectionnee(FEUILLE_PAIEMENTS)
    If r > 0 Then modPaiements.GenererAnnexePaiement r
End Sub

Public Sub Btn_Rafraichir()
    RafraichirTableauBord
End Sub


' --------------------------------------------------------------------------
' Cree (ou recree) les boutons d'action sur le Tableau de bord.
' A lancer une seule fois apres l'import des modules.
' --------------------------------------------------------------------------
Public Sub InstallerBoutons()
    Dim ws As Worksheet, b As Object
    Dim defs As Variant, i As Long
    Dim x As Double, y As Double, larg As Double, haut As Double, pas As Double

    Set ws = Feuille(FEUILLE_TDB)
    ws.Buttons.Delete

    defs = Array( _
        Array("1 - Importer les factures", "Btn_ImporterFactures"), _
        Array("2 - Confirmer les donnees (ligne)", "Btn_ConfirmerDonnees"), _
        Array("3 - Appliquer le routage (ligne)", "Btn_AppliquerRoutage"), _
        Array("4 - Preparer la signature (ligne)", "Btn_PreparerDocusign"), _
        Array("5 - Detecter les retours Docusign", "Btn_DetecterRetoursDocusign"), _
        Array("6 - Marquer signee (ligne)", "Btn_MarquerSignee"), _
        Array("7 - Classer la facture (ligne)", "Btn_ClasserFacture"), _
        Array("Nouveau fournisseur", "Btn_CreerDemandeFournisseur"), _
        Array("MaJ onboarding (ligne)", "Btn_MajOnboarding"), _
        Array("Creer un lot de paiement", "Btn_CreerLotPaiement"), _
        Array("Generer annexe paiement (ligne)", "Btn_GenererAnnexe"), _
        Array("Rafraichir le tableau de bord", "Btn_Rafraichir"))

    x = 360
    y = 30
    larg = 240
    haut = 26
    pas = 30
    For i = LBound(defs) To UBound(defs)
        Set b = ws.Buttons.Add(x, y + i * pas, larg, haut)
        b.Caption = defs(i)(0)
        b.OnAction = defs(i)(1)
        b.Font.Size = 9
    Next i

    MsgBox "Boutons installes sur le Tableau de bord.", vbInformation, _
           "Installation"
End Sub

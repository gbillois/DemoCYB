Attribute VB_Name = "modSuivi"
Option Explicit

' ==========================================================================
' modSuivi - Transitions de statut des factures et classement des fichiers
' ==========================================================================

' --------------------------------------------------------------------------
' Confirme les donnees pre-remplies d'une ligne : retire les surlignages,
' rafraichit fournisseur + routage et passe en 'En attente signature'.
' --------------------------------------------------------------------------
Public Sub ConfirmerDonnees(ByVal ligne As Long)
    Dim ws As Worksheet, c As Long
    Set ws = Feuille(FEUILLE_FACTURES)

    VerifierFournisseur ligne
    modRouteur.AppliquerRoutage ligne

    For c = F_FOURNISSEUR To F_REMARQUES
        ws.Cells(ligne, c).Interior.ColorIndex = xlNone
    Next c

    ws.Cells(ligne, F_STATUT).Value = ST_ATTENTE_SIGN
    Journaliser "Confirmation donnees", "Facture " & _
                CStr(ws.Cells(ligne, F_ID).Value) & " confirmee."
End Sub


' --------------------------------------------------------------------------
' Prepare l'envoi en signature : stockage reseau du PDF, recapitulatif des
' signataires, ouverture de Docusign, passage au statut 'Envoyee Docusign'.
' L'enveloppe Docusign reste creee manuellement (choix valide).
' --------------------------------------------------------------------------
Public Sub PreparerDocusignFacture(ByVal ligne As Long)
    Dim ws As Worksheet, dossierASigner As String, cheminFinal As String
    Dim recap As String, signataires As String

    On Error GoTo Erreur
    Set ws = Feuille(FEUILLE_FACTURES)

    ' calcule le routage si necessaire
    signataires = Trim$(CStr(ws.Cells(ligne, F_SIGNATAIRES).Value))
    If Len(signataires) = 0 Or InStr(signataires, "(") > 0 Then
        modRouteur.AppliquerRoutage ligne
        signataires = Trim$(CStr(ws.Cells(ligne, F_SIGNATAIRES).Value))
    End If

    ' stockage du PDF sur le reseau (dossier A signer)
    dossierASigner = Param("Dossier reseau - A signer")
    cheminFinal = CopierFichier(CStr(ws.Cells(ligne, F_PDF).Value), dossierASigner)
    If Len(cheminFinal) > 0 Then
        ws.Cells(ligne, F_EMPLACEMENT).Value = cheminFinal
    Else
        AjouterRemarqueFacture ligne, "Copie reseau 'A signer' echouee - a faire manuellement."
    End If

    ' recapitulatif a copier dans l'enveloppe Docusign
    recap = "RECAPITULATIF POUR L'ENVELOPPE DOCUSIGN" & vbCrLf & _
            "----------------------------------------" & vbCrLf & _
            "Facture    : " & CStr(ws.Cells(ligne, F_ID).Value) & vbCrLf & _
            "Fournisseur: " & CStr(ws.Cells(ligne, F_FOURNISSEUR).Value) & vbCrLf & _
            "N0 facture : " & CStr(ws.Cells(ligne, F_NUM_FACTURE).Value) & vbCrLf & _
            "Montant TTC: " & CStr(ws.Cells(ligne, F_TTC).Value) & vbCrLf & _
            "Type       : " & CStr(ws.Cells(ligne, F_TYPE).Value) & vbCrLf & _
            "Signataires: " & signataires & vbCrLf & _
            "Fichier    : " & CStr(ws.Cells(ligne, F_EMPLACEMENT).Value)

    ws.Cells(ligne, F_STATUT).Value = ST_ENVOYEE
    ws.Cells(ligne, F_DATE_ENVOI).Value = Now
    Journaliser "Envoi Docusign", "Facture " & CStr(ws.Cells(ligne, F_ID).Value) & _
                " - signataires : " & signataires

    If MsgBox(recap & vbCrLf & vbCrLf & _
              "Ouvrir Docusign maintenant pour creer l'enveloppe ?", _
              vbQuestion + vbYesNo, "Preparer la signature") = vbYes Then
        modUI.OuvrirDocusign
    End If
    modUI.RafraichirTableauBord
    Exit Sub
Erreur:
    MsgBox "Erreur lors de la preparation : " & Err.Description, vbCritical, _
           "Preparer la signature"
    Journaliser "ERREUR Envoi Docusign", Err.Description
End Sub


' --------------------------------------------------------------------------
' Marque une facture comme signee : copie du PDF dans le dossier 'Signees'.
' --------------------------------------------------------------------------
Public Sub MarquerSignee(ByVal ligne As Long)
    Dim ws As Worksheet, dossierSignees As String, cheminFinal As String
    Set ws = Feuille(FEUILLE_FACTURES)

    dossierSignees = Param("Dossier reseau - Signees")
    cheminFinal = CopierFichier(CStr(ws.Cells(ligne, F_PDF).Value), dossierSignees)
    If Len(cheminFinal) > 0 Then
        ws.Cells(ligne, F_PDF).Value = cheminFinal
        ws.Cells(ligne, F_EMPLACEMENT).Value = cheminFinal
    Else
        AjouterRemarqueFacture ligne, "Copie reseau 'Signees' echouee - a faire manuellement."
    End If

    ws.Cells(ligne, F_STATUT).Value = ST_SIGNEE
    ws.Cells(ligne, F_DATE_SIGN).Value = Now
    Journaliser "Facture signee", "Facture " & CStr(ws.Cells(ligne, F_ID).Value) & _
                " classee dans 'Signees'."
End Sub


' --------------------------------------------------------------------------
' Classe definitivement une facture signee : deplacement vers les archives.
' --------------------------------------------------------------------------
Public Sub ClasserFacture(ByVal ligne As Long)
    Dim ws As Worksheet, dossierArchives As String, cheminFinal As String
    Set ws = Feuille(FEUILLE_FACTURES)

    dossierArchives = Param("Dossier reseau - Archives")
    cheminFinal = DeplacerFichier(CStr(ws.Cells(ligne, F_PDF).Value), dossierArchives)
    If Len(cheminFinal) > 0 Then
        ws.Cells(ligne, F_PDF).Value = cheminFinal
        ws.Cells(ligne, F_EMPLACEMENT).Value = cheminFinal
    Else
        AjouterRemarqueFacture ligne, "Deplacement vers archives echoue - a faire manuellement."
    End If

    ws.Cells(ligne, F_STATUT).Value = ST_CLASSEE
    Journaliser "Facture classee", "Facture " & CStr(ws.Cells(ligne, F_ID).Value) & _
                " archivee."
End Sub


' --- Helper ---------------------------------------------------------------

Private Sub AjouterRemarqueFacture(ByVal ligne As Long, ByVal texte As String)
    Dim ws As Worksheet, actuel As String
    Set ws = Feuille(FEUILLE_FACTURES)
    actuel = CStr(ws.Cells(ligne, F_REMARQUES).Value)
    If Len(actuel) > 0 Then
        ws.Cells(ligne, F_REMARQUES).Value = actuel & " | " & texte
    Else
        ws.Cells(ligne, F_REMARQUES).Value = texte
    End If
End Sub

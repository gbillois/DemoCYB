Attribute VB_Name = "modOutlook"
Option Explicit

' ==========================================================================
' modOutlook - Ingestion des emails de factures et detection des retours
' Outlook est pilote via son modele objet COM (liaison tardive).
' ==========================================================================

' --------------------------------------------------------------------------
' Recupere une instance d'Outlook (existante ou nouvelle).
' --------------------------------------------------------------------------
Public Function ObtenirApplicationOutlook() As Object
    On Error Resume Next
    Set ObtenirApplicationOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If ObtenirApplicationOutlook Is Nothing Then
        Set ObtenirApplicationOutlook = CreateObject("Outlook.Application")
    End If
End Function


' --------------------------------------------------------------------------
' Recherche recursive d'un dossier Outlook par son nom.
' --------------------------------------------------------------------------
Public Function TrouverDossier(ByVal racine As Object, _
                                 ByVal nom As String) As Object
    Dim f As Object, res As Object
    On Error Resume Next
    For Each f In racine.Folders
        If LCase$(f.Name) = LCase$(nom) Then
            Set TrouverDossier = f
            Exit Function
        End If
        Set res = TrouverDossier(f, nom)
        If Not res Is Nothing Then
            Set TrouverDossier = res
            Exit Function
        End If
    Next f
End Function


' --------------------------------------------------------------------------
' Renvoie le dossier Outlook surveille (defini dans Parametres).
' --------------------------------------------------------------------------
Private Function DossierSurveille(ByVal ol As Object) As Object
    Dim ns As Object, store As Object, nom As String
    nom = Param("Dossier Outlook surveille")
    If Len(nom) = 0 Then nom = "Boite de reception"
    Set ns = ol.GetNamespace("MAPI")
    For Each store In ns.Folders
        Set DossierSurveille = TrouverDossier(store, nom)
        If Not DossierSurveille Is Nothing Then Exit Function
    Next store
    ' repli : boite de reception par defaut
    On Error Resume Next
    Set DossierSurveille = ns.GetDefaultFolder(6)   ' olFolderInbox
End Function


' --------------------------------------------------------------------------
' ACTION : importe les factures PDF non traitees du dossier surveille.
' Pour chaque PDF : enregistrement local, creation d'une ligne de suivi,
' extraction des donnees, verification du fournisseur, marquage de l'email.
' --------------------------------------------------------------------------
Public Sub ImporterFactures()
    Dim ol As Object, dossier As Object, itm As Object, att As Object
    Dim ws As Worksheet, dossierLocal As String, categorie As String
    Dim nbImporte As Long, nbMails As Long, cheminPDF As String
    Dim r As Long

    On Error GoTo Erreur
    Application.ScreenUpdating = False

    categorie = Param("Categorie Outlook - traite")
    If Len(categorie) = 0 Then categorie = "Traitee par l'outil"
    dossierLocal = Param("Dossier local extraction PDF")
    If Not AssurerDossier(dossierLocal) Then
        MsgBox "Impossible de creer le dossier local d'extraction :" & vbCrLf & _
               dossierLocal, vbExclamation, "Import des factures"
        GoTo Fin
    End If

    Set ol = ObtenirApplicationOutlook()
    Set dossier = DossierSurveille(ol)
    If dossier Is Nothing Then
        MsgBox "Dossier Outlook surveille introuvable. Verifiez le parametre" & _
               " 'Dossier Outlook surveille'.", vbExclamation, "Import des factures"
        GoTo Fin
    End If

    Set ws = Feuille(FEUILLE_FACTURES)

    For Each itm In dossier.Items
        If TypeName(itm) = "MailItem" Then
            nbMails = nbMails + 1
            If InStr(1, "|" & itm.Categories & "|", categorie, vbTextCompare) = 0 Then
                For Each att In itm.Attachments
                    If EstPDF(att.FileName) Then
                        cheminPDF = TerminerParSeparateur(dossierLocal) & _
                                    NomFichierUnique(dossierLocal, att.FileName)
                        att.SaveAsFile cheminPDF
                        r = DerniereLigne(ws, F_ID) + 1
                        ws.Cells(r, F_ID).Value = NouvelId("FAC", ws, F_ID)
                        ws.Cells(r, F_DATE_RECEPTION).Value = itm.ReceivedTime
                        ws.Cells(r, F_PDF).Value = cheminPDF
                        ws.Cells(r, F_STATUT).Value = ST_RECUE
                        ws.Cells(r, F_EMAIL_ID).Value = itm.EntryID
                        ExtraireDonneesFacture r
                        VerifierFournisseur r
                        nbImporte = nbImporte + 1
                    End If
                Next att
                ' marque l'email comme traite
                If Len(itm.Categories) = 0 Then
                    itm.Categories = categorie
                Else
                    itm.Categories = itm.Categories & "," & categorie
                End If
                itm.Save
            End If
        End If
    Next itm

    Journaliser "Import factures", nbImporte & " facture(s) importee(s) sur " & _
                nbMails & " email(s) examines."
    modUI.RafraichirTableauBord
    MsgBox nbImporte & " facture(s) importee(s)." & vbCrLf & vbCrLf & _
           "Verifiez les champs surlignes en jaune dans 'Suivi Factures'.", _
           vbInformation, "Import des factures"

Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'import : " & Err.Description, vbCritical, _
           "Import des factures"
    Journaliser "ERREUR Import factures", Err.Description
End Sub


' --------------------------------------------------------------------------
' ACTION : detecte les emails de retour Docusign (enveloppes signees) et
' met a jour les factures correspondantes au statut 'Signee'.
' --------------------------------------------------------------------------
Public Sub DetecterRetoursDocusign()
    Dim ol As Object, ns As Object, inbox As Object, itm As Object
    Dim ws As Worksheet, r As Long, dl As Long
    Dim filtreExp As String, texteMail As String
    Dim numFac As String, fournisseur As String
    Dim nbTraite As Long, nbNonRattache As Long

    On Error GoTo Erreur
    Application.ScreenUpdating = False

    filtreExp = LCase$(Param("Expediteur retour Docusign (filtre)"))
    Set ol = ObtenirApplicationOutlook()
    Set ns = ol.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(6)   ' olFolderInbox
    Set ws = Feuille(FEUILLE_FACTURES)
    dl = DerniereLigne(ws, F_ID)

    For Each itm In inbox.Items
        If TypeName(itm) = "MailItem" Then
            If EstRetourDocusignSigne(itm, filtreExp) Then
                texteMail = LCase$(itm.Subject & " " & itm.Body)
                Dim rattache As Boolean
                rattache = False
                For r = 2 To dl
                    If CStr(ws.Cells(r, F_STATUT).Value) = ST_ENVOYEE Then
                        numFac = LCase$(Trim$(CStr(ws.Cells(r, F_NUM_FACTURE).Value)))
                        fournisseur = LCase$(Trim$(CStr(ws.Cells(r, F_FOURNISSEUR).Value)))
                        If (Len(numFac) > 2 And InStr(texteMail, numFac) > 0) _
                           Or (Len(fournisseur) > 2 And InStr(texteMail, fournisseur) > 0) Then
                            modSuivi.MarquerSignee r
                            nbTraite = nbTraite + 1
                            rattache = True
                            Exit For
                        End If
                    End If
                Next r
                If Not rattache Then nbNonRattache = nbNonRattache + 1
            End If
        End If
    Next itm

    Journaliser "Retours Docusign", nbTraite & " facture(s) passee(s) a 'Signee', " & _
                nbNonRattache & " retour(s) non rattache(s)."
    modUI.RafraichirTableauBord
    MsgBox nbTraite & " facture(s) marquee(s) signee(s)." & vbCrLf & _
           nbNonRattache & " retour(s) Docusign n'ont pas pu etre rattache(s) " & _
           "automatiquement (a verifier manuellement).", vbInformation, _
           "Retours Docusign"

Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la detection des retours : " & Err.Description, _
           vbCritical, "Retours Docusign"
    Journaliser "ERREUR Retours Docusign", Err.Description
End Sub


' --------------------------------------------------------------------------
' Cree un brouillon d'email et l'affiche pour relecture par l'utilisatrice.
' --------------------------------------------------------------------------
Public Sub CreerBrouillonEmail(ByVal destinataire As String, _
                                ByVal objet As String, _
                                ByVal corps As String, _
                                Optional ByVal pieceJointe As String = "")
    Dim ol As Object, mail As Object
    Set ol = ObtenirApplicationOutlook()
    Set mail = ol.CreateItem(0)   ' olMailItem
    mail.To = destinataire
    mail.Subject = objet
    mail.Body = corps
    If Len(pieceJointe) > 0 Then
        If FichierExiste(pieceJointe) Then mail.Attachments.Add pieceJointe
    End If
    mail.Display   ' affiche le brouillon ; l'envoi reste manuel
End Sub


' --- Helpers --------------------------------------------------------------

Private Function EstPDF(ByVal nomFichier As String) As Boolean
    EstPDF = (LCase$(Right$(nomFichier, 4)) = ".pdf")
End Function

Private Function EstRetourDocusignSigne(ByVal mail As Object, _
                                         ByVal filtreExp As String) As Boolean
    Dim expediteur As String, suj As String
    On Error Resume Next
    expediteur = LCase$(mail.SenderEmailAddress)
    suj = LCase$(mail.Subject)
    On Error GoTo 0
    If Len(filtreExp) > 0 Then
        If InStr(expediteur, filtreExp) = 0 And InStr(suj, "docusign") = 0 Then
            Exit Function
        End If
    End If
    EstRetourDocusignSigne = (InStr(suj, "complet") > 0 _
                              Or InStr(suj, "completed") > 0 _
                              Or InStr(suj, "signe") > 0 _
                              Or InStr(suj, "termine") > 0)
End Function

Private Function TerminerParSeparateur(ByVal chemin As String) As String
    If Right$(chemin, 1) = "\" Then
        TerminerParSeparateur = chemin
    Else
        TerminerParSeparateur = chemin & "\"
    End If
End Function

Private Function NomFichierUnique(ByVal dossier As String, _
                                   ByVal nomFichier As String) As String
    Dim base As String, ext As String, cible As String, i As Long, p As Long
    p = InStrRev(nomFichier, ".")
    If p > 0 Then
        base = Left$(nomFichier, p - 1)
        ext = Mid$(nomFichier, p)
    Else
        base = nomFichier
        ext = ""
    End If
    cible = base & ext
    i = 1
    Do While FichierExiste(TerminerParSeparateur(dossier) & cible)
        cible = base & " (" & i & ")" & ext
        i = i + 1
    Loop
    NomFichierUnique = cible
End Function

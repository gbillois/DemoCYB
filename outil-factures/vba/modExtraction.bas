Attribute VB_Name = "modExtraction"
Option Explicit

' ==========================================================================
' modExtraction - Pipeline d'extraction des donnees de facture depuis le PDF
'
'   Niveau 1 (par defaut) : lecture de la couche texte des PDF numeriques via
'                           Word (conversion silencieuse), puis parsing.
'   Niveau 2 (optionnel)  : OCR cloud Azure AI Document Intelligence pour les
'                           PDF scannes. Active si le parametre
'                           'Extraction OCR niveau 2 (Azure)' vaut 'Oui'.
'
' Les champs pre-remplis sont surlignes en jaune : ils doivent etre verifies
' manuellement avant l'envoi en signature.
' ==========================================================================

Private Const SEUIL_TEXTE_MINIMAL As Long = 40   ' en-deca = PDF scanne


' --------------------------------------------------------------------------
' Lit la couche texte d'un PDF numerique via Word (aucune installation).
' Renvoie "" pour un PDF scanne (image) ou en cas d'erreur.
' --------------------------------------------------------------------------
Public Function LireTextePDF(ByVal cheminPDF As String) As String
    Dim wd As Object, doc As Object
    On Error GoTo Echec
    If Not FichierExiste(cheminPDF) Then Exit Function
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    wd.DisplayAlerts = 0   ' wdAlertsNone
    Set doc = wd.Documents.Open(FileName:=cheminPDF, ConfirmConversions:=False, _
                                ReadOnly:=True, AddToRecentFiles:=False, _
                                Visible:=False)
    LireTextePDF = doc.Content.Text
    doc.Close SaveChanges:=False
    wd.Quit
    Set wd = Nothing
    Exit Function
Echec:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If Not wd Is Nothing Then wd.Quit
    LireTextePDF = ""
End Function


' --------------------------------------------------------------------------
' ACTION : extrait et pre-remplit les donnees d'une ligne de Suivi Factures.
' --------------------------------------------------------------------------
Public Sub ExtraireDonneesFacture(ByVal ligne As Long)
    Dim ws As Worksheet, cheminPDF As String, texte As String

    On Error GoTo Erreur
    Set ws = Feuille(FEUILLE_FACTURES)
    cheminPDF = CStr(ws.Cells(ligne, F_PDF).Value)
    texte = LireTextePDF(cheminPDF)

    If Len(Trim$(texte)) < SEUIL_TEXTE_MINIMAL Then
        ' PDF probablement scanne : pas de couche texte exploitable.
        If LCase$(Param("Extraction OCR niveau 2 (Azure)")) = "oui" Then
            If ExtraireViaAzure(cheminPDF, ligne) Then
                Journaliser "Extraction", "Ligne " & ligne & " : OCR Azure utilise."
                Exit Sub
            End If
        End If
        ws.Cells(ligne, F_STATUT).Value = ST_A_CONFIRMER
        AjouterRemarque ws, ligne, "Extraction auto impossible (PDF scanne). " & _
            "Saisir les champs manuellement ou via Copilot."
        Journaliser "Extraction", "Ligne " & ligne & " : PDF scanne, saisie manuelle."
        Exit Sub
    End If

    ' --- Niveau 1 : parsing de la couche texte ---
    RemplirChamp ws, ligne, F_FOURNISSEUR, DetecterFournisseur(texte)
    RemplirChamp ws, ligne, F_NUM_FACTURE, ExtraireValeurRegex(texte, _
        "facture[^\r\n]*?([A-Za-z]{0,4}\d[\dA-Za-z\-\/]{3,})", 1)
    RemplirChamp ws, ligne, F_DATE_FACTURE, ExtraireValeurRegex(texte, _
        "(\d{2}[\/\.\-]\d{2}[\/\.\-]\d{2,4})", 1)
    RemplirMontant ws, ligne, F_HT, MontantApres(texte, Array("total ht", "montant ht", "base ht"))
    RemplirMontant ws, ligne, F_TVA, MontantApres(texte, Array("montant tva", "total tva", "tva"))
    RemplirMontant ws, ligne, F_TTC, MontantApres(texte, Array("total ttc", "montant ttc", "net a payer", "total a payer"))

    ws.Cells(ligne, F_STATUT).Value = ST_A_CONFIRMER
    Journaliser "Extraction", "Ligne " & ligne & " : donnees pre-remplies (niveau 1)."
    Exit Sub
Erreur:
    Journaliser "ERREUR Extraction", "Ligne " & ligne & " : " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Cells(ligne, F_STATUT).Value = ST_A_CONFIRMER
        AjouterRemarque ws, ligne, "Erreur extraction : a saisir manuellement."
    End If
End Sub


' --------------------------------------------------------------------------
' Niveau 2 : OCR via Azure AI Document Intelligence (modele prebuilt-invoice).
' Necessite les parametres 'Azure DI - Endpoint' et une cle dans la variable
' d'environnement nommee par 'Azure DI - Variable d'env. cle'.
' Renvoie True si l'extraction a abouti.
' --------------------------------------------------------------------------
Public Function ExtraireViaAzure(ByVal cheminPDF As String, _
                                  ByVal ligne As Long) As Boolean
    Dim endpoint As String, cle As String, urlAnalyse As String
    Dim http As Object, octets() As Byte, ff As Integer
    Dim urlResultat As String, json As String, essais As Long
    Dim ws As Worksheet

    On Error GoTo Echec
    endpoint = Param("Azure DI - Endpoint")
    cle = Environ$(Param("Azure DI - Variable d'env. cle"))
    If Len(endpoint) = 0 Or Len(cle) = 0 Then Exit Function

    If Right$(endpoint, 1) = "/" Then endpoint = Left$(endpoint, Len(endpoint) - 1)
    urlAnalyse = endpoint & "/documentintelligence/documentModels/" & _
                 "prebuilt-invoice:analyze?api-version=2024-11-30"

    ' lecture du PDF en binaire
    ff = FreeFile
    Open cheminPDF For Binary Access Read As #ff
    ReDim octets(0 To LOF(ff) - 1)
    Get #ff, , octets
    Close #ff

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", urlAnalyse, False
    http.SetRequestHeader "Ocp-Apim-Subscription-Key", cle
    http.SetRequestHeader "Content-Type", "application/pdf"
    http.Send octets
    If http.Status <> 202 Then GoTo Echec
    urlResultat = http.GetResponseHeader("Operation-Location")
    If Len(urlResultat) = 0 Then GoTo Echec

    ' polling du resultat (max ~30 s)
    Do
        Application.Wait Now + TimeValue("0:00:03")
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", urlResultat, False
        http.SetRequestHeader "Ocp-Apim-Subscription-Key", cle
        http.Send
        json = http.ResponseText
        essais = essais + 1
        If InStr(json, """status"":""succeeded""") > 0 Then Exit Do
        If InStr(json, """status"":""failed""") > 0 Then GoTo Echec
    Loop While essais < 10
    If InStr(json, """status"":""succeeded""") = 0 Then GoTo Echec

    Set ws = Feuille(FEUILLE_FACTURES)
    RemplirChamp ws, ligne, F_FOURNISSEUR, ValeurChampJson(json, "VendorName")
    RemplirChamp ws, ligne, F_NUM_FACTURE, ValeurChampJson(json, "InvoiceId")
    RemplirChamp ws, ligne, F_DATE_FACTURE, ValeurChampJson(json, "InvoiceDate")
    RemplirMontant ws, ligne, F_HT, ConvertirMontant(ValeurChampJson(json, "SubTotal"))
    RemplirMontant ws, ligne, F_TVA, ConvertirMontant(ValeurChampJson(json, "TotalTax"))
    RemplirMontant ws, ligne, F_TTC, ConvertirMontant(ValeurChampJson(json, "InvoiceTotal"))
    ws.Cells(ligne, F_STATUT).Value = ST_A_CONFIRMER
    ExtraireViaAzure = True
    Exit Function
Echec:
    ExtraireViaAzure = False
End Function


' --- Helpers de parsing ---------------------------------------------------

' Cherche un nom de fournisseur connu (referentiel) present dans le texte.
Private Function DetecterFournisseur(ByVal texte As String) As String
    Dim ws As Worksheet, r As Long, dl As Long, nom As String
    Set ws = Feuille(FEUILLE_FOURNISSEURS)
    dl = DerniereLigne(ws, FR_NOM)
    For r = 2 To dl
        nom = Trim$(CStr(ws.Cells(r, FR_NOM).Value))
        If Len(nom) > 2 Then
            If InStr(1, texte, nom, vbTextCompare) > 0 Then
                DetecterFournisseur = nom
                Exit Function
            End If
        End If
    Next r
End Function

' Extrait la premiere capture d'une expression reguliere.
Public Function ExtraireValeurRegex(ByVal texte As String, _
                                     ByVal motif As String, _
                                     Optional ByVal groupe As Long = 1) As String
    Dim re As Object, corr As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = motif
    If re.Test(texte) Then
        Set corr = re.Execute(texte)(0)
        If corr.SubMatches.Count >= groupe Then
            ExtraireValeurRegex = Trim$(corr.SubMatches(groupe - 1))
        End If
    End If
End Function

' Renvoie le premier montant trouve apres l'un des mots-cles fournis.
Private Function MontantApres(ByVal texte As String, ByVal motsCles As Variant) As Double
    Dim i As Long, brut As String
    For i = LBound(motsCles) To UBound(motsCles)
        brut = ExtraireValeurRegex(texte, _
            EchapperRegex(CStr(motsCles(i))) & "[^0-9\-]{0,40}([0-9][0-9\s\.,]*[0-9])", 1)
        If Len(brut) > 0 Then
            MontantApres = ConvertirMontant(brut)
            Exit Function
        End If
    Next i
End Function

' Convertit un montant texte (formats FR/EN) en Double.
Public Function ConvertirMontant(ByVal s As String) As Double
    Dim posV As Long, posP As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    s = Replace$(s, Chr$(160), "")   ' espace insecable
    s = Replace$(s, " ", "")
    s = Replace$(s, Chr$(8364), "")  ' symbole euro
    posV = InStrRev(s, ",")
    posP = InStrRev(s, ".")
    If posV > 0 And posV > posP Then
        s = Replace$(s, ".", "")
        s = Replace$(s, ",", ".")
    ElseIf posP > 0 And posP > posV Then
        s = Replace$(s, ",", "")
    Else
        s = Replace$(s, ",", "")
    End If
    ConvertirMontant = Val(s)
End Function

' Echappe les caracteres speciaux d'une chaine pour usage dans un motif regex.
Private Function EchapperRegex(ByVal s As String) As String
    Dim i As Long, c As String, res As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr(".\+*?[]^${}()|/", c) > 0 Then res = res & "\"
        res = res & c
    Next i
    EchapperRegex = res
End Function

' Extrait la valeur 'content' d'un champ de la reponse JSON Azure DI.
Private Function ValeurChampJson(ByVal json As String, _
                                  ByVal champ As String) As String
    ValeurChampJson = ExtraireValeurRegex(json, """" & champ & _
        """\s*:\s*\{[\s\S]*?""content""\s*:\s*""([^""]*)""", 1)
End Function

' --- Helpers d'ecriture ---------------------------------------------------

' Ecrit une valeur texte et surligne la cellule (a verifier).
Private Sub RemplirChamp(ByVal ws As Worksheet, ByVal ligne As Long, _
                          ByVal col As Long, ByVal valeur As String)
    If Len(Trim$(valeur)) = 0 Then Exit Sub
    ws.Cells(ligne, col).Value = valeur
    ws.Cells(ligne, col).Interior.Color = COULEUR_A_VERIFIER
End Sub

' Ecrit un montant et surligne la cellule (a verifier).
Private Sub RemplirMontant(ByVal ws As Worksheet, ByVal ligne As Long, _
                            ByVal col As Long, ByVal valeur As Double)
    If valeur = 0 Then Exit Sub
    ws.Cells(ligne, col).Value = valeur
    ws.Cells(ligne, col).Interior.Color = COULEUR_A_VERIFIER
End Sub

' Ajoute un texte dans la colonne Remarques sans ecraser l'existant.
Private Sub AjouterRemarque(ByVal ws As Worksheet, ByVal ligne As Long, _
                             ByVal texte As String)
    Dim actuel As String
    actuel = CStr(ws.Cells(ligne, F_REMARQUES).Value)
    If Len(actuel) > 0 Then
        ws.Cells(ligne, F_REMARQUES).Value = actuel & " | " & texte
    Else
        ws.Cells(ligne, F_REMARQUES).Value = texte
    End If
End Sub

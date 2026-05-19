Attribute VB_Name = "modPaiements"
Option Explicit

' ==========================================================================
' modPaiements - Constitution des lots de paiement et des annexes
'
' La verification du solde bancaire et l'execution du virement restent
' manuelles (sites bancaires non automatisables). L'outil prepare le lot,
' calcule le total, genere l'annexe et suit le statut.
' ==========================================================================

' --------------------------------------------------------------------------
' ACTION : cree un lot de paiement a partir des lignes selectionnees dans
' 'Suivi Factures'. Toutes les factures doivent avoir le meme fournisseur.
' --------------------------------------------------------------------------
Public Sub CreerLotPaiement()
    Dim wsF As Worksheet, wsP As Worksheet, ar As Range
    Dim lignes As Object, cle As Variant, r As Long, rr As Long, dl As Long
    Dim fournisseur As String, fCourant As String
    Dim total As Double, ids As String, nb As Long

    On Error GoTo Erreur
    Set wsF = Feuille(FEUILLE_FACTURES)
    If ActiveSheet.Name <> FEUILLE_FACTURES Then
        MsgBox "Selectionnez d'abord les factures a payer dans la feuille '" & _
               FEUILLE_FACTURES & "'.", vbExclamation, "Lot de paiement"
        Exit Sub
    End If

    dl = DerniereLigne(wsF, F_ID)
    Set lignes = CreateObject("Scripting.Dictionary")
    For Each ar In Selection.Areas
        For rr = ar.Row To ar.Row + ar.Rows.Count - 1
            If rr > 1 And rr <= dl Then
                If Not lignes.Exists(rr) Then lignes.Add rr, True
            End If
        Next rr
    Next ar
    If lignes.Count = 0 Then
        MsgBox "Aucune facture selectionnee.", vbExclamation, "Lot de paiement"
        Exit Sub
    End If

    For Each cle In lignes.Keys
        r = CLng(cle)
        fCourant = Trim$(CStr(wsF.Cells(r, F_FOURNISSEUR).Value))
        If Len(fournisseur) = 0 Then
            fournisseur = fCourant
        ElseIf LCase$(fCourant) <> LCase$(fournisseur) Then
            MsgBox "Toutes les factures d'un lot doivent concerner le meme " & _
                   "fournisseur." & vbCrLf & "Detecte : '" & fournisseur & _
                   "' et '" & fCourant & "'.", vbExclamation, "Lot de paiement"
            Exit Sub
        End If
        total = total + Val(CStr(wsF.Cells(r, F_TTC).Value))
        If Len(ids) > 0 Then ids = ids & ", "
        ids = ids & CStr(wsF.Cells(r, F_ID).Value)
        nb = nb + 1
    Next cle

    Set wsP = Feuille(FEUILLE_PAIEMENTS)
    r = DerniereLigne(wsP, P_ID) + 1
    wsP.Cells(r, P_ID).Value = NouvelId("PAY", wsP, P_ID)
    wsP.Cells(r, P_DATE).Value = Date
    wsP.Cells(r, P_FOURNISSEUR).Value = fournisseur
    wsP.Cells(r, P_FACTURES).Value = ids
    wsP.Cells(r, P_MONTANT).Value = total
    wsP.Cells(r, P_SIGNATAIRES).Value = SignatairesPaiement(total)
    wsP.Cells(r, P_STATUT).Value = "En preparation"

    Journaliser "Lot de paiement", "Lot " & CStr(wsP.Cells(r, P_ID).Value) & _
                " : " & nb & " facture(s), total " & Format$(total, "# ##0.00") & _
                " EUR."
    modUI.RafraichirTableauBord
    wsP.Activate
    wsP.Cells(r, P_SOCIETE).Select
    MsgBox "Lot de paiement cree (" & nb & " facture(s), total " & _
           Format$(total, "# ##0.00") & " EUR)." & vbCrLf & _
           "Completez la societe emettrice, le compte debite et le solde.", _
           vbInformation, "Lot de paiement"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Lot de paiement"
    Journaliser "ERREUR Lot de paiement", Err.Description
End Sub


' --------------------------------------------------------------------------
' ACTION : genere un fichier annexe (recapitulatif) pour un lot de paiement.
' --------------------------------------------------------------------------
Public Sub GenererAnnexePaiement(ByVal ligne As Long)
    Dim wsP As Worksheet, fso As Object, fichier As Object
    Dim dossier As String, chemin As String, contenu As String

    On Error GoTo Erreur
    Set wsP = Feuille(FEUILLE_PAIEMENTS)
    dossier = Param("Dossier local extraction PDF")
    If Not AssurerDossier(dossier) Then
        MsgBox "Dossier de sortie indisponible : " & dossier, vbExclamation, _
               "Annexe paiement"
        Exit Sub
    End If

    contenu = "ANNEXE - LOT DE PAIEMENT" & vbCrLf & _
              "========================" & vbCrLf & _
              "Lot          : " & CStr(wsP.Cells(ligne, P_ID).Value) & vbCrLf & _
              "Date         : " & CStr(wsP.Cells(ligne, P_DATE).Value) & vbCrLf & _
              "Fournisseur  : " & CStr(wsP.Cells(ligne, P_FOURNISSEUR).Value) & vbCrLf & _
              "Societe      : " & CStr(wsP.Cells(ligne, P_SOCIETE).Value) & vbCrLf & _
              "Factures     : " & CStr(wsP.Cells(ligne, P_FACTURES).Value) & vbCrLf & _
              "Montant total: " & Format$(Val(CStr(wsP.Cells(ligne, P_MONTANT).Value)), "# ##0.00") & " EUR" & vbCrLf & _
              "Compte debite: " & CStr(wsP.Cells(ligne, P_COMPTE).Value) & vbCrLf & _
              "Solde connu  : " & CStr(wsP.Cells(ligne, P_SOLDE).Value) & vbCrLf & _
              "Signataires  : " & CStr(wsP.Cells(ligne, P_SIGNATAIRES).Value) & vbCrLf

    chemin = dossier & "\Annexe_" & CStr(wsP.Cells(ligne, P_ID).Value) & ".txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fichier = fso.CreateTextFile(chemin, True, True)   ' Unicode
    fichier.Write contenu
    fichier.Close

    wsP.Cells(ligne, P_STATUT).Value = "Annexes generees"
    Journaliser "Annexe paiement", "Annexe generee : " & chemin
    MsgBox "Annexe generee :" & vbCrLf & chemin, vbInformation, "Annexe paiement"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Annexe paiement"
    Journaliser "ERREUR Annexe paiement", Err.Description
End Sub


' --------------------------------------------------------------------------
' Signataires d'un paiement selon 3 seuils de montant (a ajuster au besoin).
' --------------------------------------------------------------------------
Public Function SignatairesPaiement(ByVal montant As Double) As String
    If montant < 10000 Then
        SignatairesPaiement = "Comptable + Asset Manager + Fund Manager"
    ElseIf montant < 100000 Then
        SignatairesPaiement = "Asset Manager + Fund Manager + Directeur General"
    Else
        SignatairesPaiement = "Fund Manager + Directeur General + President"
    End If
End Function

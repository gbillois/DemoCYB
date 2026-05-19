Attribute VB_Name = "modRouteur"
Option Explicit

' ==========================================================================
' modRouteur - Determination des signataires selon les regles metier
'
' Regle : une ligne de 'Regles Signataires' s'applique si le type de depense
' correspond et si le montant TTC est dans l'intervalle [min ; max[ (borne
' superieure exclue). Les signataires non vides sont concatenes avec ' + '.
' ==========================================================================

' --------------------------------------------------------------------------
' Renvoie la liste des signataires requis pour un type de depense + montant.
' --------------------------------------------------------------------------
Public Function DeterminerSignataires(ByVal typeDepense As String, _
                                       ByVal montant As Double) As String
    Dim ws As Worksheet, r As Long, dl As Long
    Dim t As String, vmin As Double, vmax As Double
    Dim res As String, s As String, c As Long

    Set ws = Feuille(FEUILLE_REGLES)
    dl = DerniereLigne(ws, RG_TYPE)
    typeDepense = LCase$(Trim$(typeDepense))

    For r = 2 To dl
        t = LCase$(Trim$(CStr(ws.Cells(r, RG_TYPE).Value)))
        If t = typeDepense And Len(t) > 0 Then
            vmin = Val(CStr(ws.Cells(r, RG_MIN).Value))
            vmax = Val(CStr(ws.Cells(r, RG_MAX).Value))
            If montant >= vmin And montant < vmax Then
                For c = RG_SIG1 To RG_SIG3
                    s = Trim$(CStr(ws.Cells(r, c).Value))
                    If Len(s) > 0 Then
                        If Len(res) > 0 Then res = res & " + "
                        res = res & s
                    End If
                Next c
                DeterminerSignataires = res
                Exit Function
            End If
        End If
    Next r

    DeterminerSignataires = "(aucune regle - a definir)"
End Function


' --------------------------------------------------------------------------
' Applique le routage des signataires a une ligne de Suivi Factures.
' --------------------------------------------------------------------------
Public Sub AppliquerRoutage(ByVal ligne As Long)
    Dim ws As Worksheet, typeDepense As String, montant As Double
    Set ws = Feuille(FEUILLE_FACTURES)
    typeDepense = Trim$(CStr(ws.Cells(ligne, F_TYPE).Value))
    montant = Val(CStr(ws.Cells(ligne, F_TTC).Value))

    If Len(typeDepense) = 0 Then
        ws.Cells(ligne, F_SIGNATAIRES).Value = "(type de depense a renseigner)"
        Exit Sub
    End If

    ws.Cells(ligne, F_SIGNATAIRES).Value = DeterminerSignataires(typeDepense, montant)
    Journaliser "Routage", "Ligne " & ligne & " : " & _
                CStr(ws.Cells(ligne, F_SIGNATAIRES).Value)
End Sub

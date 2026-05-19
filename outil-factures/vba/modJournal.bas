Attribute VB_Name = "modJournal"
Option Explicit

' ==========================================================================
' modJournal - Journalisation horodatee des actions de l'outil
' ==========================================================================

' --------------------------------------------------------------------------
' Ajoute une ligne dans la feuille Journal.
' --------------------------------------------------------------------------
Public Sub Journaliser(ByVal action As String, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next   ' la journalisation ne doit jamais bloquer un traitement
    Set ws = Feuille(FEUILLE_JOURNAL)
    If ws Is Nothing Then Exit Sub
    r = DerniereLigne(ws, 1) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(r, 2).Value = action
    ws.Cells(r, 3).Value = detail
    ws.Cells(r, 4).Value = Param("Utilisateur")
End Sub

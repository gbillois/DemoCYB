VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' ThisWorkbook - Code de l'objet classeur.
'
' IMPORTANT : ce fichier ne s'IMPORTE PAS comme un module. Ouvrez l'objet
' 'ThisWorkbook' dans l'explorateur de projet VBA (Alt+F11) et COLLEZ-Y le
' contenu ci-dessous (a partir de 'Private Sub Workbook_Open').
' ==========================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    Application.Calculate
    ThisWorkbook.Worksheets(FEUILLE_TDB).Activate
End Sub

Attribute VB_Name = "modFichiers"
Option Explicit

' ==========================================================================
' modFichiers - Manipulation de fichiers sur le reseau / le disque local
' Utilise Scripting.FileSystemObject en liaison tardive (aucune reference).
' ==========================================================================

' --------------------------------------------------------------------------
' Cree un dossier (et ses parents) s'il n'existe pas. Renvoie True si OK.
' --------------------------------------------------------------------------
Public Function AssurerDossier(ByVal chemin As String) As Boolean
    Dim fso As Object, parent As String
    On Error GoTo Echec
    If Len(Trim$(chemin)) = 0 Then GoTo Echec
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(chemin) Then
        AssurerDossier = True
        Exit Function
    End If
    parent = fso.GetParentFolderName(chemin)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then
            If Not AssurerDossier(parent) Then GoTo Echec
        End If
    End If
    fso.CreateFolder chemin
    AssurerDossier = fso.FolderExists(chemin)
    Exit Function
Echec:
    AssurerDossier = False
End Function


' --------------------------------------------------------------------------
' Construit un chemin de destination libre (ajoute (1), (2)... si besoin).
' --------------------------------------------------------------------------
Private Function CheminLibre(ByVal fso As Object, ByVal dossier As String, _
                              ByVal nomFichier As String) As String
    Dim base As String, ext As String, cible As String, i As Long
    base = fso.GetBaseName(nomFichier)
    ext = fso.GetExtensionName(nomFichier)
    If Len(ext) > 0 Then ext = "." & ext
    cible = fso.BuildPath(dossier, base & ext)
    i = 1
    Do While fso.FileExists(cible)
        cible = fso.BuildPath(dossier, base & " (" & i & ")" & ext)
        i = i + 1
    Loop
    CheminLibre = cible
End Function


' --------------------------------------------------------------------------
' Copie un fichier vers un dossier. Renvoie le chemin final, "" si echec.
' --------------------------------------------------------------------------
Public Function CopierFichier(ByVal source As String, _
                                ByVal dossierDest As String) As String
    Dim fso As Object, cible As String
    On Error GoTo Echec
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(source) Then GoTo Echec
    If Not AssurerDossier(dossierDest) Then GoTo Echec
    cible = CheminLibre(fso, dossierDest, fso.GetFileName(source))
    fso.CopyFile source, cible, False
    CopierFichier = cible
    Exit Function
Echec:
    CopierFichier = ""
End Function


' --------------------------------------------------------------------------
' Deplace un fichier vers un dossier. Renvoie le chemin final, "" si echec.
' --------------------------------------------------------------------------
Public Function DeplacerFichier(ByVal source As String, _
                                 ByVal dossierDest As String) As String
    Dim fso As Object, cible As String
    On Error GoTo Echec
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(source) Then GoTo Echec
    If Not AssurerDossier(dossierDest) Then GoTo Echec
    cible = CheminLibre(fso, dossierDest, fso.GetFileName(source))
    fso.MoveFile source, cible
    DeplacerFichier = cible
    Exit Function
Echec:
    DeplacerFichier = ""
End Function


' --------------------------------------------------------------------------
' Vrai si le fichier existe.
' --------------------------------------------------------------------------
Public Function FichierExiste(ByVal chemin As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FichierExiste = fso.FileExists(chemin)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name: ManageParam
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind: Module
' Purpose: Gestion des paramètres de l'application
' Version: 01.00
' Author: Eric VENTALON
' Date: 08/05/2019
'
'
Private gestionFichier As Scripting.FileSystemObject
Private logSheetName  As String

Private Const NO_COL_VALEUR = 4
Private Const NO_COL_VERIF = 5


Private indexParametres As Scripting.Dictionary
Private indexParaLine As Scripting.Dictionary

Private Function isExistSheetName() As Boolean

On Error GoTo SiErreur
Dim Feuille As Worksheet
        
    isExistSheetName = False
    For Each Feuille In Worksheets
        If Feuille.name = logSheetName Then
             isExistSheetName = True
             Exit Function
         End If
    Next Feuille
    
    Exit Function

SiErreur:

    isExistSheetName = False
    
End Function

Public Function initManageParam(ByRef argNotreLog As ManageLog) As Integer
   
   On Error GoTo errorInitManageParam
     
   
    initManageParam = 0
    
    Set indexParametres = Nothing
    Set indexParametres = New Scripting.Dictionary
    
    Set indexParaLine = Nothing
    Set indexParaLine = New Scripting.Dictionary
     
    
    Set gestionFichier = Nothing
    Set gestionFichier = CreateObject("Scripting.FileSystemObject")
    
    '
    ' Est ce que la feuille des paramètres existe ?
    logSheetName = "PARAMETRES"
    If Not isExistSheetName() Then
       argNotreLog.writeError ("La feuille des paramètres n'existe pas, il est donc impossible de poursuivre.")
       initManageParam = -1
       Exit Function
    End If
   
    Dim ligne, lastLine As Long
    Dim sGeneral, sType, sSousType As String
        
    lastLine = getTheLastLine()
   
    sGeneral = Trim(Worksheets(logSheetName).Cells(2, 1).Value)
    sType = Trim(Worksheets(logSheetName).Cells(2, 2).Value)
    sSousType = Trim(Worksheets(logSheetName).Cells(2, 3).Value)
     
    If sGeneral = "" Then
       argNotreLog.writeError ("Veuillez saisir une activité en premiere ligne.")
       initManageParam = -2
       Exit Function
    End If
    
    If sType = "" Then
       argNotreLog.writeError ("Veuillez saisir un type en premiere ligne.")
       initManageParam = -2
       Exit Function
    End If
    
    If sSousType = "" Then
       argNotreLog.writeError ("Veuillez saisir un sous type en premiere ligne.")
       initManageParam = -2
       Exit Function
    End If
    
    Dim sKeyParam  As String
    
    For ligne = 2 To lastLine - 1
    
       If Trim(Worksheets(logSheetName).Cells(ligne, 1).Value) <> "" Then
          sGeneral = Trim(Worksheets(logSheetName).Cells(ligne, 1).Value)
       End If
       If Trim(Worksheets(logSheetName).Cells(ligne, 2).Value) <> "" Then
          sType = Trim(Worksheets(logSheetName).Cells(ligne, 2).Value)
       End If
       If Trim(Trim(Worksheets(logSheetName).Cells(ligne, 3).Value)) <> "" Then
          sSousType = Trim(Worksheets(logSheetName).Cells(ligne, 3).Value)
       End If
       
       sKeyParam = UCase(sGeneral + "_" + sType + "_" + sSousType)
    
       If UCase(sType) <> "SYSTEM" Then
            '
            ' Action 1
            '
            If Right(Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value, 1) <> "\" Then
                 Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value = Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value + "\"
            End If
            '
            ' Action 2
            '
            Worksheets(logSheetName).Cells(ligne, NO_COL_VERIF).Value = "KO"
            If gestionFichier.FolderExists(Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value) = False Then
              gestionFichier.CreateFolder Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value
              Call argNotreLog.writeInfo("Création du répertoire " + Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Value)
              Worksheets(logSheetName).Cells(ligne, NO_COL_VERIF).Value = "OK CREE"
            Else
              Worksheets(logSheetName).Cells(ligne, NO_COL_VERIF).Value = "OK"
            End If
       Else
            Worksheets(logSheetName).Cells(ligne, NO_COL_VERIF).Value = "PARA"
       End If
       '
       ' Action 3
       '
       If indexParametres.Exists(sKeyParam) Then
          argNotreLog.writeError ("La combinaison " + sKeyParam + " existe en ligne " + CStr(ligne) + ".")
          initManageParam = -3
          Exit For
       End If
       
       indexParametres.Add sKeyParam, Worksheets(logSheetName).Cells(ligne, NO_COL_VALEUR).Text
       indexParaLine.Add sKeyParam, CStr(ligne)
       
    Next
        
    Set gestionFichier = Nothing
    On Error GoTo 0
    Exit Function
    
errorInitManageParam:
    
    Call argNotreLog.writeError("...Ligne " & ligne & " : Numéro " & Err.Number & " : " & Err.Description)
    
    initManageParam = -1
    Set gestionFichier = Nothing
    On Error GoTo 0
    Exit Function
    
End Function



Public Function getParams(argParam) As String
On Error GoTo errGetParam

    getParams = indexParametres.Item(argParam)
    Exit Function
    
errGetParam:

    getParams = ""
    On Error GoTo 0
    
End Function
 
Public Function getNumberLineOf(argParam) As Long
On Error GoTo errGetNumberLineOf

    getNumberLineOf = indexParaLine.Item(argParam)
    Exit Function
    
errGetNumberLineOf:

    getNumberLineOf = 0
    On Error GoTo 0
    
End Function

Private Function getTheLastLine() As Long
    
    
    Dim Value As String
    Dim ligne, ligneEnd As Long
    
    getTheLastLine = 2
    
    While Worksheets(logSheetName).Cells(getTheLastLine, NO_COL_VALEUR).Value <> ""
        getTheLastLine = getTheLastLine + 1
    Wend
    
        
End Function








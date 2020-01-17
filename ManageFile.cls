VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: ManageFile
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind: Module de classe
' Purpose: Gestion des fichiers :
'            - Archivage
'            - Copier
'            - Déplacer
'            - Vérifier l'existance
'            - Convertir un XLS en CSV
'            - Lecture d'un fichier csv en format uf8 ou windows
'
' Author: Eric VENTALON
' Date: 21/05/2019
' ------------------------------------------------------
Option Explicit

Private m_objPersoFichier           As Object
Private m_sColSeparator             As String
Private m_bArgHeader                As Boolean
Private m_objLogApplicatif          As New ManageLog
Private m_objStream                 As ADODB.Stream
Private m_objWorkBook               As Workbook
Private m_objOnglet                 As New ManageOnglet
Private m_bEnteteDefine               As Boolean

' ----------------------------------------------------------------
' Procedure Name : dzipTarGz
' Purpose : Dézippe un fichier
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sTypeParam (String):
' Parameter sTypeFichierOutput (String):
' Parameter r_objNotreParam (ManageParam):
' Parameter r_objNotreLog (ManageLog):
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Sub dzipTarGz(ByVal sTypeParam As String, _
    ByVal sTypeFichierOutput As String, _
    ByRef r_objNotreParam As ManageParam, _
    ByRef r_objNotreLog As ManageLog)
    
    Dim sFileName, sMyPath, sMyExt As String
    Dim sShellStr  As String
    
    
    sTypeParam = UCase(sTypeParam)
    sTypeFichierOutput = UCase(sTypeFichierOutput)
    
    sMyPath = r_objNotreParam.getParams(sTypeParam)
    
    'File Ext to Change
    sMyExt = "*.gz*"
    
    'Target Path with Ending Extention
    sFileName = Dir(sMyPath & sMyExt)
    
    Do While Len(sFileName)
        sShellStr = Chr(34) + "C:\Program Files\7-Zip\7zG.exe\" + Chr(34) + " e " + Chr(34) + r_objNotreParam.getParams(sTypeParam) + sFileName + Chr(34) + " -o " + Chr(34) + r_objNotreParam.getParams(sTypeFichierOutput) + Chr(34)
        Call r_objNotreLog.writeInfo(sShellStr)
        Call Shell(sShellStr, vbHide)
        sFileName = Dir
    Loop
    
    
End Sub


' ----------------------------------------------------------------
' Procedure Name : ConvertToCsv
' Purpose :
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sTypeParam (String):
' Parameter sTypeFichierOutput (String):
' Parameter r_objNotreParam (ManageParam):
' Parameter r_objNotreLog (ManageLog):
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Sub ConvertToCsv(ByVal sTypeParam As String, _
    ByVal sTypeFichierOutput As String, _
    ByRef r_objNotreParam As ManageParam, _
    ByRef r_objNotreLog As ManageLog)
    
    Dim objWork As Workbook
    
    Dim sMyPath As String
    Dim sMyFile As String
    Dim sMyExt As String
    Dim sWorkName As String
             
    sTypeParam = UCase(sTypeParam)
    sTypeFichierOutput = UCase(sTypeFichierOutput)
    
    sMyPath = r_objNotreParam.getParams(sTypeParam)

    Call r_objNotreLog.writeInfo("Export des fichiers " + sMyPath + " en " + r_objNotreParam.getParams(sTypeFichierOutput))
    
    'File Ext to Change
    sMyExt = "*.xl*"
    
    'Target Path with Ending Extention
    sMyFile = Dir(sMyPath & sMyExt)
    
    'Loop through each Excel file in folder
    Do While sMyFile <> ""
        '
        ' Ouvrir le fichier XLS
        Set objWork = Workbooks.Open(filename:=sMyPath & sMyFile)
        Call r_objNotreLog.writeInfo("Ouverture fichier " + sMyPath & sMyFile)
        
        sWorkName = r_objNotreParam.getParams(sTypeFichierOutput) & Left(sMyFile, InStr(1, sMyFile, ".") - 1) & ".csv"
        '
        ' Sauvegarder ce fichier csv
        If existFile(sWorkName) = False Then
            objWork.SaveAs filename:=sWorkName, FileFormat:=xlCSV, local:=True
            Call r_objNotreLog.writeInfo("Sauvegarde en fichier csv : " + sWorkName)
        Else
            Call r_objNotreLog.writeInfo("Le fichier " + sWorkName + " est deja présent.")
        End If
        '
        ' Fermer ce fichier
        objWork.Close savechanges:=False
        
        '
        ' Passage au fichier xls suivant
        sMyFile = Dir
    Loop
  

    
End Sub

' ----------------------------------------------------------------
' Procedure Name : existFile
' Purpose : Vérifie l'existance d'un fichier. true s'il existe, false sinon
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sName (String): nom du fichier complet. repertoire + nom fichier
' Return Type: Boolean
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Function existFile(sName As String) As Boolean

    Dim objFso As New Scripting.FileSystemObject
    
    existFile = False
   
    If objFso.FileExists(sName) Then
        existFile = True
    Else
        existFile = False
    End If
    
    Set objFso = Nothing
    
End Function


' ----------------------------------------------------------------
' Procedure Name : moveFile
' Purpose : Déplace un ensemble de fichier
' Procedure Kind: Function
' Procedure Access: Private
' Parameter strSourceFilePath (): Source de départ
' Parameter strDestinationPath (): Source d'arrivée
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Function moveFile(strSourceFilePath, strDestinationPath)

    On Error GoTo moveFile_Error

    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String

    Dim objFso As New Scripting.FileSystemObject
    
    Call objFso.moveFile(strSourceFilePath, strDestinationPath)
      
    Set objFso = Nothing

    On Error GoTo 0
    Exit Function

moveFile_Error:

    lErr = Err.Number
    sDescription = Err.Description + " dans " + strDestinationPath
    sSource = Err.Source

    Set objFso = Nothing
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageFile.moveFile", sDescription
   
End Function

' ----------------------------------------------------------------
' Procedure Name : archiveFile
' Purpose : Archive des fichiers
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sTypeParamInput (String): Fichier source
' Parameter sTypeParamOutput (String): Fichier destination
' Parameter r_objNotreParam (ManageParam): Objet paramètre
' Parameter r_objNotreLog (ManageLog): Objet de trace
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Sub archiveFile(ByVal sTypeParamInput As String, _
    ByVal sTypeParamOutput As String, _
    ByRef r_objNotreParam As ManageParam, _
    ByRef r_objNotreLog As ManageLog)
    
    Dim sMyPath As String
    Dim sMyFile As String
    Dim sMyExt As String
             
    sTypeParamInput = UCase(sTypeParamInput)
    sTypeParamOutput = UCase(sTypeParamOutput)
    
    sMyPath = r_objNotreParam.getParams(sTypeParamInput)

    Call r_objNotreLog.writeInfo("Début processus archivage sur " & sMyPath)
    
    'File Ext to Change
    sMyExt = "*.*"
    
    'Target Path with Ending Extention
    sMyFile = Dir(sMyPath & sMyExt)
    
    'Loop through each Excel file in folder
    Do While sMyFile <> ""
        
        Call moveFile(sMyPath & sMyFile, r_objNotreParam.getParams(sTypeParamOutput))
        Call r_objNotreLog.writeInfo(sMyFile)
        '
        ' Passage au fichier xls suivant
        sMyFile = Dir
    Loop
    
    Call r_objNotreLog.writeInfo("Fin processus archivage.")
    
End Sub
' ----------------------------------------------------------------
' Procedure Name : openFileCsv
' Purpose : Ouverture d'un fichier texte en mode ascii ou uft_8
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePathName (String): Chemin complet du fichier
' Parameter argSeparator (String): Séparateur entre les colonnes. Par défaut c'est le point virgule
' Parameter argUTF8 (Boolean): uft8 = true, windows = false
' Parameter argHeader (Boolean): true = le fichier posséde des colonnes sinon false
' Return Type: Boolean
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Function openFileCsv(sFilePathName As String, _
    Optional ByVal argSeparator As String = ";", _
    Optional ByVal argUTF8 As Boolean = True, _
    Optional ByVal argHeader As Boolean = True) As Boolean
    
    On Error GoTo erreurGenerale
    
    Dim lErr As Long
    Dim sDescription As String
    Dim sSource As String
 
    openFileCsv = False
  
    Set m_objStream = CreateObject("ADODB.Stream")
    If argUTF8 = True Then
        m_objStream.Charset = "utf-8"
    Else
        m_objStream.Charset = "cp1256"
    End If
    m_objStream.Open
    m_objStream.Position = 0
    Call m_objStream.loadFromFile(sFilePathName)
  
    openFileCsv = True
    m_sColSeparator = argSeparator
    m_bArgHeader = argHeader
   
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageFile.openFileCsv()", sDescription
 
End Function
' ----------------------------------------------------------------
' Procedure Name : getRowCsv
' Purpose : Lecture d'une ligne du fichier csv et effectue les traitements suivants :
'           Extraire la ligne du fichier.
'           Convertion de la ligne en tableau.
'           Exécution de la fonction personalisée getLineCsv_Custom si celle-ci existe.
' Parameter r_objAdoHeritage (Object): Nom de la classe métier
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Integer
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Function getRowCsv(ByRef r_objAdoHeritage As Object) As Integer
    On Error GoTo erreurGenerale
    
    Dim lResultRow          As Long
    Dim lResultFormat       As Long
    Dim lResultHeader       As Long
    Dim lResultSetBdd       As Long
    Dim lResultBefore       As Long
    Dim lCptLigne           As Long
    Dim lErr                As Long
    Dim iCpt                As Integer
    
    Dim sLine               As String
    Dim sDescription        As String
    Dim sSource             As String
    Dim sColonne()          As String
       
    getRowCsv = -1
    lCptLigne = 1
    '
    ' Parcourir tous le fichier
    lResultRow = 0
    lResultFormat = 0
    lResultHeader = 0
    lResultSetBdd = 0
    lResultBefore = 0
    '
    '
    lResultBefore = r_objAdoHeritage.beforeReadFile()
    If lResultBefore = 2 Then
        Exit Function
    End If

    ' Impossible de placer d'autres conditions dans le while
    Do Until m_objStream.EOS
        ' Extraire la ligne
        sLine = m_objStream.ReadText(-2)
        ' Conversion en tableau de string unidimentionnel grace au séparateur
        sColonne = Split(sLine, m_sColSeparator)
        ' Suppression des guillemets qui encadrent la valeur
        For iCpt = 0 To UBound(sColonne)
            sColonne(iCpt) = supGuillemetBeginEnd(sColonne(iCpt))
        Next iCpt
        ' Exécution d une fonction personalisée d'analyse des colonnes
        ' Pour passer à fonction suivante il faut que le code retour de la précédente = 0
        ' Pour passer directement à la ligne suivante il faut un code retour = 1
        ' Pour quitter définitivement le chargement il faut un code retour = 2
        If (m_bArgHeader = True And lCptLigne = 1) Then
            lResultHeader = r_objAdoHeritage.getHeader(lCptLigne, sColonne)
            If lResultHeader = 2 Then
                Exit Do
            End If
        Else
            lResultRow = r_objAdoHeritage.getrow(lCptLigne, sColonne)
            If lResultRow = 2 Then
                Exit Do
            End If
            If lResultRow = 0 Then
                ' Exécution d une fonction personalisée d'analyse des formats des colonnes
                lResultFormat = r_objAdoHeritage.verifRowFomat(lCptLigne, sColonne)
                If lResultFormat = 2 Then
                    Exit Do
                End If
                If lResultFormat = 0 Then
                    ' Exécution d une fonction personalisée d'insertion en base de données
                    lResultSetBdd = r_objAdoHeritage.setRowInBdd(lCptLigne, sColonne)
                    If lResultSetBdd = 2 Then
                        Exit Do
                    End If
                End If
            End If
        End If
        lCptLigne = lCptLigne + 1
    Loop
    
    getRowCsv = r_objAdoHeritage.afterReadFile()
   
    If Not m_objStream Is Nothing Then
        If m_objStream.State = adStateOpen Then
            m_objStream.Close
        End If
    End If
    Set m_objStream = Nothing
   
    On Error GoTo 0
    Exit Function
   
erreurGenerale:
    
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    On Error GoTo 0
   
    Call Err.Raise(lErr, sSource + " : ManageFile.getRowCsv()", sDescription)
   
End Function
' ----------------------------------------------------------------
' Procedure Name : supGuillemetBeginEnd
' Purpose : Suppression des guillements au debut et à la fin du mot
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sValue (String): La valeur de la colonne extraite
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Private Function supGuillemetBeginEnd(ByVal sValue As String) As String
    
    sValue = Trim(sValue)
    
    If Right(sValue, 1) = Chr(34) Then
        sValue = Mid(sValue, 1, Len(sValue) - 1)
    End If
    
    If Left(sValue, 1) = Chr(34) Then
        sValue = Mid(sValue, 2, Len(sValue))
    End If
    
    supGuillemetBeginEnd = sValue
  

End Function


' ----------------------------------------------------------------
' Procedure Name: openFileToArray
' Purpose: Ouverture d'un fichier et formattage dans une feuille
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePathName (String): Nom complet du fichier
' Parameter sSheetName (String): Nom de la feuille de travail.
' Parameter sColMaxLine (String): Nom de la colonne de comptage.Exemple A.
' Parameter r_objAdoHeritage (Object): Objet métier
' Parameter sSeparateur (String): Un séparetur de colonne.
' Return Type: Workbook
' Author: Eric VENTALON
' Date: 28/06/2019
' ----------------------------------------------------------------
Public Function openFileToArray(ByVal sKeySheet, _
    ByVal sFilePathName As String, _
    ByVal sSheetName As String, _
    ByVal sColMaxLine As String, _
    ByRef r_objAdoHeritage As Object, _
    Optional lHearderColumn As Long = 1, _
    Optional sSeparateur As String = "", _
    Optional lHearderLigne As Long = 1) As ManageOnglet
    
    On Error GoTo erreurGenerale
       
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    
    Dim objOnglet       As New ManageOnglet
        
    Set openFileToArray = Nothing
    '
    ' Ouverture du fichier et active une feuille
    Set m_objWorkBook = Workbooks.Open(sFilePathName)
    '
    ' Mémorise les paramétres
    Call objOnglet.setSheet(sKeySheet, m_objWorkBook, sSheetName, sColMaxLine, r_objAdoHeritage, lHearderColumn, sSeparateur, lHearderLigne)
        
    Set openFileToArray = objOnglet
       
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    On Error GoTo 0
   
    Call Err.Raise(lErr, sSource + " : ManageFile.openFileToArray()", sDescription)
 
End Function

Public Function openCsvFileArray( _
    ByVal sKeySheet, _
    ByVal sFilePathName As String, _
    ByVal sSheetName As String, _
    ByVal sColMaxLine As String, _
    ByRef r_objAdoHeritage As Object, _
    Optional lHearderColumn As Long = 1, _
    Optional sSeparateur As String = "", _
    Optional lHearderLigne As Long = 1) As ManageOnglet
    
    On Error GoTo erreurGenerale
       
    Dim lErr            As Long
    Dim lLigneMax       As Long
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sZone           As String
    
    Dim objOnglet       As New ManageOnglet
    Dim objArray        As New ManageArray
    Dim objHeaderArray  As New Scripting.Dictionary
    
    Dim objQueryTable   As QueryTable
        
    Set openCsvFileArray = Nothing
    '
    ' Overture du fichier et active une feuille
    Application.DisplayAlerts = False
    
    Sheets(sSheetName).Select
    
    
    MsgBox "", vbOKOnly Or vbCritical, "ALERTE FRANWORK - REVOIR CETTE FONCTION"

    
    Application.DisplayAlerts = True
    
    'Call objOnglet.setSheet(sKeySheet, m_objWorkBook, sSheetName, sColMaxLine, r_objAdoHeritage, lHearderColumn, sSeparateur, lHearderLigne)
        
    'Set openFileToArray = objOnglet
    
    
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    On Error GoTo 0
   
    Call Err.Raise(lErr, sSource + " : ManageFile.openCsvFileArray()", sDescription)
 
End Function





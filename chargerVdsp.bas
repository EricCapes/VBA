Attribute VB_Name = "chargerVdsp"
' ------------------------------------------------------
' Name: chargerVdsp
' Kind: Module
' Purpose: Chagement de tous les fichiers VDSP
' Author: XMXU139
' Date: 13/12/2019
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: import_VDSP
' Purpose: Chargement de tous les fichiers VDSP
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 13/12/2019
'       16/01/2020 - Optimisation code
' ----------------------------------------------------------------
Public Sub import_VDSP()

    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim lDerlign        As Long
    Dim lCumulPresta    As Long
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sPathFiles      As String
    Dim sFeuilleName    As String
    Dim sMyExt, sMyFile As String
    
    Dim objTrace        As New ManageLog
    Dim objFile         As New ManageFile
    Dim objOnglet       As New ManageOnglet
    Dim objParam        As New ManageParam
    Dim objArray        As New ManageArray
    Dim objVdsp         As New MetierVdsp
   
    Dim objWork         As Workbook
    
    '
    ' TEST 01 - AUCUN FICHIER - OK
    ' TEST 02 - 1 FICHIER - OK
    ' TEST 03 - 2 FICHIERS MEME NATURE - OK
        
    '
    ' Récupération du nom du fichier principale.
    Set objWork = ThisWorkbook
    '
    ' Désactive les calculs automatiques
    Call objOnglet.setAllEnableCalculation(False, objWork)
    '
    ' Chargement des paramètres
    Call objTrace.InitManageLog
    If objParam.initManageParam(objTrace) <> 0 Then
        Call objTrace.displayMessage
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse VDSP ANAKIN"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Récupére le chemin d'accès au fichier
    sPathFiles = objParam.getParams("P_INPUT_VDSP")
    If sPathFiles = "" Then
        Call objTrace.writeError("Veuillez définir le répertoire des  entrées VDSP : Paramètre P_INPUT_VDSP à renseigner.")
        Call objTrace.displayMessage
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse VDSP ANAKIN"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Pour le log : début du traitement
    '
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("Chargement du fichier VDSP : DEBUT")
    Call objTrace.displayMessage
    '
    ' Extension des fichiers
    sMyExt = "*.csv"
    sMyFile = Dir(sPathFiles & sMyExt)
    If sMyFile = "" Then
        Call objTrace.writeInfo("Chargement du fichier VDSP : FIN")
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Parcourir tous les fichiers CSV
    Do While sMyFile <> ""
        ' Trace en log et dans la barre de status
        Call objTrace.writeInfo("...Ouverture du fichier " & sPathFiles & sMyFile)
        Application.StatusBar = "Lecture fichier " & sPathFiles & sMyFile & "."
        ' Ouverture d'un fichier VDSP
        Set objOnglet = objFile.openFileToArray("VDSP_LOAD", sPathFiles & sMyFile, "", "A", objVdsp, 1, ";")
        ' Compter les lignes sans l'entête
        lDerlign = Application.WorksheetFunction.CountA(Range("A:A")) - 1
        ' Validation du tableau en mémoire
        Set objArray = objOnglet.checkToArray("VDSP_LOAD")
        If lDerlign > 1 Then
            '
            ' ... Ouvre la feuille DEY ou DEN etc ...
            sFeuilleName = UCase(getNatureDocument(objArray))
            Sheets(sFeuilleName).Select
            Sheets(sFeuilleName).Cells.ClearOutline
            Sheets(sFeuilleName).AutoFilterMode = False
            '
            ' Compte le nombre de ligne existant
            lCumulPresta = Application.WorksheetFunction.CountA(Range("A:A"))
            ' Si pas d'entête ou que l'entête
            If lCumulPresta < 2 Then lCumulPresta = 2 Else lCumulPresta = lCumulPresta + 1
            Call objArray.arrayToSheet(ThisWorkbook, sFeuilleName, objOnglet, "A" & lCumulPresta & ":S" & (lCumulPresta + lDerlign - 1))
            '
            ' Suppression des lignes en double dans tous les onglets VDSP
            Call suppressionDesDoublonsPhysiques(sFeuilleName)
       
        End If
        '
        ' Libération de la mémoire
        Call objOnglet.dropArray("VDSP_LOAD")
        '
        ' Informe le log
        Call objTrace.writeInfo("...Insertion de " & lDerlign & " ligne(s).")
        '
        ' Passage au fichier csv suivant
        sMyFile = Dir
    Loop
    '
    ' Archivage
    Call objFile.archiveFile("P_INPUT_VDSP", "P_INPUT_VDSP_ARC", objParam, objTrace)
    '
    ' Pour le log : Fin du traitement
    Call objTrace.writeInfo("Chargement du fichier VDSP : FIN")
    Call objTrace.displayMessage
    '
    ' Ecriture de l'étape
    Call objOnglet.writeStepNameOnly("LOAD_VDSP")
    
    Application.StatusBar = "Fin du chargement"
    '
    ' Arrêt de l'optimisation graphique
    Call objOnglet.setAllEnableCalculation(True, objWork)

    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call objOnglet.setAllEnableCalculation(True, objWork)
      
    Call objTrace.writeError("...Module     : chargerVdsp")
    Call objTrace.writeError("...Fonction   : import_VDSP")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.displayMessage
        
    On Error GoTo 0
        
End Sub



' ----------------------------------------------------------------
' Procedure Name: getNatureDocument
' Purpose: Trouve la nature du fichier csv : DEY ou DEN ...
' Procedure Kind: Function
' Procedure Access: Private
' Parameter r_objArray (ManageArray): La tableau mémoire
' Return Type: String
' Author: XMXU139
' Date: 13/12/2019
' ----------------------------------------------------------------
Private Function getNatureDocument(ByRef r_objArray As ManageArray) As String

    getNatureDocument = ""
    If r_objArray.getLineMax > 0 Then
        getNatureDocument = r_objArray.getValue(1, 5)
    End If
    
End Function

' ----------------------------------------------------------------
' Procedure Name: suppressionDesDoublonsPhysiques
' Purpose: Suppression des lignes en double
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sSheetName (String): Nom de la feuille
' Author: XMXU139
' Date: 13/12/2019
' ----------------------------------------------------------------
Public Sub suppressionDesDoublonsPhysiques(ByVal sSheetName As String)
    
    Sheets(sSheetName).Select
    Cells.Select
    If sSheetName = "DEY" Or sSheetName = "DEX" Then
       ActiveSheet.Range("A:T").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), Header:=xlYes
    Else
       ActiveSheet.Range("A:Q").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), Header:=xlYes
    End If
    Range("A2").Select
    
End Sub



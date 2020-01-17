Attribute VB_Name = "chargerHub"
' ------------------------------------------------------
' Name: chargerHub
' Kind: Module
' Purpose: Chagement de tous les fichiers VDSP
' Author: XMXU139
' Date: 13/12/2019
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: import_HUB
' Purpose: Chargement de tous les fichiers VDSP
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 13/12/2019
' Maintenance: 02/01/2020 - Message plus explicite quand pas de fichier HUB
' ----------------------------------------------------------------
Public Sub import_HUB()

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
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse HUB"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Récupére le chemin d'accès au fichier
    sPathFiles = objParam.getParams("P_INPUT_HUB")
    If sPathFiles = "" Then
        Call objTrace.writeError("Veuillez définir le répertoire des  entrées VDSP : Paramètre P_INPUT_HUB à renseigner.")
        Call objTrace.displayMessage
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse HUB"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Pour le log : début du traitement
    '
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("Chargement du fichier HUB : DEBUT")
    Call objTrace.displayMessage
    '
    ' Extension des fichiers
    sMyExt = "*.csv"
    sMyFile = Dir(sPathFiles & sMyExt)
    If sMyFile = "" Then
        ' 02/01/2020 - Debut - Message plus explicite quand pas de fichier HUB
        Call objTrace.writeInfo("...Pas de fichier HUB")
        Call objTrace.writeInfo("Chargement du fichier HUB : FIN")
        Call objTrace.displayMessage
        ' 02/01/2020 - Fin - Message plus explicite quand pas de fichier HUB
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    sFeuilleName = "SM"
    '
    ' Parcourir tous les fichiers CSV
    Do While sMyFile <> ""
        ' Trace en log et dans la barre de status
        Call objTrace.writeInfo("...Ouverture du fichier " & sPathFiles & sMyFile)
        Application.StatusBar = "Lecture fichier " & sPathFiles & sMyFile & "."
        ' Ouverture d'un fichier VDSP
        Set objOnglet = objFile.openFileToArray("HUB_LOAD", sPathFiles & sMyFile, "", "A", objVdsp, 1, ";")
        ' Compter les lignes sans l'entête
        lDerlign = Application.WorksheetFunction.CountA(Range("A:A")) - 1
        ' Validation du tableau en mémoire
        Set objArray = objOnglet.checkToArray("HUB_LOAD")
        If lDerlign > 1 Then
            '
            ' ... Ouvre la feuille DEY ou DEN etc ...
            Sheets(sFeuilleName).Select
            Sheets(sFeuilleName).Cells.ClearOutline
            Sheets(sFeuilleName).AutoFilterMode = False
            '
            ' Compte le nombre de ligne existant
            lCumulPresta = Application.WorksheetFunction.CountA(Range("A:A"))
            ' Si pas d'entête ou que l'entête
            If lCumulPresta < 2 Then lCumulPresta = 2 Else lCumulPresta = lCumulPresta + 1
            Call objArray.arrayToSheet(ThisWorkbook, sFeuilleName, objOnglet, "A" & lCumulPresta & ":D" & (lCumulPresta + lDerlign - 1))
       
        End If
        '
        ' Libération de la mémoire
        Call objOnglet.dropArray("HUB_LOAD")
        '
        ' Informe le log
        Call objTrace.writeInfo("...Insertion de " & lDerlign & " ligne(s).")
        lCumulPresta = lCumulPresta + lDerlign
        '
        ' Passage au fichier csv suivant
        sMyFile = Dir
    Loop
    '
    ' Suppression des lignes en double
    Call suppressionDesDoublonsPhysiques("SM")
    '
    ' Archivage
    Call objFile.archiveFile("P_INPUT_HUB", "P_INPUT_HUB_ARC", objParam, objTrace)
    '
    ' Pour le log : Fin du traitement
    Call objTrace.writeInfo("Chargement du fichier HUB : FIN")
    Call objTrace.displayMessage
    '
    ' Ecriture de l'étape
    Call objOnglet.writeStepNameOnly("LOAD_HUB")
    
    
    Application.StatusBar = "Prêt"
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
      
    Call objTrace.writeError("...Module     : chargerHub")
    Call objTrace.writeError("...Fonction   : import_HUB")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.displayMessage
        
    On Error GoTo 0
        
End Sub

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
    
    ActiveSheet.Range("A:D").RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
    
    Range("A2").Select
    
End Sub


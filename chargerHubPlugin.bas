Attribute VB_Name = "chargerHubPlugin"
' ------------------------------------------------------
' Name: chargerHubPlugin
' Kind: Module
' Purpose: Chagement de tous les fichiers HUB PLUGIN
' Author: XMXU139
' Date: 16/01/2020
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: importHubPlugin
' Purpose: Chargement de tous les fichiers HUB PLUGIN
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 16/01/2020
' ----------------------------------------------------------------
Public Sub importHubPlugin()

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
    sPathFiles = objParam.getParams("P_INPUT_HUB_PLUG")
    If sPathFiles = "" Then
        Call objTrace.writeError("Veuillez définir le répertoire des  entrées VDSP : Paramètre P_INPUT_HUB_PLUG à renseigner.")
        Call objTrace.displayMessage
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse VDSP ANAKIN"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Pour le log : début du traitement
    '
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("Chargement du fichier HUB_PLUGIN : DEBUT")
    '
    ' Extension des fichiers
    sMyExt = "*.xlsx"
    sMyFile = Dir(sPathFiles & sMyExt)
    If sMyFile = "" Then
        Call objTrace.writeInfo("Chargement du fichier HUB_PLUGIN : FIN")
        Call objTrace.displayMessage
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    sFeuilleName = "HUB_PLUG"
    '
    ' Parcourir tous les fichiers CSV
    Do While sMyFile <> ""
        ' Trace en log et dans la barre de status
        Call objTrace.writeInfo("...Ouverture du fichier " & sPathFiles & sMyFile)
        Application.StatusBar = "Lecture fichier " & sPathFiles & sMyFile & "."
        ' Ouverture d'un fichier VDSP
        Set objOnglet = objFile.openFileToArray("KEY_HUBPLUG", sPathFiles & sMyFile, "", "A", objVdsp, 1, ";")
        ' Compter les lignes sans l'entête
        lDerlign = Application.WorksheetFunction.CountA(Range("A:A")) - 1
        ' Validation du tableau en mémoire
        Set objArray = objOnglet.checkToArray("KEY_HUBPLUG")
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
            '
            ' Suppression des lignes en double dans tous les onglets VDSP
            Call suppressionDesDoublonsPhysiques(sFeuilleName)
       
        End If
        '
        ' Libération de la mémoire
        Call objOnglet.dropArray("KEY_HUBPLUG")
        '
        ' Informe le log
        Call objTrace.writeInfo("...Insertion de " & lDerlign & " ligne(s).")
        '
        ' Passage au fichier csv suivant
        sMyFile = Dir
    Loop
    '
    ' Archivage
    Call objFile.archiveFile("P_INPUT_HUB_PLUG", "P_INPUT_HUB_PLUG_ARC", objParam, objTrace)
    '
    ' Pour le log : Fin du traitement
    Call objTrace.writeInfo("Chargement du fichier HUB_PLUGIN : FIN")
    Call objTrace.displayMessage
    '
    ' Ecriture de l'étape
    Call objOnglet.writeStepNameOnly("LOAD_HUBPLUG")
    
    Application.StatusBar = "Fin du chargement - Prêt"
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
      
    Call objTrace.writeError("...Module     : chargerHubPlugin")
    Call objTrace.writeError("...Fonction   : importHubPlugin")
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
' Date: 16/01/2020
' ----------------------------------------------------------------
Public Sub suppressionDesDoublonsPhysiques(ByVal sSheetName As String)
    
    Sheets(sSheetName).Select
    Cells.Select
    ActiveSheet.Range("A:D").RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
    Range("A2").Select
    
End Sub




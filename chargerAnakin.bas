Attribute VB_Name = "chargerAnakin"
' ------------------------------------------------------
' Name: chargerAnakin
' Kind: Module
' Purpose: Chargement et contrôle du fichier
' Author: XMXU139
' Date: 10/12/2019
' ------------------------------------------------------
Option Explicit

' ---------------------------------------------------------------------------------
' Procedure Name: importAnakin
' Purpose: Chargement du fichier ANAKIN
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 10/12/2019
' Maintenance : 02/01/2020 - Si pas de fichier alors quitter procédure
'               15/01/2020 - CRANAOOS_15012001 - Gestion des insertions dans ANAKIN
' ---------------------------------------------------------------------------------
Public Sub importAnakin()

    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim lLigneInsert    As Long
    Dim lLigneWork      As Long
    Dim lLigneSource    As Long
    
    Dim lCpt            As Long
    
    Dim lTabStats(2)    As Long
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sPathFiles      As String
    Dim sFeuilleName    As String
    
    Dim objTrace        As New ManageLog
    Dim objFile         As New ManageFile
    Dim objOnglet       As New ManageOnglet
    Dim objParam        As New ManageParam
    Dim objArray        As New ManageArray
    Dim objArrayTarget  As New ManageArray
    
    Dim objAnakin       As New CAnakin
    
    Dim objWork         As Workbook
    
    Dim objIndexSource  As New Scripting.Dictionary
    Dim objIndexTarget  As New Scripting.Dictionary
        
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
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse ANAKIN"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Récupére le chemin d'accès au fichier
    sPathFiles = objParam.getParams("P_INPUT_ANAKIN")
    If sPathFiles = "" Then
        Call objTrace.writeError("Veuillez définir le répertoire du fichier ANAKIN : P_INPUT_ANAKIN")
        Call objTrace.displayMessage
        MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbOKOnly Or vbCritical, "Analyse ANAKIN"
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Exit Sub
    End If
    '
    ' Désactiver les groupes dans l'onglet ANAKIN
    sFeuilleName = "ANAKIN"
    '  et les groupes
    Sheets(sFeuilleName).Cells.ClearOutline
    Sheets(sFeuilleName).AutoFilterMode = False
    
    Call objTrace.writeInfo("Chargement du fichier ANAKIN : DEBUT")
    Call objTrace.displayMessage
    
    Dim sMyExt, sMyFile As String
    
    sMyExt = "*.xlsx"
    sMyFile = Dir(sPathFiles & sMyExt)
    
    ' 02/01/2020 - Si pas de fichier alors quitter
    If sMyFile = "" Then
        Call objOnglet.setAllEnableCalculation(True, objWork)
        Call objTrace.writeInfo("...Pas de fichier ANAKIN.")
        Call objTrace.displayMessage
        Exit Sub
    End If
    ' 02/01/2020 - Fin
        
    ' Entête : ligne 3, colonne 1
    Set objOnglet = objFile.openFileToArray("ANAKIN_FILE", sPathFiles + sMyFile, "", "F", objAnakin, 1, "", 3)
    Set objArray = objOnglet.checkToArray("ANAKIN_FILE")
        
    Sheets(sFeuilleName).Select
    If Sheets(sFeuilleName).Range("A" & Rows.Count).End(xlUp).Row < 2 Then
        '
        ' Premier fois ? --> Tout copier
        Call objArray.arrayToSheet(objWork, sFeuilleName, objOnglet)
    Else
        ' non alors realiser l'update insert
        ' Index du fichier chargé sur la mission
        Call objArray.createIndexUnique(True, "idx_Mission_UUID", "selectMissionUuid", "getLineMissionUuid", "whereMissionUuid")
        '
        ' Placer l'existant en mémoire et ...
        Call objOnglet.setSheet("ANAKIN_TARGET", objWork, "ANAKIN", "A", objAnakin, 1, "")
        Set objArrayTarget = objOnglet.checkToArray("ANAKIN_TARGET", "A1:AR", False)
        ' Compte le nombre de ligne actuelle
        lLigneInsert = objArrayTarget.getLineMax + 1
        ' ... indexer les lignes sur idx_Mission_UUID
        Call objArrayTarget.createIndexUnique(True, "idx_Mission_UUID", "selectMissionUuid", "getLineMissionUuid", "whereMissionUuid")
        '
        ' Récupère les 2 index
        Set objIndexTarget = objArrayTarget.getIndex("idx_Mission_UUID")
        Set objIndexSource = objArray.getIndex("idx_Mission_UUID")
        ' J'ajoute 1000 lignes en plus pour les ajouts on ne sait jamais ...
        If objIndexSource.Count > objIndexTarget.Count Then
            Call objArrayTarget.addLine(objIndexSource.Count - objIndexTarget.Count + 1)
        End If
        '
        ' Parcourir les lignes de l'index de la feuille
        Dim vMissionUuid As Variant
        For Each vMissionUuid In objIndexSource
            
            If Not IsEmpty(vMissionUuid) Then
            
                lLigneSource = objIndexSource(vMissionUuid)
                '
                ' Existe dans celle chargée ?
                If objIndexTarget.Exists(vMissionUuid) = True Then
                    '
                    ' Oui alors update
                    lLigneWork = objIndexTarget(vMissionUuid)
                    lTabStats(0) = lTabStats(0) + 1
                Else
                    '
                    ' Non alors insertion
                    lLigneWork = lLigneInsert
                    '
                    ' 15/01/2020 - CRANAOOS_15012001 - Gestion des insertions dans ANAKIN - DEBUT
                    ' Faut-il agrandir le tableau ?
                    If objArrayTarget.getLineMax < lLigneInsert Then
                        Call objArrayTarget.addLine(60000)
                    End If
                    '
                    ' Ajoute à l'index la nouvelle mision
                    Call objIndexTarget.Add(vMissionUuid, lLigneWork)
                    ' 15/01/2020 - CRANAOOS_15012001 - Gestion des insertions dans ANAKIN - FIN
                    '
                    ' Incrémente la ligne de fin de tableau et maj des stats insertions
                    lLigneInsert = lLigneInsert + 1
                    lTabStats(1) = lTabStats(1) + 1
                End If
                '
                ' Modifier les données dans la cible
                For lCpt = 1 To objArray.getColumnMax
                    Call objArrayTarget.setValue(lLigneWork, lCpt, objArray.getValue(lLigneSource, lCpt))
                Next
            End If
        Next
        '
        ' Affiche le tableau
        Call objArrayTarget.arrayToSheet(objWork, "ANAKIN", objOnglet, "A2:AR" & lLigneInsert)
        
    End If
    
    Call objOnglet.dropArray("ANAKIN_FILE")
    Call objOnglet.dropArray("ANAKIN_TARGET")
    
    Call objTrace.writeInfo("...Nombre de modification : " & lTabStats(0))
    Call objTrace.writeInfo("...Nombre d'insertion     : " & lTabStats(1))
    '
    '
    Call objTrace.writeInfo("Chargement du fichier ANAKIN : FIN")
    Call objTrace.displayMessage
    
    Application.StatusBar = "Fin du chargement"
    '
    ' Ecriture de l'étape
    Call objOnglet.writeStepNameOnly("LOAD_ANAKIN")
    '
    ' Arrêt de l'optimisation graphique
    Call objOnglet.setAllEnableCalculation(True, objWork)
    '
    ' Archivage des fichiers
    Call objFile.archiveFile("P_INPUT_ANAKIN", "P_INPUT_ANAKIN_ARC", objParam, objTrace)
    Call objTrace.displayMessage

    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call objOnglet.setAllEnableCalculation(True, objWork)
    Call objOnglet.dropArray("ANAKIN_FILE")
    
    MsgBox "Erreur: " & lErr & " - " & sDescription & " - Module: chagerAnakin - Fonction: importAnakin", vbOKOnly Or vbCritical, "ERREUR TRAITEMENT ANAKIN"
      
    Call objTrace.writeError("...Module     : chagerAnakin")
    Call objTrace.writeError("...Fonction   : importAnakin")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.displayMessage
        
    On Error GoTo 0
        
End Sub



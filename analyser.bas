Attribute VB_Name = "analyser"
' ------------------------------------------------------
' Name: analyser
' Kind: Module
' Purpose: Chargements + calculs
' Author: XMXU139
' Date: 15/12/2019
' ------------------------------------------------------
Option Explicit


' ----------------------------------------------------------------
' Procedure Name: Main
' Purpose: Point d'entrée du programme
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 10/12/2019
' ----------------------------------------------------------------
Public Sub Main()
    Sheets("ANAKIN").Select
    Selection.AutoFilter

    Call import_VDSP
    Call importAnakin
    Call import_HUB
    Call importHubPlugin
    Call execCompute
    
    Sheets("ANAKIN").Select
    Selection.AutoFilter
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: execCompute
' Purpose: Exécution des calculs
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 10/12/2019
' Maintenance: 03/01/2013 - ERV - Correction dans la recherche du HUB
'              15/01/2020 - ERV - CRANAOOS_15012002 - DEN cloture GCP
'              15/01/2020 - ERV - CRANAOOS_15012003 - Recherche HUB PLUGIN
' ----------------------------------------------------------------
Public Sub execCompute()
    
    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim lLigneInsert    As Long
    Dim lLigneWork      As Long
    Dim lLigneSource    As Long
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sIdPresta       As String
    Dim sFeuilleName    As String
    
    Dim objTrace        As New ManageLog
    Dim objOnglet       As New ManageOnglet
    Dim objParam        As New ManageParam
    
    Dim objArrayAnakin  As New ManageArray
    Dim objArrayDey     As New ManageArray
    Dim objArrayDex     As New ManageArray
    Dim objArrayDen     As New ManageArray
    Dim objArrayDenw    As New ManageArray
    Dim objArraySm      As New ManageArray
    Dim objArrayHp      As New ManageArray
    
    Dim objAnakin       As New CAnakin
    Dim objVdsp         As New CVdsp
    Dim objCsm          As New CServiceMark
    Dim objHubPluging   As New CHubPluging
    
    Dim objWork         As Workbook
    
    Dim objHeaderAkn    As Scripting.Dictionary
    Dim objHeaderDey    As Scripting.Dictionary
    Dim objHeaderSm     As Scripting.Dictionary
    
    Dim objIndexDey     As Scripting.Dictionary
    Dim objIndexDex     As Scripting.Dictionary
    Dim objIndexDen     As Scripting.Dictionary
    Dim objIndexDenw    As Scripting.Dictionary
    Dim objIndexSm      As Scripting.Dictionary
    Dim objIndexHubPlug As Scripting.Dictionary
            
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
    ' Désactiver les groupes dans l'onglet ANAKIN
    sFeuilleName = "ANAKIN"
    '  et les groupes
    Sheets(sFeuilleName).Cells.ClearOutline
    Sheets(sFeuilleName).AutoFilterMode = False
    
    Call objTrace.writeInfo("Compute ANAKIN : DEBUT")
    Call objTrace.displayMessage
    '
    ' Mémorise une partie de la feuille. Les calculs seront en mémoire.
    Call objOnglet.setSheet("ANAKIN_TARGET", objWork, "ANAKIN", "A", objAnakin, 1, "")
    Set objArrayAnakin = objOnglet.checkToArray("ANAKIN_TARGET", "A1:AX", False)
    '
    ' Mémorise la feuille DEY
    Call objOnglet.setSheet("DEY", objWork, "DEY", "A", objVdsp, 1, "")
    Set objArrayDey = objOnglet.checkToArray("DEY", "A1:T", False)
    Call objArrayDey.createIndexUnique(True, "idxMissionDey", "selectMission", "getLineMission", "whereVdsp")
    Set objIndexDey = objArrayDey.getIndex("idxMissionDey")
    '
    ' Mémorise la feuille DEX
    Call objOnglet.setSheet("DEX", objWork, "DEX", "A", objVdsp, 1, "")
    Set objArrayDex = objOnglet.checkToArray("DEX", "A1:T", False)
    Call objArrayDex.createIndexUnique(True, "idxMissionDex", "selectMission", "getLineMission", "whereVdsp")
    Set objIndexDex = objArrayDex.getIndex("idxMissionDex")
    '
    ' Mémorise la feuille DEN
    Call objOnglet.setSheet("DEN", objWork, "DEN", "A", objVdsp, 1, "")
    Set objArrayDen = objOnglet.checkToArray("DEN", "A1:T", False)
    Call objArrayDen.createIndexUnique(True, "idxMissionDen", "selectMission", "getLineMission", "whereVdsp")
    Set objIndexDen = objArrayDen.getIndex("idxMissionDen")
    '
    ' Mémorise la feuille DENW
    Call objOnglet.setSheet("DENW", objWork, "DENW", "A", objVdsp, 1, "")
    Set objArrayDenw = objOnglet.checkToArray("DENW", "A1:T", False)
    Call objArrayDenw.createIndexUnique(True, "idxMissionDenw", "selectMission", "getLineMission", "whereVdsp")
    Set objIndexDenw = objArrayDenw.getIndex("idxMissionDenw")
    '
    ' Mémorise la feuille ServiceMark
    Call objOnglet.setSheet("KEY_SM", objWork, "SM", "A", objCsm, 1, "")
    Set objArraySm = objOnglet.checkToArray("KEY_SM", "A1:D", False)
    Call objArraySm.createIndexUnique(True, "idxSm", "selectSm", "getServiceMark", "whereSm")
    Set objIndexSm = objArraySm.getIndex("idxSm")
    '
    ' Mémorise la feuille HUB_PLUGIN
    ' 15/01/2020 - ERV - CRANAOOS_15012003 - Recherche HUB PLUGIN - DEBUT
    Call objOnglet.setSheet("KEY_HUBPLUG", objWork, "HUB_PLUG", "A", objHubPluging, 1, "")
    Set objArrayHp = objOnglet.checkToArray("KEY_HUBPLUG", "A1:D", False)
    Call objArrayHp.createIndexUnique(True, "idxHubPlug", "selectExternalId", "getHubPlugin", "whereHubPlugin")
    Set objIndexHubPlug = objArrayHp.getIndex("idxHubPlug")
    ' 15/01/2020 - ERV - CRANAOOS_15012003 - Recherche HUB PLUGIN - FIN
    
    '
    ' Parcourir chaque ligne du tableau ANAKIN
    lLigneWork = objArrayAnakin.getLineMax()
    '
    ' Les entêtes des feuilles
    Set objHeaderAkn = objArrayAnakin.getHeader()
    Set objHeaderDey = objArrayDey.getHeader()
    Set objHeaderSm = objArraySm.getHeader()
    '
    Dim bTrouve As Boolean
    For lLigneSource = 1 To lLigneWork
        '
        ' Extraire la mission
        sSource = objArrayAnakin.getValue(lLigneSource, objHeaderAkn("Mission_UUID"))
        ' Extraire l'Id prestation
        sIdPresta = objArrayAnakin.getValue(lLigneSource, objHeaderAkn("ID Prestation"))
       
        
        bTrouve = False
        '
        ' existe-t-elle dans DEN ?
        If objIndexDen.Exists(sIdPresta) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("DEN cloture GCP"), objIndexDen(sIdPresta) + 1)
            bTrouve = True
        Else
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("DEN cloture GCP"), "")
        End If
        '
        ' existe-t-elle dans DENW ?
        If objIndexDenw.Exists(sIdPresta) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("DENW cloture GCP"), objIndexDenw(sIdPresta) + 1)
            bTrouve = True
        Else
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("DENW cloture GCP"), "")
        End If
        '
        ' Attende de validation flux DJLN
        ' existe-t-elle dans DJLN ?
        'If objIndexDjln.Exists(sIdPresta) Then
        '    Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("DJLN cloture GCP"), objIndexDjln(sIdPresta) + 1)
        '    bTrouve = True
        'End If
        '
        ' 15/01/2020 - ERV - CRANAOOS_15012003 - Recherche HUB PLUGIN - Debut
        ' existe-t-elle dans HUB PLUGIN ?
        If objIndexHubPlug.Exists(sSource) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("recherche cr vide"), objIndexHubPlug(sSource) + 1)
        Else
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("recherche cr vide"), "")
        End If
        ' 15/01/2020 - ERV - CRANAOOS_15012003 - Recherche HUB PLUGIN - FIN
        
        '
        ' existe-t-elle dans DEY ?
        If objIndexDey.Exists(sSource) Then
           
            lLigneInsert = objIndexDey(sSource)
           
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("CR DEY"), lLigneInsert + 1)
           
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("nb collecté DEY"), objArrayDey.getValue(lLigneInsert, objHeaderDey("Nombre de contenants collectés (par type)")))
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("motif non real"), objArrayDey.getValue(lLigneInsert, objHeaderDey("Motif 1")))
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("nb mission"), objArrayDey.getValue(lLigneInsert, objHeaderDey("Nombre de missions")))
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("nb commandé"), objArrayDey.getValue(lLigneInsert, objHeaderDey("Nb commandé")))
           
            If objIndexSm.Exists(objArrayAnakin.getValue(lLigneSource, objHeaderAkn("Order_Id"))) Then
                Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("SM"), objIndexSm(objArrayAnakin.getValue(lLigneSource, objHeaderAkn("Order_Id"))))
            Else
                Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("SM"), "???")
            End If
           
            bTrouve = True
        End If
        '
        ' existe-t-elle dans DEX?
        If objIndexDex.Exists(sSource) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("CR DEX"), objIndexDex(sSource) + 1)
            bTrouve = True
        End If
        '
        ' existe-t-elle dans DEN ?
        If objIndexDen.Exists(sSource) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("CR DEN"), objIndexDen(sSource) + 1)
            bTrouve = True
        End If
        '
        ' existe-t-elle dans DENW ?
        If objIndexDenw.Exists(sSource) Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("CR DENW"), objIndexDenw(sSource) + 1)
            bTrouve = True
        End If
        
        If bTrouve Then
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("Recap CR trouvé"), "trouvé")
        Else
            Call objArrayAnakin.setValue(lLigneSource, objHeaderAkn("Recap CR trouvé"), "pas trouvé")
        End If
    
    Next lLigneSource
    
    Call objArrayAnakin.arrayToSheet(objWork, sFeuilleName, objOnglet, "A2:")
    
    Sheets(sFeuilleName).Select
                
    Call objOnglet.dropArray("ANAKIN_TARGET")
    Call objOnglet.dropArray("DEY")
    Call objOnglet.dropArray("DEX")
    Call objOnglet.dropArray("DEN")
    Call objOnglet.dropArray("DENW")
    Call objOnglet.dropArray("KEY_SM")
    Call objOnglet.dropArray("KEY_HUBPLUG")
    
    '
    '
    Call objTrace.writeInfo("Compute ANAKIN : FIN")
    Call objTrace.displayMessage
    
    '
    ' Ecriture de l'étape
    Call objOnglet.writeStepNameIntoModeOp("CALCUL_AKN", objTrace)
    '
    ' Arrêt de l'optimisation graphique
    Call objOnglet.setAllEnableCalculation(True, objWork)
    '
    ' Archivage des fichiers
    Call objTrace.displayMessage
    
     Application.StatusBar = "Prêt"

    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call objOnglet.setAllEnableCalculation(True, objWork)
      
    Call objTrace.writeError("...Module     : analyser")
    Call objTrace.writeError("...Fonction   : execCompute")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.displayMessage
        
    On Error GoTo 0
End Sub


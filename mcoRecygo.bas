Attribute VB_Name = "mcoRecygo"
'
'
' _   .-')
'( '.( OO )_
' ,--.   ,--.)   .-----.   .-'),-----.
' |   `.'   |   '  .--./  ( OO'  .-.  '
' |         |   |  |('-.  /   |  | |  |
' |  |'.'|  |  /_) |OO  ) \_) |  |\|  |
' |  |   |  |  ||  |`-'|    \ |  | |  |
' |  |   |  | (_'  '--'\     `'  '-'  '
' `--'   `--'    `-----'       `-----'
'
'
Option Explicit


Public Sub gotoMenuGeneral()
    Sheets("MODE OP").Select
End Sub


' ----------------------------------------------------------------
' Procedure Name : exportModuleVbs
' Purpose : Export des modules VBA
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Eric VENTALON
' Date: 17/05/2019
' ----------------------------------------------------------------
Public Sub exportModuleVbs()
    
    Dim objAppli                As New ManageParam
    Dim objApplicatif           As New ManageLog
    Dim sPathFiles              As String
    Dim LeFich                  As Variant
    
    Call objApplicatif.InitManageLog
    If objAppli.initManageParam(objApplicatif) <> 0 Then
       Call objApplicatif.displayMessage
       MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbCritical + vbOKOnly
       Exit Sub
    End If

    sPathFiles = objAppli.getParams("P_OUTPUT_VBA")
    If sPathFiles = "" Then
       Call objApplicatif.writeError("Veuillez définir le répertoire des fichiers de log de vérification : P_OUTPUT_VBA")
       Call objApplicatif.displayMessage
       MsgBox "Une erreur est survenue, il est donc impossible de continuer. Veuillez regarder le message des logs pour de plus amples informations.", vbCritical + vbOKOnly
       Exit Sub
    End If
        
    For Each LeFich In ThisWorkbook.VBProject.VBComponents
        Select Case LeFich.Type
            Case 1
                ThisWorkbook.VBProject.VBComponents(LeFich.name).Export sPathFiles & LeFich.name & ".bas"
            Case 2
                ThisWorkbook.VBProject.VBComponents(LeFich.name).Export sPathFiles & LeFich.name & ".cls"
            Case 3
                ThisWorkbook.VBProject.VBComponents(LeFich.name).Export sPathFiles & LeFich.name & ".frm"
            Case 100
                ThisWorkbook.VBProject.VBComponents(LeFich.name).Export sPathFiles & LeFich.name & ".cls"
        End Select
    Next
     
End Sub

Public Sub EmptyAllSheets()
    Dim sFeuille As String
    
    sFeuille = ActiveSheet.name
    
    If sFeuille = "MODE OP" Then
        Range("K6:M1000").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        Range("K6").Select
        Exit Sub
    End If
    
    If sFeuille = "PARAMETRES" Then
        MsgBox "Il est strictement INTERDIT d'effacer cette feuille."
        Exit Sub
    End If
    
    If sFeuille = "JOURNALISATION" Then
        Sheets(sFeuille).Rows("1:" & Rows.Count).ClearContents
        Exit Sub
    End If
    

End Sub



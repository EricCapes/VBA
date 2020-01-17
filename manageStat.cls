VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "manageStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name    : manageStat
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind    : Module de classe
' Purpose : Gestion des statistiques et de leurs persistances
' Author  : Eric VENTALON
' Date    : 22/05/2019
' ------------------------------------------------------
Option Explicit

Private m_lMatriceStat()    As Long
Private m_sLibelle()        As String
Private m_sNomDesFichiers() As String
Private m_sPath             As String
Private m_objDesFichier()   As Scripting.TextStream

Private m_objParam          As New ManageParam
Private m_objLogApplicatif  As New ManageLog

Private m_objFso            As New Scripting.FileSystemObject

' ----------------------------------------------------------------
' Procedure Name : addStat
' Purpose : Initialise un compteur de statistique
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter lIndice (Long): Numéro de stat
' Parameter sLibelle (String): Libelle d'affichage des stats
' Parameter sFichier (String): Détail des erreurs dans le fichier
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Sub addStat(lIndice As Long, sLibelle As String, Optional sFichier As String = "")

    Dim sFileTimeTs As String

    sFileTimeTs = Format(Now(), "yyyyMMdd-hhmmss")
    
    ReDim Preserve m_lMatriceStat(UBound(m_lMatriceStat) + 1) As Long
    m_lMatriceStat(lIndice) = 0
    
    ReDim Preserve m_sLibelle(UBound(m_sLibelle) + 1) As String
    m_sLibelle(lIndice) = sLibelle
    
    ReDim Preserve m_sNomDesFichiers(UBound(m_sNomDesFichiers) + 1) As String
    m_sNomDesFichiers(lIndice) = sFichier
    
    ReDim Preserve m_objDesFichier(UBound(m_objDesFichier) + 1) As Scripting.TextStream
    If sFichier <> "" Then
       Set m_objDesFichier(lIndice) = m_objFso.CreateTextFile(m_sPath + sFileTimeTs + "_" + sFichier + "_erreur.csv")
    Else
       Set m_objDesFichier(lIndice) = Nothing
    End If
    
End Sub

' ----------------------------------------------------------------
' Procedure Name : openStat
' Purpose : Initialise les stats
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sParam (String): Nom du parametre de localisation des fichiers
' Return Type: Boolean
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Function openStat(sParam As String) As Boolean

    openStat = False
    
    Call m_objLogApplicatif.InitManageLog
    If m_objParam.initManageParam(m_objLogApplicatif) <> 0 Then
        Exit Function
    End If
      
    m_sPath = m_objParam.getParams(sParam)
    If m_sPath = "" Then
        Exit Function
    End If
    
    ReDim m_lMatriceStat(0) As Long
    ReDim m_sLibelle(0) As String
    ReDim m_sNomDesFichiers(0) As String
    ReDim m_objDesFichier(0) As Scripting.TextStream
    
    openStat = True

End Function

' ----------------------------------------------------------------
' Procedure Name : sumStat
' Purpose : Addition dans les stats et ecriture dans le fichier ou pas
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter lIndice (Long): Le numro de stat
' Parameter lLigne (Long): La ligne dans le fichier
' Parameter sValueFlux (String): Valeur du flux
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Sub sumStat(lIndice As Long, Optional lLigne As Long = 0, Optional sValueFlux As String = "")
   
    m_lMatriceStat(lIndice) = m_lMatriceStat(lIndice) + 1
    If m_sNomDesFichiers(lIndice) <> "" Then
       m_objDesFichier(lIndice).Write lLigne & ";" & sValueFlux & Chr(13)
    End If
   
End Sub
' ----------------------------------------------------------------
' Procedure Name: displayStat
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter lIndiceBegin (Long):
' Parameter lIndiceEnd (Long):
' Author: Eric VENTALON
' Date: 11/06/2019
' ----------------------------------------------------------------
Public Sub displayStat(lIndiceBegin As Long, lIndiceEnd As Long)
   
   Dim iCpt As Long
   
   For iCpt = lIndiceBegin To lIndiceEnd
       Call m_objLogApplicatif.writeInfo("..." + m_sLibelle(iCpt) & m_lMatriceStat(iCpt))
   Next iCpt
   Call m_objLogApplicatif.displayMessage
   
End Sub



' ----------------------------------------------------------------
' Procedure Name : closeStat
' Purpose : Fermeture des stats
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Public Sub closeStat()
    Dim lCpt As Long
   
    For lCpt = 0 To UBound(m_objDesFichier) - 1
        If Not m_objDesFichier(lCpt) Is Nothing Then
            Call m_objDesFichier(lCpt).Close
            Set m_objDesFichier(lCpt) = Nothing
        End If
    Next lCpt
   
    ReDim m_lMatriceStat(0) As Long
    ReDim m_sLibelle(0) As String
    ReDim m_sNomDesFichiers(0) As String
    ReDim m_objDesFichier(0) As Scripting.TextStream

End Sub


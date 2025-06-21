VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manip_Tasking_AAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Manip_Tasking_AAR
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Resposible for manipulating the Tasking_AAR Word document for either closing or opening a mission.
'
' @class Manip_Tasking_AAR
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private pTeamFolderPath As String
Private pTaskingFileName As String
Private pTaskingFilePath As String

' ============================================= '
' Properties
' ============================================= '

''
' Get the file path in the team folder that contains this mission
'
' @property Resource
' @type String
''

Public Property Get TaskingFilePath() As String
    TaskingFilePath = pTaskingFilePath
End Property

' ============================================= '
' Public Methods
' ============================================= '

Public Sub NewManip(aNetworkDrivePath As String, aLastName As String, aFirstName As String, aDate As Date)
    pTeamFolderPath = aNetworkDrivePath & "\Team_Folder"
    SetTaskingFileName aLastName, aFirstName
    SetFilePath aDate
End Sub

Public Sub FillTaskingAAR(ByRef aDataArr As Variant, openFilePath As String, startAt As Integer)
    Const sProcName As String = "FillTaskingAAR"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    
    Dim nameArr As Variant
    nameArr = Split(Environ("username"), ".")
    
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(openFilePath)
    wdApp.Visible = False
    
    wdApp.Run "ThisDocument.Fill", aDataArr, startAt

    wdApp.Visible = True
    wdApp.Activate
    wdDoc.SaveAs2 (pTaskingFilePath)
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub
    
ErrorHandler:
   Call Error_Handle(sProcName, Err.Number, Err.Description)
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub SetTaskingFileName(aLastName As String, aFirstName As String)
    pTaskingFileName = aLastName & ",_" & aFirstName & ".docm"
End Sub

Private Sub SetFilePath(aDate As Date)
    pTaskingFilePath = FindMishFol("Mission", Format(aDate, "mmm"), Format(aDate, "yy"), pTeamFolderPath) & "\" & pTaskingFileName
End Sub

Private Function FindMishFol(ByVal fString As String, ByVal mmmString As String, ByVal yyString As String, ByVal fPath As String) As String
    Const sProcName As String = "FindMishFol"
   If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim subFol, monthFol, parentFol As Scripting.Folder
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    Set parentFol = fso.GetFolder(fPath)

    For Each subFol In parentFol.SubFolders
        If InStr(1, subFol.Name, fString, vbTextCompare) Then
            Set subFol = fso.GetFolder(subFol.Path)
            For Each monthFol In subFol.SubFolders
                If InStr(1, monthFol.Name, mmmString, vbTextCompare) And InStr(1, monthFol.Name, yyString, vbTextCompare) Then
                    FindMishFol = monthFol.Path
                    Exit Function
                End If
            Next monthFol
            'If the month folder isn't found, then it's made and gets returned as result of this function
            FindMishFol = fso.CreateFolder(subFol.Path & "\" & UCase(mmmString) & "_" & yyString).Path
        End If
    Next subFol
   GoTo SetNothing
ErrorHandler:
   Call Error_Handle(sProcName, Err.Number, Err.Description)
   MsgBox mmmString & " " & yyString
SetNothing:
    Set fso = Nothing
    Set parentFol = Nothing
    Set subFol = Nothing
End Function

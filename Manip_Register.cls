VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manip_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Manip_Register
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Resposible for manipulating the Register for either closing or opening a mission.
'
' @class Manip_Register
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private pFuneralDate As Date
Private pRegisterPath As String
Private pInitials As String
Private manip_p_ExcelApp As Excel.Application
Private manip_p_RegBook As Excel.Workbook
Private manip_p_RegSheet As Excel.Worksheet

' ============================================= '
' Public Methods
' ============================================= '

Public Sub NewManip(aFuneralDate As Date, aRegisterPath As String)
    pFuneralDate = aFuneralDate
    pRegisterPath = aRegisterPath
    pInitials = Initials(Environ("username"), ".")
    Set manip_p_ExcelApp = New Excel.Application
    manip_p_ExcelApp.Visible = False
End Sub

Public Function MissionToRegister(aDataArr As Variant) As String
    Const sProcName As String = "MissionToRegister"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim openSlot As Excel.Range
    Dim cellColumn As Integer
    Dim MissionNumber As Excel.Range
    
    If Not OpenRegister Then
        Exit Function
    End If
    
    With manip_p_RegSheet
        .Activate
        'find and select the first empty slot
        Set openSlot = .Cells(.Rows.Count, 2).End(xlUp).Offset(1, 0)
        .Range(openSlot, openSlot.Offset(0, 9)).Select
    End With
    
    Dim Response As Integer
    
    Set MissionNumber = GetMissionNumber(openSlot)
    If MissionNumber Is Nothing Then
        Exit Function
    End If
    
    Response = MsgBox("Assign to mission number " & openSlot.Offset(0, -1).Value & "?", vbQuestion + vbYesNo)
    If Response = vbNo Then
        Exit Function
    End If
    
    'place data in each corresponding cell
    For cellColumn = 0 To 9
        openSlot.Offset(0, cellColumn).Value = aDataArr(cellColumn)
    Next cellColumn
    
    'make the cells the correct color
    For cellColumn = 7 To 9
        Select Case openSlot.Offset(0, cellColumn)
        Case "WAARNG"
            openSlot.Offset(0, 6).Interior.Color = RGB(255, 255, 0) 'yellow
        Case "TNCVHG"
            openSlot.Offset(0, cellColumn).Interior.Color = RGB(196, 215, 155) 'light green
        Case "WSVCSG"
            openSlot.Offset(0, cellColumn).Interior.Color = RGB(118, 147, 60) 'dark green
        Case Is <> vbNullString
            openSlot.Offset(0, cellColumn).Interior.Color = RGB(230, 184, 183) 'pinkish red
        End Select
    Next cellColumn
    
    openSlot.Offset(0, 10).Value = pInitials
    MissionToRegister = GetMissionNumber(openSlot)
    manip_p_RegBook.Save
    Exit Function
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Function

Public Sub FinishRegister(MissionNumber As String, Mileage As Integer)
    Const sProcName As String = "FinishRegister"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim missionSlot As Excel.Range
    
    If Not OpenRegister Then
        Exit Sub
    End If
    
    Set missionSlot = SearchForMission(MissionNumber)
    
    missionSlot.Offset(0, 11).Value = missionSlot.Offset(0, 11).Value & "/" & pInitials
    missionSlot.Offset(0, 12).Value = pInitials
    missionSlot.Offset(0, 15).Value = Mileage
    
    manip_p_RegBook.Save
    Exit Sub
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Class_Terminate()
    manip_p_RegBook.Close SaveChanges:=False
    manip_p_ExcelApp.Quit
    Set manip_p_RegSheet = Nothing
    Set manip_p_RegBook = Nothing
    Set manip_p_ExcelApp = Nothing
End Sub

Private Function GetMissionNumber(ByRef Slot As Excel.Range) As Excel.Range
    If Not IsEmpty(Slot.Offset(0, -1).Value2) Then
        Set GetMissionNumber = Slot.Offset(0, -1)
    Else
        MsgBox "The mission numbers are not listed out far enough. Please fix this in the register.", vbExclamation
    End If
End Function

Private Function OpenRegister() As Boolean
    Const sProcName As String = "OpenRegister"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Set manip_p_RegBook = manip_p_ExcelApp.Workbooks.Open(FindRegister("Tasking", Year(pFuneralDate), pRegisterPath))
    
    If manip_p_RegBook.ReadOnly = True Then
        MsgBox "The register is in read-only mode. Someone else is probably using it. Please try again later.", vbExclamation
        OpenRegister = False
        Exit Function
    End If
    OpenRegister = True
    Set manip_p_RegSheet = manip_p_RegBook.Sheets("Register")
    Exit Function
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Function

Private Function SearchForMission(MissionNumber As String) As Excel.Range
    With manip_p_RegSheet.Range("A1:A500")
        Set SearchForMission = .Find(MissionNumber)
    End With
End Function

Private Function FindRegister(ByVal fString As String, ByVal yyyyString As String, ByVal fPath As String) As String
    Const sProcName As String = "FindRegister"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim subFol, parentFol As Scripting.Folder
    Dim subFil As Scripting.File
    Dim fso As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    Set parentFol = fso.GetFolder(fPath)

    For Each subFol In parentFol.SubFolders
        If InStr(1, subFol.Name, fString, vbTextCompare) Then
            Set subFol = fso.GetFolder(subFol.Path)
            For Each subFil In subFol.Files
                If InStr(1, subFil.Name, "Register", vbTextCompare) > 0 And InStr(1, subFil.Name, yyyyString, vbTextCompare) > 0 Then
                    FindRegister = subFil.Path
                    Exit Function
                End If
            Next subFil
        End If
    Next subFol
    GoTo SetNothing
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
SetNothing:
    Set fso = Nothing
    Set parentFol = Nothing
    Set subFol = Nothing
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Search For Mission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Search For Mission
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form allows the user to search for records in the 'Missions' table to open them in either the
' 'Intake Mission' or 'Close Mission' form.
'
' @form Search For Mission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Form Events
' ============================================= '

Private Sub Form_Load()
    'Me.Caption = Me.OpenArgs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Ctrl As Control

    For Each Ctrl In Controls
        If Ctrl.ControlType = 109 Or Ctrl.ControlType = 111 Then
            Ctrl.Value = Ctrl.DefaultValue
        End If
    Next
End Sub

' ============================================= '
' Button Events
' ============================================= '

Private Sub MishNumSearch_Click()
    FindMissionAndOpenAAR Me.MissionNumber
End Sub

Private Sub SSNSearch_Click()
    FindMissionAndOpenAAR Me.SSN
End Sub

''
' Takes a parameter and searchs the Missions table for a matching record, then either opens the 'Close Mission'
' or 'Intake Mission'
'
' @param srchParam is a Textbox with either a SSN or a Mission Number
''
Private Sub FindMissionAndOpenAAR(ByRef srchParam As TextBox)
    Dim db As DAO.Database
    Dim Rec As DAO.Recordset

    Set db = CurrentDb
    Set Rec = db.OpenRecordset("Missions")

    'srchParam has a tag with the column in Missions associated with the contents of srchParam
    'this tag will be the index (column) of the record search
    Rec.Index = srchParam.Tag
    Rec.Seek "=", Format(srchParam, srchParam.InputMask)

    If Rec.NoMatch Then
        MsgBox "No mission exists with that " & srchParam.Tag
    ElseIf Me.OpenArgs = "Close Mission" Then
        'closes AAr if it is open, and opens it, setting its current record with an SQL statement
        DoCmd.Close acForm, "Close Mission", acSaveYes
        DoCmd.OpenForm "Close Mission", acNormal, vbNullString, _
                        WhereCondition:=srchParam.Tag & "='" & Format(srchParam, srchParam.InputMask) & "'", _
                        OpenArgs:="Close"
        DoCmd.Close acForm, "Search For Mission", acSaveYes
    ElseIf Me.OpenArgs = "Existing Mission" Then
        DoCmd.Close acForm, "Intake Mission", acSaveYes
        DoCmd.OpenForm "Intake Mission", acNormal, vbNullString, _
                        WhereCondition:=srchParam.Tag & "='" & Format(srchParam, srchParam.InputMask) & "'", _
                        OpenArgs:="Modify"
        DoCmd.Close acForm, "Search For Mission", acSaveYes
    End If
End Sub


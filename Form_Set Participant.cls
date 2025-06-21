VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Set Participant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Missions_Honor Guard Members
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is an interface between MFH admin personnel and the 'Missions_Honor Guard Members' table.
' It allows them to add/edit records in the 'Missions_Honor Guard Members' table of this database in a user-friendly
' manner while preventing them from entering incomplete data or accidentally corrupting mission data.
'
' @form Add Funeral Director
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

''
' Communicates with the parent form (Close Mission)
''
Public Event ParticipantAdded(aType As String)

Private Const HGMsql = "SELECT [Honor Guard Members].[ID], [Honor Guard Members].[Last Name], [Honor Guard Members].[First Name], [Honor Guard Members].[Rank] " & _
                       "FROM [Honor Guard Members] " & _
                       "ORDER BY [Last Name], [First Name], [Rank]; "

Private p_ParticipantType As String
Private p_ParticipantField As String
Private p_SSN As String
Private p_ID As Variant

' ============================================= '
' Form Events
' ============================================= '

Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.OpenArgs) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    p_ParticipantType = Split(Me.OpenArgs, "|")(0)
    p_SSN = Split(Me.OpenArgs, "|")(1)
    p_ID = Split(Me.OpenArgs, "|")(2)
    SetParticipantField
    If p_ParticipantType = "HGM" Then
        Me.RecordSource = "Missions_Honor Guard Members"
        Participant_Label.Caption = "Team Member"
        Participant.RowSource = HGMsql
        Participant.ControlSource = "HG Member ID"
        Participant.ColumnCount = 4
        Participant.ColumnWidths = "0;1 in;1 in;1 in"
        Participant.ListWidth = "3 in"
        If p_ID <> vbNullString Then
            Me.Caption = "Edit Honor Guard Member"
            Recordset.FindFirst "[" & Me.RecordSource & "]." & "[Deceased SSN]='" & p_SSN & "' AND " & "[" & Me.RecordSource & "]." & "[" & p_ParticipantField & "]=" & p_ID
        ElseIf p_ID = vbNullString Then
            Me.Caption = "Add Honor Guard Member"
            Deceased_SSN = p_SSN
        End If
    ElseIf p_ParticipantType = "VSO" Then
        Me.RecordSource = "Missions_VSOs"
        Participant_Label.Caption = "VSO"
        Participant.RowSource = "VSOs"
        Participant.ControlSource = "VSO"
        Participant.ColumnCount = 1
        Participant.ColumnWidths = "1 in"
        Participant.ListWidth = "1 in"
        If p_ID <> vbNullString Then
            Me.Caption = "Edit VSO"
            Recordset.FindFirst "[" & Me.RecordSource & "]." & "[Deceased SSN]='" & p_SSN & "' AND " & "[" & Me.RecordSource & "]." & "[" & p_ParticipantField & "]='" & p_ID & "'"
        ElseIf p_ID = vbNullString Then
            Me.Caption = "Add VSO"
            Deceased_SSN = p_SSN
        End If
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNull(Participant) Then
        MsgBox "Please enter a participant before saving.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Form_Close()
    Me.RecordSource = vbNullString
    Me.Caption = ""
End Sub

' ============================================= '
' Button Events
' ============================================= '

Private Sub Save_Click()
On Error GoTo Save_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, vbNullString
    Else
        DoCmd.Close acForm, "Set Participant", acSaveYes
    End If
    RaiseEvent ParticipantAdded(p_ParticipantType)

Save_Click_Exit:
    Exit Sub

Save_Click_Err:
    MsgBox Error$
    Resume Save_Click_Exit

End Sub

Private Sub Delete_Click()
On Error GoTo Delete_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.newRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.newRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.newRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    Else
        DoCmd.Close acForm, "Set Participant", acSaveYes
    End If

    RaiseEvent ParticipantAdded(p_ParticipantType)
Delete_Click_Exit:
    Exit Sub

Delete_Click_Err:
    MsgBox Error$
    Resume Delete_Click_Exit

End Sub

' ============================================= '
' Private Subroutines/Functions
' ============================================= '

Private Sub GoToRecord(SSN As String, FieldName As String, ID As Variant)

        
End Sub

Private Sub SetParticipantField()
    If p_ParticipantType = "HGM" Then
        p_ParticipantField = "HG Member ID"
    ElseIf p_ParticipantType = "VSO" Then
        p_ParticipantField = "VSO"
    End If
End Sub






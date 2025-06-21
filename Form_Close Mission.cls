VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Close Mission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Close Mission

' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is an interface between MFH admin personnel and the mission records.
' It allows them to finish records in the "Missions" table of this database in a user-friendly
' manner while preventing them from entering incomplete data or accidentally corrupting mission data.
'
' @form Close Mission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private out_p_CloseMission As New CloseMission
Private out_p_HGMColl As Collection, out_p_VSOColl As Collection
Private out_p_HGMsql As String, out_p_VSOsql As String

''
' Subforms whose events must be subscribed to by this form.
''
Private WithEvents out_p_SetParticipant As [Form_Set Participant]
Attribute out_p_SetParticipant.VB_VarHelpID = -1

' ============================================= '
' Form Events
' ============================================= '

''
' This form cannot be opened unless an argument is passed to it on open.
''
Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.OpenArgs) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Recordset("Last Name") & ", " & Me.Recordset("First Name") & " (" & Me.Recordset("SSN") & ")"
    SetSQLStatements
End Sub

Private Sub Form_Close()
    Me.Caption = vbNullString
End Sub

Private Sub out_p_SetParticipant_ParticipantAdded(aType As String)
    If aType = "HGM" Then
        SMList.Requery
    ElseIf aType = "VSO" Then
        VSOList.Requery
    End If
    SetMembers aType
End Sub

' ============================================= '
' Control Events
' ============================================= '

Private Sub SMList_DblClick(Cancel As Integer)
    If SMList.ItemsSelected.Count = 0 Or SMList.ListCount = 0 Then
        Cancel = True
    Else
        OpenParticipantsForm "HGM", SSN, SMList
    End If
End Sub

Private Sub VSOList_DblClick(Cancel As Integer)
    If VSOList.ItemsSelected.Count = 0 Or VSOList.ListCount = 0 Then
        Cancel = True
    Else
        OpenParticipantsForm "VSO", SSN, VSOList
    End If
End Sub


' ============================================= '
' Button Events
' ============================================= '

Private Sub Close_Register_Click()
    SetMission
    out_p_CloseMission.FinishRegister
End Sub

Private Sub Close_Tasking_Click()
    SetMission
    out_p_CloseMission.FinishTasking
End Sub

Private Sub Command42_Click()
    DoCmd.OpenForm "Close Mission", acDesign
End Sub

Private Sub Add_SM_Click()
On Error GoTo Add_SM_Click_Err

'    DoCmd.OpenForm "Missions_Honor Guard Members", acNormal, OpenArgs:=SSN, DataMode:=acFormAdd
'    Set out_p_AddHGMForm = Forms("Missions_Honor Guard Members")

    OpenParticipantsForm "HGM", SSN
Add_SM_Click_Exit:
    Exit Sub

Add_SM_Click_Err:
    MsgBox Error$
    Resume Add_SM_Click_Exit

End Sub

Private Sub Add_VSO_Click()
On Error GoTo Add_VSO_Click_Err

'    DoCmd.OpenForm "Missions_VSOs", acNormal, OpenArgs:=SSN, DataMode:=acFormAdd
'    Set out_p_AddVSOForm = Forms("Missions_VSOs")
    OpenParticipantsForm "VSO", SSN

Add_VSO_Click_Exit:
    Exit Sub

Add_VSO_Click_Err:
    MsgBox Error$
    Resume Add_VSO_Click_Exit

End Sub

Private Sub Save_Click()
On Error GoTo Save_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, vbNullString
            End If

    NewCloseMission (Me.Recordset)
Save_Click_Exit:
    Exit Sub

Save_Click_Err:
    MsgBox Error$
    Resume Save_Click_Exit
End Sub

' ============================================= '
' Private Subroutines/Functions
' ============================================= '

Private Sub OpenParticipantsForm(aType As String, aSSN As String, Optional ID As String)
    DoCmd.Close acForm, "Set Participant", acSaveYes
    If ID = vbNullString Then
        DoCmd.OpenForm "Set Participant", acNormal, OpenArgs:=aType & "|" & aSSN, DataMode:=acFormAdd
    Else
        DoCmd.OpenForm "Set Participant", acNormal, OpenArgs:=aType & "|" & aSSN & "|" & ID, DataMode:=acFormEdit
    End If
    Set out_p_SetParticipant = Forms("Set Participant")
End Sub

Private Sub NewCloseMission(aRecord As DAO.Recordset)
    out_p_CloseMission.NewCloseMission aRecord
End Sub

''
' Runs the query that looks up the current form's record. Opens an 'CloseMission' object with that record info.
''
Private Sub SetMission()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim prm As DAO.Parameter

    DoCmd.RunCommand acCmdSaveRecord

    Set db = CurrentDb()
    Set qdf = db.QueryDefs("SSN Lookup (Close Mission)")

    For Each prm In qdf.Parameters
        prm = Eval(prm.Name)
    Next prm

    out_p_CloseMission.NewCloseMission qdf.OpenRecordset(dbOpenDynaset)
End Sub

''
' Decides which collection to load with the records in
'
' @param {String} member is either 'HG' or 'VSO'
''
Private Sub SetMembers(member As String)
    SetSQLStatements
    If member = "VSO" Then
        GetParticipants out_p_VSOColl, out_p_VSOsql
    ElseIf member = "HGM" Then
        GetParticipants out_p_HGMColl, out_p_HGMsql
    End If
End Sub
''
' Sets the SQL lookup to get the participants based on the SSN associated with this record
''
Private Sub SetSQLStatements()
    out_p_HGMsql = "SELECT [Missions_Honor Guard Members].[Deceased SSN], [Missions_Honor Guard Members].[HG Member ID],[Honor Guard Members].[Last Name], [Honor Guard Members].[First Name], [Honor Guard Members].Rank, [Missions_Honor Guard Members].[Positions] " & _
        "FROM [Honor Guard Members] INNER JOIN [Missions_Honor Guard Members] ON [Honor Guard Members].ID = [Missions_Honor Guard Members].[HG Member ID] " & _
        "WHERE ((([Missions_Honor Guard Members].[Deceased SSN])=""" & Me.Recordset("SSN") & """));"

    out_p_VSOsql = "SELECT [Missions_VSOs].[Deceased SSN], [Missions_VSOs].[VSO], [Missions_VSOs].[Positions] " & _
        "FROM [VSOs] INNER JOIN [Missions_VSOs] ON [VSOs].[Name (Register)] = [Missions_VSOs].[VSO] " & _
        "WHERE ((([Missions_VSOs].[Deceased SSN])=""" & Me.Recordset("SSN") & """));"
End Sub

''
' Grabs the records associated with this mission from the 'Missions_VSOs' and 'Missions_Honor Guard Members'
' and formats the data and adds it to the appropriate collection
'
' @param {Collection} partColl the collection (either honor guard or VSO) to load with the data from the participant records
' @param {String} sql the sql command that will call the records to be loaded into this form's collection
Private Sub GetParticipants(partColl As Collection, sql As String)
    Dim rs As DAO.Recordset, childRS As DAO.Recordset
    Dim field As Variant, pos As Field2
    Dim Participants As New Collection
    Dim dict As Scripting.Dictionary
    Dim pstns As String

    Set rs = CurrentDb.OpenRecordset(sql)

    Do While Not rs.EOF

        Set dict = New Scripting.Dictionary
        For Each field In rs.Fields
            With field
                If TypeName(.Value) = "String" Or TypeName(.Value) = "Long" Then
                    dict.Add .Name, .Value
                ElseIf TypeName(.Value) = "Recordset2" Then
                    ' Open a Recordset for the multivalued field.
                    Set childRS = .Value
                    ' Exit the loop if the multivalued field contains no records.
                    Do Until childRS.EOF
                        childRS.MoveFirst
                        pstns = vbNullString
                        ' Loop through the records in the child recordset.
                        Do Until childRS.EOF
                            If pstns = vbNullString Then
                                pstns = childRS!Value.Value
                            Else
                                pstns = pstns & vbCrLf & childRS!Value.Value
                            End If
                            childRS.MoveNext
                        Loop
                    Loop
                    dict.Add .Name, pstns
                End If
            End With
        Next field

        Participants.Add dict

        rs.MoveNext

    Loop

    Set partColl = Participants
End Sub

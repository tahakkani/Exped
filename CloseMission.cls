VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CloseMission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' CloseMission
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Resposible for manging all 'closing out' tasks.
'
' @class CloseMission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private pMission As Mission
Private pAARData(77) As Variant
Private pG1PortalParticipantData As Variant
Private pDriver As DAO.Recordset
Private pHGMembers As DAO.Recordset
Private pFirers As Integer

Private pCloseAAR As Manip_Tasking_AAR
Private pCloseRegister As Manip_Register
Private pCloseG1Portal As Manip_GovSite

' ============================================= '
' Properties
' ============================================= '

''
' Set the number of firers on this mission.
'
' @property Firers
' @type Integer
''
Private Property Set Firers(Participants As DAO.Recordset)
    Dim field As Field2
    Dim childRS As DAO.Recordset
    Dim firerCount As Integer

    firerCount = 0

    Do Until Participants.EOF
        For Each field In Participants.Fields
            With field
                If .Type = 109 Then
                    Set childRS = .Value
                        ' Exit the loop if the multivalued field contains no records.
                        Do Until childRS.EOF
                            childRS.MoveFirst
                            ' Loop through the records in the child recordset.
                            Do Until childRS.EOF
                                If childRS!Value.Value = "Firing Party" Then
                                    firerCount = firerCount + 1
                                End If
                                childRS.MoveNext
                            Loop
                        Loop
                End If
            End With
        Next field

        Participants.MoveNext
    Loop

    Participants.MoveFirst
    pFirers = firerCount
End Property

''
' Set the honor guard members on this mission
'
' @property HGMembers
' @type DAO.Recordset
''
Private Property Set HGMembers(socSecNum As Variant)

    Dim sql As String
    sql = "SELECT *" & _
            "FROM [Missions_Honor Guard Members]" & _
            "WHERE [Missions_Honor Guard Members]![Deceased SSN]='" & socSecNum & "'"

    Set pHGMembers = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    'ListParticipants pHGMembers
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Basically a workaround for a constructor
''
Sub NewCloseMission(aRecord As DAO.Recordset)
    pMission.NewMission aRecord
    Set HGMembers = pMission.Record("SSN")
    Set Firers = pHGMembers
End Sub

Sub FinishG1Portal()
    Set pCloseG1Portal = New Manip_GovSite
    SetPortalParticipantData
    pCloseG1Portal.CloseMissionG1Portal pMission.Record("SSN"), pG1PortalParticipantData, pDriver, pMission.Record("GSA Mileage")
    pCloseG1Portal = Nothing
End Sub
Sub FinishTasking()
    Set pCloseAAR = New Manip_Tasking_AAR
    pCloseAAR.NewManip pMission.NetworkDrivePath, pMission.Record("Last Name"), _
                        pMission.Record("First Name"), pMission.Record("Funeral Date")
    SetAARData
    pCloseAAR.FillTaskingAAR pAARData, pCloseAAR.TaskingFilePath, 60
    pCloseAAR = Nothing
End Sub

Sub FinishRegister()
    Set pCloseRegister = New Manip_Register
    pCloseRegister.NewManip pMission.Record("Funeral Date"), pMission.StaffFolderPath
    pCloseRegister.FinishRegister pMission.Record("Mission Number"), pMission.Record("GSA Mileage")
    Set pCloseRegister = Nothing
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Class_Initialize()
    pFirers = 0
    Set pMission = New Mission
End Sub

Private Sub SetDriver()
    Set pDriver = SeekRecord("Honor Guard Members", "Participant ID", pMission.Record("Who Drove"))
End Sub
Private Sub SetPortalParticipantData()
    Dim part As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT [Missions_Honor Guard Members].[Deceased SSN], [Honor Guard Members].[Last Name], [Honor Guard Members].[First Name], [Missions_Honor Guard Members].Positions, [Honor Guard Members].ID " & _
            "FROM [Honor Guard Members] INNER JOIN [Missions_Honor Guard Members] ON [Honor Guard Members].ID = [Missions_Honor Guard Members].[HG Member ID] " & _
            "WHERE ((([Missions_Honor Guard Members].[Deceased SSN])='" & pMission.Record("SSN") & "'))"
    
    Set part = CurrentDb.OpenRecordset(sql)
    Set pG1PortalParticipantData = LoadParticipantsToArray(part)
End Sub
Private Sub SetAARData()
    Dim i As Integer

    'CHECKBOXES ON SECOND PAGE
    For i = 1 To 38
        Select Case i
        Case 9 To 16
            pAARData(i) = vbNullString 'coordinators/POCs (don't need to fill out)
        Case 29, 30, 33
            pAARData(i) = 0 'full honors check boxes
        Case Else
            pAARData(i) = 1
        End Select
    Next i

    'COUNTY
    pAARData(39) = pMission.LookUpCounty

    'PLACE OF SERVICE
    pAARData(40) = 0
    pAARData(41) = 0
    pAARData(42) = 0

    Select Case pMission.Location("Place of Service (for Tasking)")
    Case "Government Cemetery"
        pAARData(40) = 1
    Case "Private Cemetery"
        pAARData(41) = 1
    Case "Chapel/Memorial"
        pAARData(42) = 1
    End Select

    'FUNCTIONS REQUESTED
    pAARData(43) = pMission.Record("Flag Fold")
    pAARData(44) = pMission.Record("Taps")
    pAARData(45) = pMission.Record("Pallbearer")
    pAARData(46) = pMission.Record("Firing Detail")
    pAARData(47) = pMission.Record("Chaplain")
    pAARData(48) = 0
    pAARData(49) = pMission.Record("Planeside")

    'TAPS PLAYED BY:
    pAARData(50) = 0 'Bugler
    pAARData(51) = 0 'CD Player/Cermonial Bugler

    'INDIVIDUALS ASSIGNED TO MISSION:
    pAARData(52) = 0    'OIC
    pAARData(53) = 0    'NCOIC
    pAARData(54) = 0    'Chaplain
    pAARData(55) = 0    'Bugler
    pAARData(56) = 0    'Pallbearers/FiringDetail (Integer from 0 to 7)
    pAARData(57) = 0    'Flag Fold
    pAARData(58) = 0    'Admin Support (Always 0)

    'TEAM MEMBERS
    SetAARTeamMembers

    ''NOK REQ BRASS (Y/N)

    If pMission.Record("Rounds Given") > 0 Then
        pAARData(66) = 1
        pAARData(67) = 0
    ElseIf pMission.Record("Rounds Given") = 0 Then
        pAARData(66) = 0
        pAARData(67) = 0
    End If

    'ROUNDS EXPENDED
    pAARData(68) = (3 * pFirers)
    'ROUNDS GIVEN TO FAMILY
    pAARData(69) = pMission.Record("Rounds Given")

    'TOTAL TIME ON MISSION
    pAARData(70) = pMission.Record("Time On Mission")

    'ROUND TRIP MILEAGE
    pAARData(71) = pMission.Record("GSA Mileage")

    'VSO SUPPORT (Y/N)
    If pMission.VSOs.RecordCount > 0 Then
        pAARData(72) = 1
        pAARData(73) = 0
    Else
        pAARData(72) = 0
        pAARData(73) = 1
    End If
    'VSOS:
    pAARData(74) = pMission.ListVSOs

    'REMARKS
    pAARData(75) = pMission.Record("Additional Remarks")

    'PERSON COMLETING FORM
    pAARData(76) = FormatSMName(SeekTeamMembers(pMission.Record.Fields("Person Completing AAR")), True)

    'DATE COMPLETING FORM
    pAARData(77) = Date

    RemoveNulls pAARData
End Sub

Private Sub SetAARTeamMembers()
    Dim i As Integer
    Dim field As field
    i = 59

    Do Until pHGMembers.EOF
        For Each field In pHGMembers.Fields
            With field
                If .Name = "HG Member ID" Then
                    pAARData(i) = FormatSMName(SeekTeamMembers(field.Value), True)
                    i = i + 1
                End If
            End With
        Next field

        pHGMembers.MoveNext
    Loop
End Sub

Public Function SeekTeamMembers(participantID As Integer) As DAO.Recordset
    Dim rstHGMembers As DAO.Recordset

    Set rstHGMembers = CurrentDb.OpenRecordset("Honor Guard Members")

    rstHGMembers.Index = "Participant ID"
    rstHGMembers.Seek "=", participantID

    If rstHGMembers.NoMatch Then
        Set SeekTeamMembers = Null
    Else
        Set SeekTeamMembers = rstHGMembers
    End If

    Set rstHGMembers = Nothing
End Function


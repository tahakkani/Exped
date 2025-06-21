VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenMission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' OpenMission
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Responsible for all 'new mission' tasks, such as making taskings, putting the mission on the Sharepoint
' calendar/Register/G1 Portal.
'
' @class OpenMission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit


' ============================================= '
' Constants and Private Variables
' ============================================= '

Private pMission As Mission
Private pTaskingData(59) As Variant
Private pRegisterData(10) As Variant
Private pCalendarData(8) As Variant
Private pG1PortalData(9, 3) As Variant

Private pGrade As String
Private pDirector As DAO.Recordset
Private pFuneralHome As DAO.Recordset

Private pToTasking As Manip_Tasking_AAR
Private pToRegister As Manip_Register
Private pToCalendar As Manip_GovSite
Private pToG1Portal As Manip_GovSite

' ============================================= '
' Properties
' ============================================= '

''
' Get the mission record
'
'
' @property Mission
' @type Mission
''
Public Property Get Mission() As Mission
    Set Mission = pMission
End Property

' ============================================= '
' Public Methods
' ============================================= '

Sub NewOpenMission(aRecord As DAO.Recordset)
    pMission.NewMission aRecord
    SetGrade
    SetDirector
    SetFuneralHome
End Sub

Public Sub ToG1Portal()
    Set pToG1Portal = New Manip_GovSite
    SetG1PortalData
    pToG1Portal.MissionToG1Portal pMission.Record("SSN"), pG1PortalData, pMission.Record("Remarks"), _
                                    Format(pMission.Record("Funeral Date"), "yyyymmdd"), pMission.Record("Funeral Time"), _
                                    pMission.Location("Name")
    
End Sub
Public Sub ToCalendar()
    Set pToCalendar = New Manip_GovSite
    SetCalendarData
    pToCalendar.MissionToCalendar pCalendarData
    Set pToCalendar = Nothing
End Sub
Sub MakeTasking()
    Set pToTasking = New Manip_Tasking_AAR
    SetTaskingData
    pToTasking.NewManip pMission.NetworkDrivePath, pMission.Record("Last Name"), _
                        pMission.Record("First Name"), pMission.Record("Funeral Date")
    pToTasking.FillTaskingAAR pTaskingData, pMission.ProjectPath & "\Tasking_AAR.docm", 1
    Set pToTasking = Nothing
End Sub

Public Function ToRegister() As String
    Set pToRegister = New Manip_Register
    SetRegisterData
    pToRegister.NewManip pMission.Record("Funeral Date"), pMission.StaffFolderPath
    ToRegister = pToRegister.MissionToRegister(pRegisterData)
    Set pToRegister = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Class_Initialize()
    Set pMission = New Mission
End Sub

Private Sub SetDirector()
    Dim db As DAO.Database
    Dim rstFunDirs As DAO.Recordset
    
    If g_HandleErrors Then On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rstFunDirs = db.OpenRecordset("Funeral Directors")
    
    rstFunDirs.Index = "Funeral Director ID"
    rstFunDirs.Seek "=", pMission.Record("Funeral Director ID")
    
    If rstFunDirs.NoMatch Then
      Set pDirector = Null
    Else
       Set pDirector = rstFunDirs
    End If
   
    Set rstFunDirs = Nothing
    Set db = Nothing
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub SetGrade()
    Dim db As DAO.Database
    Dim Rec As DAO.Recordset
    
    Set db = CurrentDb
    Set Rec = db.OpenRecordset("Ranks")
    
    Rec.Index = pMission.Record("Branch")
    Rec.Seek "=", pMission.Record("Rank")
    
    If Rec.NoMatch Then
        pGrade = vbNullString
    Else
        pGrade = Rec!Grade
    End If
End Sub

Private Sub SetFuneralHome()
    Dim rstFuneralHomes As DAO.Recordset
    
    If g_HandleErrors Then On Error GoTo ErrorHandler
    
    Set rstFuneralHomes = CurrentDb.OpenRecordset("Funeral Homes")
    
    rstFuneralHomes.Index = "Funeral Home ID"
    rstFuneralHomes.Seek "=", pDirector("Funeral Home ID")
    
    If rstFuneralHomes.NoMatch Then
        Set pFuneralHome = Null
    Else
        Set pFuneralHome = rstFuneralHomes
    End If
 
   Set rstFuneralHomes = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub SetG1PortalData()
    Dim thisMission As DAO.Recordset, missionLocation As DAO.Recordset
    Set thisMission = pMission.Record
    Set missionLocation = pMission.Location
    
    'set textbox values
    pG1PortalData(1, 1) = thisMission("Last Name")
    pG1PortalData(2, 1) = thisMission("First Name")
    pG1PortalData(3, 1) = thisMission("Middle Name")
    pG1PortalData(4, 1) = missionLocation("Street")
    pG1PortalData(6, 1) = missionLocation("City")
    pG1PortalData(7, 1) = missionLocation("ZIP")
    pG1PortalData(8, 1) = Format(Now(), "mm/dd/yyyy")
    pG1PortalData(9, 1) = Format(Now(), "hhmm")
    
    'set dropdown values
    pG1PortalData(1, 2) = thisMission("Suffix")
    pG1PortalData(2, 2) = missionLocation("State")
    pG1PortalData(3, 2) = thisMission("Component")
    pG1PortalData(4, 2) = thisMission("Rank")
    pG1PortalData(5, 2) = thisMission("Status")
    pG1PortalData(6, 2) = thisMission("Requester")
    pG1PortalData(7, 2) = missionLocation("Place of service (TAPS)")
    pG1PortalData(8, 2) = thisMission("Remains")
    
    'set checkboxes
    pG1PortalData(1, 3) = thisMission("Flag Fold")
    pG1PortalData(2, 3) = thisMission("Taps")
    pG1PortalData(3, 3) = thisMission("Firing Party")
    pG1PortalData(4, 3) = thisMission("Pallbearers")
    pG1PortalData(5, 3) = thisMission("Planeside Honors")
    pG1PortalData(6, 3) = 0
    pG1PortalData(7, 3) = 0
    
End Sub
Private Sub SetCalendarData()
    Dim FuneralDate As Date
    Dim formattedDate As String
    Dim FuneralTime As String
    Dim funeralTimePlusHalfHour As String
    
    pCalendarData(1) = CalendarTitle
    pCalendarData(2) = ""
    
    FuneralDate = pMission.Record("Funeral Date")
    formattedDate = Format(FuneralDate, "m/d/yyyy")
    pCalendarData(3) = formattedDate
    pCalendarData(4) = formattedDate
    
    FuneralTime = pMission.Record("Funeral Time")
    funeralTimePlusHalfHour = AddHalfHour(FuneralTime)
    pCalendarData(5) = Left(FuneralTime, 2) & ":"
    pCalendarData(6) = Right(FuneralTime, 2)
    pCalendarData(7) = Left(funeralTimePlusHalfHour, 2) & ":"
    pCalendarData(8) = Right(funeralTimePlusHalfHour, 2)
End Sub

Private Function CalendarTitle() As String
    Dim title As String, locationName As String, Team As String
    title = ""
    If pMission.Record("Status") = "Retired" Then
        title = title & "** "
    End If
    
    locationName = pMission.Location("Name")
    
    If locationName = "Washington State Veterans Cemetery" Or locationName = "Washington State Veteran's Cemetery" Then
        title = title & "WSVC "
    ElseIf locationName = "TAHOMA NATIONAL CEMETERY" Then
        title = title & "Tahoma "
    Else
        title = title & StrConv(pMission.Location("City"), vbProperCase) & " "
    End If
    
    Team = pMission.Record("Team")
    
    title = title & mid(Team, InStr(Team, "(")) & " "
    
    title = title & pMission.Record("Last Name")
End Function
Private Sub SetRegisterData()
    Dim missionRecord As DAO.Recordset
    Dim mid As String, Branch As String
    
    Set missionRecord = pMission.Record
    
    pRegisterData(0) = Format(missionRecord("Funeral Date"), "dd-MMM-yy")
    pRegisterData(1) = pMission.LookUpCounty
    pRegisterData(2) = pMission.Location("City")
    
    If missionRecord("Status") = "Veteran" Then
        mid = "/"
    ElseIf missionRecord("Status") = "Retired" Then
        mid = " "
    End If
    
    Branch = missionRecord("Branch")
    pRegisterData(3) = missionRecord("Status") & mid & pMission.Branches.Item(Branch)
    pRegisterData(4) = FormatSMName(pMission.Record, False)
    If Not IsNull(pMission.Record("Team")) Then
        pRegisterData(5) = Split(pMission.Record("Team"), " ")(0)
    End If
    
    If pMission.Record("Firing Detail") = False And pMission.VSOs.RecordCount = 0 Then
        pRegisterData(6) = "2-man"
    Else
        pRegisterData(6) = "Full"
    End If

    'VSOs
    Dim i As Integer
    Dim rst As DAO.Recordset
    
    Set rst = pMission.VSOs
    i = 7
    While Not rst.EOF
        With rst
            .MoveFirst
            While Not rst.EOF
                pRegisterData(i) = rst(1)
                .MoveNext
                i = i + 1
            Wend
        End With
    Wend
    
    If pMission.Record("Firing Detail") = True And pMission.Record("Status") = "Retired" Then
        pRegisterData(9) = pRegisterData(8)
        pRegisterData(8) = pRegisterData(7)
        pRegisterData(7) = "WAARNG"
    End If
End Sub

Private Sub SetTaskingData()
    pTaskingData(1) = pMission.Record("Mission Number")
    pTaskingData(3) = pMission.Record("Team")
    pTaskingData(4) = "xxx-xx-" & Right(pMission.Record("SSN"), 4)
    pTaskingData(5) = pMission.Record("Last Name")
    pTaskingData(6) = pMission.Record("First Name")
    pTaskingData(7) = pMission.Record("Middle Name")
    pTaskingData(8) = pMission.Record("Suffix")
    
    '
    If InStr(pGrade, "E") Then
        pTaskingData(9) = pGrade                        ' Enlisted Grade
    ElseIf InStr(pGrade, "W") Then
        pTaskingData(10) = pGrade                       ' WO Grade
    ElseIf InStr(pGrade, "O") Then
        pTaskingData(11) = pGrade                       ' Officer Grade
    End If
    
    pTaskingData(12) = pMission.Record("Status")
    pTaskingData(13) = pMission.Record("Branch") ' TURN THIS INTO COMPONENT
    pTaskingData(14) = Format(Date, "dd-mmm-yy")
    
    'funeral home
    pTaskingData(17) = pFuneralHome("Name")
    pTaskingData(18) = pFuneralHome("Street")
    pTaskingData(19) = pFuneralHome("City")
    pTaskingData(20) = pFuneralHome("State")
    pTaskingData(21) = pFuneralHome("ZIP")
    pTaskingData(22) = Format(pFuneralHome("Phone Number"), "(###) ###-####")

    'Date and Time
    pTaskingData(23) = pMission.Record("Funeral Date")
    pTaskingData(24) = Format(pMission.Record("Funeral Time"), "hhmm")
    
    pTaskingData(25) = 0
    'myArr(26) = " "
    pTaskingData(27) = 0
    pTaskingData(28) = 0
    
    Select Case pMission.Record("Remains")
        Case "Casket"
            pTaskingData(25) = 1
            pTaskingData(26) = "Casket"
        Case "Cremains"
            pTaskingData(27) = 1
        Case "No Remains"
            pTaskingData(28) = 1
    End Select
    
    'Funeral Location
    pTaskingData(29) = pMission.Location("Name")
    pTaskingData(30) = pMission.Location("Street")
    pTaskingData(31) = pMission.Location("City")
    pTaskingData(32) = pMission.Location("State")
    pTaskingData(33) = pMission.Location("ZIP")

    'POC
    pTaskingData(34) = pDirector("POC First Name") & " " & pDirector("POC Last Name")
    pTaskingData(35) = Format(pDirector("POC Phone"), "(###) ###-####")

    pTaskingData(36) = 0
    pTaskingData(37) = 0
    pTaskingData(38) = 0
    pTaskingData(39) = 0
    Select Case pMission.Record("Requester")
        Case "Family"
            pTaskingData(36) = 1
        Case "Friend"
            pTaskingData(37) = 1
        Case "Funeral Director"
            pTaskingData(38) = 1
        Case "Other"
            pTaskingData(39) = 1
    End Select
    
    
    'National/State/Private Cemetery
    pTaskingData(41) = 0
    pTaskingData(42) = 0
    pTaskingData(43) = 0
    pTaskingData(44) = 0
    Select Case pMission.Location("Place of Service (for Tasking)")
    Case "Government Cemetery"
        If InStr(pMission.Location("Name"), "National") Then
            pTaskingData(41) = 1
        ElseIf InStr(pMission.Location("Name"), "State") Then
            pTaskingData(42) = 1
        End If
    Case "Private Cemetery", "Chapel/Memorial"
        pTaskingData(43) = 1
    Case Else
        pTaskingData(44) = 1
    End Select
    
    pTaskingData(46) = pMission.Record("Flag Recipient")
    pTaskingData(47) = pMission.Record("Relationship")
    
    'Honors Requested
    pTaskingData(48) = pMission.Record("Flag Fold")
    pTaskingData(49) = pMission.Record("Pallbearer")
    pTaskingData(50) = pMission.Record("Firing Detail")
    pTaskingData(51) = pMission.Record("Officer")
    pTaskingData(52) = pMission.Record("Taps")
    pTaskingData(53) = pMission.Record("Taps By")
    pTaskingData(54) = pMission.Record("Chaplain")
    pTaskingData(55) = pMission.Record("Planeside")
    
    pTaskingData(56) = pMission.Record("Remarks")
    
    pTaskingData(57) = 0
    pTaskingData(58) = 0
    pTaskingData(59) = 0
    
    Select Case pMission.Record("Discharge PPW verified by")
        Case "Funeral Home"
            pTaskingData(57) = 1
        Case "CAC"
            pTaskingData(58) = 1
        Case "WAARNG"
            pTaskingData(59) = 1
    End Select
    
    RemoveNulls pTaskingData
End Sub






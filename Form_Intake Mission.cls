VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Intake Mission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Intake Mission
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is an interface between MFH admin personnel and the mission records.
' It allows them to add/edit records in the "Missions" table of this database in a user-friendly
' manner while preventing them from entering incomplete data or accidentally corrupting mission data.
'
' @form Intake Mission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private in_p_OpenMission As New OpenMission

''
' These collections contain the controls on this form whose fields need to be filled out to perform
' the specified action.
' For example, in order to complete the tasking, the controls contained in
' in_p_TaskingControls must be filled out by the user.
''
Private in_p_TaskingControls As New Collection
Private in_p_CalendarControls As New Collection
Private in_p_RegisterControls As New Collection
Private in_p_G1PortalControls As New Collection

''
' Subforms whose events must be subscribed to by this form.
''
Private WithEvents in_p_SetVSO As [Form_Set Participant]
Attribute in_p_SetVSO.VB_VarHelpID = -1
Private WithEvents in_p_AddFuneralHome As [Form_Add Funeral Home]
Attribute in_p_AddFuneralHome.VB_VarHelpID = -1
Private WithEvents in_p_AddFuneralDirector As [Form_Add Funeral Director]
Attribute in_p_AddFuneralDirector.VB_VarHelpID = -1

' ============================================= '
' Form Events
' ============================================= '

''
' This form can open in one of two ways: either in DataEntry mode (ie. a brand new mission) or
' 'Modify' mode (ie. where you already have the mission saved, you just need to change or add
' something). The mode is specified by the "OpenArgs" string that is passed to the form when it opens.
' If no string is passed to the form, opening the form is cancelled. This narrows the ability
' for the user to access the form, which prevent errors.
''
Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.OpenArgs) Then
        Cancel = True
    ElseIf Me.OpenArgs = "New Mission" Then
        Me.DataEntry = True
    Else
        Me.DataEntry = False
    End If
End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    DoCmd.MoveSize Down:=50
    Last_Name.SetFocus

    If Me.OpenArgs = "Modify" Then
        Me.FilterOn = True
        sqlstr = "SELECT [Funeral Homes].*" & _
                "FROM [Funeral Homes] INNER JOIN [Funeral Directors] ON [Funeral Homes].[Funeral Home ID] = [Funeral Directors].[Funeral Home ID]" & _
                "WHERE ([Funeral Directors].[Funeral Director ID])=" & Funeral_Director & ";"
        Funeral_Home = CurrentDb.OpenRecordset(sqlstr).Fields("Funeral Home ID")

        PopulateAddress Funeral_Home, FHStreet, FHCityStZIP
        FHPhone.Value = Funeral_Director.Column(2)
        Funeral_Location_AfterUpdate

        SetMission
        VSOList.Value = in_p_OpenMission.Mission.ListVSOs
        Mission_Number_AfterUpdate
    End If
    SetAllControlCollections
End Sub

''
' The form simply must have its 'SSN', 'Funeral Director', and 'Funeral Location' fields filled out before the
' record can be saved. This is technically for database integrity reasons, however it also makes sense that
' you shouldn't be able to save your work for later unless you have the contact info of whoever is requesting the
' mission so as to prevent dropped missions. Funeral location and SSN is a duh.
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNull(SSN) Then
        MsgBox "Please enter a SSN.", vbExclamation
        Cancel = True
    End If
    If IsNull(Funeral_Director) Then
        MsgBox "Please select or add a funeral director.", vbExclamation
        Cancel = True
    End If
    If IsNull(Funeral_Location) Then
        MsgBox "Please select a funeral location.", vbExclamation
        Cancel = True
    End If
End Sub

''
' This will close any potential subform and net them to Nothing.
' The ToRegister button must also be re-enabled.
Private Sub Form_Close()
    DoCmd.Close acForm, "Missions VSOs", acSaveYes
    DoCmd.Close acForm, "Add Funeral Home", acSaveYes
    DoCmd.Close acForm, "Add Funeral Director", acSaveYes
    Set in_p_AddFuneralHome = Nothing
    Set in_p_AddFuneralDirector = Nothing
    Set in_p_SetVSO = Nothing
    ToRegister.Enabled = True
End Sub

Private Sub in_p_SetVSO_ParticipantAdded(aType As String)
    If aType = "VSO" Then
        SetMission
        VSOList.Value = in_p_OpenMission.Mission.ListVSOs
    End If
End Sub

' ============================================= '
' Control Events
' ============================================= '

' --------------- '
' Mission Number
' --------------- '
''
' Once updated, it must be disabled to prevent accidental change of the mission number.
' If there is no mission number in the textbox, the 'ToRegister' button must be enabled
' to allow the user to assign a mission number to the mission programatically. Otherwise,
' the 'ToRegister' button is disabled to prevent accidental over-riding of data.
'
' @type Textbox Event
''
Private Sub Mission_Number_AfterUpdate()
    Mission_Number.Enabled = False
    If (Mission_Number <> vbNullString) Then
        ToRegister.Enabled = False
    Else
        ToRegister.Enabled = True
    End If
End Sub

''
' This event allows the user to manually enter a mission number, if need be.
'
' @type Textbox
''
Private Sub Mission_Number_Label_DblClick(Cancel As Integer)
    Dim choice As Integer
    choice = MsgBox("Are you sure you want to manually assign a mission number " & _
                    "rather than using the auto-assign feature of the 'Register' button?", vbYesNo)
    If choice = vbNo Then
        Cancel = True
        Exit Sub
    End If
    Mission_Number.Enabled = True
    Mission_Number.SetFocus
End Sub

' --------------- '
' Branch
' --------------- '
''
' When the user selects a different branch from the dropdown, the 'Rank' Combobox must be repopulated
' with ranks of that branch. This is done by querying the 'Ranks' table.
'
' @type Combobox Event
''
Private Sub Branch_Change()
    Dim sqlstr As String

    sqlstr = _
        "SELECT Ranks.[" & Branch & "] " & _
        "FROM Ranks " & _
        "WHERE ((Not (Ranks.[" & Branch & "])=""""));"

    Rank.RowSource = sqlstr
End Sub

' ---------------- '
' Funeral Home
' ---------------- '
''
'
' @type Combobox Event
''
Private Sub Funeral_Home_GotFocus()
    Funeral_Home.Requery
End Sub

''
' When the user types in a value that isn't already in the dropdown (which is just all the funeral
' home names from 'Funeral Homes' table), they're given the option of adding a new one, since new funeral homes may
' may pop up from time to time. After confirming that they would like to add one, a form allowing them to do so appears.
'
' @type Combobox Event
''
Private Sub Funeral_Home_NotInList(NewData As String, Response As Integer)
    Dim choice As Integer

    choice = MsgBox("That funeral home is not in the records. Would you like to add it?", vbYesNo + vbInformation, "Funeral Home Not Found")
    Response = acDataErrContinue

    If choice = vbYes Then
        DoCmd.Close acForm, "Add Funeral Home", acSaveYes
        DoCmd.OpenForm "Add Funeral Home", OpenArgs:=NewData
        Set in_p_AddFuneralHome = Forms("Add Funeral Home")
    Else
        Funeral_Home.Undo
    End If
End Sub

''
' Once they add a funeral home, the 'Funeral Home' combobox is updated to reflect the addition of the new record.
' For the ease of the user, the funeral home they just added is automatically selected in the combobox.
'
' @type Subform Event
''
Private Sub in_p_AddFuneralHome_Updated(FuneralHomeName As String)
    Funeral_Home = Null
    Funeral_Home.Requery
    Dim rst As DAO.Recordset

    Set rst = CurrentDb.OpenRecordset("SELECT [Funeral Home ID] FROM [Funeral Homes] WHERE [Funeral Homes]!Name='" & FuneralHomeName & "'")
    Funeral_Home = rst(0)
    Funeral_Home_AfterUpdate
End Sub

''
' Each time a new funeral home is selected, the address of that funeral home is automatically loaded into the form from
' the 'Funeral Homes' table. In addition, the list of funeral directors must be updated to only show the funeral directors
' from that funeral home.
'
' @type Combobox Event
''
Private Sub Funeral_Home_AfterUpdate()
    PopulateAddress Funeral_Home, FHStreet, FHCityStZIP
    UpdateFuneralDirectorControl
End Sub

' ---------------- '
' Funeral Director
' ---------------- '
''
' If there are no funeral directors listed for the selected funeral home, the user is FORCED to add one. There must be a funeral
' director for every mission (even if that 'funeral director' is just a family member), for database integrity reasons.
'
' @type Combobox Event
''
Private Sub Funeral_Director_GotFocus()
    If Funeral_Home <> 0 And Funeral_Director.ListCount = 0 Then
        'add a funeral director
        MsgBox "There is no funeral director listed for " & Funeral_Home.Column(1) & ". Please add one now.", vbOKOnly
        DoCmd.Close acForm, "Add Funeral Director", acSaveYes
        DoCmd.OpenForm "Add Funeral Director", OpenArgs:=Funeral_Home
        Set in_p_AddFuneralDirector = Forms("Add Funeral Director")
    End If
End Sub

''
' When a funeral director is added, the control gets updated so that the funeral director is in the dropdown
'
' @type Subform Event
''
Private Sub in_p_AddFuneralDirector_Updated(FirstName As String, LastName As String, Phone As String, FuneralHomeID As Integer)
    UpdateFuneralDirectorControl
End Sub

''
' Once a funeral director is selected, their phone number is shown in the 'Phone' Textbox
'
' @type Combobox Event
''
Private Sub Funeral_Director_AfterUpdate()
    If IsNull(Funeral_Director) Then
        FHPhone = FHPhone.DefaultValue
    Else
        FHPhone.Value = [Funeral Director].Column(2)
    End If
End Sub

' ---------------- '
' Funeral Date
' ---------------- '
''
' DatePicker pops up when this control has the focus, to make it easier to select a date
'
' @type Textbox Event
''
Private Sub Funeral_Date_GotFocus()
    Application.RunCommand (acCmdShowDatePicker)
End Sub

' ---------------- '
' Funeral Time
' ---------------- '
''
' When the user enters a time in this control, this checks that they entered a sensical
' time (eg. 14:30 is OK, 14:68 is NOT OK, neither is 87:30). If they didn't, the textbox
' turns red, signifying that they f*cked up.
'
' @param
' @type Textbox Event
''
Private Sub Funeral_Time_BeforeUpdate(Cancel As Integer)
    Dim myInput As String
    Dim myInputSplit As Variant

    If IsNull(Funeral_Time) Then
        GoTo OK
    End If

    myInput = Numerize(Funeral_Time.Value)

    If Len(myInput) <> 4 Then
        GoTo Invalid
    End If

    If Not Left(myInput, 2) < 24 Then
        GoTo Invalid
    ElseIf Not Right(myInput, 2) < 60 Then
        GoTo Invalid
    End If
OK:
    Funeral_Time.BackColor = RGB(255, 255, 255)
    Exit Sub
Invalid:
    ShowError Funeral_Time
    Cancel = True
End Sub

' ---------------- '
' Funeral Location
' ---------------- '
''
' Each time a new funeral location is selected, the address of that funeral location is automatically loaded into the form from
' the 'Funeral Locations' table.
'
' @type Combobox Event
''
Private Sub Funeral_Location_AfterUpdate()
    PopulateAddress Me.Funeral_Location, FLStreet, FLCityStZIP
End Sub

' ============================================= '
' Button Events
' ============================================= '

Private Sub Add_Funeral_Director_Click()
    DoCmd.OpenForm "Add Funeral Director", acNormal, OpenArgs:=Funeral_Home
    Set in_p_AddFuneralDirector = Forms("Add Funeral Director")
    [Funeral Director].SetFocus
End Sub

Private Sub ToPortal_Click()
    ResetAllControlCollections
    If FormHasNessecaryFieldsFor("putting the mission in the G1 Portal", in_p_G1PortalControls) Then
        SetMission
        'in_p_OpenMission.ToPortal
    End If
End Sub

Private Sub ToRegister_Click()
    ResetAllControlCollections
    If FormHasNessecaryFieldsFor("making an entry in the register", in_p_RegisterControls) Then
        SetMission
        Mission_Number = in_p_OpenMission.ToRegister
        Mission_Number_AfterUpdate
    End If
End Sub

Private Sub Make_Tasking_Click()
    ResetAllControlCollections
    If FormHasNessecaryFieldsFor("making a tasking", in_p_TaskingControls) Then
        SetMission
        in_p_OpenMission.MakeTasking
    End If
End Sub

Private Sub ToSharepoint_Click()
    ResetAllControlCollections
    If FormHasNessecaryFieldsFor("posting the mission on the calendar", in_p_CalendarControls) Then
        SetMission
        'in_p_OpenMission.ToSharepoint
    End If
End Sub

Private Sub Add_VSOs_Click()
On Error GoTo Add_VSOs_Click_Err

    DoCmd.RunCommand acCmdSaveRecord
    SetMission
    DoCmd.OpenForm "Set Participant", acNormal, vbNullString, _
                    "[Deceased SSN]=" & " '" & [SSN] & "'", , acNormal, "VSO" & "|" & SSN
    Set in_p_SetVSO = Forms("Set Participant")

Add_VSOs_Click_Exit:
    Exit Sub

Add_VSOs_Click_Err:
    If Error$ = "No current record." Then
        MsgBox Error$ & " Please enter a SSN before adding VSOs."
    ElseIf InStr(Error$, "Funeral Directors") Then
        MsgBox "Please select a funeral director before adding VSOs."
    ElseIf InStr(Error$, "Funeral Location") Then
        MsgBox "Please select a location before adding VSOs."
    End If
    Resume Add_VSOs_Click_Exit
End Sub




' ============================================= '
' Private Subroutines/Functions
' ============================================= '

''
' Defines all of the nessecary controls that need to be filled out before sending the data. Stored in arrays.
''
Private Sub SetAllControlCollections()
    SetCollection in_p_G1PortalControls, _
                    Array(Me.Last_Name, Me.First_Name, Me.SSN, Me.Funeral_Director, Me.Funeral_Date, Me.Funeral_Location, _
                            Me.Status, Me.Branch)

    SetCollection in_p_RegisterControls, _
                    Array(Me.Last_Name, Me.First_Name, Me.SSN, Me.Funeral_Director, Me.Funeral_Date, Me.Funeral_Location, _
                            Me.Status, Me.Branch, Me.Team)

    SetCollection in_p_TaskingControls, _
                    Array(Me.Mission_Number, Me.Last_Name, Me.First_Name, Me.SSN, Me.Funeral_Date, _
                            Me.Funeral_Time)

    SetCollection in_p_CalendarControls, _
                    Array(Me.Last_Name, Me.Funeral_Date, Me.SSN, Me.Funeral_Director, Me.Funeral_Location, _
                            Me.Funeral_Time, Me.Status, Me.Branch, Me.Team)
End Sub

''
' If for whatever reason the collections above get cleared, this just sets them again.
''
Private Sub ResetAllControlCollections()
    If in_p_RegisterControls.Count = 0 Or in_p_TaskingControls.Count = 0 Or _
       in_p_CalendarControls.Count = 0 Or in_p_G1PortalControls.Count = 0 Then SetAllControlCollections
End Sub

''
' Adds the specified controls to the collection.
'
' @param {Collection} coll The collection of controls
' @param {Variant} ctls Array of controls to be added the the collection
''
Private Sub SetCollection(Coll As Collection, ByVal Ctls As Variant)
    Dim Ctl As Variant
    For Each Ctl In Ctls
        Coll.Add Ctl
    Next Ctl
End Sub

''
' Checks if the form has the nessecary controls filled out to acheive the desired goal.
'
' @param {String} Goal Used for displaying error message if the user doesn't have the nessecary controls filled out
' @param {Collection} Ctls Collection of controls nessecary for executing the task
' @return {Boolean}
''
Private Function FormHasNessecaryFieldsFor(Goal As String, Ctls As Collection) As Boolean
    Dim Ctl As Control
    For Each Ctl In Ctls
        If ControlIsBlank(Ctl) Then
            HandleBlankControl Goal, Ctl
            FormHasNessecaryFieldsFor = False
            Exit Function
        End If
    Next Ctl

    FormHasNessecaryFieldsFor = True
End Function

''
' Checks if the control is blank.
'
' @param {Ctl} Control The Control to be checked
' @return {Boolean}
''
Private Function ControlIsBlank(Ctl As Control) As Boolean
    If (Ctl.Name = "Funeral Director" Or Ctl.Name = "Funeral Location") And Ctl.Value = 0 Then
        GoTo BlankControl
    ElseIf (Ctl.ControlType = acTextBox Or Ctl.ControlType = acComboBox) And IsNull(Ctl.Value) Then
        GoTo BlankControl
    End If

    ControlIsBlank = False

    Exit Function
BlankControl:
    ControlIsBlank = True
End Function

''
' If a control not being filled out is preventing an action, this method displays to the user which control it is so
' that they can fix it.
'
' @param {String} Goal Used for displaying error message if the user doesn't have the nessecary controls filled out
' @param {Control} Ctl The problem control
''
Private Sub HandleBlankControl(Goal As String, Ctl As Control)
    Dim message As String
    message = "Please enter something in the " & Ctl.Name & " field before " & Goal & "."
    If Ctl.Name = "Mission Number" Then
        message = message & " (The mission needs to be assigned to a number in the register first)."
    End If
    MsgBox message

    If Ctl.Enabled = True Then Ctl.SetFocus
End Sub

''
' Simply resets and requerys the funeral director control.
''
Private Sub UpdateFuneralDirectorControl()
    Funeral_Director = Funeral_Director.DefaultValue
    FHPhone = FHPhone.DefaultValue
    Funeral_Director.Requery
End Sub

''
' Takes the value in the Combobox, finds its "Street", "City", "State" and "ZIP" fields and displays
' them in a properly formatted way in the two 'TextBox's
'
' @param {ComboBox} Place The control whose value is being used to update the address
' @param {TextBox} Street Holds the street address
' @param {TextBox} CityStZIP Holds the City, State, ZIP
''
Private Sub PopulateAddress(ByRef Place As ComboBox, ByRef Street As TextBox, ByRef CityStZIP As TextBox)
    Dim str As String
    str = vbNullString

    If Not IsNull(Place) Then
        Street = Place.Column(2)
        If Not (Place.Column(3) = vbNullString Or Place.Column(3) = " " Or IsNull(Place.Column(3))) Then
            str = str & Place.Column(3) & ", "
        End If
        If Not (Place.Column(4) = vbNullString Or Place.Column(4) = " " Or IsNull(Place.Column(4))) Then
            str = str & Place.Column(4) & " " & Place.Column(5)
        End If
        CityStZIP = str
    Else
        Street = Null
        CityStZIP = Null
    End If
End Sub

''
' Runs the query that looks up the current form's record. Opens an 'OpenMission' object with that record info.
''
Private Sub SetMission()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim prm As DAO.Parameter

    DoCmd.RunCommand acCmdSaveRecord

    Set db = CurrentDb()
    Set qdf = db.QueryDefs("SSN Lookup")

    For Each prm In qdf.Parameters
        prm = Eval(prm.Name)
    Next prm

    in_p_OpenMission.NewOpenMission qdf.OpenRecordset(dbOpenDynaset)
End Sub

''
' Turns the specified control red.
'
' @param {Control} Ctl The control that will get turned red
''
Private Sub ShowError(Ctl As Control)
    Ctl.BackColor = RGB(255, 100, 100)
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Master Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Form_Master Form
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is the first thing the user sees when opening this Access file.
' It just has a bunch of buttons that allows them to do all sorts of cool things.
' Do I really need to explain anything else? The button names say exactly what they do,
' so I'm sure you can figure it out. ¯\_(?)_/¯
'
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Button Events
' ============================================= '

Private Sub Add_Funeral_Director_Click()
    DoCmd.Close acForm, "Add Funeral Director", acSaveYes
    DoCmd.OpenForm "Add Funeral Director", acNormal, OpenArgs:=0
End Sub

Private Sub Add_Funeral_Home_Click()
    DoCmd.Close acForm, "Add Funeral Home", acSaveYes
    DoCmd.OpenForm "Add Funeral Home", acNormal, OpenArgs:=" "
End Sub

Private Sub Create_Mission_Click()
On Error GoTo New_Mission_Click_Err

    DoCmd.Close acForm, "Intake Mission", acSaveYes
    DoCmd.OpenForm "Intake Mission", acNormal, vbNullString, vbNullString, , acNormal, OpenArgs:="New Mission"


New_Mission_Click_Exit:
    Exit Sub

New_Mission_Click_Err:
    MsgBox Error$
    Resume New_Mission_Click_Exit

End Sub

Private Sub Modify_Mission_Click()
On Error GoTo Existing_Mission_Click_Err

    DoCmd.Close acForm, "Search For Mission", acSaveYes
    DoCmd.OpenForm "Search For Mission", acNormal, OpenArgs:="Existing Mission"

Existing_Mission_Click_Exit:
    Exit Sub

Existing_Mission_Click_Err:
    MsgBox Error$
    Resume Existing_Mission_Click_Exit

End Sub

Private Sub Close_Mission_Click()
On Error GoTo Close_Click_Err

    DoCmd.Close acForm, "Search For Mission", acSaveYes
    DoCmd.OpenForm "Search For Mission", acNormal, OpenArgs:="Close Mission"

Close_Click_Exit:
    Exit Sub

Close_Click_Err:
    MsgBox Error$
    Resume Close_Click_Exit
End Sub

Private Sub Set_Intake_Defaults_Click()
    DoCmd.OpenForm "Set Defaults", acNormal
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Add Funeral Director"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Add Funeral Director

' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is an interface between MFH admin personnel and the 'Funeral Directors' table.
' It allows them to add/edit records in the 'Funeral Directors' table of this database in a user-friendly
' manner while preventing them from entering incomplete data or accidentally corrupting mission data.
'
' @form Add Funeral Director
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

''
' Communicates with the parent form (Intake Mission)
''
Public Event Updated(FirstName As String, LastName As String, Phone As String, FuneralHomeID As Integer)


' ============================================= '
' Form Events
' ============================================= '

''
' Form must be opened with an argument, which is the funeral home ID under which the director will be added;
''
Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.OpenArgs) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Funeral_Home_ID = Me.OpenArgs
End Sub

''
' Once this form's record gets updated, the parent form needs to be notified to update ITS controls
''
Private Sub Form_AfterUpdate()
    RaiseEvent Updated(POC_First_Name, POC_Last_Name, POC_Phone, Funeral_Home_ID)
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
    End If
    DoCmd.Close acForm, "Add Funeral Director"


Save_Click_Exit:
    Exit Sub

Save_Click_Err:
    MsgBox Error$
    Resume Save_Click_Exit

End Sub



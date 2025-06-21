VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Add Funeral Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Add Funeral Home

' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form is an interface between MFH admin personnel and the 'Funeral Homes' table.
' It allows them to add/edit records in the 'Funeral Homes' table of this database in a user-friendly
' manner while preventing them from entering incomplete data or accidentally corrupting mission data.
'
' @form Add Funeral Home
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

''
' Communicates with the parent form (Intake Mission)
''
Public Event Updated(FuneralHomeName As String)

' ============================================= '
' Form Events
' ============================================= '

Private Sub Form_Load()
    Me.Controls("Name") = Me.OpenArgs
End Sub

Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.OpenArgs) Then
        Cancel = True
    End If
End Sub

' ============================================= '
' Control Events
' ============================================= '

''
' Prevents the user from entering a coty that isnt recognized
''
Private Sub City_BeforeUpdate(Cancel As Integer)
    Dim rstCities As DAO.Recordset

    Set rstCities = CurrentDb.OpenRecordset("County Lookup")

    rstCities.Index = "PrimaryKey"
    rstCities.Seek "=", City.text

    If rstCities.NoMatch Then
        MsgBox "That entry is not a listed city."
        Cancel = True
    End If
End Sub

' ============================================= '
' Button Events
' ============================================= '

Private Sub Add_Click()
On Error GoTo Add_Click_Err

    On Error Resume Next

    Dim newRecord As DAO.Recordset
    Dim Ctl As Control
    Set newRecord = CurrentDb.OpenRecordset("Funeral Homes")

    newRecord.AddNew
    For Each Ctl In Form.Controls
        newRecord.Fields(Ctl.Name) = Ctl
    Next Ctl

    newRecord.Update
    RaiseEvent Updated(Me.Controls("Name"))
    DoCmd.Close acForm, "Add Funeral Home"

    newRecord.Close

Add_Click_Exit:
    Exit Sub

Add_Click_Err:
    MsgBox Error$
    Resume Add_Click_Exit

End Sub


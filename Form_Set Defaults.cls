VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Set Defaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Set Defaults
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' This form allows the user to set the default values of the 'Intake Mission' form.
'
' @form Set Defaults
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

''
' Takes the values from the controls on this form and sets them as the default values of 'Intake Mission' form.
''
Private Sub SetDefaults()
    Dim Ctl As Control

    DoCmd.OpenForm "Intake Mission", acDesign, WindowMode:=acHidden
    For Each Ctl In Me.Controls
        If Ctl.ControlType <> acLabel And Ctl.ControlType <> acCommandButton Then
            Forms("Intake Mission").Controls(Ctl.Name).DefaultValue = """" & Ctl.Value & """"
        End If
    Next Ctl
    DoCmd.Close acForm, "Intake Mission", acSave
End Sub

' ============================================= '
' Button Events
' ============================================= '

Private Sub Set_Default_Click()
On Error GoTo Set_Default_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, vbNullString
    End If

    SetDefaults
    DoCmd.Close acForm, "Set Defaults"

Set_Default_Click_Exit:
    Exit Sub

Set_Default_Click_Err:
    MsgBox Error$
    Resume Set_Default_Click_Exit

End Sub



Attribute VB_Name = "Helper"
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Helper
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Contains general-purpose helpers that are used throughout Expedition. Includes:
'
' - Removing non-numeral characters from Strings
' - Removing nulls from Arrays
' - Error handling
' - Name formatting
'
' @module Helper
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private Const g_errLogPath = "C:\Users\tahak\OneDrive\Documents\Expedition\Expedition\Expedition_Error_Log.txt"

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream
Public Const g_HandleErrors As Boolean = False

' ============================================= '
' Formatting
' ============================================= '

''
' Takes any String and returns that String with all non-numeric values removed
'
' @param {String} szInput any String whose non-numeric values are to be removed
' @return {String} szOutput the String whose non-numeric values are removed
''
Function Numerize(szInput As String) As String
    Dim iInChar As Integer
    Dim szInChar As String, szOutput As String

    ' Clear the output string
    szOutput = vbNullString
    ' For each character in the input string
    For iInChar = 1 To Len(szInput)
        ' Extract the current character
        szInChar = mid(szInput, iInChar, 1)

        ' If the current character is a number
        If (IsNumeric(szInChar)) Then
            ' Append the current character to the output
            szOutput = szOutput & szInChar
        End If
    Next iInChar

    ' Return the output string
    Numerize = szOutput
End Function

Public Function AddHalfHour(StartTimehhmm As String) As String
    Dim EndTime As Date, StartTime As String
    StartTime = Left(StartTimehhmm, 2) & ":" & Right(StartTimehhmm, 2) & ":00"
    EndTime = TimeValue(StartTime) + TimeValue("00:30:00")
    AddHalfHour = Format(EndTime, "hhmm")
End Function

''
' Takes any Array and chagnes all Null values to vbNullString
'
' @param {Variant} Arr an Array that may or may not contain Null values
''
Sub RemoveNulls(ByRef Arr As Variant)
    Dim i As Integer

    For i = 1 To UBound(Arr)
        If IsNull(Arr(i)) Or IsEmpty(Arr(i)) Then
            Arr(i) = vbNullString
        End If
    Next i
End Sub

''
' Takes a String containing a full name, with delimiters and returns the initials of that name
'
' @param {String} Name full name including delimiters between name parts
' @param {String} Delim what separates the parts of the name
''
Function Initials(Name As String, Delim As String) As String
    Dim nameArr As Variant
    Dim i As Integer
    Dim str As String
    nameArr = Split(Name, Delim, , vbTextCompare)

    str = vbNullString
    For i = 0 To UBound(nameArr)
        str = str & UCase(Left(nameArr(i), 1))
    Next i

    Initials = str
End Function

Public Sub SetWarnings(Status As Boolean)
    DoCmd.SetWarnings Status
End Sub

''
' Formats the name of service members in a nice way, either with or without rank included.
'
' @param {DAO.Recordset} SM, the recordset containing someone's full name
' @param {Boolean} IncludeRank, choice of wether to include rank in the String
''
Public Function FormatSMName(SM As DAO.Recordset, IncludeRank As Boolean) As String
    Dim str As String
    str = vbNullString

    If SM("Last Name") <> vbNullString And SM("Last Name") <> " " Then
        str = str & SM("Last Name")
    End If
    If SM("First Name") <> vbNullString And SM("First Name") <> " " Then
        str = str & ", " & SM("First Name")
    End If
    If SM("Middle Name") <> vbNullString And SM("Middle Name") <> " " Then
        str = str & " " & Left(SM("Middle Name"), 1)
    End If
    If SM("Rank") <> vbNullString And SM("Rank") <> " " And IncludeRank Then
        str = str & ", " & SM("Rank")
    End If
    FormatSMName = str
End Function
 
Function SeekRecord(RecordsetDef As String, Index As String, SearchFor As Variant) As DAO.Recordset
    Dim Record As DAO.Recordset

    Set Record = CurrentDb.OpenRecordset(RecordsetDef)

    Record.Index = Index
    Record.Seek "=", SearchFor

    If Record.NoMatch Then
        Set SeekRecord = Null
    Else
        Set SeekRecord = Record
    End If

    Set Record = Nothing
End Function

Public Function LoadParticipantsToArray(Participant As DAO.Recordset) As Variant
    Dim field As Field2
    Dim childRS As DAO.Recordset, partRS As DAO.Recordset
    Dim iParticipant As Integer, iField As Integer
    Dim TwoDimArray(9, 5) As Variant
    
    iParticipant = 0
    
    Do Until Participant.EOF
        iField = 0
        For Each field In Participant.Fields
            With field
                If .Type = 109 Then
                    Set childRS = .Value
                    ' Exit the loop if the multivalued field contains no records.
                    Do Until childRS.EOF
                        childRS.MoveFirst
                        ' Loop through the records in the child recordset.
                        Do Until childRS.EOF
                            TwoDimArray(iParticipant, iField) = childRS!Value.Value 'Debug.Print childRS!Value.Value
                            iField = iField + 1
                            childRS.MoveNext
                        Loop
                    Loop
                ElseIf field.Name <> "ID" Then
                    TwoDimArray(iParticipant, iField) = field 'Debug.Print field
                End If
            End With
            iField = iField + 1
        Next field

        Participant.MoveNext
        iParticipant = iParticipant + 1
    Loop

    Participant.MoveFirst
    
    LoadParticipantsToArray = TwoDimArray
End Function

Public Sub DisplayRecord(Record As DAO.Recordset)
    Dim field As field
    For Each field In Record.Fields
        Debug.Print field.Name, field.Value
    Next
End Sub

' ============================================= '
' Error Handling
' ============================================= '

'Generic Error Handling Subroutine
Public Sub Error_Handle(ByVal sRoutineName As String, _
                        ByVal sErrorNo As String, _
                        ByVal sErrorDescription As String)
    Dim sMessage As String
    sMessage = sErrorNo & " - " & sErrorDescription
    Call MsgBox(sMessage, vbCritical, sRoutineName & " - Error")
    Call LogFile_WriteError(sRoutineName, sMessage)
End Sub

Public Function LogFile_WriteError(ByVal sRoutineName As String, ByVal sMessage As String)
    Dim sText As String

    If g_HandleErrors Then On Error GoTo ErrorHandler
    If (g_objFSO Is Nothing) Then
        Set g_objFSO = New FileSystemObject
    End If
    If (g_scrText Is Nothing) Then
        If (g_objFSO.FileExists(g_errLogPath) = False) Then
            Set g_scrText = g_objFSO.OpenTextFile(g_errLogPath, IOMode.ForWriting, True)
        Else
            Set g_scrText = g_objFSO.OpenTextFile(g_errLogPath, IOMode.ForAppending)
        End If
    End If
    sText = sText & vbNullString & vbCrLf
    sText = sText & Format(Date, "dd MMM yyyy") & "-" & Time() & vbCrLf
    sText = sText & " " & sRoutineName & vbCrLf
    sText = sText & " " & sMessage & vbCrLf
    g_scrText.WriteLine sText
    g_scrText.Close
    Set g_scrText = Nothing
    Exit Function
ErrorHandler:
    Set g_scrText = Nothing
    Call MsgBox("Unable to write to log file", vbCritical, "LogFile_WriteError")
End Function

Public Sub ExportAllCode()

    Dim c As VBComponent
    Dim Sfx As String
 
    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            c.Export _
                FileName:=CurrentProject.Path & "\" & _
                c.Name & Sfx
        End If
    Next c

End Sub






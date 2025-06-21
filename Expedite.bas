Attribute VB_Name = "Expedite"
Option Explicit

Sub TestSometing()
    Dim part As DAO.Recordset
    Dim sql As String
    Dim SSN As String
    SSN = "111-11-1111"
    sql = "SELECT [Missions_Honor Guard Members].[Deceased SSN], [Honor Guard Members].[Last Name], [Honor Guard Members].[First Name], [Missions_Honor Guard Members].Positions, [Honor Guard Members].ID " & _
            "FROM [Honor Guard Members] INNER JOIN [Missions_Honor Guard Members] ON [Honor Guard Members].ID = [Missions_Honor Guard Members].[HG Member ID] " & _
            "WHERE ((([Missions_Honor Guard Members].[Deceased SSN])='" & SSN & "'))"
          
    
    
    Set part = CurrentDb.OpenRecordset(sql)
    LoadParticipantsToArray part
End Sub

'Dim mishFileName As String
'Public g_objFSO As Scripting.FileSystemObject
'Public g_scrText As Scripting.TextStream
'
'Public Type BSResult
'    Row As Integer
'    found As Boolean
'End Type
'
'
'Function cleanString(text As String) As String
'    Const sProcName As String = "cleanString"
'    If gcfHandleErrors Then On Error GoTo ErrorHandler
'
'    Dim output As String
'    Dim c 'since char type does not exist in vba, we have to use variant type.
'    For i = 1 To Len(text)
'        c = mid(text, i, 1) 'Select the character at the i position
'        If (c >= "0" And c <= "9") Then 'Or (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
'            output = output & c 'add the character to your output.
'        End If
'    Next
'    cleanString = output
'
'       Exit Function
'ErrorHandler:
'   Call Error_Handle(sProcName, Err.Number, Err.Description)
'End Function
'
'
'Private Function Get_IE_Window(URL As String) As SHDocVw.InternetExplorer
'    Const sProcName As String = "Get_IE_Window"
'   If gcfHandleErrors Then On Error GoTo ErrorHandler
'    'Look for an IE browser window or tab already open at the domain of the specified URL (which can start with http://, https://
'    'or nothing) and, if found, return that browser as an InternetExplorer object.  Otherwise return Nothing
'
'    Dim Domain As String
'    Dim Shell As Object
'    Dim ie As SHDocVw.InternetExplorer
'    Dim i As Variant 'Must be a Variant to index Shell.Windows.Item() array
'    Dim p1 As Integer, p2 As Integer
'
'    p1 = InStr(URL, "://")
'    If p1 = 0 Then
'        p1 = 1
'    Else
'        p1 = p1 + 3
'    End If
'    p2 = InStr(p1, URL, "/")
'    If p2 = 0 Then p2 = Len(URL) + 1
'    Domain = mid(URL, p1, p2 - p1)
'
'    Set Shell = CreateObject("Shell.Application")
'
'    i = 0
'    Set Get_IE_Window = Nothing
'    While i < Shell.Windows.Count And Get_IE_Window Is Nothing
'        Set ie = Shell.Windows.Item(i)
'        If Not ie Is Nothing Then
'            If TypeOf ie Is SHDocVw.InternetExplorer And InStr(ie.LocationURL, "file://") <> 1 Then
'                If InStr(ie.LocationURL, Domain) > 0 Then
'                    Set Get_IE_Window = ie
'                End If
'            End If
'        End If
'        i = i + 1
'    Wend
'       Exit Function
'ErrorHandler:
'    Call Error_Handle(sProcName, Err.Number, Err.Description)
'End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manip_GovSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Manip_IE_G1Portal
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Resposible for manipulating Internet Explorer to either enter a new mission or close out a mission in the
' G1 Portal.
'
' @class Manip_Register
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '



' ============================================= '
' Public Methods
' ============================================= '

Public Sub MissionToG1Portal(SSN As String, Data As Variant, Remarks As String, FuneralDate As String, FuneralTime As String, FuneralLocationName As String)
    Const sProcName As String = "MissionToG1Portal"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim ie As SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim htmlDivs As MSHTML.IHTMLElementCollection
    Dim htmlSpans As MSHTML.IHTMLElementCollection
    Dim htmlClick As MSHTML.IHTMLElement
    Dim htmlClicks As MSHTML.IHTMLElementCollection
    Dim htmlAs As MSHTML.IHTMLElementCollection
    Dim htmlInput As Object
    Dim htmlInputs As MSHTML.IHTMLElementCollection
    Dim htmlTexts As MSHTML.IHTMLElementCollection
    Dim htmlSelects As MSHTML.IHTMLElementCollection

    Dim myURL As String

    myURL = ""

    NavigateToGovSite myURL, ie, HTMLDoc

''''''''''''''''''''Using Known SSN Page'''''''''''''''''''''''''''''''''''''''''

    Set htmlClick = HTMLDoc.getElementById("tbSSN")
    Do While htmlClick Is Nothing
        Set htmlClick = HTMLDoc.getElementById("tbSSN")
        Debug.Print 1
    Loop
    htmlClick.Value = SSN   '''

    Set htmlClick = HTMLDoc.getElementById("tbSSNcompare")
    htmlClick.Value = SSN   ''''
    HTMLDoc.getElementById("lbSSNInfo").Click

    WaitForPageToLoad ie
''''''''''''''''''''''''''''Collect Basic Mission Data Page'''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer

    'Go through each text box ('tb') and fill in the appropriate info
    Set htmlInputs = HTMLDoc.getElementsByTagName("input")
    i = 1
    For Each htmlInput In htmlInputs
        If htmlInput.getAttribute("type") = "text" Then
            If htmlInput.ID <> "tbSSN" And htmlInput.ID <> "tbSSNcompare" And htmlInput.ID <> "tbNewSSN" Then
                htmlInput.Value = Data(i, 1)    '''
                i = i + 1
            End If
        End If
    Next htmlInput

    'Now, we must cycle through each of the dropdown bars
    Set htmlInputs = HTMLDoc.getElementsByTagName("select")
    i = 1
    For Each htmlInput In htmlInputs
        htmlInput.Value = Data(i, 2)     '''
        i = i + 1
    Next htmlInput

    'Now, we must cycle through each of the check boxes
    Set htmlInputs = HTMLDoc.getElementsByTagName("input")
    i = 1
    For Each htmlInput In htmlInputs
        If htmlInput.getAttribute("type") = "checkbox" Then
            If htmlInput.ID <> "chkInternational" And htmlInput.ID <> "cblHonors_7" Then
                If Data(i, 3) = True Then
                    htmlInput.Checked = True
                End If
                i = i + 1
            End If
        End If
    Next htmlInput

    'put an error catcher here in case SSN is already in TAPS
    Set htmlInput = HTMLDoc.getElementById("tbRequests")
    If htmlInput Is Nothing Then
        MsgBox ("Hmmmm....It seems that SSN has already been entered into TAPS. Check that you didn't make a typo and try again.")
        Exit Sub
    Else
        htmlInput.Value = Remarks 'Remarks
    End If

'''''''''''''''''''''wait for user to click Save'''''''''''''''''''''''''''''''''''''''''
    Do While ie.ReadyState = READYSTATE_COMPLETE
        If ie.Busy Or HTMLDoc.URL = "" Then
            Exit Do
        End If
    Loop

    WaitForPageToLoad ie
    'Application.Wait (Now() + TimeValue("00:00:01"))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'click on  mission location page
    Dim htmlA As Object
    Set htmlAs = HTMLDoc.getElementsByTagName("a")
    For Each htmlA In htmlAs
        If htmlA.ID = "MissionTaskListCtrl_m_rSectionRepeater_ctl00_m_rTaskRepeater_ctl01_hl" Then
            htmlA.Click
        End If
    Next htmlA

    WaitForPageToLoad ie

''''''''''''''''''''''MissionLocations Page'''''''''''''''''''''''''''''''

    Dim htmlText As Object
    Set htmlTexts = HTMLDoc.getElementsByTagName("input")
    i = 0
    For Each htmlText In htmlTexts
        If htmlText.getAttribute("type") = "text" Then
            If InStr(htmlText.ID, "Date") And InStr(htmlText.ID, "_ev") Then
                htmlText.Value = FuneralDate    ''''
            ElseIf InStr(htmlText.ID, "StartTime") And InStr(htmlText.ID, "_ev") Then
                htmlText.Value = FuneralTime    ''''
            End If
        End If
    Next htmlText

    Dim htmlselect As Object
    Set htmlSelects = HTMLDoc.getElementsByTagName("select")
    For Each htmlselect In htmlSelects
        If htmlselect.ID = "_evBurialSched_ddlCemDuration" Or htmlselect.ID = "_evFuneralSched_ddlFHDuration" Then
            htmlselect.Value = 1    '''
        End If
    Next htmlselect

    Set htmlInputs = HTMLDoc.getElementsByTagName("input")
    For Each htmlInput In htmlInputs
        If htmlInput.ID = "_ddlFuneralHomeInput" Or htmlInput.ID = "_ddlCemeteryInput" Then
            htmlInput.Value = FuneralLocationName   ''''
            Exit For
        End If
    Next htmlInput

   Exit Sub
ErrorHandler:
    'Call Error_Handle(sProcName, Err.Number, Err.Description)
End Sub

Sub MissionToCalendar(aData As Variant)
    Const sProcName As String = "MissionToCalendar"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim ie As SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim htmlDivs As MSHTML.IHTMLElementCollection
    Dim htmlSpans As MSHTML.IHTMLElementCollection
    Dim htmlClicks As MSHTML.IHTMLElementCollection
    Dim htmlAs As MSHTML.IHTMLElementCollection
    Dim htmlInputs As MSHTML.IHTMLElementCollection
    Dim htmlSelects As MSHTML.IHTMLElementCollection
    

    Dim calendarURL As String
    calendarURL = ""

    NavigateToGovSite calendarURL, ie, HTMLDoc

    Dim i As Integer
    i = 1
    Dim htmlInput As Object
    Set htmlInputs = HTMLDoc.getElementsByTagName("input")
    For Each htmlInput In htmlInputs
        If htmlInput.title = "Title Required Field" Or htmlInput.title = "Location" Or htmlInput.title = "Start Time Required Field" Or htmlInput.title = "End Time Required Field" Then
            htmlInput.Value = aData(i)
            i = i + 1
            If i = 5 Then
                Exit For
            End If
        End If
    Next htmlInput

    Dim htmlselect As Object
    Set htmlSelects = HTMLDoc.getElementsByTagName("select")
    For Each htmlselect In htmlSelects
        htmlselect.Value = aData(i)
        i = i + 1
    Next htmlselect
   Exit Sub
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Sub

Sub CloseMissionG1Portal(SSN As String, ParticipantArr As Variant, Driver As DAO.Recordset, Mileage As Integer)
    Const sProcName As String = "CloseMissionInTAPS"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    Dim ie As SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim htmlDivs As MSHTML.IHTMLElementCollection
    Dim htmlSpans As MSHTML.IHTMLElementCollection
    Dim htmlClicks As MSHTML.IHTMLElementCollection
    Dim htmlClick As MSHTML.IHTMLElement
    Dim htmlAs As MSHTML.IHTMLElementCollection
    Dim htmlInputs As MSHTML.IHTMLElementCollection
    Dim htmlTexts As MSHTML.IHTMLElementCollection
    Dim htmlSelects As MSHTML.IHTMLElementCollection
    Dim htmlselect As MSHTML.IHTMLElement
    Dim htmlOption As MSHTML.IHTMLElement
    Dim table As MSHTML.IHTMLTable

    Dim destURL As String

    destURL = "https://arngg1.ngb.army.mil/TAPS/MissionSearch.aspx?key=0&level=State"
    NavigateToGovSite destURL, ie, HTMLDoc

    Set htmlClick = HTMLDoc.getElementById("SearchCtrl1_m_txtSSN")
    htmlClick.Value = Numerize(SSN) 'WorksheetFunction.Substitute(Intake.[C10], "-", "")

    HTMLDoc.getElementById("SearchCtrl1_m_btnSrchSSN").Click

    WaitForPageToLoadExtra ie

    HTMLDoc.getElementById("MissionTaskListCtrl_m_rSectionRepeater_ctl00_m_rTaskRepeater_ctl02_hl").Click

    WaitForPageToLoadExtra ie

    HTMLDoc.getElementById("btnAdd").Click

    WaitForPageToLoadExtra ie
'''''''''''''''''''''''''''''''selecting participants and assigning roles'''''''''''''''''''''''''''''''''''
    Dim iPtcpt As Long, iFld As Long
    Dim fName As String, LName As String
    Dim foundSM, nextSM As Boolean

    For iPtcpt = 0 To 9
        On Error GoTo nSM
        fName = ParticipantArr(iPtcpt, 2) 'AAR.Cells(i, 4).Value
        LName = ParticipantArr(iPtcpt, 1) 'AAR.Cells(i, 3).Value
        foundSM = False

        If LName <> " " And LName <> "" Then
            Set htmlselect = HTMLDoc.getElementById("ddlGeneral")

            'loops through the ddl and finds the SM who went on the service
            For Each htmlOption In htmlselect.Children
                If InStr(1, htmlOption.innerText, LName, vbTextCompare) Then
                    If InStr(1, htmlOption.innerText, fName, vbTextCompare) Then
                        htmlselect.Value = htmlOption.Value
                        htmlselect.OnChange
                        foundSM = True
                        Exit For
                    End If
                End If
            Next htmlOption

            If Not foundSM Then
                MsgBox "It seems TAPS couldn't find " & fName & " " & LName & " in the TAPS database."
                Exit For
            End If

            WaitForPageToLoadWithRedirect ie

            'loops through the sm's positions and assigns them, one by one
            For iFld = 3 To 5     ' iFld represents role index

                If IsEmpty(ParticipantArr(iPtcpt, iFld)) Then
                    Exit For
                End If

                Set htmlselect = HTMLDoc.getElementById("ddlPartcipantRole")

                htmlselect.Value = ParticipantArr(iPtcpt, iFld) 'AAR.Cells(i, j).Value
                htmlselect.OnChange

                WaitForPageToLoadWithRedirect ie
            Next iFld

            'wait for user to assign Component and hit the 'Save Record' button, making sure they end up on the right page
            If iPtcpt < 7 Then
                HTMLDoc.getElementById("ddlPartcipantSubFunction").Value = ""
            End If
nSM:
            'this loop makes sure they don't land back on the same page due to an error
            nextSM = False
            Do While nextSM = False
                nextSM = True
                Do While ie.ReadyState = READYSTATE_COMPLETE
                    If ie.Busy Then
                        Exit Do
                    End If
                Loop
                WaitForPageToLoadExtra ie
                For Each table In HTMLDoc.getElementsByTagName("table")
                    If table.ID = "tblCompleteRecord" Then
                        nextSM = False
                    End If
                Next table
            Loop

            Set htmlselect = Nothing
        End If
    Next iPtcpt

    HTMLDoc.getElementById("M_profile1_td2").Click

    WaitForPageToLoadExtra ie

'''''''''''''''''''''''''''   AAR PAGE   ''''''''''''''''''''''''''''''''''''
    Dim Row As MSHTML.IHTMLTableRow
    Dim Data As MSHTML.IHTMLTableCell
    Dim DDL As MSHTML.IHTMLElement

    Set table = HTMLDoc.getElementById("tblParticipant")

    For Each Row In table.Rows
        If Row.RowIndex = 0 Then: GoTo nextRow
        Dim Rank, last, first As String
        Dim nextRow As Boolean

        For Each Data In Row.Cells
            Select Case Data.nodeIndex
            Case 2
                Rank = Data.innerText
            Case 3
                LName = Data.innerText
            Case 4
                fName = Data.innerText
            Case 5
                For Each DDL In Data.Children 'there's only one dropdown list
                    If Rank = "" Then
                        DDL.Value = 11 'LOCAL NO EXPENSE
                        GoTo nextRow
                    ElseIf LName = UCase(Driver("Last Name")) And fName = UCase(Driver("First Name")) Then
                        DDL.Value = 2 'GSA
                        DDL.OnChange
                    Else
                        DDL.Value = 3 'PASSENGER
                        GoTo nextRow
                    End If
                Next DDL
            Case 6
                Data.FirstChild.Value = Mileage
            Case 7
                Data.FirstChild.Value = "$ 0.00"
            Case 8
                Data.FirstChild.Value = "$ 0.00"
                GoTo nextRow
            End Select
        Next Data
        Rank = ""
nextRow:
    Next Row

    'HTMLDoc.getElementById("btnComplete").Click
'at this point the user is free to 'Mark AAR as complete' and SAVE
   Exit Sub
ErrorHandler:
    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub NavigateToGovSite(ByVal myURL As String, ByRef myIE As SHDocVw.InternetExplorerMedium, ByRef myHTMLDoc As MSHTML.HTMLDocument)
    Const sProcName As String = "NavigateToGovSite"
    If g_HandleErrors Then On Error GoTo ErrorHandler
    
    If myIE Is Nothing Then
        Set myIE = New SHDocVw.InternetExplorerMedium
    End If

    myIE.Visible = True
    myIE.Navigate myURL

    WaitForPageToLoad myIE
    Set myHTMLDoc = myIE.Document

'if redirect to SSO logon then...
    If InStr(myHTMLDoc.URL, "federation.eams") Then
        WaitForPageToLoadExtra myIE
        Set myHTMLDoc = myIE.Document
        If InStr(myHTMLDoc.URL, "federation.eams") Then
            myHTMLDoc.getElementById("pki-login").Click
        End If
        WaitForPageToLoad myIE
    End If
    
    Dim htmlDiv As MSHTML.IHTMLDivElement
    Dim htmlDivs As MSHTML.IHTMLElementCollection
    Dim htmlButton As MSHTML.IHTMLButtonElement
    Dim htmlButtons As MSHTML.IHTMLElementCollection
    
    Set htmlDivs = myHTMLDoc.getElementsByClassName("nav_bar")
    For Each htmlDiv In htmlDivs
'if redirect to .DOD Monitoring Consent, then click 'Accept'
        If htmlDiv.innerText = "Department of Defense Warning" Then
            Set htmlButtons = myHTMLDoc.getElementsByTagName("button")
            For Each htmlButton In htmlButtons
                If htmlButton.innerText = "Accept" Then
                    htmlButton.Click
                    Exit For
                End If
            Next htmlButton
        End If
    Next htmlDiv

    WaitForPageToLoad myIE
       Exit Sub
ErrorHandler:
    If Err.Number = 70 Then
        MsgBox "Please open Internet Explorer to continue."
    End If

    Call Error_Handle(sProcName, Err.Number, Err.Description)
End Sub

Private Sub WaitForPageToLoadExtra(ByRef iedoc As SHDocVw.InternetExplorerMedium)
    Do While iedoc.Busy Or iedoc.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
End Sub

Private Sub WaitForPageToLoad(ByRef iedoc As SHDocVw.InternetExplorerMedium)
    Do While iedoc.Busy
        DoEvents
    Loop
End Sub

Private Sub AccountForRedirect(ByRef iedoc As SHDocVw.InternetExplorerMedium)
    Dim i As Integer
    i = 0
    Do While iedoc.Busy = False And i < 2000
        DoEvents
        i = i + 1
    Loop
End Sub

Private Sub WaitForPageToLoadWithRedirect(ByRef iedoc As SHDocVw.InternetExplorerMedium)
    WaitForPageToLoad iedoc
    AccountForRedirect iedoc
    WaitForPageToLoad iedoc
End Sub

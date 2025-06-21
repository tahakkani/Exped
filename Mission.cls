VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expedition")
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Mission
'
' Taha Hakkani - https://github.com/tahakkani/Expedition
'
' Resposible for holding general data about missions and tasks that are used by both 'CloseMission' and
' 'OpenMission' objects
'
' @class Mission
' @author tahakkani@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Explicit

' ============================================= '
' Constants and Private Variables
' ============================================= '

Private Const p_NETWORK_DRIVE = "CC:\Users\tahak\OneDrive\Documents\Expedition\Expedition"
Private Const p_STAFF_FOL_PATH = p_NETWORK_DRIVE & "\Staff_Folder"

Private pRecord As DAO.Recordset
Private pFuneralLocation As DAO.Recordset

Private pProjPath As String

Private pVSOs As DAO.Recordset
Private pBranches As New Dictionary

' ============================================= '
' Properties
' ============================================= '

Property Set Record(aRecord As DAO.Recordset)
    Set pRecord = aRecord
End Property

Property Get Record() As DAO.Recordset
    Set Record = pRecord
End Property

Property Get NetworkDrivePath()
    NetworkDrivePath = p_NETWORK_DRIVE
End Property

Property Get StaffFolderPath()
    StaffFolderPath = p_STAFF_FOL_PATH
End Property

Public Property Get ProjectPath()
    ProjectPath = pProjPath
End Property

Public Property Get VSOs() As DAO.Recordset
    Set VSOs = pVSOs
End Property

Public Property Get Branches() As Dictionary
    Set Branches = pBranches
End Property
Public Property Get Location() As DAO.Recordset
    Set Location = pFuneralLocation
End Property

Private Property Set ProjectPath(aPath As Variant)
    pProjPath = aPath
End Property

' ============================================= '
' Public Methods
' ============================================= '

Public Sub NewMission(aRecord As DAO.Recordset)
    Set Record = aRecord
    SetLocation
    SetVSOs
    SetBranches
End Sub

Public Sub SetVSOs()
    Dim sql As String
    sql = "SELECT *" & _
            "FROM [Missions_VSOs]" & _
            "WHERE [Missions_VSOs]![Deceased SSN]='" & pRecord("SSN") & "'"
            
    Set pVSOs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)
    
    'ListParticipants pVSOs
End Sub

Public Function ListVSOs() As String
    Dim field As Field2
    Dim str As String
    
    str = vbNullString
    If Not pVSOs.EOF Then
        Do Until pVSOs.EOF
            For Each field In pVSOs.Fields
                With field
                    If .Name = "VSO" Then
                        If str <> vbNullString Then
                            str = str & ", "
                        End If
                        str = str & .Value
                    End If
                End With
            Next field
            
            pVSOs.MoveNext
        Loop
        
        pVSOs.MoveFirst
    End If
    
    ListVSOs = str
End Function

Public Function LookUpCounty() As String
    Dim rstCity As DAO.Recordset
    
    Set rstCity = CurrentDb.OpenRecordset("County Lookup")
    
    rstCity.Index = "PrimaryKey"
    rstCity.Seek "=", pFuneralLocation("City").Value
    
    If rstCity.NoMatch Then
        LookUpCounty = vbNullString
    Else
        LookUpCounty = rstCity("County")
    End If
 
    Set rstCity = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Class_Initialize()
    pProjPath = Environ("UserProfile") & "\OneDrive\Desktop\Expedition"
End Sub

Private Sub SetBranches()
    Dim i As Integer
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("select * from  " & "Ranks" & " where 1=0")

    pBranches.RemoveAll
    
    For i = 0 To rst.Fields.Count - 1
        With rst.Fields(i)
            pBranches.Add .Name, "US" & Initials(.Name, " ")
        End With
    Next i
End Sub

Private Sub SetLocation()
    Dim rstFuneralLocs As DAO.Recordset
    
    If g_HandleErrors Then On Error GoTo ErrorHandler
    
    Set rstFuneralLocs = CurrentDb.OpenRecordset("Funeral Locations")
    
    rstFuneralLocs.Index = "Key"
    rstFuneralLocs.Seek "=", pRecord("Funeral Location ID")
    
    If rstFuneralLocs.NoMatch Then
        Set pFuneralLocation = Null
    Else
        Set pFuneralLocation = rstFuneralLocs
    End If
 
    'rstFuneralHomes.Close
    Set rstFuneralLocs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

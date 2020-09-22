VERSION 5.00
Begin VB.Form frmSQLServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " SQL Servers"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSQLServers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboServers 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   45
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "frmSQLServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------
' Module     : frmSQLServers
'
' Description: This is my modified version of Get_SQL_Server by Mike G (mikeg@ivbnet.com)
'              that I found on the excellent VB site PlanetSourceCode:
'              http://www.planetsourcecode.com/xq/ASP/txtCodeId.9912/lngWId.1/qx/vb/scripts/ShowCode.htm
'
'              There seemed to be a problem with an API call (lstrlenw), so
'                I changed it to use a basic Len() command to get the string length;
'                Also, the way the code was written originally, it only
'                displayed the first 4 characters of the SQL Servers on my LAN,
'                so I changed the nLen line from * 2 to * 4 and that made it all better!
'              See if it works on your company's LAN if it has multiple SQL Servers
'                you can see from your workstation
'
' Procedures : GetSQLServers()
'              Form_Load()
'              Pointer2stringw(ByVal L As Long)
'
' Modified   : Apr 16, 2001 by B Battles  WS1O
' --------------------------------------------------
Option Explicit

' API declarations

' kernel32 declaration
Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (Destination As Any, _
                         Source As Any, _
                         ByVal Length As Long)

' netapi declarations
Private Declare Function NetServerEnum Lib "netapi32" ( _
        strServername As Any, _
        ByVal Level As Long, _
        bufPtr As Long, _
        ByVal PrefMaxLen As Long, _
        EntriesRead As Long, _
        TotalEntries As Long, _
        ByVal ServerType As Long, _
        strDomain As Any, _
        ResumeHandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" _
        (ByVal lpBuffer As Long) As Long

Private Const SV_TYPE_SERVER    As Long = &H2
Private Const SV_TYPE_SQLSERVER As Long = &H4

Private Type SV_100
    Platform As Long
    Name     As Long
End Type
Public Sub GetSQLServers()
    
    ' --------------------------------------------------
    ' Comments  : hunts for SQL Server boxes on your domain's LAN
    ' Modified  : Apr 16, 2001 by B Battles  WS1O
    ' --------------------------------------------------
    
    On Error GoTo Err_GetSQLServers
    
    Dim L            As Long
    Dim EntriesRead  As Long
    Dim TotalEntries As Long
    Dim hResume      As Long
    Dim bufPtr       As Long
    Dim Level        As Long
    Dim PrefMaxLen   As Long
    Dim lType        As Long
    Dim Domain()     As Byte
    Dim I            As Long
    Dim sv100        As SV_100
    Dim strDomain    As String
    
    Level = 100
    PrefMaxLen = -1
    lType = SV_TYPE_SQLSERVER
    
    ' MAKE SURE YOU CHANGE THIS!!!
    ' use your own domain name
    'Domain = "MYNETWORKDOMAIN" & vbNullChar
    
        If (IsNull(Domain)) Or (Len(Format(Domain)) < 1) Then
TryAgain:
            strDomain = InputBox("Please enter your network's Domain Name", "     DOMAIN NAME NEEDED", "MYCOMPANYDOMAIN")
            DoEvents
            Screen.MousePointer = vbHourglass
            Domain = Trim$(strDomain) & vbNullChar
            If Len(Format(Domain)) < 1 Then
                ' no value entered, or user cancelled
                MsgBox "No Domain Name value entered," & vbCrLf & "            or user cancelled", vbInformation, "     Exiting Program"
                Unload Me
                End
            Else
                ' use value entered in inputbox
                Domain = strDomain & vbNullChar
            End If
        End If
    L = NetServerEnum(ByVal 0&, _
            Level, _
            bufPtr, _
            PrefMaxLen, _
            EntriesRead, _
            TotalEntries, _
            lType, _
            Domain(0), _
            hResume)
        If L = 0 Or L = 234& Then
            For I = 0 To EntriesRead - 1
                CopyMemory sv100, ByVal bufPtr, Len(sv100)
                cboServers.AddItem Pointer2StringW(sv100.Name)
                bufPtr = bufPtr + Len(sv100)
            Next I
        End If
    NetApiBufferFree bufPtr
    
Exit_GetSQLServers:
    
    On Error GoTo 0
    Exit Sub
    
Err_GetSQLServers:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & Err.Source, vbInformation, "GetSQLServers - Advisory"
            Resume Exit_GetSQLServers
    End Select
    
End Sub
Private Sub Form_Load()
    
    ' --------------------------------------------------
    ' Comments  : runs the routines to get the SQL Servers
    '                and fill the combobox
    ' Modified  : Apr 16, 2001 by B Battles  WS1O
    ' --------------------------------------------------
    
    On Error GoTo Err_Form_Load
    
    Screen.MousePointer = vbHourglass
    GetSQLServers
    
Exit_Form_Load:
    
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_Form_Load:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & Err.Source, vbInformation, "Form_Load - Advisory"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Function Pointer2StringW(ByVal L As Long) As String
    
    ' --------------------------------------------------
    ' Comments  : converts pointers returned by API call
    '              to string containing the SQL Servers' names
    ' Parameters: L
    ' Returns   : String
    ' Modified  : Apr 16, 2001 by B Battles  WS1O
    ' --------------------------------------------------
    
    On Error GoTo Err_Pointer2StringW
    
    Dim Buffer() As Byte
    Dim nLen     As Long
    
    nLen = (Len(L)) * 4
    If nLen Then
        ReDim Buffer(0 To (nLen - 1)) As Byte
        CopyMemory Buffer(0), ByVal L, nLen
        Pointer2StringW = Buffer
    End If
    
Exit_Pointer2StringW:
    
    On Error GoTo 0
    Exit Function
    
Err_Pointer2StringW:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & Err.Source, vbInformation, "Pointer2StringW - Advisory"
            Resume Exit_Pointer2StringW
    End Select
    
End Function

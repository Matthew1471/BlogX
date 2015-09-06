VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form CheckForUpdate 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matthew1471 Update"
   ClientHeight    =   1830
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6480
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Update"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4905
      Top             =   135
   End
   Begin VB.Timer Timer2 
      Left            =   5310
      Top             =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Check For Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4875
      Picture         =   "Update.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   570
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1455
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5745
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Begin to Check For An Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   4695
   End
End
Attribute VB_Name = "CheckForUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub Command1_Click()

    Dim TransferSuccess As Boolean
    
    UpdateTime = 0
    Timer2.Interval = 1000
    Command1.Enabled = False
    ProgressBar1.Value = 1
    status$ = "Checking for updated version."
    TransferSuccess = GetInternetFile(Inet1, "http://BlogX.co.uk/Download/Update.asp", App.Path)

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        Exit Sub
    End If
       
    ProgressBar1.Value = 2
    
    status$ = "Version Check Complete"

    On Error Resume Next
    Open App.Path & "\Update.asp" For Input As #1
         Input #1, updatever$
    Close #1
    
    If Err.Number <> 0 Then MsgBox ("Version Information Could Not Be Read")
    
    On Error GoTo 0
      
    If updatever$ > myVer Then
        Label1.Caption = "New Version " & updatever & " Now Available"
    Else
        Label1.Caption = "No New Version"
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If

    status$ = "Getting updated file."
    TransferSuccess = GetInternetFile(Inet1, "http://BlogX.co.uk/Download/WinBlogX Setup.exe", App.Path)

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If
    
    ProgressBar1.Value = 3
    Timer2.Interval = 0
    Command1.Enabled = True
        
    X = ShellExecute(Me.hwnd, "open", App.Path & "\WinBlogX Setup.exe", vbNullString, "", 0)

    Dim oFrm As Form

    For Each oFrm In Forms
        Unload oFrm
    Next

End Sub

Private Sub Form_Load()

status$ = "Idle"
UpdateTime = 0

myVer = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Timer1_Timer()
If Inet1.StillExecuting = False Then
    StatusBar1.Panels(1).Text = "Status: Idle"
Else
    StatusBar1.Panels(1).Text = "Status: " & status$
End If

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub

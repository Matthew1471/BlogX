VERSION 5.00
Begin VB.Form Logon 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login To Your WebBlog"
   ClientHeight    =   3240
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4620
   Icon            =   "Logon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   705
      Picture         =   "Logon.frx":0442
      ScaleHeight     =   345
      ScaleWidth      =   315
      TabIndex        =   14
      ToolTipText     =   "Click For Help"
      Top             =   1710
      Width           =   315
   End
   Begin VB.TextBox txtFolder 
      Height          =   300
      Left            =   1425
      TabIndex        =   2
      Text            =   "BlogX"
      Top             =   360
      Width           =   2835
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2520
      TabIndex        =   12
      Top             =   2655
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   915
      TabIndex        =   11
      Top             =   2655
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   2415
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   4230
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   60
         Picture         =   "Logon.frx":0632
         ScaleHeight     =   345
         ScaleWidth      =   315
         TabIndex        =   13
         ToolTipText     =   "Click For Help"
         Top             =   225
         Width           =   315
      End
      Begin VB.CheckBox CheckBox 
         BeginProperty DataFormat 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   7
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   2130
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1515
         TabIndex        =   4
         Text            =   "admin"
         Top             =   600
         Width           =   2625
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "ilovehannah"
         Top             =   930
         Width           =   2625
      End
      Begin VB.TextBox txtServer 
         Height          =   330
         Left            =   1500
         TabIndex        =   8
         Text            =   "Yoursite.com"
         Top             =   1605
         Width           =   2640
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Save &Details:"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   9
         Top             =   2130
         Width           =   945
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Blog Folder:"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   1
         Top             =   300
         Width           =   840
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   585
         TabIndex        =   3
         Top             =   645
         Width           =   915
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   975
         Width           =   885
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   855
         TabIndex        =   7
         Top             =   1665
         Width           =   630
      End
   End
End
Attribute VB_Name = "Logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'Variables
Public Username, Password, Folder, Server As String

Private Sub cmdCancel_Click()
If CheckBox.Value = 1 Then SaveChanges
Unload Me
End Sub

Private Sub cmdOK_Click()
If CheckBox.Value = 1 Then SaveChanges

Username = txtUID.Text
Password = txtPWD.Text
Folder = txtFolder.Text
Server = txtServer.Text

Load WinBlogX
WinBlogX.Show
Unload Me
End Sub

Private Sub Form_Load()

'FileName
Dim m_File As String
m_File = App.Path & "\WinBlog.ini"

'read from the file
Username = ReadIni(m_File, "Login", "Username")
Password = ReadIni(m_File, "Login", "Password")
Folder = ReadIni(m_File, "Login", "Folder")
Server = ReadIni(m_File, "Login", "Server")

If (Username <> "") And (Password <> "") And (Server <> "") Then
Load WinBlogX
WinBlogX.Show
Unload Me
End If

'txtUID.Text = Username
'txtPWD.Text = Password
'txtFolder.Text = Folder
'txtServer.Text = Server

End Sub

Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
ReadIni = Left(RetVal, v)
End Function

Sub SaveChanges()
'FileName
Dim m_File As String
m_File = App.Path & "\WinBlog.ini"

'Write To File Saved Details
WritePrivateProfileString "Login", "UserName", txtUID.Text, m_File
WritePrivateProfileString "Login", "Password", txtPWD.Text, m_File
WritePrivateProfileString "Login", "Folder", txtFolder.Text, m_File
WritePrivateProfileString "Login", "Server", txtServer.Text, m_File
End Sub

Private Sub Picture1_Click()
Load Help
Help.Show
End Sub

Private Sub Picture2_Click()
Load Help
Help.Show
End Sub

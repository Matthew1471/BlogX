VERSION 5.00
Begin VB.Form Step1 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Step1 - Setup Blog"
   ClientHeight    =   2880
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4620
   Icon            =   "Logon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   750
      Picture         =   "Logon.frx":0442
      ScaleHeight     =   345
      ScaleWidth      =   315
      TabIndex        =   12
      ToolTipText     =   "Click For Help"
      Top             =   1710
      Width           =   315
   End
   Begin VB.TextBox txtFolder 
      Height          =   300
      Left            =   1425
      TabIndex        =   2
      Top             =   360
      Width           =   2835
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2430
      TabIndex        =   10
      Top             =   2280
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   870
      TabIndex        =   9
      Top             =   2295
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   2070
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
         TabIndex        =   11
         ToolTipText     =   "Click For Help"
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1515
         TabIndex        =   4
         Top             =   600
         Width           =   2625
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   930
         Width           =   2625
      End
      Begin VB.TextBox txtServer 
         Height          =   330
         Left            =   1500
         TabIndex        =   8
         Top             =   1605
         Width           =   2640
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
         Height          =   195
         Index           =   1
         Left            =   735
         TabIndex        =   3
         Top             =   645
         Width           =   765
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Index           =   2
         Left            =   765
         TabIndex        =   5
         Top             =   975
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Index           =   3
         Left            =   945
         TabIndex        =   7
         Top             =   1665
         Width           =   510
      End
   End
End
Attribute VB_Name = "Step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables
Private m_File As String

Private Sub cmdOK_Click()
SaveChanges
Load Step2
Step2.Show
Unload Me
End Sub

Private Sub Form_Load()

' Filename
m_File = App.Path & "\Mail2BlogX.ini"

'read from the file
Dim Username, Password, Folder, Server As String
Username = ReadIni(m_File, "Login", "Username")
Password = ReadIni(m_File, "Login", "Password")
Folder = ReadIni(m_File, "Login", "Folder")
Server = ReadIni(m_File, "Login", "Server")

txtUID.Text = Username
txtPWD.Text = Password
txtFolder.Text = Folder
txtServer.Text = Server

End Sub

Sub SaveChanges()
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

Private Sub cmdCancel_Click()
Unload Me
End Sub


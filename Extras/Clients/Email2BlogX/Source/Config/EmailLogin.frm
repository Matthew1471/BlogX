VERSION 5.00
Begin VB.Form Step2 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Step2 - Setup Email"
   ClientHeight    =   2385
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4515
   Icon            =   "EmailLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2355
      TabIndex        =   7
      Top             =   1800
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   645
      TabIndex        =   6
      Top             =   1830
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   1590
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   4230
      Begin VB.TextBox txtServer 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         TabIndex        =   4
         Top             =   975
         Width           =   2625
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   945
         TabIndex        =   2
         Top             =   240
         Width           =   2625
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   945
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   555
         Width           =   2625
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Mail Server:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Username:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   285
         Width           =   765
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   3
         Top             =   555
         Width           =   735
      End
   End
End
Attribute VB_Name = "Step2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables
Private m_File As String

Private Sub cmdOK_Click()
SaveChanges
Unload Me
End Sub

Private Sub Form_Load()

' Filename
m_File = App.Path & "\Mail2BlogX.ini"

'Read from the file
Dim POP3Username, POP3Password, POP3Server As String
POP3Username = ReadIni(m_File, "Login", "POP3Username")
POP3Password = ReadIni(m_File, "Login", "POP3Password")
POP3Server = ReadIni(m_File, "Login", "POP3Server")

txtUID.Text = POP3Username
txtPWD.Text = POP3Password
txtServer.Text = POP3Server

End Sub

Sub SaveChanges()
'Write To File Saved Details
WritePrivateProfileString "Login", "POP3UserName", txtUID.Text, m_File
WritePrivateProfileString "Login", "POP3Password", txtPWD.Text, m_File
WritePrivateProfileString "Login", "POP3Server", txtServer.Text, m_File
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

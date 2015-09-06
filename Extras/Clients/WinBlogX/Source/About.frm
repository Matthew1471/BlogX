VERSION 5.00
Begin VB.Form Dialogue 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WinBlogX"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "About.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Frame About 
      BackColor       =   &H00C0C0FF&
      Caption         =   "About WinBlogX"
      Height          =   2655
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   5025
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "About.frx":0442
         Top             =   2250
         Width           =   5115
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Version : 1.04.14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1335
         TabIndex        =   5
         Top             =   660
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Matthew1471(C) 2004 Weblog Publisher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   555
         TabIndex        =   4
         Top             =   315
         Width           =   4560
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   105
         Picture         =   "About.frx":048D
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "http://matthew1471.co.uk/Blog/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   390
         TabIndex        =   2
         Top             =   1245
         Width           =   4320
      End
   End
End
Attribute VB_Name = "Dialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Sub Label2_Click()
If ShellExecute(Me.hwnd, "open", "http://BlogX.co.uk/Refer.asp?Refer=WinBlogX&Version=1.04.14", vbNullString, "", 0) < 33 Then
'--- Exception ---
MsgBox "Please Inform The Webmaster An Exception Occured"
'--- Exception ---
End If
End Sub
Private Sub OKButton_Click()
Unload Me
End Sub
Private Sub CancelButton_Click()
Unload Me
End Sub

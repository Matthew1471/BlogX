VERSION 5.00
Begin VB.Form Help 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Help"
   ClientHeight    =   2820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6915
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame About 
      BackColor       =   &H00C0C0FF&
      Caption         =   "WinBlogX Help"
      Height          =   2580
      Left            =   105
      TabIndex        =   1
      Top             =   135
      Width           =   5310
      Begin VB.TextBox Help 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "Help.frx":0742
         Top             =   300
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Help.frx":081B
         Top             =   2100
         Width           =   5265
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   105
         Picture         =   "Help.frx":0891
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5535
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub

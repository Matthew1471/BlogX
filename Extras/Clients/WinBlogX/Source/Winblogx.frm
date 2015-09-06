VERSION 5.00
Begin VB.Form WinBlogX 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinBlogX"
   ClientHeight    =   4425
   ClientLeft      =   2340
   ClientTop       =   2520
   ClientWidth     =   6720
   Icon            =   "WinBlogX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton StrikeOut 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   330
      Left            =   2130
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   705
      Width           =   375
   End
   Begin VB.CommandButton Underline 
      Caption         =   "&U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1755
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   705
      Width           =   375
   End
   Begin VB.CommandButton Italics 
      Caption         =   "&I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   705
      Width           =   375
   End
   Begin VB.CommandButton Bold 
      Caption         =   "&B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1005
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   705
      Width           =   375
   End
   Begin VB.CommandButton Line 
      Caption         =   "&Line"
      Height          =   330
      Left            =   6090
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   705
      Width           =   480
   End
   Begin VB.CommandButton Link 
      Caption         =   "&Link"
      Height          =   330
      Left            =   5610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   705
      Width           =   480
   End
   Begin VB.TextBox FormCategory 
      Height          =   375
      Left            =   4680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox FormText 
      Height          =   2415
      Left            =   585
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1170
      Width           =   6015
   End
   Begin VB.TextBox FormTitle 
      Height          =   375
      Left            =   600
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "&Save New Entry"
      Height          =   615
      Left            =   2565
      MaskColor       =   &H00FF0000&
      TabIndex        =   6
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Formating :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   690
      Width           =   6480
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Category :"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2565
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Text :"
      Height          =   2430
      Left            =   105
      TabIndex        =   4
      Top             =   1185
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T&itle :"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu Verify 
         Caption         =   "Check Password"
         Index           =   0
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Index           =   0
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Entry 
      Caption         =   "Entry"
      Index           =   1
      Begin VB.Menu Clear 
         Caption         =   "Clear"
         Index           =   0
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Index           =   0
      Begin VB.Menu Update 
         Caption         =   "Check For Update"
         Index           =   0
         Shortcut        =   ^U
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "WinBlogX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objXMLHTTP As Object
Dim Response

Private Sub About_Click(Index As Integer)
Load Dialogue
Dialogue.Show
End Sub

Private Sub Bold_Click()

If FormText.SelText <> "" Then
FormText.SelText = "<b>" & FormText.SelText & "</b>"
Else
FormText.Text = FormText.Text & "<b> </b>"
End If

End Sub

Private Sub Clear_Click(Index As Integer)
FormTitle.Text = ""
FormText.Text = ""
FormCategory.Text = ""
End Sub

Private Sub Exit_Click(Index As Integer)
Quit ("Are You Sure You Want To Quit?")
End Sub

   Private Function Quit(Value)
   Dim MyString
   Response = MsgBox(Value, vbYesNo + vbQuestion + vbDefaultButton2, "Quit")
   If Response = vbYes Then
    Dim oFrm As Form

    For Each oFrm In Forms
        Unload oFrm
    Next
   Else
   MyString = "No"
   End If
   End Function
   
   Public Function Encode(Value As String) As String
   
   Encode = Replace(Value, vbCrLf, "%0D%0A")

   For i = 0 To 31
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 33 To 36
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 38 To 47
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 58 To 64
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 91 To 96
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 123 To 255
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   Encode = Replace(Encode, " ", "+")
   End Function
   
Private Sub Command1_Click()
On Error Resume Next
Dim URL As String
URL = "http://" & Logon.Server & "/" & Logon.Folder & "/" & "Application.asp"

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
'Set objXMLHTTP = CreateObject("Msxml2.XMLHTTP")

' Set the method of request which is POST and the URL,and set the Async parameter to false
objXMLHTTP.Open "POST", URL, False

' Sets the header so that the web server knows a form is going to be posted
objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLHTTP.setRequestHeader "Referer", "VisualBasic"
objXMLHTTP.setRequestHeader "User-Agent", "Matthew1471 Blogging Utility"


'--- Encode The Text ---
Dim Content
Content = Encode(FormText.Text)

'--- Encode The Ttitle ---
Dim Title
Title = Encode(FormTitle.Text)

' Construct the message body first before we send, it is a name/value pair,separated by ampersands
' which looks like "username=admin&password=letmein"
strbody = "Username=" & Logon.Username & "&Password=" & Logon.Password & "&Content=" & Content & "&Title=" & Title & "&Category=" & Replace(FormCategory.Text, " ", "+")

' Send It Baby!
objXMLHTTP.send strbody
Response = objXMLHTTP.ResponseText

' Let Them Know We Failed
If Error <> "" Then MsgBox "Error: " & Error

' So Did The Submission Go Ok?
If Response = "Entry Submission Successfull" Then
FormTitle.Text = ""
FormText.Text = ""
FormCategory.Text = ""
Quit ("Finished? Do You Want To Exit The Program?")

ElseIf Response = "No Text Entered" Then
MsgBox "No Text Entered"
ElseIf Response = "" Then
MsgBox "Server Not Found"
Kill (App.Path & "\WinBlog.ini")
ElseIf Response = "User/Password Error" Then
MsgBox "Invalid Username/Password"
Kill (App.Path & "\WinBlog.ini")
ElseIf (InStr(Response, "404") <> 0) Or (InStr(Response, "Not Found") <> 0) Then
MsgBox "Blog Not Found"
Kill (App.Path & "\WinBlog.ini")
Else: MsgBox "Server Error"
MsgBox Response
Kill (App.Path & "\WinBlog.ini")
End If

End Sub

Private Sub Italics_Click()

If FormText.SelText <> "" Then
FormText.SelText = "<i>" & FormText.SelText & "</i>"
Else
FormText.Text = FormText.Text & "<i> </i>"
End If

End Sub

Private Sub Line_Click()
FormText.Text = FormText.Text & "<hr>"
End Sub

Private Sub Link_Click()

URLAddress = InputBox("URL For The Link", "Insert URL")

If (FormText.SelText <> "") And (URLAddress <> "") Then
FormText.SelText = "<a href=""" & URLAddress & """>" & FormText.SelText & "</a>"

ElseIf URLAddress = "" Then
MsgBox "No URL Specified", vbExclamation, "No URL"

Else
FormText.Text = FormText.Text & "<a href=""" & URLAddress & """>Link</a>"
End If

End Sub

Private Sub StrikeOut_Click()

If FormText.SelText <> "" Then
FormText.SelText = "<s>" & FormText.SelText & "</s>"
Else
FormText.Text = FormText.Text & "<s> </s>"
End If

End Sub

Private Sub Underline_Click()

If FormText.SelText <> "" Then
FormText.SelText = "<u>" & FormText.SelText & "</u>"
Else
FormText.Text = FormText.Text & "<u> </u>"
End If

End Sub

Private Sub Update_Click(Index As Integer)
If (FormTitle.Text <> "") Or (FormText.Text <> "") Or (FormCategory.Text <> "") Then
   Response = MsgBox("By clicking ""YES"" you will loose your current entry", vbYesNo + vbQuestion + vbDefaultButton2, "Quit")
   If Response <> vbYes Then Exit Sub
End If

   Load CheckForUpdate
   CheckForUpdate.Show
   
End Sub

Private Sub Verify_Click(Index As Integer)
On Error Resume Next
URL = "http://" & Logon.Server & "/" & Logon.Folder & "/" & "Application.asp"

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

' Set the method of request which is POST and the URL,and set the Async parameter to false
objXMLHTTP.Open "POST", URL, False

' Sets the header so that the web server knows a form is going to be posted
objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLHTTP.setRequestHeader "Referer", "VisualBasic"
objXMLHTTP.setRequestHeader "User-Agent", "Matthew1471 Blogging Utility"

' Construct the message body first before we send, it is a name/value pair,separated by ampersands
' which looks like "username=admin&password=letmein"
strbody = "Username=" & Logon.Username & "&Password=" & Logon.Password

' Send It Baby!
objXMLHTTP.send strbody
Response = objXMLHTTP.ResponseText

If Response = "No Text Entered" Then
MsgBox "Password OK"
ElseIf Response = "" Then
MsgBox "Server Not Found"
Kill (App.Path & "\WinBlog.ini")
ElseIf Response = "User/Password Error" Then
MsgBox "Invalid Username/Password"
Kill (App.Path & "\WinBlog.ini")
ElseIf (InStr(Response, "404") <> 0) Or (InStr(Response, "Not Found") <> 0) Then
MsgBox "Blog Not Found"
Kill (App.Path & "\WinBlog.ini")
Else
MsgBox "Server Error"
Kill (App.Path & "\WinBlog.ini")
End If

' Let Them Know We Failed
If Error <> "" Then MsgBox "Error: " & Error
End Sub

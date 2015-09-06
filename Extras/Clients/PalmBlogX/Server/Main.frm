VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PalmBlogX Server"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4950
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin NTService.NTService NTService 
      Left            =   3645
      Top             =   2265
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "PalmBlogX"
      ServiceName     =   "PalmBlogX Server"
      StartMode       =   2
   End
   Begin VB.Timer Disconnecter 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4065
      Top             =   2250
   End
   Begin VB.TextBox Edit_Title 
      Height          =   285
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   4485
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   4485
      Top             =   2235
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Temp 
      Height          =   2190
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   435
      Width           =   4905
   End
   Begin VB.Label Label_Title 
      Caption         =   "Title : "
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   90
      Width           =   465
   End
   Begin VB.Menu Menu_File 
      Caption         =   "File"
      Begin VB.Menu Menu_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecievedData As String

   
   Function Encode(Value As String) As String
   Dim i
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

Private Sub Disconnecter_Timer()
LogFile ("Connection Timed Out")
Winsock_Close
End Sub

Private Sub Form_Load()

'--- NT Service Management --'
On Error GoTo Err_Load
    Dim strDisplayName As String
    Dim bStarted As Boolean
    
    strDisplayName = NTService.DisplayName

    If Command = "-install" Then
        If NTService.Install Then
            MsgBox strDisplayName & " Installed Successfully"
        Else
            MsgBox strDisplayName & " Failed To Install"
        End If
        End
    ElseIf Command = "-uninstall" Then
        If NTService.Uninstall Then
            MsgBox strDisplayName & " Uninstalled Successfully"
        Else
            MsgBox strDisplayName & " Failed To Uninstall"
        End If
        End
    ElseIf Command = "-debug" Then
        NTService.Debug = True
    ElseIf Command <> "" Then
        MsgBox "Invalid Command Option"
        End
    End If
    
    If Not App.PrevInstance = True Then
        Winsock.LocalPort = 28 'This can be any Valid Port Number
        'Wait for Clients to Connect with Your Server.
        Winsock.Listen
        
        'Connect Service To Windows NT services controller
        NTService.StartService
        
    Else
        End
    End If

    GoTo Normal
    
Err_Load:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
'-- Done --'

Normal:
    
   If Winsock.State = sckListening Then Me.Caption = "PlamBlogX : Listening"

End Sub

Private Sub Menu_Exit_Click()
Unload Me
End Sub


Private Sub NTService_Start(Success As Boolean)
On Error GoTo Err_Start
Me.Hide
Success = True
    
Err_Start:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub

Private Sub NTService_Stop()
On Error GoTo Err_Stop
Unload Me
GoTo Normal

Err_Stop:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    
Normal:
End Sub

Private Sub Winsock_Close()
    Disconnecter.Enabled = False
    Winsock.Close
    Winsock.Listen
    Me.Caption = "PlamBlogX : Listening"
    LogFile ("")
End Sub

Private Sub Winsock_ConnectionRequest(ByVal RequestID As Long)
    On Error Resume Next
    
    'First Check if the Winsock Control is Connected or not
    'If connected then Close it
    If Winsock.State <> sckClosed Then Winsock.Close
    
    Disconnecter.Enabled = True
    Winsock.Accept RequestID
    
    Me.Caption = "PlamBlogX : Connected " & Winsock.RemoteHostIP
      LogFile ("----- " & Now() & " -----")
      LogFile ("Connection Request : " & Winsock.RemoteHostIP)
      LogFile ("-------------------------------")
      
    '-- RESET POST --'
    Temp.Text = ""
    Edit_Title.Text = ""
    RecievedData = ""
    '-- END OF POST --'
    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

    Dim Str As String
    Dim Username, Password, Server, Folder
    Dim Title As String, Content As String
    Dim Serial As String
    Dim strbody, Response As String
    
    Winsock.GetData Str
    
    Disconnecter.Enabled = False
    Disconnecter.Enabled = True
    
    Debug.Print "RAW : " & Str

    Select Case Str
    Case "CurrentVersionCHECK" & Chr(13) & Chr(10)
                LogFile ("Update Check")
                Winsock.SendData "V1" & Chr(10)
    Case ""
                Winsock_Close
    Case Else
                RecievedData = RecievedData & Str
                
                If InStr(RecievedData, "<Send POST>") <> 0 Then
                                
                RecievedData = Left(RecievedData, Len(RecievedData) - 11)
                                
                If ((InStr(RecievedData, "&") <> 0) And (InStr(RecievedData, Chr(10)) <> 0)) Then
                RecievedData = Replace(RecievedData, "Username=", "")
                Username = Left(RecievedData, InStr(RecievedData, "&") - 1)
                RecievedData = Right(RecievedData, Len(RecievedData) - InStr(RecievedData, "&"))
                LogFile ("Username : " & Username)
                
                RecievedData = Replace(RecievedData, "Password=", "")
                Password = Left(RecievedData, InStr(RecievedData, Chr(13) & Chr(10)) - 1)
                RecievedData = Right(RecievedData, Len(RecievedData) - InStr(RecievedData, Chr(10)))
                LogFile ("Password : " & Password)
                End If
                                
                 If RecievedData <> "" Then
                                                                          
                    If InStr(RecievedData, "Serial : ") <> 0 Then
                    Serial = Left(RecievedData, InStr(RecievedData, Chr(10)) - 2)
                    RecievedData = Replace(RecievedData, Serial & Chr(13) & Chr(10), "")
                    Serial = Right(Serial, Len(Serial) - 9)
                    LogFile ("Serial   : " & Serial)
                    End If
                                                                          
                    If InStr(RecievedData, "Server : ") <> 0 Then
                    Server = Left(RecievedData, InStr(RecievedData, Chr(10)) - 2)
                    RecievedData = Replace(RecievedData, Server & Chr(13) & Chr(10), "")
                    Server = Right(Server, Len(Server) - 9)
                    LogFile ("Server   : " & Server)
                    End If
                    
                    If InStr(RecievedData, "Folder : ") <> 0 Then
                    Folder = Left(RecievedData, InStr(RecievedData, Chr(10)) - 2)
                    RecievedData = Replace(RecievedData, Folder & Chr(13) & Chr(10), "")
                    Folder = Right(Folder, Len(Folder) - 9)
                    If Folder <> "" Then LogFile ("Folder   : " & Folder)
                    End If

                    If InStr(RecievedData, "Title : ") <> 0 Then
                    Title = Left(RecievedData, InStr(RecievedData, Chr(10)) - 2)
                    RecievedData = Replace(RecievedData, Title & Chr(10), "")
                    Title = Right(Title, Len(Title) - 8)
                    End If
                    
                    If InStr(RecievedData, Chr(10)) <> 0 Then
                    RecievedData = Replace(RecievedData, Chr(10), vbCrLf)
                    Content = Right(RecievedData, Len(RecievedData) - InStr(RecievedData, Chr(10)))
                    End If
                
                    If Serial <> "" Then
                    
                    '-- POST data --'
                    Edit_Title.Text = Title
                    Temp.Text = Content
                    
                    On Error Resume Next
                    Dim objXMLHTTP As Object
                    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

                    ' Set the method of request which is POST and the URL,and set the Async parameter to false
                    Dim URL As String
                    
                    URL = "http://" & Server & "/"
                    If Folder <> "" Then URL = URL & Folder & "/"
                    URL = URL & "Application.asp"
                    
                    Debug.Print URL
                    objXMLHTTP.Open "POST", URL, False

                    ' Sets the header so that the web server knows a form is going to be posted
                    objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    objXMLHTTP.setRequestHeader "Referer", "VisualBasic"
                    objXMLHTTP.setRequestHeader "User-Agent", "Matthew1471 Blogging Utility"
                                       
                    '--- Encode The Text ---
                    Content = Encode(Content)
                    If Content = "PasswordCHECK" Then
                    Content = ""
                    LogFile ("User/Password CHECK")
                    End If

                    '--- Encode The Title ---
                    Title = Encode(Title)

                    ' Construct the message body first before we send, it is a name/value pair,separated by ampersands
                    ' which looks like "username=admin&password=letmein"
                    strbody = "Username=" & Username & "&Password=" & Password & "&Content=" & Content & "&Title=" & Title

                    ' Send It Baby!
                    objXMLHTTP.send strbody
                    Response = objXMLHTTP.ResponseText

                    ' Let Them Know We Failed
                    If Error <> "" Then LogFile "Error    : " & Error

                    ' So Did The Submission Go Ok?
                    If Response = "Entry Submission Successfull" Then
                        Winsock.SendData "Message Relayed" & Chr(10)
                        LogFile ("Entry Relayed")
                    ElseIf Response = "No Text Entered" Then
                        Winsock.SendData "No Text Entered" & Chr(10)
                        LogFile ("No Text Entered")
                    ElseIf Response = "" Then
                        Winsock.SendData "Server Not Found" & Chr(10)
                        LogFile ("Server Not Found")
                    ElseIf Response = "User/Password Error" Then
                        Winsock.SendData "Invalid Username/Password" & Chr(10)
                        LogFile ("Failed Authentication")
                    ElseIf (InStr(Response, "404") <> 0) Or (InStr(Response, "Not Found") <> 0) Then
                        Winsock.SendData "Blog Not Found" & Chr(10)
                        LogFile ("Blog Not Found")
                    Else:
                        Winsock.SendData "Server Error" & Chr(10)
                        LogFile ("Server Error")
                        LogFile ("Debug : " & Response)
                    End If
                    
                    Set objXMLHTTP = Nothing
                    
                    On Error GoTo 0
                    Else
                    Winsock.SendData "Unregistered" & Chr(10)
                    LogFile ("Unregistered (" & Serial & ")")
                    End If
                    '-- END OF POST --'
                 Else
                 Winsock.SendData "No Text Entered" & Chr(10)
                 LogFile ("No Text Entered")
                 End If
                 
                 End If
    End Select

End Sub

Private Function LogFile(Message As String)

    Dim Path As String
    Path = App.Path & "\Debug.log"

    Open Path For Append As #1
      Print #1, Message
    Close #1
End Function

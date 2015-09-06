VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CheckMail 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email2BlogX"
   ClientHeight    =   720
   ClientLeft      =   2340
   ClientTop       =   2220
   ClientWidth     =   1980
   Icon            =   "CheckMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   1980
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock POP3 
      Left            =   660
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CheckMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_Retr
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State         As POP3States
Private m_oMessage       As CMessage
Private m_colMessages   As New CMessages

Public Title, Content, Category As String
Public SplittedData

Private Sub Form_Load()
'On Error Resume Next

Me.Hide
LoadVariables

If (Username = "") Or (Password = "") Or (Folder = "") Or _
   (Server = "") Or (POP3Server = "") Or (POP3Username = "") Then GoTo EmergencyExit

POP3.Close
POP3.LocalPort = 0
POP3.Connect POP3Server, 110
    
GoTo Normal

EmergencyExit:
Dim X
MsgBox "Please Make Sure Email2BlogX Is Set-up Before Running This Porgram Again"
X = ShellExecute(Me.hwnd, "open", App.Path & "\Config.exe", vbNullString, "", 0)

Normal:

End Sub

Private Sub POP3_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    '
    'Save the received data into strData variable
    POP3.GetData strData
    'Debug.Print strData
    
    If Left$(strData, 1) = "+" Or m_State = POP3_Retr Then
        'If the first character of the server's response is "+" then
        'server accepted the client's command and waits for the next one
        'If this symbol is "-" then here we can do nothing
        'and execution skips to the Else section of the code
        'The first symbol may differ from "+" or "-" if the received
        'data are the part of the message's body, i.e. when
        'm_State = POP3_RETR (the loading of the message state)
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                POP3.SendData "USER " & POP3Username & vbCrLf
                Debug.Print "USER " & POP3Username
                'Here is the end of POP3_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                POP3.SendData "PASS " & POP3Password & vbCrLf
                Debug.Print "PASS " & POP3Password
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                POP3.SendData "STAT" & vbCrLf
                Debug.Print "STAT"
                
            Case POP3_DELE
                '
                'The server is now erasing the message
                '
                m_State = POP3_Retr
                
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                intMessages = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                If intMessages > 0 Then
                    '
                    'Oops. There is something in the mailbox!
                    'Change the session state
                    m_State = POP3_Retr
                    '
                    'Increment the number of messages by one
                    intCurrentMessage = intCurrentMessage + 1
                    '
                    'and we're sending to the server the RETR command in
                    'order to retrieve the first message
                    POP3.SendData "RETR 1" & vbCrLf
                    Debug.Print "RETR 1"
                Else
                    'The mailbox is empty. Send the QUIT command to the
                    'server in order to close the session
                    m_State = POP3_QUIT
                    POP3.SendData "QUIT" & vbCrLf
                    Debug.Print "QUIT"
                End If
            Case POP3_Retr
                'This code executes while the retrieving of the mail body
                'The size of the message could be quite big and the
                'DataArrival event may rise several time. All the received
                'data stores at the strBuffer variable:
                strBuffer = strBuffer & strData
                '
                'If case of presence of the point in the buffer it indicates
                'the end of the message (look at SMTP protocol)
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    '
                    'Done! The message has loaded
                    '
                    'Delete the first string-the server's response
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    '
                    'Delete the last string. It contains only the "." symbol,
                    'which indicates the end of the message
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    '
                    'Add new message to m_colMessages collection
                    Set m_oMessage = New CMessage
                    m_oMessage.CreateFromText strBuffer
                    m_colMessages.Add m_oMessage, m_oMessage.MessageID
                    
                    m_State = POP3_DELE
                    Debug.Print "DELE " & intCurrentMessage
                    POP3.SendData "DELE " & intCurrentMessage & vbCrLf
                                        
                    Set m_oMessage = Nothing
                    '
                    'Clear buffer for next message
                    strBuffer = ""
                    'Now we comparing the number of loaded messages with the
                    'one returned as a response to the STAT command
                    If intCurrentMessage = intMessages Then
                        'If these values are equal then all the messages
                        'have loaded. Now we can finish the session. Due to
                        'this reason we send the QUIT command to the server
                        m_State = POP3_QUIT
                        POP3.SendData "QUIT" & vbCrLf
                        Debug.Print "QUIT"
                    Else
                        'If these values aren't equal then there are
                        'remain messages. According with that
                        'we increment the messages' counter
                        intCurrentMessage = intCurrentMessage + 1
                        '
                        'Change current state of session
                        m_State = POP3_Retr
                        '
                        'Send RETR command to download next message
                        POP3.SendData "RETR " & CStr(intCurrentMessage) & vbCrLf
                        Debug.Print "RETR " & intCurrentMessage
                    End If
                End If
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                POP3.Close
                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                Call ListMessages
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            POP3.Close
            MsgBox "POP3 Error: " & strData, vbExclamation, "POP3 Error"
    End If
End Sub

Private Sub POP3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    MsgBox "Winsock Error: #" & Number & vbCrLf & Description
            
End Sub

Private Sub ListMessages()

    Dim Category As String
    Dim oMes As CMessage
    
    For Each oMes In m_colMessages
        Dim I As Integer
        Dim Attachments, AttachmentData
        
        MsgBox oMes.MessageBody
        
        If InStr(1, LCase(oMes.MessageBody), "category: ") <> 0 Then
            SplittedData = Split(oMes.MessageBody, vbCrLf)
            For I = 0 To UBound(SplittedData)
            
                If InStr(1, SplittedData(I), "category: ", vbTextCompare) <> 0 Then
                Category = Replace(SplittedData(I), "category: ", "", 1, -1, vbTextCompare)
                SplittedData(I) = ""
                oMes.MessageBody = Join(SplittedData, vbCrLf)
                End If
                
            Next I
        Else
                Category = ""
        End If
        
        If InStr(1, LCase(oMes.MessageBody), "content-transfer-encoding:") <> 0 Then
        
            SplittedData = Split(oMes.MessageBody, vbCrLf)
            
            For I = 0 To UBound(SplittedData)
            
                If InStr(1, LCase(SplittedData(I)), "content-transfer-encoding: base64") <> 0 Then
                Attachments = Mid(SplittedData(I - 4), 43)
                AttachmentData = Mid(oMes.MessageBody, InStr(1, oMes.MessageBody, "content-transfer-encoding: base64", vbTextCompare))
                oMes.MessageBody = Left(oMes.MessageBody, InStr(1, oMes.MessageBody, "content-transfer-encoding: base64", vbTextCompare) - 1)
                End If
                
                If InStr(1, LCase(SplittedData(I)), "content-transfer-encoding: 8bit") <> 0 Then
                oMes.MessageBody = DeleteElement(I)
                oMes.MessageBody = DeleteElement(I - 1)
                oMes.MessageBody = DeleteElement(I - 1)
                oMes.MessageBody = DeleteElement(I - 1)
                oMes.MessageBody = DeleteElement(I - 1)
                End If
            Next I
            
            
           SplittedData = Split(oMes.MessageBody, vbCrLf)
            
            For I = 0 To UBound(SplittedData)
                If InStr(1, LCase(SplittedData(I)), "content-transfer-encoding: ") <> 0 Then
                oMes.MessageBody = DeleteElement(I)
                End If
            Next I
            
           SplittedData = Split(oMes.MessageBody, vbCrLf)
            
            For I = 0 To UBound(SplittedData)
                If InStr(1, LCase(SplittedData(I)), "content-type: ") <> 0 Then
                oMes.MessageBody = DeleteElement(I + 2)
                End If
            Next I
            
        End If
        
                oMes.MessageBody = Replace(oMes.MessageBody, "<BR>", vbCrLf)
                oMes.MessageBody = Replace(oMes.MessageBody, "=20", "")
                oMes.MessageBody = Replace(oMes.MessageBody, "=" & vbCrLf, "")
                
        Debug.Print "Message ID : " & oMes.MessageID
        Debug.Print "From : " & oMes.From
        Debug.Print "Subject : " & oMes.Subject
        Debug.Print "Date : " & oMes.SendDate
        Debug.Print "Size : " & oMes.Size & "kb"
        Debug.Print "Body : " & Replace(oMes.MessageBody, "=" & vbCrLf, vbCrLf)
        'Debug.Print "Attachments : " & Attachments
        Debug.Print "Category: " & Category
        
        Title = oMes.Subject
        Content = StripTags(oMes.MessageBody)
        
        If InStr(1, oMes.MessageBody, "message in MIME format") <> 0 Then

        Dim Filename As String
        
            On Error Resume Next
            If Dir$(App.Path & "\Failed Messages\*.*") = vbNullString Then MkDir App.Path & "\Failed Messages"
            On Error GoTo 0
        
        Filename = Replace(Time(), ":", "") & ".eml"
        Open App.Path & "\Failed Messages\" & Filename For Output As #1
        Print #1, oMes.MessageBody
        Close #1
        
        MsgBox "The Message Downloaded Was Not In Plain Text, It may not display correctly, See " & Filename & " for details"
        End If
        
        PostToBlog
    Next
    
    Unload Me
    
End Sub

Private Function DeleteElement(Element As Integer)

 ReDim DeleteArray(0)
 Dim Elecount As Integer
 Dim ElementNo As Integer
 
            Elecount = 1
            
            For ElementNo = 1 To UBound(SplittedData)
            
                If (ElementNo < Element - 5) Or (ElementNo > Element) Then
                    ReDim Preserve DeleteArray(Elecount)
                    DeleteArray(Elecount) = SplittedData(ElementNo)
                    Elecount = Elecount + 1
                End If
                
            Next ElementNo
            
            DeleteElement = Join(DeleteArray(), vbCrLf)
End Function

Public Function StripTags(ByVal HTML As String) As String
    ' Removes tags from passed HTML
    Dim objRegEx As New RegExp
    objRegEx.Pattern = "<[^>]*>"
    objRegEx.IgnoreCase = True
    objRegEx.Global = True
    StripTags = objRegEx.Replace(HTML, "")
    
End Function

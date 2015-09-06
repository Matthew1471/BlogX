VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "RFC822 e-mail message."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variables to hold property values
Private m_strReturnPath      As String
Private m_strReceived        As String
Private m_strSendDate        As String
Private m_strMessageID       As String
Private m_strMessageTo       As String
Private m_strFrom            As String
Private m_strSubject         As String
Private m_strReplyTo         As String
Private m_strSender          As String
Private m_strCC              As String
Private m_strBCC             As String
Private m_strInReplyTo       As String
Private m_strReferences      As String
Private m_strKeywords        As String
Private m_strComments        As String
Private m_strEncrypted       As String
Private m_strMessageText     As String
Private m_strMessageBody     As String
Private m_strHeaders         As String
Private m_lSize              As Long

Public Sub CreateFromText(strMessage As String)

    Dim intPosA         As Integer
    Dim vHeaders        As Variant
    Dim vField          As Variant
    Dim strHeader       As String
    Dim strHeaderName   As String
    Dim strHeaderValue  As String
    
    intPosA = InStr(1, strMessage, vbCrLf & vbCrLf)
    If intPosA Then
        m_strHeaders = Left$(strMessage, intPosA - 1)
        m_strMessageBody = Right$(strMessage, Len(strMessage) - intPosA - 3)
        m_strMessageText = strMessage
    Else
        Err.Raise vbObjectError + 512 + 101, "CMessage.CreateFromText", _
                    "Invalid message format"
        Exit Sub
    End If
    
    vHeaders = Split(m_strHeaders, vbCrLf)
    For Each vField In vHeaders
        strHeader = CStr(vField)
        intPosA = InStr(1, strHeader, ":")
        If intPosA Then
            strHeaderName = LCase(Left$(strHeader, intPosA - 1))
        Else
            strHeaderName = ""
        End If
        strHeaderValue = Trim$(Right$(strHeader, Len(strHeader) - intPosA))
        Select Case strHeaderName
            Case "return-path"
                m_strReturnPath = strHeaderValue
            Case "received"
                m_strReceived = strHeaderValue
            Case "from"
                m_strFrom = strHeaderValue
            Case "sender"
                m_strSender = strHeaderValue
            Case "reply-to"
                m_strReplyTo = strHeaderValue
            Case "to"
                m_strMessageTo = strHeaderValue
            Case "cc"
                m_strCC = strHeaderValue
            Case "bcc"
                m_strBCC = strHeaderValue
            Case "message-id"
                m_strMessageID = strHeaderValue
            Case "in-reply-to"
                m_strInReplyTo = strHeaderValue
            Case "references"
                m_strReferences = strHeaderValue
            Case "keywords"
                m_strKeywords = strHeaderValue
            Case "subject"
                m_strSubject = strHeaderValue
            Case "comments"
                m_strComments = strHeaderValue
            Case "encrypted"
                m_strEncrypted = strHeaderValue
            Case "date"
                m_strSendDate = strHeaderValue
        End Select
    Next
    
End Sub


Public Function CombineMessage() As String

End Function


Public Property Let Headers(ByVal vData As String)
Attribute Headers.VB_Description = "Headers text."
    m_strHeaders = vData
End Property


Public Property Get Headers() As String
    Headers = m_strHeaders
End Property

Public Property Let MessageBody(ByVal vData As String)
Attribute MessageBody.VB_Description = "Message text without headers."
    m_strMessageBody = vData
End Property

Public Property Get MessageBody() As String
    MessageBody = m_strMessageBody
End Property

Public Property Let MessageText(ByVal vData As String)
Attribute MessageText.VB_Description = "Message text with headers."
    m_strMessageText = vData
End Property

Public Property Get MessageText() As String
    MessageText = m_strMessageText
End Property

Public Property Let Encrypted(ByVal vData As String)
Attribute Encrypted.VB_Description = "Value of Encrypted header field."
    m_strEncrypted = vData
End Property

Public Property Get Encrypted() As String
    Encrypted = m_strEncrypted
End Property

Public Property Let Comments(ByVal vData As String)
Attribute Comments.VB_Description = "Value of Comments header field."
    m_strComments = vData
End Property

Public Property Get Comments() As String
    Comments = m_strComments
End Property

Public Property Let Keywords(ByVal vData As String)
Attribute Keywords.VB_Description = "Value of Keywords header field."
    m_strKeywords = vData
End Property

Public Property Get Keywords() As String
    Keywords = m_strKeywords
End Property

Public Property Let References(ByVal vData As String)
Attribute References.VB_Description = "Value of References header field."
    m_strReferences = vData
End Property

Public Property Get References() As String
    References = m_strReferences
End Property

Public Property Let InReplyTo(ByVal vData As String)
Attribute InReplyTo.VB_Description = "Value of In-Reply-To header field."
    m_strInReplyTo = vData
End Property

Public Property Get InReplyTo() As String
    InReplyTo = m_strInReplyTo
End Property

Public Property Let BCC(ByVal vData As String)
Attribute BCC.VB_Description = "Value of BCC header field."
    m_strBCC = vData
End Property

Public Property Get BCC() As String
    BCC = m_strBCC
End Property

Public Property Let CC(ByVal vData As String)
Attribute CC.VB_Description = "Value of CC header field."
    m_strCC = vData
End Property

Public Property Get CC() As String
    CC = m_strCC
End Property

Public Property Let Sender(ByVal vData As String)
Attribute Sender.VB_Description = "Value of Sender header field."
    m_strSender = vData
End Property

Public Property Get Sender() As String
    Sender = m_strSender
End Property

Public Property Let ReplyTo(ByVal vData As String)
Attribute ReplyTo.VB_Description = "Value of ReplyTo header field."
    m_strReplyTo = vData
End Property

Public Property Get ReplyTo() As String
    ReplyTo = m_strReplyTo
End Property

Public Property Let Subject(ByVal vData As String)
Attribute Subject.VB_Description = "Value of Subject header field."
    m_strSubject = vData
End Property

Public Property Get Subject() As String
    Subject = m_strSubject
End Property

Public Property Let From(ByVal vData As String)
Attribute From.VB_Description = "Value of From header field."
    m_strFrom = vData
End Property

Public Property Get From() As String
    From = m_strFrom
End Property

Public Property Let MessageTo(ByVal vData As String)
Attribute MessageTo.VB_Description = "Value of To header field."
    m_strMessageTo = vData
End Property

Public Property Get MessageTo() As String
    MessageTo = m_strMessageTo
End Property

Public Property Let MessageID(ByVal vData As String)
Attribute MessageID.VB_Description = "Value of Message-ID header field."
    m_strMessageID = vData
End Property

Public Property Get MessageID() As String
    MessageID = m_strMessageID
End Property

Public Property Let SendDate(ByVal vData As String)
Attribute SendDate.VB_Description = "Value of Date hesder field."
    m_strSendDate = vData
End Property

Public Property Get SendDate() As String
    SendDate = m_strSendDate
End Property

Public Property Let Received(ByVal vData As String)
Attribute Received.VB_Description = "Value of Received header field."
    m_strReceived = vData
End Property

Public Property Get Received() As String
    Received = m_strReceived
End Property

Public Property Let ReturnPath(ByVal vData As String)
Attribute ReturnPath.VB_Description = "Value of Return-Path header field."
    m_strReturnPath = vData
End Property

Public Property Get ReturnPath() As String
    ReturnPath = m_strReturnPath
End Property

Public Property Get Size() As Long
    Size = Len(m_strMessageText)
End Property

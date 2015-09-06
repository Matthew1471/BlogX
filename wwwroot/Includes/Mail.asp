<%
' --------------------------------------------------------------------------
'¦Introduction : Mail Component Handler.                                    ¦
'¦Purpose      : Acts as a mailer API and sends the e-mail.                 ¦
'¦Requires     : Nothing.                                                   ¦
'¦Used By      : Admin/MailingListMembers.asp, CommentNotify.asp,           ¦
'¦               Comments_Validate.asp.                                     ¦
'---------------------------------------------------------------------------

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************

Select Case LCase(EmailComponent)

 Case "abmailer"
  Set Mail = Server.CreateObject("ABMailer.Mailman")
   Mail.ServerAddr = EmailServer
   Mail.FromName = Name
   Mail.FromAddress = From
   Mail.SendTo = ToEmail
   Mail.MailSubject = Subject
   Mail.MailMessage = Body
   On Error Resume Next
    Mail.SendMail
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "aspemail"
  Set Mail = Server.CreateObject("Persits.MailSender")
   Mail.FromName = Name
   Mail.From = From
   Mail.AddReplyTo From
   Mail.Host = EmailServer
   Mail.AddAddress ToEmail, ToName
   Mail.Subject = Subject
   Mail.Body = Body
   On Error Resume Next
    Mail.Send
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "aspmail"
  Set Mail = Server.CreateObject("SMTPsvg.Mailer")
   Mail.ContentType = "text/html"
   Mail.FromName = Name
   Mail.FromAddress = From
   Mail.ReplyTo = From
   Mail.RemoteHost = EmailServer
   Mail.AddRecipient ToName, ToEmail
   Mail.Subject = Subject
   Mail.BodyText = Body
   On Error Resume Next
    SendOk = Mail.SendMail
    If NOT(SendOk) <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"

 Case "aspqmail"
  Set Mail = Server.CreateObject("SMTPsvg.Mailer")
   Mail.ContentType = "text/html"
   Mail.QMessage = 1
   Mail.FromName = Name
   Mail.FromAddress = From
   Mail.ReplyTo = From
   Mail.RemoteHost = EmailServer
   Mail.AddRecipient ToName, ToEmail
   Mail.Subject = Subject
   Mail.BodyText = Body
   On Error Resume Next
    Mail.SendMail
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "cdonts"
  Set Mail = Server.CreateObject("CDONTS.NewMail")
   Mail.BodyFormat = 0
   Mail.MailFormat = 0
   On Error Resume Next
    Mail.Send From, ToEmail, Subject, Body
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "chilicdonts"
  Set Mail = Server.CreateObject("CDONTS.NewMail")
   Mail.Host = EmailServer
   Mail.To = ToName & "<" & ToEmail & ">"
   Mail.From = Name & "<" & From & ">"
   Mail.Subject = Subject
   Mail.Body = Body
   On Error Resume Next
    Mail.Send
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "cdosys"
  Set iConf = Server.CreateObject("CDO.Configuration")

  Set Flds = iConf.Fields 
   '-- Set and update fields properties --'
   Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
   Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailServer
   'Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
   'Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
   'Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
   Flds.Update

    Set Mail = Server.CreateObject("CDO.Message")
     Set Mail.Configuration = iConf

     '-- Format and send message --'
     Err.Clear 

     Mail.To = ToName & "<" & ToEmail & ">"
     Mail.From = Name & "<" & From & ">"
     If ReplyTo <> "" Then Mail.ReplyTo = ReplyTo
     Mail.Subject = Subject
     Mail.HTMLBody = Body

     '-- Forcing a character set will fix a Nokia e-mail client bug --'
     '-- http://msdn.microsoft.com/en-us/library/ms526296(EXCHG.10).aspx --'
     Mail.HTMLBodyPart.Charset = "us-ascii"

     On Error Resume Next
      Mail.Send
	  If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

    Set Flds = Nothing
    Set iConf = Nothing

 Case "dkqmail"
	Set Mail = Server.CreateObject("dkQmail.Qmail")
	 Mail.FromEmail = From
	 Mail.ToEmail = ToEmail
	 Mail.Subject = Subject
	 Mail.Body = Body
	 Mail.CC = ""
	 Mail.MessageType = "TEXT"
	 On Error Resume Next
	  Mail.SendMail()
	  If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "dundasmailq"
  Set Mail = Server.CreateObject("Dundas.Mailer")
   On Error Resume Next
    Mail.QuickSend From, ToEmail, Subject, Body
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "dundasmails"
  Set Mail = Server.CreateObject("Dundas.Mailer")
   Mail.TOs.Add ToEmail
   Mail.FromAddress = From
   Mail.Subject = Subject
   Mail.Body = Body
   On Error Resume Next
	Mail.SendMail
	If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "geocel"
  Set Mail = Server.CreateObject("Geocel.Mailer")
  Mail.AddServer EmailServer, 25
  Mail.AddRecipient ToEmail, ToName
  Mail.FromName = Name
  Mail.FromAddress = From
  Mail.Subject = Subject
  Mail.Body = Body
  On Error Resume Next
   Mail.Send()
	If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "iismail"
  Set Mail = Server.CreateObject("iismail.iismail.1")
   Mail.Server = EmailServer
   Mail.addRecipient(ToEmail)
   Mail.From = From
   Mail.Subject = Subject
   Mail.body = Body
   On Error Resume Next
    Mail.Send
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
 
 Case "jmail"
  Set Mail = Server.CreateObject("Jmail.smtpmail")
   Mail.ServerAddress = EmailServer
   Mail.AddRecipient ToEmail
   Mail.Sender = From
   Mail.Subject = Subject
   Mail.body = Body
   Mail.priority = 3
   On Error Resume Next
    Mail.execute
	If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "jmail4"
  Set Mail = Server.CreateObject("Jmail.Message")
   'Mail.MailServerUserName = "myUserName"
   'Mail.MailServerPassword = "MyPassword"
   Mail.From = From
   Mail.FromName = Name
   Mail.AddRecipient ToEmail, ToName
   Mail.Subject = Subject
   Mail.Body = Body
   On Error Resume Next
	Mail.Send(EmailServer)
	If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "mdaemon"
  Set gMDUser = Server.CreateObject("MDUserCom.MDUser")
  mbDllLoaded = gMDUser.LoadUserDll

  If mbDllLoaded = False Then
   Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: Could not load MDUSER.DLL! Program will exit." & "<br />"
  Else
   Set gMDMessageInfo = Server.CreateObject("MDUserCom.MDMessageInfo")
	gMDUser.InitMessageInfo gMDMessageInfo
	gMDMessageInfo.To = ToEmail
	gMDMessageInfo.From = From
	gMDMessageInfo.Subject = Subject
	gMDMessageInfo.MessageBody = Body
	gMDMessageInfo.Priority = 0
    On Error Resume Next
     gMDUser.SpoolMessage gMDMessageInfo
     If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
	mbDllLoaded = gMDUser.FreeUserDll
   Set gMDMessageInfo = Nothing
  End If

 Case "ocxmail"
  Set Mail = Server.CreateObject("ASPMail.ASPMailCtrl.1")
   On Error Resume Next
    Result = Mail.SendMail(EmailServer, ToEmail, From, Subject, Body)
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "ocxqmail"
  Set Mail = Server.CreateObject("ocxQmail.ocxQmailCtrl.1")
   On Error Resume Next
    Mail.Q EmailServer, Name, From, "", "", ToEmail, "", "", "", Subject, Body
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "sasmtpmail"
  Set Mail = Server.CreateObject("SoftArtisans.SMTPMail")
   Mail.FromName = Name
   Mail.FromAddress = From
   Mail.AddRecipient ToName, ToEmail
   'Mail.AddReplyTo From
   Mail.BodyText = Body
   Mail.organization = SiteDescription
   Mail.Subject = Subject
   Mail.RemoteHost = EmailServer
   On Error Resume Next
    SendOk = Mail.SendMail
    If Not(SendOk) <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"

 Case "smtp"
  Set Mail = Server.CreateObject("SmtpMail.SmtpMail.1")
   Mail.MailServer = EmailServer
   Mail.Recipients = ToEmail
   Mail.Sender = From
   Mail.Subject = Subject
   Mail.Message = Body
   On Error Resume Next
    Mail.SendMail2
    If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"

 Case "vsemail"
  Set Mail = CreateObject("VSEmail.SMTPSendMail")
   Mail.Host = EmailServer
   Mail.From = From
   Mail.SendTo = ToEmail
   Mail.Subject = Subject
   Mail.Body = Body
   On Error Resume Next
	Mail.Connect
	Mail.Send
	Mail.Disconnect
	If Err <> 0 Then Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
End Select

Set Mail = Nothing
On Error GoTo 0
%>
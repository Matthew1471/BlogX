<%
select case lcase(EmailComponent) 
	case "abmailer"
		Set Mail = Server.CreateObject("ABMailer.Mailman")
		Mail.ServerAddr = EmailServer
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.SendTo = ToEmail
		Mail.MailSubject = Subject
		Mail.MailMessage = Body
		on error resume next '## Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "aspemail"
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.FromName = Name
		Mail.From = From
		Mail.AddReplyTo From
		Mail.Host = EmailServer
		Mail.AddAddress ToEmail, ToName
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '## Ignore Errors
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "aspmail"
		Set Mail = Server.CreateObject("SMTPsvg.Mailer")
                Mail.ContentType = "text/html"
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.ReplyTo = From
		Mail.RemoteHost = EmailServer
		Mail.AddRecipient ToName, ToEmail
		Mail.Subject = Subject
		Mail.BodyText = Body
		on error resume next '## Ignore Errors
		SendOk = Mail.SendMail
		If not(SendOk) <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"
		End if
	case "aspqmail"
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
		on error resume next '## Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "cdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		Mail.BodyFormat = 0
		Mail.MailFormat = 0
		on error resume next '## Ignore Errors
		Mail.Send From, ToEmail, Subject, Body
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "chilicdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		on error resume next '## Ignore Errors
		Mail.Host = EmailServer
		Mail.To = ToName & "<" & ToEmail & ">"
		Mail.From = Name & "<" & From & ">"
		Mail.Subject = Subject
		Mail.Body = Body
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "cdosys"
	        Set iConf = Server.CreateObject ("CDO.Configuration")
        	Set Flds = iConf.Fields 

	        'Set and update fields properties
        	Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	        Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailServer
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
        	Flds.Update

	        Set Mail = Server.CreateObject("CDO.Message")
        	Set Mail.Configuration = iConf

	        'Format and send message
        	Err.Clear 

		Mail.To = ToName & "<" & ToEmail & ">"
		Mail.From = Name & "<" & From & ">"
		Mail.Subject = Subject
		Mail.HTMLBody = Body
        	On Error Resume Next
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	 case "dkqmail"
		Set Mail = Server.CreateObject("dkQmail.Qmail")
		Mail.FromEmail = From
		Mail.ToEmail = ToEmail
		Mail.Subject = Subject
		Mail.Body = Body
		Mail.CC = ""
		Mail.MessageType = "TEXT"
		on error resume next '## Ignore Errors
		Mail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "dundasmailq"
		set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.QuickSend From, ToEmail, Subject, Body
		on error resume next '##Ignore Errors
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "dundasmails"
		set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.TOs.Add ToEmail
		Mail.FromAddress = From
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '##Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "geocel"
		set Mail = Server.CreateObject("Geocel.Mailer")
		Mail.AddServer EmailServer, 25
		Mail.AddRecipient ToEmail, ToName
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '##  Ignore Errors
		Mail.Send()
		if Err <> 0 then 
			Response.Write "Your request was not sent due to the following error: " & Err.Description 
		else
			Response.Write "Your mail has been sent..."
		end if
	case "iismail"
		Set Mail = Server.CreateObject("iismail.iismail.1")
		MailServer = EmailServer
		Mail.Server = EmailServer
		Mail.addRecipient(ToEmail)
		Mail.From = From
		Mail.Subject = Subject
		Mail.body = Body
		on error resume next '## Ignore Errors
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "jmail"
		Set Mail = Server.CreateObject("Jmail.smtpmail")
		Mail.ServerAddress = EmailServer
		Mail.AddRecipient ToEmail
		Mail.Sender = From
		Mail.Subject = Subject
		Mail.body = Body
		Mail.priority = 3
		on error resume next '## Ignore Errors
		Mail.execute
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "jmail4"
		Set Mail = Server.CreateObject("Jmail.Message")
		'Mail.MailServerUserName = "myUserName"
		'Mail.MailServerPassword = "MyPassword"
		Mail.From = From
		Mail.FromName = Name
		Mail.AddRecipient ToEmail, ToName
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '## Ignore Errors
		Mail.Send(EmailServer)
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "mdaemon"
		Set gMDUser = Server.CreateObject("MDUserCom.MDUser")
		mbDllLoaded = gMDUser.LoadUserDll
		if mbDllLoaded = False then
			response.write "Could not load MDUSER.DLL! Program will exit." & "<br />"
		else
			Set gMDMessageInfo = Server.CreateObject("MDUserCom.MDMessageInfo")
			gMDUser.InitMessageInfo gMDMessageInfo
			gMDMessageInfo.To = ToEmail
			gMDMessageInfo.From = From
			gMDMessageInfo.Subject = Subject
			gMDMessageInfo.MessageBody = Body
			gMDMessageInfo.Priority = 0
			gMDUser.SpoolMessage gMDMessageInfo
			mbDllLoaded = gMDUser.FreeUserDll
		end if
		if Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		end if
	case "ocxmail"
		Set Mail = Server.CreateObject("ASPMail.ASPMailCtrl.1")
		recipient = ToEmail
		sender = From
		subject = Subject
		message = Body
		mailserver = EmailServer
		on error resume next '## Ignore Errors
		result = Mail.SendMail(mailserver, recipient, sender, subject, message)
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "ocxqmail"
		Set Mail = Server.CreateObject("ocxQmail.ocxQmailCtrl.1")
		mailServer = EmailServer
		FromName = Name
		FromAddress = From
		priority = ""
		returnReceipt = ""
		ToAddressList = ToEmail
		ccAddressList = ""
		bccAddressList = ""
		attachmentList = ""
		messageSubject = Subject
		messageText = Body
		on error resume next '## Ignore Errors
		Mail.Q mailServer,      _
			fromName,      _
		        fromAddress,      _
		        priority,      _
		        returnReceipt,      _
		        ToAddressList,      _
		        ccAddressList,      _
		        bccAddressList,      _
		        attachmentList,      _
		        messageSubject,      _
		        messageText
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "sasmtpmail"
		Set Mail = Server.CreateObject("SoftArtisans.SMTPMail")
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.AddRecipient ToName, ToEmail
		'Mail.AddReplyTo From
		Mail.BodyText = Body
		Mail.organization = strForumTitle
		Mail.Subject = Subject
		Mail.RemoteHost = EmailServer
		on error resume next
		SendOk = Mail.SendMail
		If not(SendOk) <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"
		End if
	case "smtp"
		Set Mail = Server.CreateObject("SmtpMail.SmtpMail.1")
		Mail.MailServer = EmailServer
		Mail.Recipients = ToEmail
		Mail.Sender = From
		Mail.Subject = Subject
		Mail.Message = Body
		on error resume next '## Ignore Errors
		Mail.SendMail2
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "vsemail"
		Set Mail = CreateObject("VSEmail.SMTPSendMail")
		Mail.Host = EmailServer
		Mail.From = From
		Mail.SendTo = ToEmail
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '## Ignore Errors
		Mail.Connect
		Mail.Send
		Mail.Disconnect
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
end select

Set Mail = Nothing
%>
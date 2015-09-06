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
		On Error Resume Next '## Ignore Errors
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
		On Error Resume Next '## Ignore Errors
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
		On Error Resume Next '## Ignore Errors
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
		On Error Resume Next '## Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "cdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		Mail.BodyFormat = 0
		Mail.MailFormat = 0
		On Error Resume Next '## Ignore Errors
		Mail.Send From, ToEmail, Subject, Body
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "chilicdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		On Error Resume Next '## Ignore Errors
		Mail.Host = EmailServer
		Mail.To = ToName & "<" & ToEmail & ">"
		Mail.From = Name & "<" & From & ">"
		Mail.Subject = Subject
		Mail.Body = Body
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

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
		On Error Resume Next '## Ignore Errors
		Mail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "dundasmailq"
		Set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.QuickSend From, ToEmail, Subject, Body
		On Error Resume Next '##Ignore Errors
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "dundasmails"
		Set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.TOs.Add ToEmail
		Mail.FromAddress = From
		Mail.Subject = Subject
		Mail.Body = Body
		On Error Resume Next '##Ignore Errors
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
		On Error Resume Next '##  Ignore Errors
		Mail.Send()
		If Err <> 0 then 
			Response.Write "Your request was not sent due to the following error: " & Err.Description 
		Else
			Response.Write "Your mail has been sent..."
		End If

	case "iismail"
		Set Mail = Server.CreateObject("iismail.iismail.1")
		MailServer = EmailServer
		Mail.Server = EmailServer
		Mail.addRecipient(ToEmail)
		Mail.From = From
		Mail.Subject = Subject
		Mail.body = Body
		On Error Resume Next '## Ignore Errors
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
		On Error Resume Next '## Ignore Errors
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

		If mbDllLoaded = False then
			response.write "Could not load MDUSER.DLL! Program will exit." & "<br />"
		Else
			Set gMDMessageInfo = Server.CreateObject("MDUserCom.MDMessageInfo")
			gMDUser.InitMessageInfo gMDMessageInfo
			gMDMessageInfo.To = ToEmail
			gMDMessageInfo.From = From
			gMDMessageInfo.Subject = Subject
			gMDMessageInfo.MessageBody = Body
			gMDMessageInfo.Priority = 0
			gMDUser.SpoolMessage gMDMessageInfo
			mbDllLoaded = gMDUser.FreeUserDll
		End if

		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End If

	case "ocxmail"
		Set Mail = Server.CreateObject("ASPMail.ASPMailCtrl.1")
		On Error Resume Next '## Ignore Errors
		Result = Mail.SendMail(EmailServer, ToEmail, From, Subject, Body)
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "ocxqmail"
		Set Mail = Server.CreateObject("ocxQmail.ocxQmailCtrl.1")
		On Error Resume Next '## Ignore Errors
		Mail.Q EmailServer,      _
			Name,      _
		        From,      _
		        "",      _
		        "",      _
		        ToEmail,      _
		        "",      _
		        "",      _
		        "",      _
		        Subject,      _
		        Body
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
		Mail.organization = SiteDescription
		Mail.Subject = Subject
		Mail.RemoteHost = EmailServer
		On Error Resume Next
		SendOk = Mail.SendMail
		If Not(SendOk) <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"
		End if

	case "smtp"
		Set Mail = Server.CreateObject("SmtpMail.SmtpMail.1")
		Mail.MailServer = EmailServer
		Mail.Recipients = ToEmail
		Mail.Sender = From
		Mail.Subject = Subject
		Mail.Message = Body
		On Error Resume Next '## Ignore Errors
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
		On Error Resume Next '## Ignore Errors
		Mail.Connect
		Mail.Send
		Mail.Disconnect
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

End Select

Set Mail = Nothing
On Error Goto 0
%>
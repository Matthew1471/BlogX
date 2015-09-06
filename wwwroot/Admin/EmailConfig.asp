<%
' --------------------------------------------------------------------------
'¦Introduction : Edit E-mail Config Page.                                   ¦
'¦Purpose      : Allows the administrator to setup e-mail functionality.    ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp.                                       ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
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
AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<% If Request.Form("Action") <> "Post" Then
	Dim TheComponent(18)
	Dim TheComponentName(18)
	Dim TheComponentValue(18)

	'-- The components --'
	TheComponent(0) = "ABMailer.Mailman"
	TheComponent(1) = "Persits.MailSender"
	TheComponent(2) = "SMTPsvg.Mailer"
	TheComponent(3) = "SMTPsvg.Mailer"
	TheComponent(4) = "CDONTS.NewMail"
	TheComponent(5) = "CDONTS.NewMail"
	TheComponent(6) = "CDO.Message"
	TheComponent(7) = "dkQmail.Qmail"
	TheComponent(8) = "Dundas.Mailer"
	TheComponent(9) = "Dundas.Mailer"
	TheComponent(10) = "Geocel.Mailer"
	TheComponent(11) = "iismail.iismail.1"
	TheComponent(12) = "Jmail.smtpmail"
	TheComponent(13) = "MDUserCom.MDUser"
	TheComponent(14) = "ASPMail.ASPMailCtrl.1"
	TheComponent(15) = "ocxQmail.ocxQmailCtrl.1"
	TheComponent(16) = "SoftArtisans.SMTPMail"
	TheComponent(17) = "SmtpMail.SmtpMail.1"
	TheComponent(18) = "VSEmail.SMTPSendMail"

	'-- The name of the components --'
	TheComponentName(0) = "ABMailer v2.2+"
	TheComponentName(1) = "ASPEMail"
	TheComponentName(2) = "ASPMail"
	TheComponentName(3) = "ASPQMail"
	TheComponentName(4) = "CDONTS (IIS 3/4/5)"
	TheComponentName(5) = "Chili!Mail (Chili!Soft ASP)"
	TheComponentName(6) = "CDOSYS (IIS 5/5.1/6)"
	TheComponentName(7) = "dkQMail"
	TheComponentName(8) = "Dundas Mail (QuickSend)"
	TheComponentName(9) = "Dundas Mail (SendMail)"
	TheComponentName(10) = "GeoCel"
	TheComponentName(11) = "IISMail"
	TheComponentName(12) = "JMail"
	TheComponentName(13) = "MDaemon"
	TheComponentName(14) = "OCXMail"
	TheComponentName(15) = "OCXQMail"
	TheComponentName(16) = "SA-Smtp Mail"
	TheComponentName(17) = "SMTP"
	TheComponentName(18) = "VSEmail"

	'-- The config value of the components --'
	TheComponentValue(0) = "abmailer"
	TheComponentValue(1) = "aspemail"
	TheComponentValue(2) = "aspmail"
	TheComponentValue(3) = "aspqmail"
	TheComponentValue(4) = "cdonts"
	TheComponentValue(5) = "chilicdonts"
	TheComponentValue(6) = "cdosys"
	TheComponentValue(7) = "dkqmail"
	TheComponentValue(8) = "dundasmailq"
	TheComponentValue(9) = "dundasmails"
	TheComponentValue(10) = "geocel"
	TheComponentValue(11) = "iismail"
	TheComponentValue(12) = "jmail"
	TheComponentValue(13) = "mdaemon"
	TheComponentValue(14) = "ocxmail"
	TheComponentValue(15) = "ocxqmail"
	TheComponentValue(16) = "sasmtpmail"
	TheComponentValue(17) = "smtp"
	TheComponentValue(18) = "vsemail"
%>

 <form id="Config" method="post" action="EmailConfig.asp" onsubmit="return setVar()">
 <p>
  <input name="Action" type="hidden" value="Post"/>
 </p>
 <p class="config">
  Enable Email<span style="color:red">*</span> : <input name="EnableEmail" type="checkbox" value="True" onchange="return setVarChange()" <%If EnableEmail = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  E-mail Component<span style="color:red">*</span> :
  <select name="EmailComponent" onchange="return setVarChange()">
  <%
	Dim I, J
	J = 0
	For I = 0 to UBound(TheComponent)
	 If IsObjInstalled(TheComponent(i)) Then 
	  Response.Write "  <option value=""" & TheComponentValue(i) & """ "
       If TheComponentValue(i) = EmailComponent Then Response.Write "selected=""selected"""
      Response.Write ">" & TheComponentName(i) & "</option>" & vbCrlf
	 Else
	  J = J + 1
	 End If
	Next
	
	If J > UBound(TheComponent) Then Response.Write	"   <option value=""None"">No Compatible Component Found</option>" & vbCrlf
	
	Response.Write	"  </select>" %>
 </p>
 <p class="config">
  Email Server<span style="color:red">*</span> : <input name="EmailServer" type="text" style="width:90%;" value="<%=EmailServer%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  Email Address<span style="color:red">*</span> : <input name="EmailAddress" type="text" style="width:90%;" value="<%=EmailAddress%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p>
  <input type="submit" value="Save"/>
 </p>
 <p class="config" style="text-align:Center">
  <span style="color:red">*</span> - Indicates a required field.
 </p>
 </form>
<% Else

 '-- CheckBox check --'
 EnableEmail    = Request.Form("EnableEmail")
 If EnableEmail = "" Then EnableEmail = False

 '### Open The Records Ready To Write ###
 Records.Open "SELECT EnableEmail, EmailAddress, EmailComponent, EmailServer FROM Config", Database, 0, 2
  Records("EnableEmail") = EnableEmail
  Records("EmailAddress") = Request.Form("EmailAddress")
  Records("EmailComponent") = Request.Form("EmailComponent")
  Records("EmailServer") = Request.Form("EmailServer")
  Records.Update
 Records.Close

 Response.Write "<p style=""text-align:center"">E-mail Settings Update Successful.</p>"
 Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
<%
Response.End

Function IsObjInstalled(strClassString)
	On Error Resume Next

	'## Default Values
	IsObjInstalled = false
	Err = 0

	'## Test Each Component
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err then IsObjInstalled = true

	'## Destroy Object
	Set xTestObj = nothing
	Err = 0
	On Error GoTo 0
End Function
%>
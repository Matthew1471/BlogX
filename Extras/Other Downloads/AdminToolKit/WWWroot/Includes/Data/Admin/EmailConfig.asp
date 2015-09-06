<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<%
	Dim TheComponent(18)
	Dim TheComponentName(18)
	Dim TheComponentValue(18)

	'## the components
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

	'## the name of the components
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

	'## the value of the components
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
<Form Name="Config" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Enable Email<font color="Red">*</Font> : </span><input Name="EnableEmail" type="Checkbox" Value="True" onChange="return setVarChange()" <%If EnableEmail = True Then Response.Write "CHECKED"%>></P>
            <P><span id="Label1">E-mail Component<font color="Red">*</Font> : </span><select name="EmailComponent" onChange="return setVarChange()">
<%
	dim i, j
	j = 0
	for i=0 to UBound(TheComponent)
		if IsObjInstalled(TheComponent(i)) Then 
			Response.Write	"<option value=""" & TheComponentValue(i) & """ "
                        If TheComponentValue(i) = EmailComponent Then Response.Write "SELECTED"
                        Response.Write ">" & TheComponentName(i) & "</option>" & vbNewline
		Else
			j = j + 1
		End if
	Next
	If j > UBound(TheComponent) Then
		Response.Write	"<option value=""None"">No Compatible Component Found</option>" & vbNewline
	End If 

	Response.Write	"</select>"%></P>
            <P><span id="Label1">Email Server<font color="Red">*</Font> : </span><input Name="EmailServer" type="text" style="width:90%;" Value="<%=EmailServer%>" onChange="return setVarChange()"></P>
            <P><span id="Label1">Email Address<font color="Red">*</Font> : </span><input Name="EmailAddress" type="text" style="width:90%;" Value="<%=EmailAddress%>" onChange="return setVarChange()"></P>
            <P></P>
            <Input Type="submit" Value="Save">
            <P align="Center"><font color="Red">*</Font> - Indicates a required field.</P>
        </form>
<% Else

'### CheckBox Check! ###'
EnableEmail     = Request.Form("EnableEmail")
If EnableEmail     = "" Then EnableEmail = False

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Config", Database
Records("EnableEmail") = EnableEmail
Records("EmailAddress") = Request.Form("EmailAddress")
Records("EmailComponent") = Request.Form("EmailComponent")
Records("EmailServer") = Request.Form("EmailServer")
Records.Update

'#### Close Objects ###
Records.Close

Response.Write "<p align=""Center"">Email Settings Update Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
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
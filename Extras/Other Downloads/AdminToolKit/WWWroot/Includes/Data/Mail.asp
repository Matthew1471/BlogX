<% AlertBack = True %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Your Name : </span><input Name="Name" type="text" style="width:80%;" onChange="return setVarChange()"></P>
            <P><span id="Label1">Your E-mail : </span><input Name="Email" type="text" style="width:80%;" onChange="return setVarChange()"></P>
            <P><span id="Label1">Subject : </span><input Name="Subject" type="text" onChange="return setVarChange()" value="<%=HTML2Text(Request.Querystring())%>" style="width:80%;"></P>
            <P>Message : <textarea Name="Body" DESIGNTIMEDRAGDROP="96" onChange="return setVarChange()" style="height:10em;width:100%;"></textarea></P>
            <P></P>
            <Input Type="submit" Value="Send">
        </form>
<% Else
on Error Resume Next

Email = Request.Form("Email")

' ---------- Test if the form was properly filled in ----------
If InStr(Email,"@") > 0 Then
			Dim ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg

			ToName = "Webmaster"
			ToEmail = EmailAddress
			From = Request.Form("Email")
			Name = Request.Form("Name")
			Subject = "Blog : " & Request.Form("Subject")

MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>E-Mail From " & Name & "</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody align=""left"">" & VbCrlf

MailBody = MailBody & "<p>"" " & Replace(Request.Form("Body"),VbCrlf,"<br>" & VbCrlf) & " ""</p>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>From <a class=""standardsButton"" href=""http://ws.arin.net/cgi-bin/whois.pl?queryinput=" & Request.ServerVariables("REMOTE_ADDR") & """>" & Request.ServerVariables("REMOTE_ADDR") & "</a></p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			Body = MailBody
%>
			<!--#INCLUDE FILE="Includes/Mail.asp" -->
<%

    
      ' ---------- Mail Error Checking ----------
	If Err_Msg <> "" Then 
  		Response.Write Err_Msg
	Else
  	Response.Write "<p align=""Center"">Thank You, " & Name & ", Your Message Was Sent...</p>"
	End if

  Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"

ElseIf (EnableEmail <> True) Then
Response.Write "The Website Owner Has Disabled E-mail"
Else
Response.Write "The E-Mail Address You Specified Did Not Pass Validation"
End If

End If
%>
</DIV>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
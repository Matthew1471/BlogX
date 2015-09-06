<!-- #INCLUDE FILE="Includes/Header.asp" -->
<% Dim TheAt, TheDot, Mailbuddy, Subject, iConf, Message, flds, Mail%>
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Blog Support For <%=Domain%></font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
      <font color="#000444">

      <%If Request.Form("Action") <> "DoIt" then%>
      <BLOCKQUOTE><center>
      <font face=arial color="DarkRed" size="2">
      If You are e-mailing about a lost password, please provide the Blog creation date.<P>
      </font>

<form method="POST">
<input type="hidden" name="Action" value="DoIt">
<input type="hidden" name="To" value="<%If Request.Querystring("To") <> "" Then Response.Write Request.Querystring("To") Else Response.Write "Webmaster"%>">
  <table border="1" cellpadding="0" cellspacing="0" width="1%" height="205" bordercolordark="#000000" bordercolorlight="#000000">
    <tr>
      <td width="30%" bgcolor="#3366FF" height="1"><b>To:</b></td>
      <td width="70%" bgcolor="#666699" height="1"><b><%If Request.Querystring("To") <> "" Then Response.Write Request.Querystring("To") Else Response.Write "Webmaster"%></b></td>
    </tr>
    <tr>
      <td width="30%" bgcolor="#3366FF" height="21"><b>Your Name:</b></td>
      <td width="70%" bgcolor="#666699" height="21"><input type="text" name="FromName" size="20"></td>
    </tr>
    <tr>
      <td width="30%" bgcolor="#3366FF" height="21"><b>Your Address:</b></td>
      <td width="70%" bgcolor="#666699" height="21"><input type="text" name="Email" size="20"></td>
    </tr>
    <tr>
      <td width="17%" bgcolor="#3366FF" height="21"><b>Blog Username:</b></td>
      <td width="80%" bgcolor="#666699" height="21"><input type="text" name="Username" size="20" Value="<%=Request.Querystring("Subject")%>"></td>
    </tr>
    <tr>
      <td width="17%" bgcolor="#3366FF" height="21"><b>Blog Created:</b></td>
      <td width="80%" bgcolor="#666699" height="21"><input type="text" name="Date" size="10" Value="<%=Date()%>"></td>
    </tr>
    <tr>
      <td width="17%" bgcolor="#3366FF" height="21"><b>Subject : </b></td>
      <td width="80%" bgcolor="#666699" height="21">
      <select size="1" name="Subject">
    <option selected value=" ">Select One.....</option>
    <option>Delete A Contraversial Entry</option>
    <option>How To Use (How Do I...?)</option>
    <option>I've Forgotten My Password</option>
    <option>My Account Has Been "disabled"</option>
    <option>My Blog Is Erroring Or There's A Bug</option>
    <option>Other</option>
  </select>
      </td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#666699"colspan="2" height="169"><textarea rows="10" name="Text" cols="51"></textarea></td>
    </tr>
    <tr>
    <td colspan="2" bgcolor="#DAE9F5">
    <center><input type="Submit" Value="Send"><input type="Reset" value="Reset"></center></font>
    </td>
    </tr>
  </table><p>
<font face=arial color="DarkGreen" size="2">
Please be advised, that we will not tolerate abusive or malicious e-mail.<br>
By contacting a member of staff, you agree that you <b>Accept</b> our <a href="Terms.asp"><font color="darkred">Terms & Conditions</font></a>.<p>
On sending your e-mail we may disclose your electonic message to other members, This page will also log your "IP Address"
</font>
</center>
      </BLOCKQUOTE>
<%Else
' --- Test For @ ---'
function EmailField(fTestString) 
	TheAt = Instr(2, fTestString, "@")
	if TheAt = 0 then 
		EmailField = 0
	else
		TheDot = Instr(cint(TheAt) + 2, fTestString, ".")
		if TheDot = 0 then
			EmailField = 0
		else
			if cint(TheDot) + 1 > Len(fTestString) then
				EmailField = 0
			else
				EmailField = -1
			end if
		end if
	end if
end function

If EmailField(Request.Form("Email")) = 0 then 
Response.Write "<center>You Must Enter A <u>Valid</u> E-mail Address<br><br>"
Response.Write "<a href=""Support.asp"">Back</a></center>"
%>
	  </td>
          <%
WriteFooter
Response.End
End if

If Request.Form("Subject") <> "" Then
Subject = Request.Form("Subject")
Else
Subject = "(None)"
End If         

'--- End Of Test For @ ---'
			ToName = "Webmaster"
			ToEmail = EmailAddress
			From = Request.Form("Email")
			Name = Request.Form("FromName")

			Subject = "[BlogX Support] : " & Subject
                        Body = "<b>Name:</b> " & Request.Form ("FromName") & "<br>"
                        Body = Body & "<b>From Address:</b> " & Request.Form("Email") & "<p>"
                        Body = Body & "<b>Username :</b> " & Request.Form("Username") & "<br>"
                        Body = Body & "<b>Created :</b> " & Request.Form("Date") & "<p>"
                        Body = Body & "<b>Thinks:</b> " & Replace(Request.Form("Text"), vbcrlf, "<br>" & vbcrlf) & "<p>"
                        Body = Body & "<b>DebugInfo:</b> <a href=""http://ws.arin.net/cgi-bin/whois.pl?queryinput=" & Request.ServerVariables("REMOTE_ADDR") & """>" & Request.ServerVariables("REMOTE_ADDR") & "</a>"
%>
			<!--#INCLUDE FILE="Includes/Mail.asp" -->
<%
                If Err_Msg <> "" then
                Response.Write "<div align=""center""><acronym title=""" & Err_Msg & """>Mailing Failed</acronym>...<br>Please Contact the <a href=""mailto:Webmaster@matthew1471.co.uk"">Webmaster</a></div>"
                Else
                Response.Write "<div align=""center"">Thank you " & Request.Form("FromName") & ",<br>"
                Response.Write " Your E-mail has been succesfully sent.<br><br>"
                Response.Write "<a href=""Default.asp"">Back</a></div>"
                End If

End if
%>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
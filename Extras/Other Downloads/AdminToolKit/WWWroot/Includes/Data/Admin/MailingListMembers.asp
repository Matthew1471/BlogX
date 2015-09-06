<%
AlertBack = True
Server.ScriptTimeout = 6000 %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script language="JavaScript" type="text/javascript" src="../Includes/RTF.js"></script>
<DIV id=content>
<%
'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM MailingList ORDER BY EmailID;", Database

'### Set Them Up ###
Set SubscriberAddress = Records("SubscriberAddress")
Set PUK = Records("PUK")
Set IP = Records("SubscriberIP")
Set Active = Records("Active")

If Request.Form("Action") <> "Send" Then %>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Send">
            <P>Subject : <input Name="Title" type="text" style="width:80%;" maxlength="80" onChange="return setVarChange()"></P>

            <P>Content :<br>
            <table border="0" cellpadding="1" cellspacing="0" width="100%">
			<tr>
			<td bgcolor="<%=CalendarBackground%>" align="left">

		<% If UseImagesInEditor <> 0 Then %>
			<img src="<%=SiteURL%>Images/Editor/Bold.gif" title="Bold" onclick="boldThis()">
			<img src="<%=SiteURL%>Images/Editor/Italicize.gif" title="Italics" onclick="italicsThis()">
			<img src="<%=SiteURL%>Images/Editor/Underline.gif"  title="Underline" onclick="underlineThis()">
			<img src="<%=SiteURL%>Images/Editor/Strike.gif"title="CrossOut" onclick="crossThis()">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<img src="<%=SiteURL%>Images/Editor/Left.gif" title="Left" onclick="leftThis()">
			<img src="<%=SiteURL%>Images/Editor/Center.gif" title="Center" onclick="centerThis()">
			<img src="<%=SiteURL%>Images/Editor/Right.gif" title="Left" onclick="rightThis()">
			<img src="<%=SiteURL%>Images/Editor/Photo.gif" title="Style the image as a photo" onclick="photoThis()">
			</td>

			<td bgcolor="<%=CalendarBackground%>" align="right">
			<img src="<%=SiteURL%>Images/Editor/SpellCheck.gif" title="Spell Check" onclick="SpellThis()">
			<img src="<%=SiteURL%>Images/Editor/URL.gif" title="Link" onclick="linkThis()">
			<img src="<%=SiteURL%>Images/Editor/Image.gif" title="Image" onclick="imageThis('')">
			&nbsp;
			<img src="<%=SiteURL%>Images/Editor/Line.gif" title="Line" onclick="lineThis()">
		<% Else %>
			<input type="button" value="Bold" onclick="boldThis()">
			<input type="button" value="Italics" onclick="italicsThis()">
			<input type="button" value="Underline" onclick="underlineThis()">
			<input type="button" value="CrossOut" onclick="crossThis()">
			</td>

			<td bgcolor="<%=CalendarBackground%>" align="right">
			<input type="button" value="Link" onclick="linkThis()">
			<input type="button" value="Image" onclick="imageThis('')">
			&nbsp;
			<input type="button" value="Line" onclick="lineThis()">
		<% End If %>

			</td>
			</tr>

            <tr>
            <td colspan="2">
            <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:10em;width:100%;" onChange="return setVarChange()"></textarea>
            </tr>
			</table>
            </P>
            <P></P>
            <Input Type="submit" Value="Send">
        </form>

<table border="0">
<tr><td><b>Mailing List Members</td></tr>
<%
Count = 0
Do Until (Records.EOF)
Count = Count + 1
%>
<tr><td><Acronym Title="<%=IP%>"><a href="mailto:<%=SubscriberAddress%>"><%=SubscriberAddress%></a></Acronym>
<% If Active = True Then Response.Write "<A class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & SubscriberAddress & "&PUK=" & PUK & """>Unsubscribe</A>"%>
</td></tr>
<%
Records.MoveNext
Loop
%>
<tr><td><b>Total</b> : <%=Count%> Members</td></tr>
</table>
<% Else

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
Response.Write "<p align=""Center"">No Text Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
<%
Response.End
End If

Response.Write "<p align=""center"">"

on Error Resume Next

Do Until (Records.EOF)
Set SubscriberAddress = Records("SubscriberAddress")
Set PUK = Records("PUK")
Set IP = Records("SubscriberIP")
Set Active = Records("Active")

If Active = True Then
MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>" & Request.Form("Title") & "</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody align=""left"">" & VbCrlf

MailBody = MailBody & Replace(Request.Form("Content"), vbcrlf, "<br>" & vbcrlf) & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>To Unsubscribe click <a class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & SubscriberAddress & "&PUK=" & PUK & """>Unsubscribe</a></p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			ToName = "Member Of " & SiteDescription
			ToEmail = SubscriberAddress
			From = EmailAddress
			Name = SiteDescription

			Subject = "Blog : " & Request.Form("Title")
			Body = MailBody
%>
                        <!-- #INCLUDE FILE="../Includes/Mail.asp" -->
<%
Else
SubscriberAddress = ".. Disabled .."
End If

If Err_Msg <> "" Then
Response.Write Err_Msg & "<br>"
Else
Response.Write "Mail Sent To <b>" & SubscriberAddress & "</b> <br>"
End If

Records.MoveNext
Loop

Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If

'#### Close Objects ###	
Records.Close
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
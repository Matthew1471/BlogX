<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<DIV id=content>
<%
Dim EntryID, Requested
EntryID = Request.Querystring("Entry")
If (IsNumeric(EntryID) = False) OR (EntryID = "") Then EntryID = 0

Requested = Request.Querystring("Email")
Requested = Replace(Requested,"'","")

If (Request.Querystring("PUK") = "") OR (Request.Querystring("Confirm") = "") Then

'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM Comments WHERE (Email='" & Requested & "') AND (Subscribe = True) AND EntryID=" & EntryID, Database

If Records.EOF = False Then
Dim PUK
Set PUK = Records("PUK")
%>
<!--- Start Content For Found Address --->
<DIV class=entry>
<H3 class=entryTitle>Unsubscribe "<%=Requested%>" From Entry <%=EntryID%></H3>
<DIV class=entryBody>
<P align="Center">Are You Sure?</P>
<p align="Center"><a href="CommentNotify.asp?Entry=<%=EntryID%>&Email=<%=Requested%>&PUK=<%If Request.Querystring("PUK") <> "" Then Response.Write Request.Querystring("PUK") Else Response.Write "0"%>&Confirm=Yes">YES!</a>&nbsp;&nbsp;&nbsp;<a href="Default.asp">NO!</a></P>
</DIV>
</DIV>
<!--- End Content --->
<% Else %>
<!--- Start Content For Not Found Address --->
<DIV class=entry>
<H3 class=entryTitle>Address Not Found</H3>
<DIV class=entryBody>
<P align="Center">Unfortunatly we could not find "<b><%=Requested%></b>" in our subscribed members.</P>
<P align="Center">That address may have already been removed</P>
</DIV>
</DIV>
<!--- End Content --->
<% End If

Records.Close

Else

PUK = Request.Querystring("PUK")
PUK = Replace(PUK,"'","")

'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM Comments WHERE (Email='" & Requested & "') AND (PUK =" & PUK & ") AND EntryID=" & EntryID, Database, 1, 3

If Records.EOF = False Then
Records("Subscribe") = False
Records.Update

Dim MailBody, ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg

MailBody = MailBody & "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>Notification Of Unsubscription</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody>" & VbCrlf

MailBody = MailBody & "<p>You are recieving this e-mail as confirmation, that you have successfully unsubscribed to be notified of updates for entry " & EntryID & " on " & SiteDescription & ".</p>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			ToName = Records("Name")
			ToEmail = Requested
			From = EmailAddress
			Name = SiteDescription

			Subject = "Blog : UnSubscription"
			Body = MailBody
%>
<!--#INCLUDE FILE="Includes/Mail.asp" -->
<!--- Start Content For Accepted Address & PUK --->
<DIV class=entry>
<H3 class=entryTitle>Unsubscribed</H3>
<DIV class=entryBody>
<P align="Center"><%=Requested%> Has been successfully unsubscribed.</P>
</DIV>
</DIV>
<!--- End Content --->
<% Else %>
<!--- Start Content For Not Found Address --->
<DIV class=entry>
<H3 class=entryTitle>Invalid PUK</H3>
<DIV class=entryBody>
<P align="Center">Invalid security response number, Please click the <b>EXACT</b> link in your e-mail.</P>
</DIV>
</DIV>
<!--- End Content --->
<% End If

Records.Close

End If
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
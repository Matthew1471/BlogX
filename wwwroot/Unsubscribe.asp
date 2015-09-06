<%
' --------------------------------------------------------------------------
'¦Introduction : Mailing List Unsubscription Page                           ¦
'¦Purpose      : Allows the user to opt-out of the mailing list they        ¦
'¦               previously subscribed to.                                  ¦
'¦Used By      : User Mailing List E-mails                                  ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Mail.asp, Includes/ViewerPass.asp                 ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT 

PageTitle = "Unsubscribe"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<div id="content">
<%
Dim Requested
Requested = Request.Querystring("Email")
Requested = Replace(Requested,"'","")

If (Request.Querystring("PUK") = "") OR (Request.Querystring("Confirm") = "") Then

'### Open The Records Ready To Write ###
Records.Open "SELECT SubscriberAddress, SubscriberIP, PUK, Active FROM MailingList WHERE (SubscriberAddress='" & Requested & "') AND (Active = True)", Database

If Records.EOF = False Then
Dim PUK
Set PUK = Records("PUK")
%>
<!-- Start Content For Found Address -->
<div class="entry">
<h3 class="entryTitle">Unsubscribe "<%=Requested%>" From The Mailing List</h3>
 <div class="entryBody">
  <p style="text-align:Center">Are You Sure?</p>
  <p style="text-align:Center"><a href="Unsubscribe.asp?Email=<%=Requested%>&amp;PUK=<%If Request.Querystring("PUK") <> "" Then Response.Write Request.Querystring("PUK") Else Response.Write "0"%>&amp;Confirm=Yes">YES!</a>&nbsp;&nbsp;&nbsp;<a href="Default.asp">NO!</a></p>
 </div>
</div>
<!-- End Content -->
<% Else %>
<!-- Start Content For Not Found Address -->
<div class="entry">
<h3 class="entryTitle">Address Not Found</h3>
 <div class="entryBody">
  <p style="text-align:Center">Unfortunatly we could not find "<b><%=Requested%></b>" in our subscribed members.</p>
  <p style="text-align:Center">That address may have already been removed.</p>
 </div>
</div>
<!-- End Content -->
<% End If

Records.Close

Else

PUK = Request.Querystring("PUK")
PUK = Replace(PUK,"'","")

'### Open The Records Ready To Write ###
Records.Open "SELECT SubscriberAddress, SubscriberIP, PUK, Active FROM MailingList WHERE (SubscriberAddress='" & Requested & "') AND (PUK =" & PUK & ")", Database, 1, 3

If Records.EOF = False Then
Records("Active") = False
Records.Update

Dim MailBody

MailBody = MailBody & "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br/>" & VbCrlf
MailBody = MailBody & "<div class=""content"">" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<div class=""entry"" style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<h3 class=""entryTitle"">Notification Of Unsubscription</h3>" & VbCrlf
MailBody = MailBody & "<div class=""entryBody"">" & VbCrlf

MailBody = MailBody & "<p>You are receiving this e-mail as confirmation, that you have successfully unsubscribed to be notified of updates on " & SiteDescription & ".</p>" & VbCrlf
MailBody = MailBody & "</div>" & VbCrlf
MailBody = MailBody & "</div>" & VbCrlf

MailBody = MailBody & "<p>To Subscribe at a later date, click <a class=""standardsButton"" href=""" & SiteURL & "Mail.asp?Please%20Could%20You%20Resubscribe%20me" & Request.Form("EmailAddress") & """>Subscribe</a> and ask to be resubscribed.</p>" & VbCrlf

MailBody = MailBody & "</center>" & VbCrlf
MailBody = MailBody & "</div>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			Dim ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo

			ToName = "Member Of " & SiteDescription
			ToEmail = Records("SubscriberAddress")
			From = NoEmailAddress
			Name = SiteDescription

			Subject = "Blog : UnSubscription"
			Body = MailBody
%>
<!--#INCLUDE FILE="Includes/Mail.asp" -->
<!-- Start Content For Accepted Address & PUK -->
<div class="entry">
 <h3 class="entryTitle">Unsubscribed</h3>
  <div class="entryBody">
   <p style="text-align:Center"><%=Requested%> has been successfully unsubscribed.</p>
   <p style="text-align:Center">To resubscribe in future, please contact the Webmaster.</p>
  </div>
</div>
<!-- End Content -->
<% Else %>
<!-- Start Content For Not Found Address -->
<div class="entry">
 <h3 class="entryTitle">Invalid PUK</h3>
 <div class="entryBody">
  <p style="text-align:Center">Invalid security response number, Please click the <b>EXACT</b> link in your e-mail.</p>
 </div>
</div>
<!-- End Content -->
<% End If

Records.Close

End If
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
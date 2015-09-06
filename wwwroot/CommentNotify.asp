<%
' --------------------------------------------------------------------------
'¦Introduction : Comment Unsubscription Page                                ¦
'¦Purpose      : Allows the user to opt-out of subscribed comments          ¦
'¦               they previously subscribed to.                             ¦
'¦Used By      : User Comment Notification E-mails                          ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Mail.asp                                          ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

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
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<div id="content">
<%
Dim EntryID, Requested
EntryID = Request.Querystring("Entry")
If (IsNumeric(EntryID) = False) OR (EntryID = "") Then EntryID = 0

Requested = Request.Querystring("Email")
Requested = Replace(Requested,"'","")

If (Request.Querystring("PUK") = "") OR (Request.Querystring("Confirm") = "") Then

 '--- Open The Records Ready To Write ---'
 Records.Open "SELECT EntryID, Email, Subscribe, PUK FROM Comments WHERE (Email='" & Requested & "') AND (Subscribe = True) AND EntryID=" & EntryID, Database

  If Records.EOF = False Then
   Dim PUK
   Set PUK = Records("PUK")
   %>
    <!-- Start Content For Found Address -->
     <div class="entry">
      <h3 class="entryTitle">Unsubscribe "<%=Requested%>" From Entry <%=EntryID%></H3>
      <div class="entryBody">
       <p align="Center">Are You Sure?</p>
       <p align="Center"><a href="CommentNotify.asp?Entry=<%=EntryID%>&Email=<%=Requested%>&PUK=<%If Request.Querystring("PUK") <> "" Then Response.Write Request.Querystring("PUK") Else Response.Write "0"%>&Confirm=Yes">YES!</a>&nbsp;&nbsp;&nbsp;<a href="Default.asp">NO!</a></p>
      </div>
     </div>
    <!-- End Content -->
  <% ElseIf (Requested <> "") AND (EntryID <> 0) Then %>
    <!-- Start Content For Not Found Address -->
     <div class="entry">
      <h3 class="entryTitle">Address Not Found</h3>
      <div class="entryBody">
       <p style="text-align:Center">Unfortunatly we could not find "<b><%=Requested%></b>" in our subscribed members.</p>
       <p style="text-align:Center">That address may have already been removed.</p>
      </div>
     </div>
     <!-- End Content -->
  <% Else %>
    <!-- Start Content For Invalid Input -->
     <div class="entry">
      <h3 class="entryTitle">No Address Sent</h3>
      <div class="entryBody">
       <p style="text-align:Center">Please click the link in your e-mail or copy and paste it as some of the information is missing.</p>
       <p style="text-align:Center">I need both an EntryID and an e-mail address to unsubscribe you.</p>
      </div>
     </div>
    <!-- End Content -->
  <% End If

Records.Close

Else

 PUK = Request.Querystring("PUK")
 PUK = Replace(PUK,"'","")

 '### Open The Records Ready To Write ###
 Records.Open "SELECT EntryID, Name, Email, Subscribe, PUK FROM Comments WHERE (Email='" & Requested & "') AND (PUK =" & PUK & ") AND EntryID=" & EntryID, Database, 1, 3

 If Records.EOF = False Then

  Records("Subscribe") = False
  Records.Update

  If Session(CookieName) = False Then

   Dim MailBody, ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo

   MailBody = MailBody & "<html>" & VbCrlf
   MailBody = MailBody & "<head>" & VbCrlf
   MailBody = MailBody & "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet/>" & VbCrlf
   MailBody = MailBody & "</head>" & VbCrlf
   MailBody = MailBody & "<body bgcolor=""" & BackgroundColor & """>" & VbCrlf

   MailBody = MailBody & "<br/>" & VbCrlf
   MailBody = MailBody & "<div class=""content"">" & VbCrlf
   MailBody = MailBody & "<center>" & VbCrlf

   MailBody = MailBody & "<div class=""entry"" style=""width: 50%"">" & VbCrlf
   MailBody = MailBody & "<h3 class=""entryTitle"">Notification Of Unsubscription</H3>" & VbCrlf
   MailBody = MailBody & "<div class=""entryBody"">" & VbCrlf

   MailBody = MailBody & "<p>You are receiving this e-mail as confirmation, that you have successfully unsubscribed from being notified of updates for the entry " & EntryID & " on " & SiteDescription & ".</p>" & VbCrlf
   MailBody = MailBody & "</div>" & VbCrlf
   MailBody = MailBody & "</div>" & VbCrlf

   MailBody = MailBody & "</center>" & VbCrlf
   MailBody = MailBody & "</div>" & VbCrlf
   MailBody = MailBody & "</html>" & VbCrlf

    ToName = Records("Name")
    ToEmail = Requested
    From = NoEmailAddress
    Name = SiteDescription

    Subject = "Blog : UnSubscription"
    Body = MailBody
 %>
 <!-- #INCLUDE FILE="Includes/Mail.asp" -->
 <% End If %>
 <!-- Start Content For Accepted Address & PUK -->
  <div class="entry">
   <h3 class="entryTitle">Unsubscribed</h3>
   <div class="entryBody">
    <p style="text-align:Center"><%=Requested%> has been successfully unsubscribed.</p>
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
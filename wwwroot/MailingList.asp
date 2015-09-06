<%
' --------------------------------------------------------------------------
'¦Introduction : Mailing List Subscription Page                             ¦
'¦Purpose      : Allows the user to opt-in to the mailing list              ¦
'¦Used By      : Includes/Footer.asp                                        ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Mail.asp, Includes/ViewerPass.asp                 ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT 

'-- Proxy Handler --'
If (NOT DontSetModified) AND (Session(CookieName) = False) Then CacheHandle(GeneralModifiedDate)

PageTitle = "Mailing List"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">

<% If Request.Form("Action") <> "Post" Then %>
 <div class="entry">
 <h3 class="entryTitle">What Is The "<%=SiteName%>" Mailing List</h3><br/>

 <form method="post" action="MailingList.asp">
 <div class="entryBody">
  <% If InStr(1,Request.ServerVariables("HTTP_Host"),"blogx.co.uk",1) <> 0 Then%>
  <p>This mailing list will notify you the <b><span style="text-decoration:underline">MOMENT</span></b> there is any exciting news on Matthew1471's BlogX available for release.</p>
  <p style="text-align: center">Click <a href="Download.asp">here</a> to download BlogX.</p>

  <p style="text-align:center">This will keep up to date on the latest information, news and other information.</p>
  <p style="text-align:center; font-size: xx-small">Note : You are free to un-subscribe at any time, I will not disclose your e-mail address to any third parties.</p>
<% Else %>
  <!--- *************** CHANGE ME *********************************************************** -->
  <p style="text-align:center">
   If ever there was a time for the site owner to open up "MailingList.asp"<br/>
   in Notepad and tell people what this mailing list is actually about, now would be it
   <img src="images/emoticons/wink.gif">.
  </p>
  <!--- *************** Leave the rest alone ************************************************ -->
<%End If %>

  <input name="Action" type="hidden" value="Post"/>
  <p style="text-align:center"><b>Your Email Address : </b> <input name="EmailAddress" type="text" size="40" maxlength="40"/>
  <input type="submit" value="Subscribe"/></p>
 
  <br/>

  <p style="text-align:center; color:blue"><b>Important : You can only subscribe ONCE from ONE IP address<br/>"Test" e-mail subscriptions invalidate your ability to join in future!</b></p>
 </div>
 </form>

</div>

<% Else
 'Dimension variables
 Dim EntryCat 'Category
 Dim Email
 Dim Returned

 Email = Request.Form("EmailAddress")
 Email = Replace(Email,"'","")

 ' ---------- Did We Forget to type our address?  ----------
 If Email = "" Then
  Response.Write "<p style=""text-align:center"">No E-mail Address Entered</p>"
  Response.Write "<p style=""text-align:center""><a href=""javascript:history.back()"">Back</a></p>"
  Response.Write "</div>"
  %>
  <!-- #INCLUDE FILE="Includes/Footer.asp" -->
  <%
  Response.End
 End If

 ' ---------- Test if the form was properly filled in ----------
 If InStr(Email,"@") = 0 Then
  Response.Write "<p style=""text-align:center"">Invalid E-mail Address Entered</p>"
  Response.Write "<p style=""text-align:center""><a href=""javascript:history.back()"">Back</a></p>"
  Response.Write "</div>"
  %>
  <!-- #INCLUDE FILE="Includes/Footer.asp" -->
  <%
  Response.End
 End If

 ' ---------- Test if the user is AOL ----------
 If InStr(Email,"aol.com") <> 0 Then
  Response.Write "<p style=""text-align:center"">AOL bans mail servers with dynamic IPs (such as this), please use a different address.</p>"
  Response.Write "<p style=""text-align:center""><a href=""javascript:history.back()"">Back</a></p>"
  Response.Write "</Div>"
  %>
  <!-- #INCLUDE FILE="Includes/Footer.asp" -->
  <%
  Response.End
 End If

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT SubscriberAddress, SubscriberIP, PUK, Active FROM MailingList WHERE SubscriberAddress ='" & Email & "' OR SubscriberIP='" & Request.ServerVariables("REMOTE_ADDR") & "'", Database

 If Records.EOF Then
  Records.AddNew

  Records("SubscriberAddress") = Request.Form("EmailAddress")
  Records("SubscriberIP") = Request.ServerVariables("REMOTE_ADDR")
  Randomize
  Records("PUK") = Int(Rnd()*99999999)
  Records("Active") = True

  Records.Update
  Returned = "Mailing List Subscription Successful"

  Dim MailBody

  MailBody = "<html>" & VbCrlf
  MailBody = MailBody & "<head>" & VbCrlf
   If TemplateURL = "" Then
    MailBody = MailBody & "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
   Else
    MailBody = MailBody & "<link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
   End If

  MailBody = MailBody & "</head>" & VbCrlf
  MailBody = MailBody & "<body bgcolor=""" & BackgroundColor & """>" & VbCrlf

  MailBody = MailBody & "<br/>" & VbCrlf
  MailBody = MailBody & "<div class=""content"">" & VbCrlf
  MailBody = MailBody & "<center>" & VbCrlf

  MailBody = MailBody & "<DIV class=""entry"" style=""width: 50%"">" & VbCrlf
  MailBody = MailBody & "<H3 class=""entryTitle"">Notification Of Subscription</H3>" & VbCrlf
  MailBody = MailBody & "<div class=""entryBody"" style=""text-align:left"">" & VbCrlf

  If InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0 Then
   MailBody = MailBody & "<p>Welcome to the new BlogX mailing list designed to help collaborate BlogX'ers so together we can fix bugs & be kept up to date.</p>" & VbCrlf
   MailBody = MailBody & "<p>I guess the first thing you want to do is <a href=""http://blogx.co.uk/Download.asp"">Download BlogX</a>.</p>" & VbCrlf
  End If

  MailBody = MailBody & "<p>You are receiving this e-mail as you, or someone posing as you, have subscribed to be notified of updates on " & SiteDescription & ".</p>" & VbCrlf
  MailBody = MailBody & "</DIV>" & VbCrlf
  MailBody = MailBody & "</DIV>" & VbCrlf

  MailBody = MailBody & "<p>To Unsubscribe click <a class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & Request.Form("EmailAddress") & "&PUK=" & Records("PUK") & """>Unsubscribe</a></p>" & VbCrlf

  MailBody = MailBody & "</Center>" & VbCrlf
  MailBody = MailBody & "</DIV>" & VbCrlf
  MailBody = MailBody & "</html>" & VbCrlf

   Dim ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo
   ToName = "Member Of " & SiteDescription
   ToEmail = Request.Form("EmailAddress")
   From = NoEmailAddress
   Name = SiteDescription
   Subject = "Blog : Subscription"
   Body = MailBody
   %>
   <!--#INCLUDE FILE="Includes/Mail.asp" -->
   <%

 '---- Notify Webmaster ----'
 MailBody = "<html>" & VbCrlf
 MailBody = MailBody & "<head>" & VbCrlf
  If TemplateURL = "" Then
   MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
  Else
   MailBody = MailBody & "<Link href=""" & TemplateURL & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
  End If

 MailBody = MailBody & "</head>" & VbCrlf
 MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

 MailBody = MailBody & "<br/>" & VbCrlf
 MailBody = MailBody & "<div class=""content"">" & VbCrlf
 MailBody = MailBody & "<center>" & VbCrlf

 MailBody = MailBody & "<div class=""entry"" style=""width: 50%"">" & VbCrlf
 MailBody = MailBody & "<h3 class=""entryTitle"">Notification Of Subscription</h3>" & VbCrlf
 MailBody = MailBody & "<div class=""entryBody"">" & VbCrlf

 MailBody = MailBody & "<p>A new member <b>" & Request.Form("EmailAddress") & "</b> has subscribed to " & SiteDescription & "'s Mailing list.</p>" & VbCrlf
 MailBody = MailBody & "</div>" & VbCrlf
 MailBody = MailBody & "</div>" & VbCrlf

 MailBody = MailBody & "<p>To view all subscriptions click <a class=""standardsButton"" href=""" & SiteURL & "Admin/MailingListMembers.asp"">Admin</a></p>" & VbCrlf
 MailBody = MailBody & "</center>" & VbCrlf
 MailBody = MailBody & "</div>" & VbCrlf
 MailBody = MailBody & "</html>" & VbCrlf

  ToName = "Webmaster"
  ToEmail = EmailAddress
  From = Request.Form("EmailAddress")
  Name = SiteDescription
  Subject = "Blog : New Member"
  Body = MailBody
  %>
  <!--#INCLUDE FILE="Includes/Mail.asp" -->
  <%
Else
 Returned = "You (or someone using your IP) are already subscribed"
 Returned = Returned & VbCrlf & "<br/><br/>Please <a href=""Mail.asp"">contact</a> the webmaster if you have changed / or need to change your address."
End If

'#### Close Objects ###
Records.Close

Response.Write "<p style=""text-align:center"">" & Returned & "</p>"
Response.Write "<p style=""text-align:center""><a href=""" & PageName & """>Back</a></p>"

End If
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
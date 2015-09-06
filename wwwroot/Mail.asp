<% 
' --------------------------------------------------------------------------
'¦Introduction : Mail The Author Page                                       ¦
'¦Purpose      : Allows visitors to e-mail the blog author                  ¦
'¦Used By      : Comments.asp, Main.asp, Discalimer.asp, ViewItem.asp.      ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Replace.asp, Includes/Cache.asp, Mail_Validate.asp¦
'¦Notes        : This page is for contacting the blog author but requires   ¦
'¦               e-mail settings to be properly configured in the config.   ¦
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

'-- Proxy Handler --'
If (NOT DontSetModified) AND (Session(CookieName) = False) Then CacheHandle(GeneralModifiedDate)

PageTitle = "Contact Me"

AlertBack = True %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">
<%
 '-- Check For A Proxy --'
 Dim MyIPAddress
 If Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" Then 
  MyIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 Else
  MyIPAddress = Request.ServerVariables("REMOTE_ADDR")
 End If

 MyIPAddress = Replace(MyIPAddress,"'","")

 '-- Check If We Are Banned (Also check the proxy) --'
 Records.Open "SELECT IP, BannedDate FROM BannedIP WHERE IP='" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","") & "' OR IP='" & MyIPAddress & "';",Database, 1, 3
  Dim Banned
  If Records.EOF = False Then Banned = True
 Records.Close

 '-- Are they multi-spamming and already in our unvalidated mail? --'
 Database.Execute "DELETE FROM Mail_Unvalidated WHERE IP='" & MyIPAddress & "';"


If Banned Then
  Response.Write "<p align=""Center"">The website owner has banned you from e-mailing them.</p>"
  Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"
ElseIf (EnableEmail <> True) Then
 Response.Write "<p align=""Center"">The Website owner has disabled e-mail.</p>"
 Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"
ElseIf Request.Form("Action") <> "Post" Then %>
<form action="Mail.asp" method="post" onsubmit="return setVar()">
  <p>
   <input name="Action" type="hidden" value="Post"/>
   <input name="Refer" type="hidden" value="<%=Request.ServerVariables("HTTP_REFERER")%>"/>
   Your Name : <input name="Name" type="text" style="width:80%;" onchange="return setVarChange()"/>
  </p>
  <p>Your E-mail : <input name="Email" type="text" style="width:80%;" onchange="return setVarChange()"/></p>
  <p>Subject : <input name="Subject" type="text" onchange="return setVarChange()" value="<%=HTML2Text(Request.Querystring())%>" style="width:80%;"/></p>
  <p>Message : <textarea name="Body" onchange="return setVarChange()" rows="7" cols="112" style="height:10em;width:100%;"></textarea></p>
  <p><input type="submit" value="Send"/></p>
 </form>
<% Else
'On Error Resume Next
 Dim Email, MailBody
 Email = Request.Form("Email")

 '## Anti-HTML ###
 Dim Content
 Content = Request.Form("Body")
 Content = Replace(Content, "&","&amp;")
 Content = Replace(Content, "<","&lt;")
 Content = Replace(Content, ">","&gt;")

 Dim SenderName
 SenderName = Request.Form("Name")
 SenderName = Replace(SenderName, "<","&lt;")
 SenderName = Replace(SenderName, ">","&gt;")
 SenderName = Replace(SenderName, "&","&amp;")

' ---------- Test if the form was properly filled in ----------
If InStr(Email,"@") > 0 Then

 MailBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf

 MailBody = MailBody & "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">" & VbCrlf

 MailBody = MailBody & "<head>" & VbCrlf
 MailBody = MailBody & " <title>Message from " & SenderName & " (" & Email & ")" & "</title>" & VbCrlf

  If TemplateURL = "" Then
   MailBody = MailBody & " <link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
  Else
   MailBody = MailBody & " <link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
  End If

 MailBody = MailBody & "</head>" & VbCrlf
 MailBody = MailBody & "<body style=""background :" & BackgroundColor & """>" & VbCrlf & VbCrlf

 MailBody = MailBody & " <p>&nbsp;</p>" & VbCrlf & VbCrlf

 MailBody = MailBody & " <div class=""content"" style=""text-align: center"">" & VbCrlf & VbCrlf

 MailBody = MailBody & "  <div class=""entry"" style=""width: 50%; margin:0 auto"">" & VbCrlf
 MailBody = MailBody & "   <h3 class=""entryTitle"">E-Mail From " & SenderName & " (" & Email & ")</h3>" & VbCrlf
 MailBody = MailBody & "   <div class=""entryBody"" style=""align: left"">" & VbCrlf

 MailBody = MailBody & "    <p>"" " & Replace(Content,VbCrlf,"<br/>" & VbCrlf) & " ""</p>" & VbCrlf

 '-- Now the blog doesn't querystring all the mail links (it was confusing the spiders) it would be a good idea to include the referer --'
 If (Request.Form("Refer") <> "") Then MailBody = MailBody & "   <p>Contacted after reading : " & Request.Form("Refer") & "</p>" & VbCrlf

 MailBody = MailBody & "   </div>" & VbCrlf
 MailBody = MailBody & "  </div>" & VbCrlf & VbCrlf

 MailBody = MailBody & "  <p>" & VbCrlf
 MailBody = MailBody & "   From <a class=""standardsButton"" href=""http://whois.domaintools.com/" & MyIPAddress & """>" & MyIPAddress & "</a> " & VbCrlf
 MailBody = MailBody & "   <acronym title=""Ban User""><a href=""" & SiteURL & "Admin/EditBan.asp?Ban=" & MyIPAddress & """><img title=""Ban User"" alt=""Color Icon"" src=""" & SiteURL & "Images/Color.gif"" style=""border: none""/></a></acronym>" & VbCrlf

 If Request.ServerVariables("REMOTE_ADDR") <> MyIPAddress Then
  MailBody = MailBody & "<br/>" & VbCrlf
  MailBody = MailBody & " (Proxy : <a class=""standardsButton"" href=""http://whois.domaintools.com/" & Request.ServerVariables("REMOTE_ADDR") & """>" & Request.ServerVariables("REMOTE_ADDR") & "</a> " & VbCrlf
  MailBody = MailBody & "   <acronym title=""Ban User""><a href=""" & SiteURL & "Admin/EditBan.asp?Ban=" & Request.ServerVariables("REMOTE_ADDR") & """><img title=""Ban Proxy"" alt=""Color Icon"" src=""" & SiteURL & "Images/Color.gif"" style=""border: none""/></a></acronym>)" & VbCrlf
 End If

 MailBody = MailBody & "  </p>" & VbCrlf & VbCrlf

 MailBody = MailBody & " </div>" & VbCrlf & VbCrlf

 MailBody = MailBody & "</body>" & VbCrlf
 MailBody = MailBody & "</html>"

 Records.Open "SELECT RecordID, FromEmail, Subject, Body, IP, PUK FROM Mail_Unvalidated",Database, 1, 3

 '-- Add Record --'
 Records.AddNew

  Records("FromEmail") = Left(SenderName & "<" & Email & ">",255)
  Records("Subject")   = "Blog : " & Request.Form("Subject")
  Records("Body") = MailBody

  '-- Important --'
  Records("IP") = MyIPAddress

  Randomize
  Records("PUK") = Int(Rnd()*99999999)

 Records.Update


  '-- Mail Error Checking --'
  If Err <> 0 Then
   Response.Write "<p align=""Center"">Mailing Failed...<br>Please notify via alternative means.</p>"
   Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"
  Else
   Dim RedirectURL
   RedirectURL = "Mail_Validate.asp?RecordID=" & Records("RecordID") & "&PUK=" & Records("PUK") & "&Name=" & Request.Form("Name")

   '--- Close The Database ---'
   'Records.Close
   'Database.Close

   'Set Records = Nothing
   'Set Database = Nothing

   '-- This page performs IP validation --'
   'Response.Redirect RedirectURL

   Response.Write "<div style=""text-align:center; font-size: large; font-weight: bold""><img width=""32"" height=""32"" alt=""Hourglass Icon"" src=""Images/hourglass.gif"" style=""position: relative; bottom: -8px""/>Notice</div>"
   Response.Write "<div style=""text-align:center;""><ul><li>Your message has not yet been sent, please <a href=""" & RedirectURL & """>click here to submit your message</a>.</li></ul></div>" & VbCrlf

  End If

  Records.Close

Else
 Response.Write "<p align=""Center"">The e-mail address you specified did not pass validation.</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
End If

End If
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
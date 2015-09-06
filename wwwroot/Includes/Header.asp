<%
' --------------------------------------------------------------------------
'¦Introduction : Page Header.                                               ¦
'¦Purpose      : Loads the config and inserts the header to start a page.   ¦
'¦Requires     : Config.asp, Languages.asp.                                 ¦
'¦Used By      : Most pages.                                                ¦
'---------------------------------------------------------------------------

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
%><!-- #INCLUDE FILE="Config.asp" -->
<%
'-- Enter here the date your header and footer was modified, some relatively static pages will use this date. --'
GeneralModifiedDate = CDate("14/08/08 17:17:00")

'-- Did the user temporarily request legacy mode? --'
If Instr(Request.Querystring(),"LegacyMode") <> 0 Then LegacyMode = True

'-- Has legacy mode been turned on by an admin or user? --'
If LegacyMode = True Then

 '-- Set all the options to the default Chris Anderson BlogX; disabling any extras. --'
 ArgoSoftMailServer = 0
 BackgroundColor = "#FFFFFF"
 CalendarCheck = 1
 CommentNotify = 0
 NoticeText = ""
 NotifyPingOMatic = 0
 NoDate = 0
 MailingList = 0
 Polls = 0
 RSSImage = 0
 SortByDay = True
 Template = "Default"

End If

'-- This routine checks for the user disabling cookies. --'
If Len(Request.Form("Username")) = 0 AND Len(Request.Form("Password")) = 0 Then
 Session("CookieTest") = "AOK"
ElseIf Session("CookieTest") <> "AOK" Then
 CookiesDisabled = True
End If

'-- If the user is an administrative user then it is likely the page will change frequently, so do not send a modification date. --'
If Session(CookieName) = True Then DontSetModified = True

'## Security handling ##'

 '-- Check for an IP address ban on this user. --'
 Records.Open "SELECT BannedIP, LoginFailCount, LastLoginFail FROM BannedLoginIP WHERE BannedIP='" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","") & "'",Database, 0, 2

  '-- Do we have a record for this user? --'
  If Records.EOF = False Then

   '-- This address is in our list of banned addresses, have they made 3 attempts in less than 15 minutes. --' 
   If DateDiff("n",Records("LastLoginFail"),Now()) < 15 AND (Records("LoginFailCount") MOD 3 = 0) Then 
    Blacklisted = True

    '-- This page denying login is going to change in 15 minutes, so do not send a modification date/time. --'
    DontSetModified = True

   End If

  End If

 If BlackListed Then
  '-- We do nothing (apart from set a diagnostic HTML comment) --'
  Response.Write "<!-- Blacklisted due to " & DateDiff("n",Records("LastLoginFail"),Now()) & " minutes since ban and " & Records("LoginFailCount") & " (MOD 3 = " &  (Records("LoginFailCount") MOD 3) & ") attempts. -->" & VbCrlf

 '-- Is the user authenticated (or should they now be authenticated?) --'
 ElseIf (Ucase(Request.Form("Username")) = Ucase(AdminUsername)) AND (Ucase(Request.Form("Password")) = UCase(AdminPassword)) OR (Request.Cookies(CookieName) = "True") Then

  '-- Set our server side session cookie to indicate they are authenticated. --'
  Session(CookieName) = True

  '-- Did the user click the checkbox stating we should set a cookie? --'
  If Request.Form("Remember") = "True" then
   Response.Cookies(CookieName) = "True"
   Response.Cookies(CookieName).Expires = "July 31, 2024"
  End If

 '-- Did the user specify a username and password but by this point it is incorrect? --'
 ElseIf Len(Request.Form("Username")) <> 0 OR Len(Request.Form("Password")) <> 0 Then

   '-- If this user (IP address) has not failed to login before then create a record for them. --'
   If Records.EOF Then 
    Records.AddNew
    Records("BannedIP") = Request.ServerVariables("REMOTE_ADDR")
   End If

   '-- Set the last time they failed to login to the time now. --'
   Records("LastLoginFail") = Now()

   '-- Increment the number of times they failed to login. --'
   Records("LoginFailCount") = Records("LoginFailCount") + 1

   '-- Save this record --'
   Records.Update

   '-- At this point have they made 3 consecutive attempts? --'
   If Records("LoginFailCount") MOD 3 = 0 Then BlackListed = True

 End If

 '-- Close the recordset for the BannedLoginIP table --'
 Records.Close

'-- Non-SSL --'
If Request.ServerVariables("HTTPS") = "on" AND CookiesDisabled <> True Then
 sURL = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
 If Request.ServerVariables("QUERY_STRING") <> "" Then sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
 Response.Redirect sURL
End If

'## End of Security ##'
%>
<!-- #INCLUDE FILE="Languages.asp" -->
<%
'-- clients detect can Pingback URIs from headers.
Response.AddHeader "X-Pingback", SiteURL & "RSS/PingBack/Default.asp"

'## Referrer handling ##'
If Logging = True Then

 'Dimension variables
 Dim RSSRefer, ReferURL

 '-- Locate the user supplied referrer. --'
 If Len(Request.ServerVariables("HTTP_REFERER")) > 0 AND InStr(1,Request.ServerVariables("HTTP_REFERER"),SiteUrl,1) = 0 Then
  ReferURL = Left(Replace(Request.ServerVariables("HTTP_REFERER"),"'", "&#39;"),255)
 Else
  ReferURL = "(None)"
 End If

 '-- Are we a client on the LAN? --'
 If Instr(Request.ServerVariables("REMOTE_ADDR"),"192.168") > 0 Then ReferURL = "Local Address"

 '-- Did the user find our page via Google cache, so we are actually our own referrer? --'
 If Instr(Request.ServerVariables("HTTP_REFERER"),"cache:") > 0 Then ReferURL = "Cache"

 '-- Open the Refer records ready to write to them. --'
 'CursorType: can be one: adOpenForwardOnly (default), adOpenStatic, adOpenDynamic, adOpenKeyset
 'LockType: can be one of: adLockReadOnly (default), adLockOptimistic, adLockPessimistic, adLockBatchOptimistic

  '-- Logging sometimes hits concurrency issues, so it is best to resume. --'
  On Error Resume Next

  Records.Open "SELECT ReferHits, ReferURL FROM Refer WHERE ReferURL='" & ReferURL & "';", Database, , 3

  '-- Has this site already been seen before? --'
  If Not Records.EOF = True Then
   Records("ReferHits") = Int(Records("ReferHits")) + 1
  Else
   Records.AddNew
   Records("ReferURL") = ReferURL
   Records("ReferHits") = 1
  End If

  Records.Update

  '-- If for whatever reason the above command failed, cancel the update. --'
  Records.CancelUpdate

  Records.Close
  On Error Goto 0
 End If

'## End of referrer handling ##'

'## Calendar querystring handling ##'

szYearMonth = Request("YearMonth")
DataDay = Request("Day")

'-- Did the user request a specific year and month? --'
If Len(szYearMonth) = 0 Then
 SpecificRequest = False
 DataYear = Year(Now())
 DataMonth = Month(Now())
Else
 SpecificRequest = True
 DataYear = Left(szYearMonth,4)
 DataMonth = Right(szYearMonth,2)

  '-- Is our user trying to make up months of the year? --'
  If IsNumeric(DataMonth) = True Then 
   If (DataMonth > 12) OR (DataMonth < 1) Then DataMonth = 1
  End If

End If

'-- Did the user specify any parameters that were not actually numeric (often for malicious gain)? --'
If (IsNumeric(DataYear) <> True) OR (IsNumeric(DataMonth) <> True) OR (IsNumeric(DataDay) <> True) Then Response.Redirect("Hacker.asp")

'-- Handle NEXT and LAST month.--'
If Request("POS") = "NEXT" Then
 DataMonth = DataMonth + 1
ElseIf Request("POS") = "LAST" Then 
 DataMonth = DataMonth - 1
End If

'-- Handle spurious months. --'
If DataMonth = 0 Then
 DataMonth = 12
 DataYear = DataYear - 1
ElseIf DataMonth = 13 Then
 DataMonth = 1
 DataYear = DataYear + 1
End If

'## End of calendar querystring handling ##'

'-- Admin logout. --'
If Request.Querystring = "ClearCookie" Then 
 Session(CookieName) = False
 If Request.Cookies(CookieName) = "True" Then Response.Cookies(CookieName) = ""
End If

'-- Handle page title set request. --'
If Len(PageTitleEntryRequest) > 0 AND IsNumeric(PageTitleEntryRequest) Then
 PageTitleEntryRequest = Replace(PageTitleEntryRequest,"-","")
 PageTitleEntryRequest = Replace(PageTitleEntryRequest,",","")
 Records.Open "SELECT RecordID, Title FROM Data WHERE RecordID=" & PageTitleEntryRequest,Database, 0, 1
  If Records.EOF = False Then PageTitle = Records("Title") & PageTitle
 Records.Close
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%If Len(PageTitle) > 0 Then Response.Write SiteDescription & " - " & PageTitle Else Response.Write SiteDescription%></title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

 <!--
 //= - - - - - - - 
 // Copyright 2004-08, Matthew Roberts
 // Copyright 2003, Chris Anderson
 // 
 // Usage Of This Software Is Subject To The Terms Of The License
 //= - - - - - - -
 -->

 <meta property="og:title" content="<%If Len(PageTitle) > 0 Then Response.Write SiteDescription & " - " & Replace(PageTitle,"""","&quot;") Else Response.Write SiteDescription%>"/>
 <meta property="og:image" content="http://blogx.co.uk/Images/Articles/BlogX%28260606%29.gif"/>
 <meta property="og:site_name" content="<%=SiteDescription%>"/>

 <script type="text/javascript">
 <!-- Hide javascript so W3C doesn't choke on it
 <% If (Instr(Request.ServerVariables("HTTP_USER_AGENT"),"web2messenger") = 0) AND (Request.Querystring("Editor") <> "True") Then Response.Write "if (parent.frames.length) top.location.href= document.location;"%>

 function PrintPopup(url) {
  var winl = (screen.width-450)/2
  var wint = (screen.height-525)/2
  popupWin = window.open(url,'Print','toolbar=no,scrollbars=yes,resizable=yes,menubar=no,width=450,height=525,top=' + wint + ',left=' + winl);
 }

 <%
 If (PingBackPage = True) AND (Request.Form("Content") <> "") Then %>
  var winl = (screen.width-275)/2
  var wint = (screen.height-200)/2
  myWindow = window.open("PingBack.asp", "PingBack",'toolbar=no,statusbar=yes,location=no,scrollbars=no,resizable=yes,width=275,height=200,top=' + wint + ',left=' + winl);
 <%
 End If

 If (AlertBack = True) AND (Request.Form("Action") <> "Post") Then
 %> var bolIsSubmitted = true;

  //Inda : The onbeforeunload event is new to Mozilla (27/12/04); not everyone will have it.
  window.onbeforeunload = window_onbeforeunload;

  function window_onbeforeunload()
  {         
    if(!bolIsSubmitted) return "You've modified a textbox or checkbox but haven't saved your changes!";
  }

 function setVar(){
   bolIsSubmitted = true;
   return true;
 }

 function setVarChange(){
   alert
   bolIsSubmitted = false;
   return true;
 }
 <% End If %>

 // Done hiding -->
 </script>
<%
If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<!-- #INCLUDE FILE="../Templates/Config.asp" -->
<% If TemplateURL = "" Then
Response.Write "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>"
Else
Response.Write "<link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>"
End If %>

 <!-- This fixes our CSS not validating... Take THAT W3C! -->

 <!--[if IE]>
 <style type="text/css">
  /* A fake viewplot for IE information bar */
  #viewplot {
    overflow-x: hidden;
    overflow-y: auto;
    height: expression(this.parentNode.offsetHeight - this.offsetTop);
  }
 </style>
 <![endif]-->

 <link rel="Alternate" type="application/rss+xml" title="RSS" href="<%=SiteURL%>Rss/<%If (ReaderPassword <> "") AND Session("Reader") = True Then Response.Write "?" & ReaderPassword%>"/>
 <link rel="pingback" href="<%=SiteURL%>RSS/PingBack/Default.asp"/>

 <!-- This makes FireFox grab the page in advance -->
 <link rel="prefetch" href="/Main.asp"/>

</head>
<body style="background-color: <%=BackgroundColor%>">

<!-- The Info Bar -->
<div id="infobar" style="display: none;"><a href="http://blogx.co.uk/Official/SNS.asp" title="You may not have an SNS client installed">There was a problem with the SNS request. You might not have an SNS client installed. Learn more...</a><br/><br/></div>
<script src="<%=SiteURL%>Includes/InfoBar.js" type="text/javascript"></script>
<!-- End Of The Bar -->

<% If (Instr(lcase(DataFile), "wwwroot") > 0) AND Session(CookieName) = True AND (Instr(lcase(DataFile), "blogx.mdb") > 0) Then
     Response.Write "<!-- You are an admin, here is an admin alert -->" & VbCrlf
     Response.Write "<table border=""1"" width=""100%"" style=""background:red; margin-top: 10px; margin-bottom: 10px"">" & vbCrlf
     Response.Write "  <tr>" & vbCrlf
     Response.Write "   <td align=""center"" style=""color:white; size:2"">"
     Response.Write "     <b>WARNING:</b> The location of your access database may not be secure.<br/><br/>"
     Response.Write "     You should consider moving the database from <b>" & DataFile & "</b> to a directory not directly accessible via a URL" & VbCrlf
     Response.Write "     and/or renaming the database to another name." 
     Response.Write "     <br/><br/><i>(After moving or renaming your database, remember to change the DataFile setting in Includes/Config.asp.)</i>"
     Response.Write "   </td>" & VbCrlf
     Response.Write "  </tr>" & VbCrlf
     Response.Write "</table>"
     Response.Write "<!-- End of admin alert -->" & VbCrlf

   ElseIf (LegacyMode = False) AND Session(CookieName) = True AND (SpamAttempts > 0) then

    '-- This can change rapidly, so make sure we do not set any modified dates later on --'
    DontSetModified = True
     Response.Write "<!-- You are an admin, here is an admin alert -->" & VbCrlf
    Response.Write "<table border=""1"" width=""100%"" style=""background:red; margin-top: 10px; margin-bottom: 10px"">" & vbCrlf
    Response.Write "  <tr>" & vbCrlf
    Response.Write "   <td align=""center"" style=""color:white; size:2"">" & VbCrlf
    Response.Write "     <b>WARNING:</b> " & SpamAttempts & " spam attempts have been logged, <br/>Please" & VbCrlf
    Response.Write "     <a href=""" & SiteURL & "Admin/Purge.asp"">purge the comments database</a> of invalid comments." & VbCrlf
    Response.Write "   </td>" & VbCrlf
    Response.Write "  </tr>" & VbCrlf
    Response.Write "</table>" & VbCrlf
    Response.Write "<!-- End of admin alert -->" & VbCrlf

   End If

 If (NoticeText <> "") Then Response.Write "<div style=""text-align: center; background-color: Red; COLOR: White; font-size: small; font-weight: bold; MARGIN-BOTTOM: 0px;"">Notice : " & NoticeText & "</div>"
 %>

<!-- Header Begin -->
<div id="header">

 <div style="text-align: right;">
 <% If (LegacyMode = False) AND (PhotoMode <> True) Then 
     Response.Write "<a href=""" & SiteURL & "PhotoAlbum.asp"">Switch To ""Photo Album Mode""</a>"
    ElseIf (LegacyMode = False) Then
     Response.Write "<a href=""" & SiteURL & "Default.asp"">Switch Back To ""Blog Mode""</a>"
    End If
  %>
 </div>

 <h1 id="title"><a style="TEXT-DECORATION: none" href="<%=SiteURL%>"><%=SiteName%></a></h1>
 <p id="byline"><%=SiteDescription%></p>
 <p id="sideTitle"><span class="blogTitleSub"><%=SiteSubTitle%></span>
 <span class="blogTitleSubDisclaimer">Please read my <a href="<%=SiteURL%>Disclaimer.asp">disclaimer</a>.</span></p>
</div>
<!-- Header End -->


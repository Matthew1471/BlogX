<%
' --------------------------------------------------------------------------
'¦Introduction : Comment Page                                               ¦
'¦Purpose      : Shows all the comments we have on file for a particular    ¦
'¦               entry, with deletion and ban options for the admin.        ¦
'¦Used By      : Main.asp, ViewItem.asp, E-mail.                            ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp,¦
'¦		         Includes/Replace.asp, Includes/ViewerPass.asp,     ¦
'¦               Mail.asp, ProtectedEntry.asp, Comments_Validate.asp        ¦
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

'-- This page will alert if a box is changed and then the page is left --'
AlertBack = True

'-- We want to control when we flush to client.. this page could be long --'
Response.Buffer = True

'-- The header looks this up for the page title --'
Dim Requested '-- The EntryID of the entry the user wants to view --'
Requested = Request.Querystring("Entry")
If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) OR (Instr(Requested,"-") > 0) Then Requested = 0
PageTitleEntryRequest = Requested
PageTitle = " - Comments"

'-- Once a day has passed should we just clear all un-validated comments? --'
Dim AutoErase
AutoErase = True
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<%
If Len(Request.Querystring("Error")) <> 0 Then
 '-- We don't want HTML injects --'
 NoticeText = Replace(Request.Querystring("Error"),"<","&lt;")
 NoticeText = Replace(NoticeText,">","&gt;") 
End If
%>

<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">
<%           
Dim Ban       '-- The IP address of the user we want to ban --'
Dim DelRecNo  '-- The CommentID we want to delete --'
Dim EntryID   '-- The EntryID of the entry the user is commenting on --'

'-- Filter & Clean --'
EntryID = Request.Form("EntryID")
If (IsNumeric(EntryID) = False) OR (EntryID = "") Then EntryID = 0 Else EntryID = Int(EntryID)

Ban = Request.Querystring("Ban")
Ban = Replace(Ban,"'","")    

'-- Check for a proxy --'
Dim MyIPAddress
If (Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "") Then 
 MyIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
Else
 MyIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If

MyIPAddress = Replace(MyIPAddress,"'","")

'-- Check if we are banned (also check the proxy) --'
Records.Open "SELECT IP FROM BannedIP WHERE IP='" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","") & "' OR IP='" & MyIPAddress & "';",Database, 0, 1
 Dim Banned
 Banned = NOT Records.EOF
Records.Close

'-- Decide whether to dump our poster who has already posted or when comments are banned --'
If (EntryID <> 0) Then

'-- Open record set --'
    Records.Open "SELECT CommentID, IP FROM Comments WHERE EntryID="& EntryID & " ORDER BY CommentID DESC",Database, 2, 1

     Dim LastIP

     '-- Create an exception for ONE proxy on *MY* server --'
     If (MyIPAddress = "213.48.73.94") AND (InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0) Then

      '-- By clearing LastIP it won't match the current IP even if it is technically the same --' 
      LastIP = ""

     Else

      '-- What was the last IP? --'
      If Records.EOF = False Then LastIP = Records("IP")
      
       '-- Are they multi-spamming and already in our unvalidated comments? --'
       Const adExecuteNoRecords = &H00000080

       '-- Under high load this sometimes fails as OLEDB likes to cache results ergo "Record is deleted" --'
       On Error Resume Next
        Database.Execute "DELETE FROM Comments_Unvalidated WHERE EntryID="& EntryID & " AND IP='" & MyIPAddress & "';",,adExecuteNoRecords
       On Error GoTo 0

     End If

     '-- Are they trying to submit a HTML form even though the blog does not allow their comment (various reasons see above) --'
     If (EnableComments = False) OR (LastIP = MyIPAddress) Then
      Records.Close
      Database.Close
      Set Records = Nothing
      Set Database = Nothing
      Response.Clear
      Response.Redirect("Comments.asp?Entry=" & EntryID & "&Error=Either you have already posted a valid comment or entries are not enabled on this blog/entry.")
     End If
    
    Records.Close

End If

'***************** Show Entry ********************'

'-- Display entry if we're not posting data --'
If Request.Form("Action") <> "Post" Then

 '--- Open set ---'
 Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, StopComments, Enclosure, EntryPUK, LastModified FROM Data WHERE RecordID=" & Requested, Database, 0, 1

 If NOT Records.EOF Then

  Dim RecordID, Title, Text, Password, DayPosted, MonthPosted, YearPosted, TimePosted, Enclosure
  Dim EntryUTCTimeZoneOffset, LastModified, NewTime, StopComments

  '-- Entry security --'
  Dim EntryPUK
  EntryPUK = Records("EntryPUK")

  '--- Setup variables ---'
  RecordID = Records("RecordID")
  Title    = Records("Title")
  Text     = Records("Text")
  Category = Records("Category")
  Password = Records("Password")

  DayPosted   =  Records("Day")
  MonthPosted =  Records("Month")
  YearPosted  =  Records("Year")

  TimePosted  =  Records("Time")
  
  EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

  LastModified = Records("LastModified")

  StopComments = Records("StopComments")
  Enclosure    = Records("Enclosure")

  '-- Check this entry doesn't have a password --'
  If Len(Password) > 0 Then
   Text = "<form action=""ProtectedEntry.asp"" method=""get"" style=""text-align:center""><p>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """/>" & VbCrlf 
   Text = Text & "<img src=""Images/Key.gif"" alt=""Key icon""/> Password Protected Entry <br/>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""/> <input type=""submit"" name=""Submit"" value=""Submit""/>" & VbCrlf
   Text = Text & "</p></form>"
  End If

'-- Have we already set the LastModified header? --'
Dim SetLastModifiedHeader
If (NOT SetLastModifiedHeader) AND (NOT DontSetModified) AND (Session(CookieName) = False) Then

 '-- Not every post has been modified --'
 If IsNull(LastModified) Then LastModified = CDate(DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & TimePosted)

 '-- Proxy Handler --'
 CacheHandle(LastModified)

 'Sun, 12 Aug 2007 09:58:50 GMT
 'Response.Write "<!-- Page Last Modified.. " & PubDate & "-->"

 '-- We don't want to set it twice.. only once, records are descending remember! --'
 SetLastModifiedHeader = True

End If

Records.Close

'--- We're British, Let's 12Hour Clock Ourselves ---'
If TimeFormat <> False Then
 If Hour(TimePosted) > 12 Then 
  NewTime = Hour(TimePosted) - 12 & ":"
 Else
  NewTime = Hour(TimePosted) & ":"
 End If
 
 If Minute(TimePosted) < 10 Then
  NewTime = NewTime & "0" & Minute(TimePosted) 
 Else
  NewTime = NewTime & Minute(TimePosted)
 End If

 If (Hour(TimePosted) < 12) AND (Hour(TimePosted) <> 12) Then
  NewTime = NewTime & " AM"
 Else
  NewTime = NewTime & " PM"
 End If

Else
 If Hour(TimePosted) < 10 Then NewTime = "0"
 NewTime = NewTime & Hour(TimePosted) & ":"
 If Minute(TimePosted) < 10 Then NewTime = NewTime & "0"
 NewTime = NewTime & Minute(TimePosted)
End If

'************ OPTIONAL FLAG ********************
'* To automatically stop comments on entries   *
'* older than one week, Uncomment the next line*

'If DateDiff("d", MonthPosted & "/" & DayPosted & "/" & YearPosted,Now) > 7 Then StopComments = True
%>
<!-- Start ID Header -->
<div class="date" id="Record<%=YearPosted%>-<%=MonthPosted%>-<%=DayPosted%>">
<h2 class="dateHeader">Comments For Entry #<%=RecordID%></h2>
<!-- End ID Header -->

<!-- Start Content For (Entry #<%=RecordID%>) -->
<div class="entry">
<div class="entryIcon">
<h3 class="entryTitle"><%=Title%> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><img src=""Images/Edit.gif"" alt=""Edit Your Entry"" style=""border: none""/></a></acronym> "%>(<a href="RSS/Comments/?Entry=<%=Requested%>">Comments RSS</a>)</h3>
</div>
<div class="entryBody">
        <% 
        Response.Write LinkURLs(Replace(Text, vbcrlf, "<br/>" & vbcrlf))

        If (Enclosure <> "") AND (Len(Password) = 0) Then 
	      If Instr(Enclosure,"http://") = 0 Then Enclosure = "Sounds/" & Enclosure
	 %>
         <br/><br/><br/>

	 <!-- Start Podcast Object (http://www.skylab.ws/?p=116) -->
	 <object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="192" height="29" id="mp3player" align="middle">
	 <param name="allowScriptAccess" value="sameDomain" />
	 <param name="movie" value="Includes/mp3player.swf?id=1.2" />
	 <param name="quality" value="high" />
	 <param name="bgcolor" value="#ffffff" />
	 <param name=FlashVars value="zipURL=<%=Enclosure%>&songURL=<%=Enclosure%>">
	 <embed src="Includes/mp3player.swf?id=1.2" FlashVars="zipURL=<%=Enclosure%>&amp;songURL=<%=Enclosure%>" quality="high" bgcolor="#ffffff" width="192" height="29" name="mp3player" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
	</object>
	 <!-- End Podcast Object -->
	<br/><small><a href="<%=enclosure%>">Download this file</a></small>
	<% End If %>
</div>
<p class="entryFooter">
<acronym title="Printer Friendly Version"><a href="javascript:PrintPopup('Printer_Friendly.asp?Entry=<%=RecordID%>')"><img alt="Printer Friendly Version" src="Images/Print.gif" style="border: none"/></a></acronym> <% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""Email The Author"" src=""Images/Email.gif"" style=""border: none""/></a></acronym>"%>
<a class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime%></a> 
<% If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <span class=""categories"">#<a href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</a></span>"%></p></div>
<!-- End Content -->
<%
    '-- If we own this blog --'
    If Session(CookieName) = True Then

     '-- Check we're not trying to SQL exploit the delete --'
     DelRecNo = Request.Querystring("Delete")
     If (IsNumeric(DelRecNo) = False) OR (DelRecNo = "") Then DelRecNo = 0 Else DelRecNo = Int(DelRecNo)  

     '-- We are welcome to delete comments --'
     If DelRecNo <> 0 Then
    	Records.Open "SELECT EntryID, CommentID FROM Comments WHERE EntryID=" & Requested & " AND CommentID=" & DelRecNo & ";",Database, 0, 3
    	
    	If NOT Records.EOF Then 
    	 Database.Execute "DELETE FROM Comments WHERE CommentID=" & DelRecNo

    	 Records.Close

    	 '-- Update Comment Count --'    	
         Records.Open "SELECT RecordID, Comments, LastModified FROM Data WHERE RecordID=" & Requested,Database
          Records("Comments") = Records("Comments") - 1
	      Records("LastModified") = Now()
    	  Records.Update
    	End If
    	
    	Records.Close
    	Database.Close
	Set Records = Nothing
    	Set Database = Nothing

        Response.Clear
        Response.Redirect("Comments.asp?Entry=" & Requested) 

     End If

     '-- We are welcome to ban people :D --'
     On Error Resume Next
      If Ban <> "" Then Database.Execute "INSERT INTO BannedIP (IP) VALUES ('" & Ban & "')" & ";"
     On Error Goto 0

    End If

'***************** Entry Comments ********************'
Records.Open "SELECT CommentID, EntryID, Name, Email, Homepage, Content, CommentedDate, IP, Subscribe FROM Comments WHERE EntryID=" & Requested & " ORDER BY CommentID ASC;",Database, 0, 1

' Loop through records until it's a next page or End of Records
Dim CommentID, Email, CommentedDate, CommentedTime, NewDate
Dim AlreadySubscribed

Do Until (Records.EOF)

 '--- Setup Variables ---'
 Set CommentID = Records("CommentID")
 Set Name      = Records("Name")
 Set Email     = Records("Email")
 Set Homepage  =  Records("Homepage")

 Content       = LinkURLs(Replace(Records("Content"), vbcrlf, "<br/>" & vbcrlf))
 
 '-- This will turn the proxy HTML comment into a ban icon --'
 If (Session(CookieName) = True) Then

  Dim ProxyPosition 
  ProxyPosition = Instr(Content,"<!-- Proxy Servers ORIGINAL Address : ")

  If ProxyPosition <> 0 Then
   Dim ProxyAddress
   ProxyAddress = Mid(Content,ProxyPosition+38)
   ProxyAddress = Left(ProxyAddress,Len(ProxyAddress)-3)
   'Response.Write "<!-- Debug " & ProxyPosition & ": """ & ProxyAddress & """ -->"
  Else
   ProxyAddress = ""
  End If

 End If

 Set Subscribe = Records("Subscribe")
   
 If (Subscribe = True) AND (Records("IP") = MyIPAddress) Then AlreadySubscribed = True
 
 Set CommentedDate = Records("CommentedDate")
 CommentedTime = FormatDateTime(CommentedDate,vbLongTime)

 '--- We're British, Let's 12Hour Clock Ourselves ---'
 NewTime = ""

 If TimeFormat <> False Then
  If Hour(CommentedTime) > 12 Then 
   NewTime = Hour(CommentedTime) - 12 & ":"
  Else
   NewTime = Hour(CommentedTime) & ":"
  End If
 
  If Minute(CommentedTime) < 10 Then
   NewTime = NewTime & "0" & Minute(CommentedTime)
  Else
   NewTime = NewTime & Minute(CommentedTime)
  End If

  If (Hour(CommentedTime) < 12) AND (Hour(CommentedTime) <> 12) Then
   NewTime = NewTime & " AM"
  Else
   NewTime = NewTime & " PM"
  End If

 Else

  If Hour(CommentedTime) < 10 Then NewTime = "0"
  NewTime = NewTime & Hour(CommentedTime) & ":"
  If Minute(CommentedTime) < 10 Then NewTime = NewTime & "0"
  NewTime = NewTime & Minute(CommentedTime)

 End If

 NewDate = Day(CommentedDate) & "/" & Month(CommentedDate) & "/" & Year(CommentedDate)
%>
<!-- Start Content For Comment <%=CommentID%> -->
<div class="comment">
 <a id="Comment<%=CommentID%>"></a>
 <h3 class="commentTitle"><%
  If (Session(CookieName) = True) Then

   Response.Write " <acronym title=""Users Using This IP""><a href=""#"" onclick=""javascript:PrintPopup('Admin/IPWhois.asp?IP=" & Records("IP") & "');""><img alt=""Users Using This IP"" src=""Images/Emoticons/Profile.gif"" style=""border: none""/></a></acronym> "
  If ProxyAddress <> "" Then Response.Write "<acronym title=""List User's Proxy Information""><a href=""http://whois.domaintools.com/" & Records("IP") & """><img alt=""List User's Proxy Information"" src=""Images/Print.gif"" style=""border: none""/></a></acronym> "
   Response.Write "<acronym title=""Ban User""><a "
   If ProxyAddress <> "" Then Response.Write "onclick=""return confirm('Are you *sure* you want to ban this user?\n\nThis user was behind a proxy. Check the address is creditable before banning.');"" "
   Response.Write "href=""Comments.asp?Entry=" & Requested & "&amp;Ban=" & Records("IP") & "#Comment" & CommentID & """><img title=""Ban User"" alt=""Color Icon"" src=""Images/Color.gif"" style=""border: none""/></a></acronym> "
   If ProxyAddress <> "" Then Response.Write "<acronym title=""Ban User's Proxy""><a href=""Comments.asp?Entry=" & Requested & "&amp;Ban=" & ProxyAddress & "#Comment" & CommentID & """ onclick=""return confirm('Are you *sure* you want to ban all users behind this proxy?\n\nProxies are sometimes used by big internet providers, other times they can be created by spammers.');""><img title=""Ban User's Proxy"" alt=""Color Icon"" src=""Images/Color.gif"" style=""border: none""/></a></acronym> "
   Response.Write "<acronym title=""Delete Comment""><a onclick=""return confirm('Are you *sure* you want to delete this comment?');"" href=""Comments.asp?Entry=" & Requested & "&amp;Delete=" & CommentID & """><img title=""Delete Comment"" alt=""Key Icon"" src=""Images/Key.gif"" style=""border: none""/></a></acronym>"
  End If
  %><%=NewDate%>&nbsp;<%=NewTime%></h3>
 <span class="commentBody"><%=Content%></span>
 <p class="commentFooter"><%If HomePage <> "" Then Response.Write "<a class=""permalink"" rel=""nofollow"" href=""" & HTML2Text(Homepage) & """>"%><%=HTML2Text(Name)%><%If HomePage <> "" Then Response.Write "</a>"%>
 <%If (Session(CookieName) = True) AND (Email <> "") Then
    Response.Write " | "
    If Records("Subscribe") = True Then Response.Write "<img alt=""Notification Enabled Icon"" title=""This user recieves e-mail notifications"" src=""Images/Emoticons/Profile.gif"" style=""border: none""/>" Else Response.Write "<img alt=""Notification Disabled Icon"" title=""This user recieves no e-mail notifications"" src=""Images/Emoticons/Post.gif"" style=""border: none""/>"
    Response.Write "<span class=""comments""><a href=""mailto:" & HTML2Text(Email) & """>" & Email & "</a></span>"
   End If
   %></p>
</div>
<!-- End Content -->
<%

 '-- When we have reached the last comment, this will be the LastIP --'
 LastIP = Records("IP")

 '-- Create an exception for ONE proxy on *MY* server --'
 If (MyIPAddress = "213.48.73.94") AND (InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0) Then LastIP = ""

 Records.MoveNext

Loop      

'--- Close The Records & Database ---
Records.Close
%>
</div>
<%

'***************** Entry Pingbacks ********************'

'-- Check If We Have Been Pinged Back --'
Records.Open "SELECT SourceURI FROM Pingback WHERE EntryID=" & Requested & ";",Database, 0, 1
If Records.EOF = False Then
%>
<!-- Start Content For PingBacks -->
<div class="comment">
<h3 class="commentTitle">Pingbacks For Entry #<%=Requested%></h3>
<div class="commentBody">
  <ul>
  <%
  Do Until (Records.EOF)
   Response.Write "  <li><a href=""" & Records("SourceURI") & """>" & Records("SourceURI") & "</a></li>" & VbCrlf
   Records.MoveNext
  Loop
  %>
  </ul>
</div>

</div>
<!-- End Content -->
<% 
    End If 
%>

<div id="AddNew" class="date">
                    <div class="comment">
                    <h3 class="commentTitle">Add New Comment</h3>
                    <div class="commentBody">
<%
If Banned = True Then
 Response.Write "<div style=""text-align:center""><b>You have been banned from making comments!</b></div>" & VbCrlf
ElseIf EnableComments <> True Then
 Response.Write "<div style=""text-align:center""><b>Comments have been disabled by the Blog administrator!</b></div>" & VbCrlf
ElseIf StopComments <> False Then
 Response.Write "<div style=""text-align:center""><b>Comments for this entry have been locked by the Blog administrator!</b></div>" & VbCrlf
ElseIf LastIP <> MyIPAddress Then 

 If LegacyMode <> True Then %>
<script type="text/javascript">
<!-- Hide javascript so W3C doesn't choke on it
function openPopup() {

 var left = (screen.width/2)-(600/2);
 var top = (screen.height/2)-(400/2);

 TheNewWin = window.open('','name','height=400,width=600, toolbar=no,directories=no,status=yes,menubar=no,top='+top+',left='+left+',scrollbars=yes,resizable=yes');
 TheNewWin.document.open;

 TheNewWin.document.write('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">\r');
 TheNewWin.document.write('<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">\r');
 TheNewWin.document.write('<head>\r');
 TheNewWin.document.write(' <title><%=Replace(SiteDescription,"'","\'")%> - Comment Preview<\/title>\r');
 TheNewWin.document.write(' <meta http-equiv="Content-Type" content="text\/html; charset=utf-8"\/>\r');

 TheNewWin.document.write(' <!-' + '-\r');
 TheNewWin.document.write(' \/\/= - - - - - - - \r');
 TheNewWin.document.write(' \/\/ Copyright 2004, Matthew Roberts\r');
 TheNewWin.document.write(' \/\/ Copyright 2003, Chris Anderson\r');
 TheNewWin.document.write(' \/\/ \r');
 TheNewWin.document.write(' \/\/ Usage Of This Software Is Subject To The Terms Of The License\r');
 TheNewWin.document.write(' \/\/= - - - - - - -\r');
 TheNewWin.document.write(' -' + '->\r');

 <% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")
 If TemplateURL = "" Then
  Response.Write " TheNewWin.document.write(' <link href=""" & SiteURL & "Templates\/" & Template & "\/Blogx.css"" type=""text\/css"" rel=""stylesheet""/>\r');"
 Else
  Response.Write " TheNewWin.document.write(' <link href=""" & TemplateURL & Template & "\/Blogx.css"" type=""text/css"" rel=""stylesheet""/>\r');"
 End If %>

 TheNewWin.document.write('<\/head>\r');
 TheNewWin.document.write('<body style="background-color:<%=BackgroundColor%>">\r');

 TheNewWin.document.write(' <p style="text-align:center; color:red">The following is a preview of your comment, that has yet to be added.<br\/>\r');
 TheNewWin.document.write(' Links and emoticons are not processed yet.<\/p><hr\/>\r\r');

 TheNewWin.document.write(' <div class="comment" style="vertical-align: middle;">\r')
 TheNewWin.document.write('  <h3 class="commentTitle">Add New Comment<\/h3>\r\r')
 TheNewWin.document.write('  <div class="commentBody">\r')                 

 // Escape HTML like submitted comments will do.
 var textPreview = document.forms['AddComment'].Content.value.replace(new RegExp ('<', 'gi'), '&lt;');
 
 // Convert new lines to HTML new lines.
 textPreview = textPreview.replace(new RegExp ('\n', 'gi'), '<br\/>\r   ');
 
 TheNewWin.document.write('   ' + textPreview + '\r');

 TheNewWin.document.write('  <\/div>\r\r');
 TheNewWin.document.write('  <hr\/>\r\r');
 TheNewWin.document.write('  <p style="text-align:center"><a href="#" onclick="self.close();return false;">Close Window<\/a><\/p>\r');
 TheNewWin.document.write(' <\/div>\r');

 TheNewWin.document.write('<\/body>\r');
 TheNewWin.document.write('<\/html>');
 
 TheNewWin.document.close();
 return false;
} // -->
</script>
<% End If %>
                        <form id="AddComment" method="post" action="Comments.asp" onsubmit="return setVar()">
                        <p>
                         <input name="Action" type="hidden" value="Post"/>  
                         <input name="EntryID" type="hidden" value="<%=RecordID%>"/>

			 <!-- Now With MORE Security -->
			 <input name="EntryPUK" type="hidden" value="<%=EntryPUK%>"/>
			 <!-- End of MORE Security -->
                        
                            Name<input name="Name" type="text" value="<%=Request.Cookies("Visitor")("Name")%>" maxlength="50"/></p>
                            <p>E-mail<input name="Email" type="text" value="<%=Request.Cookies("Visitor")("Email")%>" maxlength="50"/></p>
                            <p>Homepage<input name="Homepage" type="text" value="<%=Request.Cookies("Visitor")("Homepage")%>" maxlength="50"/></p>
                            <p><input name="RememberMe" type="checkbox" value="True" checked="checked"/>Remember Me
<% If (AlreadySubscribed = False) AND (LegacyMode <> True) AND (Session(CookieName) <> True) Then%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="Subscribe" type="checkbox" value="True" checked="checked"/>E-mail me replies
<% End If %></p>
                            <p>Content (HTML not allowed)</p>
                            <p><textarea name="Content" rows="12" cols="40" onchange="return setVarChange()"></textarea></p>
                            <p><input type="submit" value="Add Comment" accesskey="s"/><% If (LegacyMode <> True) Then Response.Write "&nbsp;<input onclick=""openPopup();"" type=""button"" value=""Preview Comment"" accesskey=""p""/>"%></p>
                        </form>
<%
Else
 Response.Write "<div style=""text-align:center""><b>You were already the last person to comment!</b></div>" & VbCrlf
End If
%>
                    </div>
                </div>
</div>

<%
Else
 '--- We're British, Let's 12Hour Clock Ourselves ---'
 If TimeFormat <> False Then

  If Hour(Time()) > 12 Then 
   NewTime = Hour(Time()) - 12 & ":"
  Else
   NewTime = Hour(Time()) & ":"
  End If
 
  If Minute(Time()) < 10 Then
   NewTime = NewTime & "0" & Minute(Time())
  Else
   NewTime = NewTime & Minute(Time())
  End If

  If (Hour(Time()) < 12) AND (Hour(Time()) <> 12) Then
   NewTime = NewTime & " AM"
  Else
   NewTime = NewTime & " PM"
  End If

 Else

  If Hour(Time()) < 10 Then NewTime = "0"
  NewTime = NewTime & Hour(Time()) & ":"
  If Minute(Time()) < 10 Then NewTime = NewTime & "0"
  NewTime = NewTime & Minute(Time())

 End If
%>
<!-- Start EOF Content -->
<div class="entry">
<h3 class="entryTitle">Error</h3>
<div class="entryBody"><p style="text-align:center"><%
 If Requested = 0 Then
  If Len(Request.Querystring("Entry")) > 0 Then
   Response.Write "Sorry, the record number you specified in the address was not a number."
  Else
   Response.Write "You did not specify a record number of an entry you wished to view, please check the link again."
  End If
 Else
  Response.Write "Sorry, the record number you requested is either invalid or has been removed."
 End If %></p>
<p style="text-align:center"><a href="<%=PageName%>">Back To The Main Page</a></p>
</div>
<p class="entryFooter"><%=NewTime%> 
| <span class="comments"><a href="Mail.asp?Whatever%20happened%20to%20record%20<%=Requested%>?">Report Error</a></span> 
| <span class="categories">#Error</span></p></div>
<!-- End EOF Content -->
<%End If

Records.Close

Else

'Dimension variables
Dim Subscribe
Subscribe = Request.Form("Subscribe")
If (Subscribe = "") OR Request.Form("Email") = EmailAddress Then Subscribe = False Else Subscribe = True

'### Did We Type In Name? ###'
If Request.Form("Name") = "" Then
 Response.Write "<p align=""Center"">No name entered</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
 Response.Write "</div>"
 %>
 <!-- #INCLUDE FILE="Includes/Footer.asp" -->
 <%
 Response.End
End If

'### Are they banned? ###'
If Banned Then
 Response.Write "<p align=""Center"">Nice try, but you are banned!</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
 %>
 <!-- #INCLUDE FILE="Includes/Footer.asp" -->
 <%
 Response.End
End If

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
 Response.Write "<p align=""Center"">No text entered.</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
 Response.Write "</div>"
 %>
 <!-- #INCLUDE FILE="Includes/Footer.asp" -->
 <%
 Response.End
End If

'### Did We Hack Hackity Hack? ###'
If EntryID = 0 Then
 Response.Write "<p align=""Center"">No hacking.</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
 Response.Write "</div>"
 %>
 <!-- #INCLUDE FILE="Includes/Footer.asp" -->
 <%
 Response.End
End If

 If Request.Form("RememberMe") = "True" then
  Response.Cookies("Visitor")("Name") = Left(Request.Form("Name"),50)
  Response.Cookies("Visitor")("Email") = Left(Request.Form("Email"),50)
  Response.Cookies("Visitor")("Homepage") = Left(Request.Form("Homepage"),80)
  Response.Cookies("Visitor").Expires = "July 31, 2012"
 End If

Randomize Timer

'-- Anti-HTML --'
Dim Content
Content = Request.Form("Content")
Content = Replace(Content, "&","&amp;")
Content = Replace(Content, "<","&lt;")
Content = Replace(Content, ">","&gt;")

Name = Request.Form("Name")
Name = Replace(Name, "<","&lt;")
Name = Replace(Name, ">","&gt;")
Name = Replace(Name, "&","&amp;")

'-- Add a http:// to non http'd links --'
Dim Homepage
Homepage = Request.Form("Homepage")
If (Instr(Homepage,"http://") = 0) AND (Len(Homepage) > 0) Then Homepage = "http://" & Homepage

'-- Perform an EntrySecurity test --'
Dim RecievedEntryPUK
RecievedEntryPUK = Request.Form("EntryPUK")
If IsNumeric(RecievedEntryPUK) Then RecievedEntryPUK = Int(RecievedEntryPUK) Else RecievedEntryPUK = 0

Records.Open "SELECT RecordID, EntryPUK FROM Data Where RecordID=" & EntryID & " AND EntryPUK=" & RecievedEntryPUK & " OR EntryPUK IS NULL", Database
If Records.EOF Then
 Response.Clear
 Response.Write "Not Authorised -- Entry Security Not Matched!..<br/><br/>If this was recieved in error, <a href=""Mail.asp"">e-mail the webmaster</a>"
 Records.Close
 Database.Close
 Set Records = Nothing
 Set Database = Nothing
 Response.End
End If
Records.Close

'-- Write In Comments --'
Records.Open "SELECT CommentID, EntryID, Name, Email, Homepage, Content, Subscribe, PUK, CommentedDate, UTCTimeZoneOffset, IP FROM Comments_Unvalidated", Database, 0, 3

On Error Resume Next
Records.AddNew

  '-- Record locking problems --'
  If Err.Number = -2147217887 Then
      
   '-- Keep trying for 3 seconds --'
   Dim EndTime
   EndTime = DateAdd("s", 3, Now())
   Do While (Now() < EndTime)
    Err.Clear
    Records.AddNew
	If Err.Number = 0 Then Exit Do
   Loop
  End If
  
  Dim WasError  
  If Err.Number <> 0 Then WasError = True 
 On Error GoTo 0

'-- Force it again so we get our server error page if needs be --'
If WasError Then Records.AddNew

Records("EntryID") = EntryID
Records("Name") = Left(Name,50)
Records("Email") = Left(Request.Form("Email"),50)
Records("Homepage") = Left(Homepage,50)

'## This will automatically add the REMOTE_ADDR just incase the "Forwarded For" Address Is Spoofed ##'
If MyIPAddress <> Request.ServerVariables("REMOTE_ADDR") Then
 Records("Content") = Content & " <!-- Proxy Servers ORIGINAL Address : " & Request.ServerVariables("REMOTE_ADDR") & "-->"
Else
 Records("Content") = Content
End If

Records("Subscribe") = Subscribe
Records("PUK") = Int(Rnd()*99999999)

Records("CommentedDate") = DateAdd("h",ServerTimeOffset,Now())

    '-- Work out Time Offset --'
    Dim Hours
    Hours = Abs(Int(UTCTimeZoneOffset / 60))
    If Hours < 10 Then Hours = "0" & Hours

    Dim Minutes
    Minutes = Abs(UTCTimeZoneOffset Mod 60)
    If Minutes < 10 Then Minutes = "0" & Minutes

    Dim OffsetTime
    If UTCTimeZoneOffset > 0 Then OffsetTime = "+" Else OffsetTime = "-"
    OffsetTime = OffsetTime & Hours & Minutes

Records("UTCTimeZoneOffset") = OffsetTime

Records("IP") = MyIPAddress

Records.Update
 
CommentID = Records("CommentID")

Dim PUK
PUK = Records("PUK")

Records.Close

'-- Dump any old spam comments --'

 '-- Auto Erasing Code (Removed old AND Date <> #" & DateValue(DateAdd("h",ServerTimeOffset,Now())) & "#" as it stopped working) --'
 If (AutoErase = True) Then Database.Execute "DELETE FROM Comments_Unvalidated WHERE DateDiff(""d"",CommentedDate,""" & DateValue(DateAdd("h",ServerTimeOffset,Now())) & """) > 0"

 '-- Clever SPAM Comment Validation --'
 Dim RedirectURL
 RedirectURL = "Comments_Validate.asp?CommentID=" & CommentID & "&PUK=" & PUK

 '--- Close The Database ---'
 'Database.Close
 'Set Records = Nothing
 'Set Database = Nothing

 '-- This page performs IP validation --'
 'Response.Redirect RedirectURL

 Response.Write "<div style=""text-align:center; font-size: large; font-weight: bold""><img width=""32"" height=""32"" alt=""Hourglass Icon"" src=""Images/hourglass.gif"" style=""position: relative; bottom: -8px""/>Notice</div>"
 Response.Write "<div style=""text-align:center;""><ul><li>Your comment has not yet been posted, please <a href=""" & RedirectURL & """>click here to submit your comment</a>.</li></ul></div>" & VbCrlf

End If
%>

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
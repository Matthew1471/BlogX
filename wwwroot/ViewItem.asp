<%
' --------------------------------------------------------------------------
'¦Introduction : View Item.                                                 ¦
'¦Purpose      : Views an individual blog entry item.                       ¦
'¦               Can be used as a "permalink".                              ¦
'¦Used By      : Main.asp, Comments.asp, ProtectedEntry.asp, ViewCat.asp.   ¦
'¦Requires     : Includes/Replace.asp, Includes/Header.asp,                 ¦
'¦               Includes/ViewerPass.asp, Includes/NAV.asp,                 ¦
'¦               Includes/Footer.asp.                                       ¦
'¦Notes        : This page also checks for entry passwords and SQL exploits.¦
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

Dim Requested
Requested = Request.Querystring("Entry")
If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) OR (Instr(Requested,"-") > 0) Then Requested = 0

'-- The header looks this up for the page title --'
PageTitleEntryRequest = Requested
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">
<%
'--- Open set ---'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, Enclosure, LastModified FROM Data WHERE RecordID=" & Requested,Database, 0, 1

If NOT Records.EOF Then

 '--- Setup Variables ---'
 Dim RecordID, Title, Text, Password, CommentsCount
 Dim DayPosted, MonthPosted, YearPosted, TimePosted
 Dim NewTime, Enclosure, LastModified

 RecordID = Records("RecordID")
 Title = Records("Title")
 Text = Records("Text")
 Category =  Records("Category")
 Password =  Records("Password")
 CommentsCount = Records("Comments")

 DayPosted =  Records("Day")
 MonthPosted =  Records("Month")
 YearPosted =  Records("Year")
 TimePosted =  Records("Time")

 Enclosure = Records("Enclosure")

 LastModified = Records("LastModified")

 If Len(Password) > 0 Then
  Text = "<form action=""ProtectedEntry.asp"" method=""get"" style=""text-align: center""><p>" & VbCrlf
  Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """/>" & VbCrlf
  Text = Text & "<img alt=""Key Icon"" src=""Images/Key.gif""/> Password Protected Entry <br/>" & VbCrlf
  Text = Text & "This post is password protected. To view it please enter your password below:"
  Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""/> <input type=""submit"" name=""Submit"" value=""Submit""/>" & VbCrlf
  Text = Text & "</p></form>"
 End If

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
%>
<!--- Start ID Header -->
<div class="date" id="D<%=YearPosted%>-<%=MonthPosted%>-<%=DayPosted%>">
<h2 class="dateHeader">Permanant Link For Entry #<%=RecordID%></h2>
<!--- End ID Header -->

<!--- Start Content For -->
<div class="entry">
<div class="entryIcon">
<h3 class="entryTitle"><%=Title%> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><img alt=""Edit Your Entry"" src=""Images/Edit.gif"" style=""border: none;""/></a></acronym>"%></h3>
</div>
<div class="entryBody"><%
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
	 <embed src="Includes/mp3player.swf?id=1.2" FlashVars="zipURL=<%=Enclosure%>&songURL=<%=Enclosure%>" quality="high" bgcolor="#ffffff" width="192" height="29" name="mp3player" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
	</object>
	 <!-- End Podcast Object -->
	<br/><small><a href="<%=enclosure%>">Download this file</a></small>
	<% End If %></div>
<p class="entryFooter">
<acronym title="Printer Friendly Version"><a href="javascript:PrintPopup('Printer_Friendly.asp?Entry=<%=RecordID%>')"><img alt="Printer Friendly Version" src="Images/Print.gif" style="border: none;"/></a></acronym> <% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""Email the Author"" src=""Images/Email.gif"" style=""border: none;""/></a></acronym>"%>
<b><%=DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime%></b>
<% If EnableComments <> False Then Response.Write " | <span class=""comments""><a href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</a></span>"%>
<% If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <span class=""categories"">#<a href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</a></span>"%></p></div>
<!--- End Content -->
</div>
<% Else
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
<!--- Start EOF Content -->
<div class="entry">
 <h3 class="entryTitle">Error</h3>
 <div class="entryBody">
  <p style="text-align:center">Sorry, the record number you requested was either invalid or has been removed.</p>
  <p style="text-align:center"><a href="<%=PageName%>">Back To The Main Page</a></p>
 </div>
 <p class="entryFooter"><%=NewTime%> 
 | <span class="comments"><a href="Mail.asp?Whatever%20happened%20to%20record%20<%=Requested%>?">Report Error</a></span> 
 | <span class="categories">#Error</span></p>
</div>
<!--- End EOF Content -->
<% End If

'--- Close The Records ---
Records.Close

'-- Have we already set the LastModified header? --'
If (NOT DontSetModified) AND (Session(CookieName) = False) Then

 '-- Not every post has been modified --'
 If IsNull(LastModified) Then LastModified = CDate(DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & TimePosted)

 '-- Proxy Handler --'
 CacheHandle(LastModified)

 'Sun, 12 Aug 2007 09:58:50 GMT
 'Response.Write "<!-- Page Last Modified.. " & PubDate & "-->"

End If
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
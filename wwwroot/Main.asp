<%
' --------------------------------------------------------------------------
'¦Introduction : Main Blog Page.                                            ¦
'¦Purpose      : Lists multiple blog entries and provides navigation.       ¦
'¦Used By      : Default.asp, IIS.                                          ¦
'¦Requires     : Includes/Replace.asp, Includes/Header.asp,                 ¦
'¦               Includes/ViewerPass.asp, Includes/Cache.asp                ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
'¦Standards    : XHTML.                                                     ¦
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

'-- The following setting will allow you to keep your main pages shorter with links to the full --'
ParagraphTruncate = False
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<%

Dim nPage, nRecCount, nPageCount, PreviousDay

'--- Open set ---'
If SpecificRequest = False Then
 Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Comments, Enclosure, LastModified FROM Data ORDER BY RecordID DESC",Database, 1, 1
Else
 Dim DaySelected
 If DataDay <> "" Then DaySelected = "AND Day=" & DataDay
 Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Comments, Enclosure, LastModified FROM Data WHERE Month=" & DataMonth & " AND Year=" & DataYear & " " & DaySelected & " ORDER BY RecordID DESC;",Database, 1, 1
End If

'-- Get Records Count --'
nRecCount = Records.RecordCount

'-- Tell recordset to split records in the pages of our size --'
Records.PageSize = EntriesPerPage

'-- How many pages we have collected --'
nPageCount = Records.PageCount

'-- Let us see what page are we looking at right now --'
nPage = Request.QueryString("Page")

'-- Make sure that the Page parameter passed to us is within the range --'
If (IsNumeric(nPage) = False) OR (nPage = "") Then 
 nPage = 1
Else
 nPage = Int(nPage)
 If (nPage < 1) Or (nPage > nPageCount) Then nPage = 1
End If

'****************************************************************

Response.Write "<div id=""content"">" & VbCrlf

'-- Time to tell user what we've got so far --'
Response.Write "<p style=""text-align:Right"">Page : " & nPage & "/" & nPageCount & "</p>"

'-- Give user some navigation --'

'-- First page --'
Response.Write "<p style=""text-align: center"">"

If nPage > 1 Then Response.Write "<a href=""" & PageName 
If (nPage > 1) AND (szYearMonth <> "") Then Response.Write "?YearMonth=" & DataYear & Right("00" & DataMonth, 2)
If nPage > 1 Then  Response.Write """>" Else Response.Write "<span style=""color:gray"">"
Response.Write "First Page"
If nPage > 1 Then Response.Write "</a>" Else Response.Write "</span>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

'-- Previous Page --'
If nPage > 1 Then Response.Write "<a href=""" & PageName & "?Page=" & nPage - 1
If (nPage > 1) AND (szYearMonth <> "") Then Response.Write "&amp;YearMonth=" & DataYear & Right("00" & DataMonth, 2)
If nPage > 1 Then  Response.Write """>" Else Response.Write "<span style=""color:gray"">"
Response.Write "Prev. Page"
If nPage > 1 Then Response.Write "</a>" Else Response.Write "</span>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
	
'-- Next Page --'
If nPage < nPageCount Then Response.Write "<a href=""" & PageName & "?Page=" & nPage + 1
If (nPage < nPageCount) AND (szYearMonth <> "") Then Response.Write "&amp;YearMonth=" & DataYear & Right("00" & DataMonth, 2)
If nPage < nPageCount Then  Response.Write """>" Else Response.Write "<span style=""color:gray"">"
Response.Write "Next Page"
If nPage < nPageCount Then Response.Write "</a>" Else Response.Write "</span>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

'-- Last Page --'
If nPage < nPageCount Then Response.Write "<a href=""" & PageName & "?Page=" & nPageCount 
If (nPage < nPageCount) AND (szYearMonth <> "") Then Response.Write "&amp;YearMonth=" & DataYear & Right("00" & DataMonth, 2)
If nPage < nPageCount Then  Response.Write """>" Else Response.Write "<span style=""color:gray"">"
Response.Write "Last Page"
If nPage < nPageCount Then Response.Write "</a>" Else Response.Write "</span>"
Response.Write "</p><br/>" & VbCrlf

'-- Position recordset to the page we want to see --'
If nRecCount > 0 Then Records.AbsolutePage = nPage

'--- Setup Day Posted ---'
PreviousDay = "0"
				
' Loop through records until it's a next page or End of Records
Dim RecordID, Title, Text, Password, DayPosted, MonthPosted, YearPosted, TimePosted, EntryUTCTimeZoneOffset, CommentsCount, Enclosure
Dim LastModified, NewTime, JustDoIt

Do Until (Records.EOF or Records.AbsolutePage <> nPage )

'--- Setup Variables ---'
   Set RecordID = Records("RecordID")
   Set Title = Records("Title")
   Set Text = Records("Text")
   Set Category =  Records("Category")
   Set Password =  Records("Password")

   Set DayPosted =  Records("Day")
   Set MonthPosted =  Records("Month")
   Set YearPosted =  Records("Year")
   Set TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

   Set CommentsCount = Records("Comments")

   Set Enclosure = Records("Enclosure")

   Set LastModified = Records("LastModified")

   If Len(Password) > 0 Then
   Text = "<form action=""ProtectedEntry.asp"" method=""get"" style=""text-align: center""><p>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """/>" & VbCrlf 
   Text = Text & "<img alt=""Login Key"" src=""Images/Key.gif""/> Password Protected Entry <br/>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""/> <input type=""submit"" name=""Submit"" value=""Submit""/>" & VbCrlf
   Text = Text & "</p></form>"
   End If

'--- We're British, Let's 12Hour Clock Ourselves ---'
NewTime = ""

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

If (DayPosted <> PreviousDay) AND (NoDate <> 1) Then

 Dim EntryWeekDay
 EntryWeekDay = WeekdayName(Weekday(MonthName(MonthPosted) & " " & DayPosted & ", " & YearPosted))

 'Friday, 6th August 2004
 Response.Write vbcrlf & "<!-- Start Date Header -->" & vbcrlf
 Response.Write "<div class=""date"" id=""Records" & YearPosted & "-" & MonthPosted & "-" & DayPosted & """>" & vbcrlf
 Response.Write "<h2 class=""dateHeader"">" & EntryWeekDay & ", " & DayPosted & " " & Left(MonthName(MonthPosted),3) & " " & YearPosted & "</h2>" & vbcrlf
 Response.Write "<!-- End Date Header -->" & vbcrlf

 JustDoIt = True
Else
 JustDoIt = False
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

 Dim ParagraphTruncate
 If (ParagraphTruncate = True) Then Text = ShortenEntry(Text,RecordID)
%>
<!-- Start Content For Entry <%=RecordID%> -->
<div class="entry">
 <div class="entryIcon">
  <h3 class="entryTitle"><a href="ViewItem.asp?Entry=<%=RecordID%>"><%=Title%></a><%If Session(CookieName) = True Then Response.Write "<acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><img alt=""Edit Your Entry"" src=""Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3><br/>
 </div>

 <div class="entryBody">
  <%
  Response.Write LinkURLs(Replace(Text, vbcrlf, "<br/>" & vbcrlf))
  If (Enclosure <> "") AND (Len(Password) = 0) Then
   If Instr(Enclosure,"http://") = 0 Then Enclosure = "Sounds/" & Enclosure
   Response.Write "<br/><br/><br/>"
   %>
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
 <% End If   %>
 </div>

<p class="entryFooter">
<% 
If LegacyMode <> True Then Response.Write "<acronym title=""Printer Friendly Version""><a href=""javascript:PrintPopup('Printer_Friendly.asp?Entry=" & RecordID & "')""><img alt=""Printer Friendly Version"" src=""Images/Print.gif"" style=""border-style: none""/></a></acronym>"
If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""Email The Author"" src=""Images/Email.gif"" style=""border-style: none""/></a></acronym>"%>
<a class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTime%></a> 
<% If EnableComments <> False Then Response.Write " | <span class=""comments""><a href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</a></span>"%>
<% If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <span class=""categories"">#<a href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</a></span>"%></p></div>
<!-- End Content -->
<%
 PreviousDay = DayPosted
 Records.MoveNext
 If JustDoIt = True Then Response.Write "</div>"
Loop

'--- Close The Records ---'
Records.Close
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
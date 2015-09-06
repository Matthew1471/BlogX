<%
' --------------------------------------------------------------------------
'¦Introduction : Search Page.                                               ¦
'¦Purpose      : Lets the visitor search the database for specific entries. ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Replace.asp, Includes/Header.asp,                 ¦
'¦               Includes/ViewerPass.asp, Includes/Cache.asp                ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
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

Dim Requested
Requested = Request("Search")

'-- Filter & Clean ### --'
Requested = Replace(Requested,"'","''")

'-- If anyone has a quick fix for these, E-mail me, I'm bored of playing with it ---'
Requested = Replace(Requested,"%","")
Requested = Replace(Requested,"_","")
Requested = Replace(Requested,"[","")
Requested = Replace(Requested,"]","")

If Len(Requested) > 0 Then
 PageTitle = "Searched for &quot;" & Requested & "&quot;"
Else
 PageTitle = "Search"
End If
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<%
'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string into VisualBasic

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2

'VB Function - FreeVBCode.com
'Converted to ASP By - Matthew1471
'IIF Function For ASP - http://www.developerfusion.com/show/1606
''**********************************************

'------------------------------------------------------------
Public Function IIf(blnExpression, vTrueResult, vFalseResult)
  If blnExpression Then
    IIf = vTrueResult
  Else
    IIf = vFalseResult
  End If
End Function

'------------------------------------------------------------
Function CharCount(OrigString, Chars, CaseSensitive)
Dim lLen, lCharLen, lAns, sInput, sChar, lCtr
Dim lEndOfLoop, bytCompareType

sInput = OrigString

If sInput <> "" Then

lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then lAns = lAns + 1
    Next

CharCount = lAns

End If

End Function

Response.Write "<div id=""content"">" & VbCrlf

'--- Should We Process The Search? ---'
If Len(Requested) <> 0 Then
 Dim sarySearch, strSQL, NewTitle
 sarySearch = Split(Trim(Requested), " ")

 '--- Build up the SQL Query on whether it's an "AnyOrder"(Tm) search or whether it is an EXACT match ---
 Dim Spaces
 If Request("NoAutoComplete") = "" Then Spaces = "" Else Spaces = " "

 If Request("Mode") = "Any" Then
  '-- Search for the first search word in the URL titles --'
  strSQL = "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, Enclosure, LastModified FROM Data WHERE Text LIKE '%" & Spaces & sarySearch(0) & Spaces & "%'"

  '-- Loop to search for each search word entered by the user --'
  For intSQLLoopCounter = 0 To UBound(sarySearch)
   strSQL = strSQL & " AND Text LIKE '%" & Spaces & sarySearch(intSQLLoopCounter) & Spaces & "%'"
  Next
		
  '-- Order the search results by the RecordID --'
  strSQL = strSQL & " ORDER BY RecordID DESC;"
 Else
  strSQL = "SELECT * FROM Data WHERE (Title LIKE '%" & Spaces & Requested & Spaces & "%') OR (Text LIKE '%" & Spaces & Requested & Spaces & "%')"
  strSQL = strSQL & " ORDER BY RecordID DESC;"
 End If

 '--- Open set ---'
 Records.CursorLocation = 3 ' adUseClient
 Records.Open strSQL,Database, 1, 3

 '-- UnFilter & Scruffisize --'
 Requested = Replace(Requested,"''","'")

 '-- Let's see what page are we looking at right now --'
 Dim nPage
 nPage = CLng(Request.QueryString("Page"))

 '****************************************************************
 '-- Get Records Count --'
 Dim nRecCount
 nRecCount = Records.RecordCount

 '-- Tell recordset to split records in the pages of our size --'
 Records.PageSize = EntriesPerPage

 '-- How many pages have we got --'
 Dim nPageCount
 nPageCount = Records.PageCount

 '-- Make sure that the Page parameter passed to us is within the range --'
 If nPage < 1 Or nPage > nPageCount Then nPage = 1

 If nRecCount > 0 Then

  '-- Time to tell user what we've got so far --'
  Response.Write "<p style=""text-align:Right"">Page : " & nPage & "/" & nPageCount & "</p>"

  '-- Give user some navigation --'

  '-- First page --'
  Response.Write "<p style=""text-align: center"">"
  Response.Write "<a href=""Search.asp?Search=" & StandardURL(Requested) & "&amp;Page=" &  1 & """>First Page</a>"
  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

  '-- Previous Page --'
  Response.Write "<a href=""Search.asp?Search=" & StandardURL(Requested) & "&amp;Page=" & nPage - 1 & """>Prev. Page</a>"
  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

  '-- Next Page --'
  Response.Write "<a href=""Search.asp?Search=" & StandardURL(Requested) & "&amp;Page=" & nPage + 1 & """>Next Page</a>"
  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

  '-- Last Page --'
  Response.Write "<a href=""Search.asp?Search=" & StandardURL(Requested) & "&amp;Page=" & nPageCount & """>Last Page</a>"
  Response.Write "</p><br/>" & VbCrlf
 End If

 '-- Position recordset to the page we want to see --'
 If nRecCount > 0 Then Records.AbsolutePage = nPage

 '--- Setup Day Posted ---'
 Dim PreviousDay
 PreviousDay = "0"

 Dim RecordID, Title, Text, Password, Enclosure

 '-- Loop through records until it's a next page or End of Records --'
 Do Until (Records.EOF or Records.AbsolutePage <> nPage)

  '--- Setup Variables ---'
  Set RecordID = Records("RecordID")
  Set Title = Records("Title")
  Set Text = Records("Text")
  Set Password = Records("Password")

  If Len(Password) > 0 Then
   Text = "<form action=""ProtectedEntry.asp"" method=""GET""><center>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """>" & VbCrlf 
   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br/>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
   Text = Text & "</center></form>"
  End If

  '********** Replace(TEXT,NEW, START, NUMBER OF OCC, CASE) ********
  If Request("Mode") = "Any" Then 
   		
   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   		       
   'Loop to search for each search word entered by the user
   Dim intSQLLoopCounter
   For intSQLLoopCounter = 0 To UBound(sarySearch)
    If InStr(sarySearch(intSQLLoopCounter),"http://") = 0 Then Text = Replace(Text, " " & sarySearch(intSQLLoopCounter) , " <span style=""color:#800000;font-weight:bold;background-color: #FFFF00"">" & sarySearch(intSQLLoopCounter) & "</span>",1,-1,1)
   Next                  
   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  
		
   NewTitle = Title

  Else
      
   NewTitle = Replace(Title, Requested, " <span style=""background-color: #FFFF00; color: black; font-weight: bold"">" & Requested & "</span>",1,-1,1)
   		
   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   'Loop to search for each search word entered by the user
   For intSQLLoopCounter = 0 To UBound(sarySearch)
    If InStr(sarySearch(intSQLLoopCounter),"http://") = 0 Then Text = Replace(Text, " " & Requested , " <b><font color=""#800000""><span style=""background-color: #FFFF00"">" & Requested & "</span></font></b>",1,-1,1)
   Next    
   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	
  End If
  
  Dim CommentsCount, DayPosted, MonthPosted, YearPosted, TimePosted, NewTime, JustDoIt, LastModified
  Set Category =  Records("Category")
  Set CommentsCount = Records("Comments")

  Set DayPosted =  Records("Day")
  Set MonthPosted =  Records("Month")
  Set YearPosted =  Records("Year")
  Set TimePosted =  Records("Time")

  Set Enclosure = Records("Enclosure")

  Set LastModified = Records("LastModified")

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

  If (DayPosted <> PreviousDay) AND (NoDate <> 1) Then
   Response.Write vbcrlf & "<!-- Start Date Header -->" & vbcrlf
   Response.Write "<div class=""date"" id=""Records" & YearPosted & "-" & MonthPosted & "-" & DayPosted & """>" & vbcrlf
   Response.Write "<h2 class=""dateHeader"">" & Left(MonthName(MonthPosted),3) & " " & DayPosted & ", " & YearPosted & " (Only containing "
   If Request("Mode") <> "Any" Then Response.Write "<span style=""text-decoration: underline"">EXACTLY</span> (in this order) "
   Response.Write """<b> " & Replace(HTML2Text(Requested), "%20", " ") & " </b>"")</h2>" & vbcrlf
   Response.Write "<!-- End Date Header -->" & vbcrlf
   JustDoit = True
  Else
   JustDoIt = False
  End If
%>
 <!-- Start Content For Search List<%=RecordID%> -->
 <div class="entry">
  <div class="entryIcon">
   <h3 class="entryTitle"><a href="ViewItem.asp?Entry=<%=RecordID%>"><%=Title%></a> <%If Session(CookieName) = True Then Response.Write "<acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><img alt=""Edit Your Entry"" src=""Images/Edit.gif"" style=""border-style: none""/></a></acronym> "%><% If Request("Mode") <> "Any" Then Response.Write "(" & CharCount(Text,Requested,False) & " Occurrences)"%></h3><br/>
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
     <param name="allowScriptAccess" value="sameDomain"/>
     <param name="movie" value="Includes/mp3player.swf?id=1.2"/>
     <param name="quality" value="high"/>
     <param name="bgcolor" value="#ffffff"/>
     <param name="FlashVars" value="zipURL=<%=StandardURL(Enclosure)%>&amp;songURL=<%=StandardURL(Enclosure)%>"/>
     <embed src="Includes/mp3player.swf?id=1.2" FlashVars="zipURL=<%=Enclosure%>&amp;songURL=<%=Enclosure%>" quality="high" bgcolor="#ffffff" width="192" height="29" name="mp3player" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer"/>
    </object>
    <!-- End Podcast Object -->
    <br/><small><a href="<%=enclosure%>">Download this file</a></small>
    <% End If  %>
  </div>

  <p class="entryFooter"><% 
  If LegacyMode <> True Then Response.Write "<acronym title=""Printer Friendly Version""><a href=""javascript:PrintPopup('Printer_Friendly.asp?Entry=" & RecordID & "')""><img alt=""Printer Friendly Version"" src=""Images/Print.gif"" style=""border-style: none""/></a></acronym>"
  If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""Email The Author"" src=""Images/Email.gif"" style=""border-style: none""/></a></acronym>"%>
  <a class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTime%></a>
  <% 
   If EnableComments <> False Then Response.Write " | <span class=""comments""><a href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</a></span>"
   If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write "| <span class=""categories"">#<a href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</a></span>"%>
  </p></div>
  <!-- End Content -->
  <%
  PreviousDay = DayPosted
  Records.MoveNext
  If JustDoIt = True Then Response.Write "</div>"
 Loop

'--- Close The Records & Database ---'
Records.Close

End If

If nRecCount < 1 Then

 '-- Proxy Handler --'
 If (NOT DontSetModified) AND (Session(CookieName) = False) Then CacheHandle(GeneralModifiedDate)
%>
<!-- Start No|Invalid Text / Default page load / EOF content -->
<div class="entry">
<h3 class="entryTitle"><% If Request("Search") = "" Then Response.Write "No Text entered" Else Response.Write "No Entries Found With That Criteria"%></h3><br/>
<div class="entryBody">

<% If (Request("Search") <> "") AND (Requested = "") Then %>
 <p align="center">You cannot perform a search on the criteria you selected.</P>
<% ElseIf (Requested = "") AND (Request("Search") = "") Then %>
 <p align="center">Welcome, Please enter your query below<br/>
<% ElseIf (nRecCount < 1) AND (Requested <> "") Then %>
 <p align="center">No Entries found with the criteria "<b><%=HTML2Text(Requested)%></b>" in either the text's sentences <b>OR</b> <% If Request("Mode") <> "Any" Then Response.Write "the title" ELSE Response.Write "in any order in the text"%>.<br/></p>
<% End If
If Request("Search") <> "" Then Response.Write "<P align=""center"">Please try again with a different criteria:" 
%>

 <form name="Search" method="post" action="Search.asp" style="text-align: center">
  <input Name="Search" type="text" value="<%=Replace(Requested,"""","&quot;")%>" size="13" maxlength="70"><input type="submit" value="Search"><br/>
  <br/>Words Can Be In Any Order In The Text (so long as they <b>ALL</b> appear) : <input name="Mode" type="Checkbox" Value="Any" <%If Request("Mode") = "Any" Then Response.Write "CHECKED"%>>
  <br/>Don't Complete words ("Holl" wont return "Holly") : <input name="NoAutoComplete" type="Checkbox" value="True" <%If Request("NoAutoComplete") = "True" Then Response.Write "CHECKED"%>>
 </form>

 </p>
 </div>

</div>

<!-- End No Text Content -->
<% End If%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
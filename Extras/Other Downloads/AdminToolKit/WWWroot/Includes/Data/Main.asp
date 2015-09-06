<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%

Dim nPage, nRecCount, nPageCount, PreviousDay

'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient

If SpecificRequest = False Then
    Records.Open "SELECT * FROM Data ORDER BY RecordID DESC",Database, 1, 3
Else

    Dim DataMonth, DataYear, DaySelected

    DataMonth = nMonth
    DataYear  = nYear

    If Request("POS") = "NEXT" Then DataMonth = DataMonth + 1
    If Request("POS") = "LAST" Then DataMonth = DataMonth - 1
 
    If DataMonth = 0 Then
    DataMonth = 12
    DataYear = DataYear - 1
    End If

    If DataMonth = 13 Then
    DataMonth = 1
    DataYear = DataYear + 1
    End If

    If nDay <> "" Then DaySelected = "AND Day=" & nDay
    Records.Open "SELECT * FROM Data WHERE Month=" & DataMonth & " AND Year=" & DataYear & " " & DaySelected & " ORDER BY RecordID DESC;",Database, 1, 3
End If

' Let's see what page are we looking at right now
nPage = CLng(Request.QueryString("Page"))

'****************************************************************
' Get Records Count
nRecCount = Records.RecordCount

' Tell recordset to split records in the pages of our size
Records.PageSize = EntriesPerPage

' How many pages we've got
nPageCount = Records.PageCount

' Make sure that the Page parameter passed to us is within the range
If nPage < 1 Or nPage > nPageCount Then nPage = 1

Response.Write "<DIV id=content>" & VbCrlf

' Time to tell user what we've got so far
Response.Write "<p align=""Right"">Page : " & nPage & "/" & nPageCount & "</p><p>"

' Give user some navigation

' First page
Response.Write "<Center>"

If nPage > 1 Then Response.Write "<A HREF=""" & PageName & "?Page=" & 1 
If (nPage > 1) AND (szYearMonth <> "") Then Response.Write "&YearMonth=" & szYearMonth
If nPage > 1 Then  Response.Write """>" Else Response.Write "<font Color=""Gray"">"
Response.Write "First Page"
If nPage > 1 Then Response.Write "</A>" Else Response.Write "</font>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

' Previous Page
If nPage > 1 Then Response.Write "<A HREF=""" & PageName & "?Page=" & nPage - 1
If (nPage > 1) AND (szYearMonth <> "") Then Response.Write "&YearMonth=" & szYearMonth
If nPage > 1 Then  Response.Write """>" Else Response.Write "<font Color=""Gray"">"
Response.Write "Prev. Page"
If nPage > 1 Then Response.Write "</A>" Else Response.Write "</font>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
	
' Next Page
If nPage < nPageCount Then Response.Write "<A HREF=""" & PageName & "?Page=" & nPage + 1
If (nPage < nPageCount) AND (szYearMonth <> "") Then Response.Write "&YearMonth=" & szYearMonth
If nPage < nPageCount Then  Response.Write """>" Else Response.Write "<font Color=""Gray"">"
Response.Write "Next Page"
If nPage < nPageCount Then Response.Write "</A>" Else Response.Write "</font>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

' Last Page
If nPage < nPageCount Then Response.Write "<A HREF=""" & PageName & "?Page=" & nPageCount 
If (nPage < nPageCount) AND (szYearMonth <> "") Then Response.Write "&YearMonth=" & szYearMonth
If nPage < nPageCount Then  Response.Write """>" Else Response.Write "<font Color=""Gray"">"
Response.Write "Last Page"
If nPage < nPageCount Then Response.Write "</A>" Else Response.Write "</font>"
Response.Write "</center><br>" & VbCrlf

' Position recordset to the page we want to see
If nRecCount > 0 Then Records.AbsolutePage = nPage

'--- Setup Day Posted ---'
PreviousDay = "0"
				
' Loop through records until it's a next page or End of Records
Dim RecordID, Title, Text, Password, DayPosted, MonthPosted, YearPosted, TimePosted, CommentsCount
Dim NewTime, JustDoIt

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

   Set CommentsCount = Records("Comments")

   If Len(Password) > 0 Then
   Text = "<form action=""ProtectedEntry.asp"" method=""GET""><center>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """>" & VbCrlf 
   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br><br>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
   Text = Text & "</center></form>"
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

Response.Write vbcrlf & "<!--- Start Date Header --->" & vbcrlf
Response.Write "<DIV class=date id=2003-11-30>" & vbcrlf
Response.Write "<H2 class=dateHeader>" & EntryWeekDay & ", " & DayPosted & " " & Left(MonthName(MonthPosted),3) & " " & YearPosted & "</H2>" & vbcrlf
Response.Write "<!--- End Date Header --->" & vbcrlf

JustDoit = True
Else
JustDoIt = False
End If
%>
<!--- Start Content For Entry <%=RecordID%> --->
<DIV class=entry>
<H3 class=entryTitle><A href="ViewItem.asp?Entry=<%=RecordID%>"><%=Title%></A><%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym>"%></H3>
<DIV class=entryBody><P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P></DIV>
<P class=entryFooter>
<% 
If LegacyMode <> True Then Response.Write "<acronym title=""Printer Friendly Version""><a href=""javascript:PrintPopup('Printer_Friendly.asp?Entry=" & RecordID & "')""><Img Border=""0"" Src=""Images/Print.gif""></a></acronym>"
If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%>
<A class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTime%></A> 
<% If EnableComments <> False Then Response.Write " | <SPAN class=""comments""><A href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</A></SPAN>"%>
<% If (ShowCat <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <SPAN class=""categories"">#<A href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</A></SPAN>"%></P></DIV>
<!--- End Content --->
<%
PreviousDay = DayPosted
Records.MoveNext
If JustDoIt = True Then Response.Write "</Div>"
Loop

'--- Close The Records ---
Records.Close
%>
</DIV>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->


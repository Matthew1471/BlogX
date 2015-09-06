<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
Category = Request.Querystring("Cat")
Category = Replace(Category,"'","&#39;")
Category = Replace(Category," ","%20")

'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient

    Records.Open "SELECT * FROM Data WHERE Category='" & Category & "' ORDER BY RecordID DESC;",Database, 1, 3

' Let's see what page are we looking at right now
Dim nPage
If IsNumeric(Request.QueryString("Page")) Then nPage = Int(Request.QueryString("Page"))

'****************************************************************
' Get Records Count
Dim nRecCount
nRecCount = Records.RecordCount

' Tell recordset to split records in the pages of our size
Records.PageSize = 10

' How many pages we've got
Dim nPageCount
nPageCount = Records.PageCount

' Make sure that the Page parameter passed to us is within the range
If nPage < 1 Or nPage > nPageCount Then nPage = 1

Response.Write "<DIV id=content>" & VbCrlf

' Time to tell user what we've got so far
Response.Write "<p align=""Right"">Page : " & nPage & "/" & nPageCount & "</p><p>"

' Give user some navigation

' First page
Response.Write "<Center>"
Response.Write 	"<A HREF=""ViewCat.asp?Cat=" & Request.Querystring("Cat") & "&Page=" &  1 & """>First Page</A>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

' Previous Page
Response.Write 	"<A HREF=""ViewCat.asp?Cat=" & Request.Querystring("Cat") & "&Page=" & nPage - 1 & """>Prev. Page</A>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
	
' Next Page
Response.Write 	"<A HREF=""ViewCat.asp?Cat=" & Request.Querystring("Cat") & "&Page=" & nPage + 1 & """>Next Page</A>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

' Last Page
Response.Write 	"<A HREF=""ViewCat.asp?Cat=" & Request.Querystring("Cat") & "&Page=" & nPageCount & """>Last Page</A>"
Response.Write "</center><br>" & VbCrlf

' Position recordset to the page we want to see
If nRecCount > 0 Then Records.AbsolutePage = nPage

'--- Setup Day Posted ---'
Dim PreviousDay
PreviousDay = "0"

Dim RecordID, Title, Text, Password, CommentsCount
Dim DayPosted, MonthPosted, YearPosted, TimePosted
Dim NewTime, JustDoIt
		
' Loop through records until it's a next page or End of Records
Do Until (Records.EOF or Records.AbsolutePage <> nPage )

'--- Setup Variables ---'
   Set RecordID = Records("RecordID")
   Set Title = Records("Title")
   Set Text = Records("Text")
   Set Category = Records("Category")
   Set Password = Records("Password")
   Set CommentsCount = Records("Comments")

   Set DayPosted =  Records("Day")
   Set MonthPosted =  Records("Month")
   Set YearPosted =  Records("Year")
   Set TimePosted =  Records("Time")

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
Response.Write vbcrlf & "<!--- Start Date Header --->" & vbcrlf
Response.Write "<DIV class=date id=2003-11-30>" & vbcrlf
Response.Write "<H2 class=dateHeader>" & Left(MonthName(MonthPosted),3) & " " & DayPosted & ", " & YearPosted & " (Only #" & Replace(Category, "%20", " ") & ")</H2>" & vbcrlf
Response.Write "<!--- End Date Header --->" & vbcrlf
JustDoit = True
Else
JustDoIt = False
End If
%>
<!--- Start Content For Category List (<%=DayPosted%>)--->
<DIV class=entry>
<H3 class=entryTitle><A href="ViewItem.asp?Entry=<%=RecordID%>"><%=Title%></A> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym>"%></H3>
<DIV class=entryBody><P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P></DIV>
<P class=entryFooter>
<% 
If LegacyMode <> True Then Response.Write "<acronym title=""Printer Friendly Version""><a href=""javascript:PrintPopup('Printer_Friendly.asp?Entry=" & RecordID & "')""><Img Border=""0"" Src=""Images/Print.gif""></a></acronym>"
If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%>
<A class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTime%></A> 
<% If EnableComments <> False Then Response.Write " | <SPAN class=""comments""><A href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</A></SPAN>"%>
| <SPAN class=categories>#<%=Replace(Category, "%20", " ")%></SPAN></P></DIV>
<!--- End Content --->
<%
PreviousDay = DayPosted
Records.MoveNext
If JustDoIt = True Then Response.Write "</Div>"
Loop

Records.Close
%>
</DIV>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
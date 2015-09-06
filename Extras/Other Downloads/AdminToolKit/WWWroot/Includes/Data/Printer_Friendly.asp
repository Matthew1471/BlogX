<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
Dim Requested
Requested = Request.Querystring("Entry")
If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then Requested = 0

'--- Open set ---'
Records.Open "SELECT * FROM Data WHERE RecordID=" & Requested,Database, 1, 3

If NOT Records.EOF Then

   Dim RecordID, Title, Text, Category, CommentsCount, Password
   Dim DayPosted, MonthPosted, YearPosted, TimePosted, NewTime

'--- Setup Variables ---'
   RecordID = Records("RecordID")
   Title = Records("Title")
   Text = Records("Text")
   Category =  Records("Category")
   CommentsCount = Records("Comments")
   Password = Records("Password")

   DayPosted =  Records("Day")
   MonthPosted =  Records("Month")
   YearPosted =  Records("Year")
   TimePosted =  Records("Time")

   If (Len(Password) > 0) AND (Ucase(Request.Querystring("Password")) <> Ucase(Password)) Then
   Text = "<form action=""Printer_Friendly.asp"" method=""GET""><center>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """>" & VbCrlf 
   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br><br>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
   Text = Text & "</center></form>"
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
<html>

<head>
<title><%=SiteDescription & " - " & Title %></title>
</head>
<body bgColor="#FFFFFF" text="#000000" onLoad="window.print()" >
<center>
<p><font face="Verdana, Arial, Helvetica" size="2"><a href="javascript:onClick=window.print()">Print Page</a> | <a href="JavaScript:onClick=window.close()">Close Window</a></font></p>
</center>

<p><font face="Verdana, Arial, Helvetica" size="2"><b><%=Title%></b></font></p>
<b>Topic:</b> <a href="<%=SiteURL%>ViewItem.asp?Entry=<%=RecordID%>"><%=SiteURL%>ViewItem.asp?Entry=<%=RecordID%></a><br>
<b>Date:</b> <%=FormatDateTime(Now(),vblongdate)%></p>
<% If (ShowCat <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write "<b>Category:</b> #<A href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</A><BR>"%>
<p><b>Subject:</b> <%=Title%><br>
<b>Posted on:</b> <%=DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime%><br>
<b>Message:</b><br>
<P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P>
<hr></p>
<p><font face="Verdana, Arial, Helvetica" size="2"><b><%=SiteDescription%> </b>: <a href="<%=SiteURL%>"><%=SiteURL%></a></p>
<p><font face="Verdana, Arial, Helvetica" size="2"><b><%=Copyright%></b> </p>
</font>
</body>
</html>
<%
Else
Response.Write "<html>"
Response.Write "<head>"
Response.Write "<script>window.close()</script>"
Response.Write "</head>"
Response.Write "</html>"
End If

'--- Close The Records ---
Records.Close
Set Records = Nothing

Database.Close
Set Database = Nothing
%>
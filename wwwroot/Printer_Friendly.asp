<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
Dim Requested
Requested = Request.Querystring("Entry")
If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then Requested = 0

'--- Open set ---'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments FROM Data WHERE RecordID=" & Requested,Database, 1, 3

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
  Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br/>" & VbCrlf
  Text = Text & "This post is password protected. To view it please enter your password below:"
  Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
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

 Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf
 Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">" & VbCrlf
 Response.Write " <head>" & VbCrlf
 Response.Write "  <title>" & SiteDescription & " - " & Title & "</title>" & VbCrlf
 Response.Write "  <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" & VbCrlf
 Response.Write " </head>" & VbCrlf
 Response.Write " <body style=""background-color: #FFFFFF; color: #000000"" onload=""window.print()"">" & VbCrlf
 Response.Write "  <p style=""text-align:center; font-family: Verdana, Arial, Helvetica; font-size: small""><a href=""javascript:onClick=window.print()"">Print Page</a> | <a href=""JavaScript:onClick=window.close()"">Close Window</a></p>" & VbCrlf
 Response.Write "  <p style=""font-family: Verdana, Arial, Helvetica; font-size: small""><b>" & Title & "</b></p>" & VbCrlf
 Response.Write "  <p><b>Topic:</b> <a href=""" & SiteURL & "ViewItem.asp?Entry=" & RecordID & """>" & SiteURL & "ViewItem.asp?Entry=" & RecordID & "</a><br/>" & VbCrlf
 Response.Write "   <b>Date:</b> " & FormatDateTime(Now(),vblongdate) & "<br/>" & VbCrlf
 If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write "<b>Category:</b> #<a href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</a><br/>"
 Response.Write "   <b>Subject:</b> " & Title & "<br/>" & VbCrlf
 Response.Write "   <b>Posted on:</b> " & DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime & "<br/>" & VbCrlf
 Response.write "   <b>Message:</b></p>" & VbCrlf & VbCrlf

 Response.Write " <div>" & LinkURLs(Replace(Text, vbcrlf, "<br/>" & vbcrlf)) & "</div>" & VbCrlf
 Response.Write " <hr/>" & VbCrlf
 Response.Write " <p style=""font-family: Verdana, Arial, Helvetica; font-size: small""><b>" & SiteDescription & " </b>: <a href=""" & SiteURL & """>" & SiteURL & "</a></p>" & VbCrlf
 Response.Write " <p style=""font-family: Verdana, Arial, Helvetica; font-size: small""><b>" & Copyright & "</b></p>" & VbCrlf

 Response.Write " </body>" & VbCrlf
 Response.Write "</html>" & VbCrlf
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
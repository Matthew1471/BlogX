<% AlertBack = True %>
<!-- #INCLUDE FILE="../../Includes/Header.asp" -->
<!-- #INCLUDE FILE="../../Admin/Admin.asp" -->
<DIV id=content>
<%
'--- Querish Querystring ---'
Dim Title2, Text2, Category2, Password2, sDay2, sMonth2, sYear2, TimePosted2, Comments2, StopComments2
Dim Requested
Requested = Request.Querystring("Entry")

If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then 
Database.Close
Set Database = Nothing
Set Records  = Nothing
Response.Write "No Entry Specified"
Response.End
End If

'--- Open New entry And Hold A TEMP Copy ---'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, StopComments FROM Data WHERE RecordID=" & Requested & " ORDER BY RecordID DESC",Database, 1, 3
	If NOT Records.EOF Then
	 '--- Setup Variables ---'
   	 Title = Records("Title")
   	 Text = Records("Text")
   	 Category =  Records("Category")
   	 Password =  Records("Password")
   	 sDay = Records("Day")
   	 sMonth = Records("Month")
   	 sYear = Records("Year")
   	 TimePosted = Records("Time")
   	 Comments = Records("Comments")
   	 StopComments = Records("StopComments")
	End If
Records.Close

'--- Open Previous Entry And Hold A TEMP2 Copy ---'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, StopComments FROM Data WHERE RecordID=" & Requested - 1 & " ORDER BY RecordID DESC",Database, 1, 3
	If NOT Records.EOF Then
	 '--- Setup Variables ---'
   	 Title2 = Records("Title")
   	 Text2 = Records("Text")
   	 Category2 =  Records("Category")
   	 Password2 =  Records("Password")
   	 sDay2 = Records("Day")
   	 sMonth2 = Records("Month")
   	 sYear2 = Records("Year")
   	 TimePosted2 = Records("Time")
   	 Comments2 = Records("Comments")
   	 StopComments2 = Records("StopComments")
	End If
Records.Close

'-- Write NEW Last But One Entry --'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, StopComments FROM Data WHERE RecordID=" & Requested & " ORDER BY RecordID DESC",Database, 1, 3
Records("Title") = Title2
Records("Text") = Text2
Records("Category") = Category2
Records("Password") = Password2
Records("Day") = sDay2
Records("Month") = sMonth2
Records("Year") = sYear2
Records("Time") = TimePosted2
Records("Comments") = Comments2
Records("StopComments") = StopComments2
Records.Update
Records.Close

'-- Write NEW Last Entry --'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, StopComments FROM Data WHERE RecordID=" & Requested - 1 & " ORDER BY RecordID DESC",Database, 1, 3
Records("Title") = Title
Records("Text") = Text
Records("Category") = Category
Records("Password") = Password
Records("Day") = sDay
Records("Month") = sMonth
Records("Year") = sYear
Records("Time") = TimePosted
Records("Comments") = Comments
Records("StopComments") = StopComments
Records.Update
Records.Close

Response.Write "<p align=""Center"">Entry Moved</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"
%>
</Div>
<!-- #INCLUDE FILE="../../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../../Includes/Footer.asp" -->
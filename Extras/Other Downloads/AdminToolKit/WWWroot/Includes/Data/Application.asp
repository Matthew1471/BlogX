<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
If (UCase(Request.Form("Username")) = UCase(AdminUsername)) AND (UCase(Request.Form("Password")) = Ucase(AdminPassword)) Then

'Dimension variables
Dim EntryCat

EntryCat = Request.Form("Category")

'### Filter & Clean ###
EntryCat = Replace(EntryCat,"'","&#39;")
EntryCat = Replace(EntryCat," ","%20")

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
Response.Write "No Text Entered"
Response.End
End If

'### Create a recordset object ###
Set Records = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Data", Database
Records.AddNew
Records("Title") = Request.Form("Title")
Records("Text") = Request.Form("Content")
Records("Category") = EntryCat

Records("Day") = Day(DateAdd("h",TimeOffset,Now()))
Records("Month") = Month(DateAdd("h",TimeOffset,Now()))
Records("Year") = Year(DateAdd("h",TimeOffset,Now()))
Records("Time") = TimeValue(DateAdd("h",TimeOffset,Time()))
Records.Update

'#### Close Objects ###	
Records.Close
Set Records = Nothing
Database.Close
Set Database = Nothing

Response.Write "Entry Submission Successfull"

Else

Response.Write "User/Password Error"

End If
%>
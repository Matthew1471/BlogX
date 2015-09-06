<%
Dim Database, Records

'### Create a connection odject ###
Set Database = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
Database.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set Records = Server.CreateObject("ADODB.Recordset")
    Records.CursorLocation = 3

'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM News ORDER BY ID DESC;", Database

Dim Title, Content, TimePosted, DatePosted

Title = Records("Title")
Content = Records("Content")
TimePosted = Records("Time")
DatePosted = Records("Date")

'--- We're British, Let's 12Hour Clock Ourselves ---'
Dim NewTime

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
NewTime = ""
If Hour(TimePosted) < 10 Then NewTime = "0"
NewTime = NewTime & Hour(TimePosted) & ":"
If Minute(TimePosted) < 10 Then NewTime = NewTime & "0"
NewTime = NewTime & Minute(TimePosted)


End If

TimePosted = NewTime

Records.Close
Database.Close

'#### Close Objects ###
Set Database = Nothing
Set Records = Nothing
%>
      <!--- News Sidebar --->
      <td width="*" bgcolor="#FFFFFF" height="194" style="PADDING-LEFT: 5px;">
      <font face="Verdana" size="2" color="#000444">
      <B><%=Title%></B>
      <BR><BR><%=Replace(Content,VbCrlf & VbCrlf,"<br><br>" & Vbcrlf)%>
      <BR><BR><B><font color="Orange"><%=DatePosted%>&nbsp;<%=TimePosted%></font></B>
      </font>
      </td>
      <!--- End Of News --->
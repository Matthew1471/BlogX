<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<table border="0">
<%

'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM BannedIP ORDER BY IP;", Database

    DelIP = Request.Querystring("Delete")
    DelIP = Replace(DelIP,"'","")

    If (DelIP <> "") Then
    Database.Execute "DELETE FROM BannedIP WHERE IP='" & DelIP & "'"
    If Records.RecordCount > 0 Then Records.Update
    If Records.RecordCount > 0 Then Records.MoveFirst
    End If

Do Until (Records.EOF) 

NotEmpty = True

Set IP = Records("IP")

Set DateBanned = Records("Date")
Set TimeBanned = Records("Time")
%>
                 <tr>
                 <td bgcolor="#FF0000"><Font color="#FFFFFF"><B><%=IP%></B></Font></td>
                 <td bgcolor="#0000FF"><Font Color="#FFFFFF"><B><%=FormatDateTime(DateBanned,vblongdate) & " (" & TimeBanned & ")" %></B></Font></td>
                 <td><acronym title="Unban User"><a href="?Delete=<%=IP%>"><Img Border="0" Src="../Images/Key.gif"></a></acronym></td>
                 </tr>
<%
Records.MoveNext
Loop

'#### Close Objects ###
Records.Close
%>
</table>
<% If NotEmpty = False Then Response.Write "<P align=""Center"">No Banned Users Found</P>" & VbCrlf & "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>" %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
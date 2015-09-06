<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<table border="0">
<%

'### Open The Records Ready To Write ###
Records.Open "SELECT * FROM Refer ORDER BY ReferURL;", Database

Count = 0

Do Until (Records.EOF)
Set URL = Records("ReferURL")
Count = Count + Records("ReferHits")
%>
                 <tr><td>
                 <% If InStr(URL,"http://") <> 0 Then Response.Write "<a href=""" & URL & """>" %>
                 <%=URL%>
                 <% If InStr(URL,"http://") <> 0 Then Response.Write "</a>" %>
                 </td><td><%=Records("ReferHits")%></td></tr>
<%
Records.MoveNext
Loop

'#### Close Objects ###	
Records.Close
%>
<tr><td>Total</td><td><%=Count%></td></tr>
</table>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
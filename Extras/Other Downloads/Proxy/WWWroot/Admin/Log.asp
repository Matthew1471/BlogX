<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<table border="0">
<tr><td>IP Address</td>
<td>Action</td>
<td>Date</td></tr>
<%

'### Open The Records Ready To Write ###
Records.Open "SELECT ID, IP, Action, Date FROM Log ORDER BY Date DESC;", Database

Do Until (Records.EOF)
%>
		 <tr><td><%=Records("IP")%></td>
                 <td><%=Records("Action")%></td>
		 <td><%=Records("Date")%></td></tr>
<%
Records.MoveNext
Loop

'#### Close Objects ###	
Records.Close
%>
</tr>
</table>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
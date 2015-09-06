<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
Records.CursorLocation = 3 ' adUseClient
Records.Open "SELECT DISTINCT Category FROM Data ORDER BY Category",Database, 1, 3
 Response.Write "Category list (" & Records.RecordCount & "):" & VbCrlf
Records.Close

'-- Output some diagnostic info --'
Response.Write "<hr/>Version of ADO : " & Database.Version & "<br/>" & VbCrlf
Response.Write "DBMS Version : " & Database.Properties("DBMS Version") & "<br/>" & VbCrlf
Response.Write "Provider Name : " & Database.Properties("Provider Name") & "<br/>" & VbCrlf
Response.Write "OLE DB Version : " & Database.Properties("OLE DB Version") & VbCrlf

Database.Close
Set Records = Nothing
Set Database = Nothing

Response.Write Session(CookieName)
%>
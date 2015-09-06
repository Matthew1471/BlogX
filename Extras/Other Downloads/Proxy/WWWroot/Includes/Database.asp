<%
'--- Open Database ---'
Set Database = Server.CreateObject("ADODB.connection")
'Database.Open  "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=" & DataPassword & "; DBQ=" & DataFile
'Database.Open  "DSN=BlogX;"
Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Datafile & ";"

'--- Open Recordset ---'
set Records = Server.CreateObject("ADODB.recordset")
%>
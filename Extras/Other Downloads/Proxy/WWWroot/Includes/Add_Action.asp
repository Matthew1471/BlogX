<%
Action = Replace(Action,"'","")

Dim LogConnection, LogRecords

'### Create a connection object ###
Set LogConnection = Server.CreateObject("ADODB.Connection")
			 
'### Database connection info and driver ###
'### Set an active connection to the Connection object ###
LogConnection.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set LogRecords = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
LogRecords.CursorType = 2
LogRecords.LockType = 3
LogRecords.Open "SELECT IP, Action, Date FROM Log", LogConnection
LogRecords.AddNew
LogRecords("IP") = Request.ServerVariables("REMOTE_ADDR")
LogRecords("Action") = Left(Action,80)
LogRecords("Date") = Now()
LogRecords.Update

'#### Close Objects ###
LogRecords.Close
LogConnection.Close

Set LogConnection = Nothing
Set LogRecords = Nothing
%>
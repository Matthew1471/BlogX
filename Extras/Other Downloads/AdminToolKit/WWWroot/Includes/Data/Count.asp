<%
OPTION EXPLICIT

On Error Resume Next

Dim DataFile, DataPassword

DataFile = "C:\Inetpub\database\BlogX.mdb"
DataPassword = "****"

'Dimension variables
Dim ReferDatabase 		'Database Connection Variable
Dim ReferRecords           	'Database Recordset Variable
Dim ReferURL                    'The URL

'### Find Out Refer ###'
If Request.ServerVariables("HTTP_REFERER") <> "" Then

ReferURL = Replace(Left(Request.ServerVariables("HTTP_REFERER"),100),"'", "&#39;")
ReferURL = Replace(ReferURL,"Admin/", "")

Dim Last
Last = InStrRev(ReferURL,"/")
ReferURL = Left(ReferURL,Last)
Else
ReferURL = "(None)"
End If

If Instr(Request.ServerVariables("REMOTE_ADDR"),"192.168") <> 0 Then ReferURL = "Local Address"
If Instr(Request.ServerVariables("REMOTE_ADDR"),"localhost") <> 0 Then ReferURL = "Local Address"
If Instr(Request.ServerVariables("HTTP_REFERER"),"cache:") <> 0 Then ReferURL = "Google Cache"

'### Create a connection odject ###
Set ReferDatabase = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
ReferDatabase.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=" & DataPassword & "; DBQ=" & DataFile

'### Create a recordset object ###
Set ReferRecords = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
ReferRecords.CursorType = 2
ReferRecords.LockType = 3
ReferRecords.Open "SELECT * FROM ScriptRefer WHERE ReferURL='" & ReferURL & "';", ReferDatabase

If Not ReferRecords.EOF = True Then
ReferRecords("ReferHits") = Int(ReferRecords("ReferHits")) + 1
Else
ReferRecords.AddNew
ReferRecords("ReferURL") = ReferURL
ReferRecords("ReferHits") = 1
End If
ReferRecords("IP") = Request.ServerVariables("REMOTE_ADDR")
ReferRecords.Update

ReferRecords.Close
ReferDatabase.Close

'#### Close Objects ###
Set ReferRecords = Nothing
Set ReferDatabase = Nothing

Server.Transfer "Images\Blank.gif"
%>
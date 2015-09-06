<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
'Dimension variables
Dim ReferDatabase 		'Database Connection Variable
Dim ReferRecords           	'Database Recordset Variable
Dim Refer, ReferURL, Last, Length'The URL

Dim Domain
Domain = Request.ServerVariables("HTTP_Host")
Domain = Replace(Domain,"www.","")

'### Find Out Refer ###'
Refer = Request.ServerVariables("HTTP_REFERER")

If (Instr(Refer,Domain) <> 0) AND (Instr(Refer,Root) <> 0) Then

ReferURL = Replace(Left(Refer,100),"'", "&#39;")
ReferURL = Replace(ReferURL,"%27", "&#39;")
ReferURL = Replace(ReferURL,"http://", "")
ReferURL = Replace(ReferURL,"www.", "")
ReferURL = Replace(ReferURL,Domain, "")
ReferURL = Replace(ReferURL,Root, "")

Length = Len(ReferURL) - 1

Last = InStrRev(ReferURL,"/")

ReferURL = Left(ReferURL,Last-1)

'### Create a connection odject ###
Set ReferDatabase = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
ReferDatabase.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set ReferRecords = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
ReferRecords.CursorType = 2
ReferRecords.LockType = 3
ReferRecords.Open "SELECT * FROM Top10 WHERE Blog='" & ReferURL & "';", ReferDatabase

If Not ReferRecords.EOF = True Then
ReferRecords("Hits") = Int(ReferRecords("Hits")) + 1
Else
ReferRecords.AddNew
ReferRecords("Blog") = ReferURL
ReferRecords("Hits") = "1"
End If

ReferRecords.Update

'### Close Objects ##
ReferRecords.Close
ReferDatabase.Close

'#### Kill Objects ###	
Set ReferDatabase = Nothing
Set ReferRecords = Nothing

End If

Server.Transfer "Includes\Images\Blank.gif"
%>
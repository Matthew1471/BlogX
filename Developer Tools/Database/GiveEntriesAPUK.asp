<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="../../Includes/Header.asp" -->
<%
'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, Enclosure, EntryPUK FROM Data Where EntryPUK IS NULL",Database, 1, 3

    Do Until Records.EOF
	Randomize Timer
	Records("EntryPUK") = Int(Rnd()*99999999)
	Records.Update

	Response.Write "Updated " & Records("RecordID") & " (PUK:" & Records("EntryPUK") & ")<br>" & VbCrlf
	Records.MoveNext
    Loop

    Records.Close %>
<!-- #INCLUDE FILE="../../Includes/Footer.asp" -->
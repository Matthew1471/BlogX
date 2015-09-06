<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="../../Includes/Header.asp" -->
<%
'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT CommentID, PUK FROM Comments Where PUK IS NULL",Database, 1, 3

    Do Until Records.EOF
	Randomize Timer
	Records("PUK") = Int(Rnd()*99999999)
	Records.Update

	Response.Write "Updated " & Records("CommentID") & " (PUK:" & Records("PUK") & ")<br>" & VbCrlf
	Records.MoveNext
    Loop

    Records.Close %>
<!-- #INCLUDE FILE="../../Includes/Footer.asp" -->
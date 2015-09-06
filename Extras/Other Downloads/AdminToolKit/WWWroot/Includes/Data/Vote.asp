<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
Dim PollID

If Request.Form("Vote") <> "" Then

'--- Open set ---'
Records.CursorLocation = 3 ' adUseClient

    '-- Find out the current Poll --'
    Records.Open "SELECT PollID FROM Poll ORDER BY PollID DESC",Database, 1, 3
    If Records.EOF = False Then PollID = Records("PollID") Else PollID = 0
    Records.Close

    '-- Have we already voted --'?
    Records.Open "SELECT VoteID FROM Votes WHERE PollID="& PollID & "AND IP='" & Request.ServerVariables("REMOTE_ADDR") & "'",Database, 1, 3
    If Records.EOF = False Then
    Records.Close
    Database.Close
    Set Records = Nothing
    Set Database = Nothing
    Response.Redirect(PageName)
    End If
    Records.Close

    '### Open The Records Ready To Write ###
    Records.CursorType = 2
    Records.LockType = 3

    '### Write In Comments ###'
    Records.Open "SELECT PollID, IP, Option FROM Votes", Database
    Records.AddNew
    Records("PollID") = PollID
    Records("IP") = Request.ServerVariables("REMOTE_ADDR")
    Records("Option") = Request.Form("Vote")
    Records.Update
    Records.Close
        
    '### Write In Comments ###'
    Records.Open "SELECT PollID, Op1, Op2, Op3, Op4, Total FROM Poll WHERE PollID=" & PollID, Database
    If Request.Form("Vote") = 1 Then Records("Op1") = Records("Op1") + 1
    If Request.Form("Vote") = 2 Then Records("Op2") = Records("Op2") + 1
    If Request.Form("Vote") = 3 Then Records("Op3") = Records("Op3") + 1
    If Request.Form("Vote") = 4 Then Records("Op4") = Records("Op4") + 1
    If Request.Form("Vote") <= 4 Then Records("Total") = Records("Total") + 1
    Records.Update
    Records.Close

End If

Database.Close
Set Records = Nothing
Set Database = Nothing
Response.Redirect(PageName)
%>
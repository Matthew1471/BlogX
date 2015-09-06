<%
'Dimension variables
Dim RSSRefer, ReferURL

'### Find Out Refer ###'
If Request.ServerVariables("HTTP_REFERER") <> "" AND InStr(Request.ServerVariables("HTTP_REFERER"),SiteUrl) = 0 Then
ReferURL = Replace(Left(Request.ServerVariables("HTTP_REFERER"),100),"'", "&#39;")
Else
ReferURL = "(None)"
End If

'### What If We Are RSS? ###'
If RSSRefer <> "" Then ReferURL = RSSRefer

If Instr(Request.ServerVariables("REMOTE_ADDR"),"192.168") = 0 Then ReferURL = "Local Address"
If Instr(Request.ServerVariables("REMOTE_ADDR"),"cache:") = 0 Then ReferURL = "Cache"

'### Open The Records Ready To Write ###

'CursorType: can be one: adOpenForwardOnly (default), adOpenStatic, adOpenDynamic, adOpenKeyset
'LockType: can be one of: adLockReadOnly (default), adLockOptimistic, adLockPessimistic, adLockBatchOptimistic

Records.LockType = 3

	On Error Resume Next

	Records.Open "SELECT ReferHits, ReferURL FROM Refer WHERE ReferURL='" & ReferURL & "';", Database

	If Not Records.EOF = True Then
	Records("ReferHits") = Int(Records("ReferHits")) + 1
	Else
	Records.AddNew
	Records("ReferURL") = ReferURL
	Records("ReferHits") = 1
	End If

	Records.Update

	Records.CancelUpdate

	Records.Close
	On Error Goto 0
%>
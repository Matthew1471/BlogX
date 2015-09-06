<%
'---------- Last 5 Entry Titles Plugin (V1.0) --------
'//= - - - - - - - 
'// Copyright 2004, Matthew Roberts
'// 
'// Usage Of This Software Is Subject To The Terms Of The License
'//= - - - - - - -

PluginTitle = "Last 5 Entry Titles"

Records.Open "SELECT RecordID, Title FROM Data ORDER BY RecordID DESC",Database, 1, 3

For Count = 0 to 4
	If NOT Records.EOF Then 
	PluginText = PluginText & "<Li><A Href=""" & SiteURL & "ViewItem.asp?Entry=" & Records("RecordID") & """> "
	PluginText = PluginText & Replace(Records("Title"),"""","&quot;") & "</a></li>" & VbCrlf
	Records.MoveNext
	End If
Next

Records.Close
%>
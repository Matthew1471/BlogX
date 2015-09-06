<%
'--- File Format Of OtherLinks.Txt ----'
'*
'<TITLE>
'<URL>
'--- End Of Help ----'

'-- Maybe Stop Webhosts not supporting FSO from erroring --'
On Error Resume Next


' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

'---Map The Physical System Path---'

If FSO.FileExists(LinksPath) Then

	' Get a handle to the file
	Dim File	
	set File = FSO.GetFile(LinksPath)

	' Open the file
	Dim TextStream, Count
        Count = 0

	'Check It's Divisable By 3
        Set TextStream = File.OpenAsTextStream(1, -2)
	Do While Not TextStream.AtEndOfStream
        TextStream.Skipline
        Count = Count + 1
	Loop
        TextStream.Close


        ' Read the file line by line
	Set TextStream = File.OpenAsTextStream(1, -2)
        If (Count > 0) AND (InStr(Count/3, ".") = 0) Then
        Dim Name
        Dim URL
	Do While Not TextStream.AtEndOfStream
		If TextStream.Readline = "*" Then
                Name = TextStream.Readline
                URL = TextStream.Readline
                Response.Write "<Li><a href=""" & URL & """>" & Name & "</a></Li>"
                End If
	Loop
        Else
        Response.Write "<Li>No Links (File Broken)</Li>"
        End If
        
        TextStream.Close
	Set TextStream = nothing
	
Else
	Response.Write "<Li>No Links (File Not Found)</Li>"
End If

Set FSO = nothing

'-- Don't Ignore Any Other Errors --'
On Error GoTO 0
%>
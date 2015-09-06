<%
'---------- Random Quote Plugin (V1.02) --------
'//= - - - - - - - 
'// Copyright 2004, Matthew Roberts
'// 
'// Usage Of This Software Is Subject To The Terms Of The License
'//= - - - - - - -

Dim Random, QuotesFile, thisLine

QuotesFile = "D:\inetpub\wwwroot\BlogX\Includes\Quotes.txt"
PluginTitle = "Random Quotes"
Count = 0

Randomize

On Error Resume Next
Set FSO = CreateObject("Scripting.FileSystemObject")

'-- Attempt to guess --'
If NOT FSO.FileExists(QuotesFile) Then QuotesFile = Replace(Server.MapPath("Includes\Quotes.txt"),"Admin\","",1)

If FSO.FileExists(QuotesFile) Then
	' Get a handle to the file	
	Set File = FSO.GetFile(QuotesFile)

        ' Read the file line by line
	Set TextStream = File.OpenAsTextStream(1, -2)
	Do While Not TextStream.AtEndOfStream 
	ThisLine = TextStream.ReadLine
	Count = Count + 1
	Loop
	TextStream.Close

	'-- Pick A Random Line --'
	Random = Int(Rnd * Count)

	Set TextStream = File.OpenAsTextStream(1, -2)
	Count = 0

	Do While Not TextStream.AtEndOfStream
	ThisLine = TextStream.ReadLine

		If Count = Random Then 
		
		SplitText = Split(Trim(ThisLine)," ") 

		For WordLoopCounter = 0 To UBound(SplitText)
		 PluginText = PluginText & " " & SplitText(WordLoopCounter)
       		 If (Int(WordLoopCounter / 4) = (WordLoopCounter / 4)) AND (WordLoopCounter > 0) Then PluginText = PluginText & "<br/>" & VbCrlf
		Next

		Exit Do

		End If
  
	Count = Count + 1
	Loop

	PluginText = "<p style=""text-align: center"">" & PluginText & "</p>"	

	TextStream.Close
	Set TextStream = Nothing
	Set File = Nothing
Else
	PluginText = "<li>No ""Quotes.txt"" File Found (Edit Plugin.asp)" & VbCrlf & "<!-- Looked in : " & QuotesFile & VbCrlf & "Probably meant: " & Replace(Server.MapPath("Includes\Quotes.txt"),"Admin\","",1) & "--></li>"
End If

Set FSO = Nothing
%>
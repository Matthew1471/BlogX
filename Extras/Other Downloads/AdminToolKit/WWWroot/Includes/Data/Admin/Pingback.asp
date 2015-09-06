<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/xmlrpc.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<html>
<head>
<title>PingBack!</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<% If Request.Querystring("Step") < 3 Then Response.Write "<META HTTP-EQUIV=Refresh CONTENT=""0; URL=PingBack.asp?Step=" & Int(Request.Querystring("Step")) + 1 & """>" %>
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<br>
<Center>
<%
Select Case Request.Querystring("Step")

 Case 0
	Response.Write "<b>Gathering Links..</b><br><br>"
 Case 1
	Response.Write "<b>Gathering Content Providers..</b><br><br>"

		'--- Open Input Text ---'
    		Records.Open "SELECT RecordID, Text FROM Data ORDER BY RecordID DESC",Database, 1, 3

    		If NOT Records.EOF Then 
		strInput = Records("Text")
		RecordID = Records("RecordID")
		End If

    		Records.Close

		'--- End Input Gathering ---'

		Dim iCurrentLocation  ' Our current position in the input string
		Dim iLinkStart        ' Beginning position of the current link
		Dim iLinkEnd          ' Ending position of the current link
		Dim strLinkText       ' Text we're converting to a link

		' Start at the first character in the string
		iCurrentLocation = 1

		' Look for http:// in the text from the current position to
		' the end of the string.  If we find it then we start the
		' linking process otherwise we're done because there are no
		' more http://'s in the string.
		Do While (InStr(iCurrentLocation, strInput, "http://", 1) <> 0)

		' Set the position of the beginning of the link
		iLinkStart = InStr(iCurrentLocation, strInput, "http://", 1)

			' Set the position of the end of the link.  I use the
			' first space as the determining factor.
                	SpacePos = InStr(iLinkStart + 1, strInput, " ", 1)
                	EnterPos = InStr(iLinkStart + 1, strInput, VbCrlf, 1)
                	HrefPos  = InStr(iLinkStart + 1, strInput, """", 1)

	  		'SpacePos = 24
			'EnterPos = 32
			'HrefPos = 20

			If (SpacePos < EnterPos) OR (EnterPos = 0) Then
			   If (SpacePos < HrefPos) OR (HrefPos = 0) Then iLinkEnd = SpacePos Else iLinkEnd = HrefPos
			Else
			   If (EnterPos < HrefPos) OR (HrefPos = 0) Then iLinkEnd = EnterPos Else iLinkEnd = HrefPos
                	End If
		
			' If we didn't find a space then we link to the
			' end of the string
			If iLinkEnd = 0 Then iLinkEnd = Len(strInput) + 1

			' Take care of any punctuation we picked up
			Select Case Mid(strInput, iLinkEnd - 1, 1)
				Case ".", "!", "?", ")", "(", ","
					iLinkEnd = iLinkEnd - 1
			End Select
	
			' Get the text we're linking and store it in a variable
			strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)
			
			' Build our link and append it to the output string
			If StrLinkTest <> "" Then StrLinkTest = StrLinkTest & VbCrlf
			strLinkTest = strLinkTest & strLinkText

			' Some good old debugging
			'Response.Write iLinkStart & "," & iLinkEnd & "<BR>" & vbCrLf
		
			' Reset our current location to the end of that link
			iCurrentLocation = iLinkEnd
		Loop
	
		' Set the return value
		If strLinkTest <> "" Then Response.Write Replace(strLinkTest,VbCrlf,"<br>" & VbCrlf) & "<br><br>"

                Dim FSO, File 
                Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
                Set File = FSO.CreateTextFile(Server.MapPath("..\Images\Articles\Temp") & "\LastPingBack.txt",true) 
                File.Write(strLinkTest)
                File.Close
                Set File = nothing
                Set FSO = nothing

 Case 2
	Response.Write "<b>Pinging Back Resources..</b><br><br>" 
		Dim objXMLHTTP, PingBackServer, CurrentLine, CurrentLine2
		
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		Set File = FSO.OpenTextFile(Server.MapPath("..\Images\Articles\Temp") & "\LastPingBack.txt", 1)

		'--- Open Input Text ---'
    		Records.Open "SELECT * FROM Data ORDER BY RecordID DESC",Database, 1, 3

    		If NOT Records.EOF Then 
		strInput = Replace(Records("Text"),VbCrlf,"<br>")
		RecordID = Records("RecordID")
		End If

    		Records.Close

		'--- End Input Gathering ---'

		Do While NOT File.AtEndOfStream
		PingBackServer = ""
		CurrentLine = File.ReadLine

			'---- Verify The Site Is PingBack'd! ----
			' If you don't have MSXML3 installed you can revert to the old line:
			Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
			'Set objXMLHTTP=Server.CreateObject("MICROSOFT.XMLHTTP")

			objXMLHTTP.open "GET", CurrentLine, true
  			objXMLHTTP.SetRequestHeader "User-Agent", "Matthew1471 BlogX"
  			On Error Resume Next 
  			objXMLhttp.send()

			'Wait for up to 3 seconds if we've not gotten the data yet
  			If objXMLHTTP.readyState <> 4 then
    			objXMLHTTP.waitForResponse 5
 			End If

  			If Err.Number <> 0 Then 
    			Response.Write "XMLhttp error " & Hex(Err.Number) & " " & Err.Description 
  	       		ElseIf objXMLhttp.status <> 200 Then 
    	       		Response.Write "http error " & CStr(objXMLhttp.status) & " " & objXMLhttp.statusText 
  			Else 

			'Abort the XMLHttp request
			If (objXMLhttp.readyState <> 4) Or (objXMLhttp.Status <> 200) Then objXMLhttp.Abort

                        ' -- Debugging --'
			'Response.Write objXMLhttp.ResponseText

                	 If objXMLHTTP.getResponseHeader("X-Pingback") <> "" Then 
                	 PingBackServer = objXMLHTTP.getResponseHeader("X-Pingback")

	      		 '--- QuickCheck Tm ---'
                         ElseIf Instr(1, objXMLhttp.responseText,"pingback", 1) <> 0 Then 

                         '--- Save The Downloaded Page ---'
                         Set objFSO = CreateObject("Scripting.FileSystemObject") 
    			 Set objTS = objFSO.CreateTextFile(Server.MapPath("..\Images\Articles\Temp") & "\LastPingBackSite.txt", True) 
    			 objTS.Write Replace(objXMLhttp.ResponseText, vbLf, vbNewLine) 
                         objTS.Close 
    		         Set objTS = Nothing 
    	            	 Set objFSO = Nothing
  		         Set objXMLHTTP = Nothing
		         '--- End of saving ---'

		 	 '--- It's Parsing Time! ---'
                         '<link rel="pingback" href="pingback server">
	                 '<link rel="pingback" href="pingback server" />

		         Set File2 = FSO.OpenTextFile(Server.MapPath("..\Images\Articles\Temp") & "\LastPingBackSite.txt", 1)
		 
		             Do While (NOT File2.AtEndOfStream) AND (PingbackServer = "")

                                      CurrentLine2 = File2.ReadLine
                                      If Instr(1, CurrentLine2,"pingback",1) <> 0 Then
                                      TempPingBackServer = Replace(CurrentLine2,"<link rel=""pingback"" href=""","",1,-1,1)
                                      TempPingBackServer = Replace(TempPingBackServer,""" />","")
                                      PingBackServer = Replace(TempPingBackServer,""">","")
                                      End If

                             Loop
			 File2.Close
			 Set File2=Nothing

		         End If
			End If

                        '--- End of Parsing Time! ---'

				     If Len(PingBackServer) > 0 Then
	 			     ReDim paramList(2)
	 			     paramList(0)= SiteURL & "ViewItem.asp?Entry=" & RecordID
	 		             paramList(1)= CurrentLine

         			     Call(xmlRPC (PingBackServer, "pingback.ping", paramList))

         			     Response.Write(CurrentLine & VbCrlf)
	 		             Response.Write("<br>")
				     End If
		Loop

                Response.Write("<br>")
                
		File.Close
		Set File=Nothing
		Set FSO=Nothing

		'-- More Debugging --
		'Response.write PingBackServer
		'Response.write("<pre>" & Replace(serverResponseText, "<", "&lt;", 1, -1, 1) & "</pre>")

 Case 3
	Response.Write "<b>Done...</b><br><br>"
	Response.Write "<Script>JavaScript:self.close();</script>"
End Select

Database.Close
Set Records = Nothing
Set Database = Nothing
%>

<a href="JavaScript:self.close();">Close</a>
</Center>

</Body>
</html>
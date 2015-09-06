<%
Function HTML2Text(strInput)
strInput  = Replace(strInput,"<","&lt;")
strInput  = Replace(strInput,">","&gt;")
strInput  = Replace(strInput,"%20"," ")
HTML2Text = Replace(strInput,"""","&quot;")
End Function

Function LinkURLs(strInput)
	Dim iCurrentLocation  ' Our current position in the input string
	Dim iLinkStart        ' Beginning position of the current link
	Dim iLinkEnd          ' Ending position of the current link
	Dim strLinkText       ' Text we're converting to a link
	Dim strOutput         ' Return string with links in it

	Dim SpacePos, EnterPos, BracketPos

	' Start at the first character in the string
	iCurrentLocation = 1

	' Look for http:// in the text from the current position to
	' the end of the string.  If we find it then we start the
	' linking process otherwise we're done because there are no
	' more http://'s in the string.
	Do While _
	(InStr(iCurrentLocation, strInput, " http://", 1) <> 0) OR _
        (InStr(iCurrentLocation, strInput, VbCrlf & "http://", 1) <> 0) OR _
        (Instr(iCurrentLocation, strInput, "(http://", 1) <> 0) OR _
	(Instr(iCurrentLocation, strInput, " irc://", 1) <> 0)

		' Set the position of the beginning of the link
		SpacePos = InStr(iCurrentLocation, strInput, " http://", 1)
		EnterPos = InStr(iCurrentLocation, strInput, VbCrlf & "http://", 1)
	        BracketPos = InStr(iCurrentLocation, strInput, Chr(40) & "http://", 1)

                If ((SpacePos < EnterPos) OR (EnterPos = 0)) AND (SpacePos <> 0) Then
                    If ((SpacePos < BracketPos) OR (BracketPos = 0)) AND (SpacePos <> 0) Then iLinkStart = SpacePos + 1 Else iLinkStart = BracketPos + 1
                Else
                    If ((EnterPos < BracketPos) OR (BracketPos = 0)) AND (EnterPos <> 0) Then iLinkStart = EnterPos + 2 Else iLinkStart = BracketPos + 1
                End If

		If (Instr(iCurrentLocation, strInput, " irc://", 1) > iLinkStart) AND (iLinkStart <> 0) Then iLinkStart = Instr(iCurrentLocation, strInput, " irc://", 1) + 1

		'------- DEBUG -----------------------------------------------------'
		'If Request.Querystring <> "Test" Then
		'Response.Write "Testing (Sorry for any inconvenience) :" & iLinkStart
		'Records.Close
		'Database.close
		'Set Records = Nothing
		'Set Database = Nothing
		'Response.End
		'End If 

	        'iLinkStart = InStr(iCurrentLocation, strInput, "http://", 1)
		'--- END OF DEBUGGING -------------------------------------------'

		' Set the position of the end of the link.  I use the
		' first space as the determining factor.
                SpacePos = InStr(iLinkStart + 1, strInput, " ", 1)
                EnterPos = InStr(iLinkStart + 1, strInput, VbCrlf, 1)

		If ((SpacePos < EnterPos) OR (EnterPos = 0)) AND (SpacePos <> 0) Then
		    iLinkEnd = SpacePos
                Else
                    iLinkEnd = EnterPos - 4
                End If

		' If we didn't find a space then we link to the
		' end of the string
		If iLinkEnd < 0 Then iLinkEnd = Len(strInput) + 1

		' Take care of any punctuation we picked up
		Select Case Mid(strInput, iLinkEnd - 1, 1)
			Case ".", "!", "?", ")", ",", VbCrlf
				iLinkEnd = iLinkEnd - 1
		End Select

		' This adds to the output string all the non linked stuff
		' up to the link we're curently processing.
		strOutput = strOutput & Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation)
		'Response.Write iCurrentLocation & " " & iLinkEnd - iLinkStart

		' Get the text we're linking and store it in a variable
		strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)

		' Build our link and append it to the output string
		strOutput = strOutput & "<A HREF=""" & strLinkText & """ Target=""_New"">" & strLinkText & "</A>"

		' Some good old debugging
		'Response.Write iLinkStart & "," & iLinkEnd & "<BR>" & vbCrLf

		' Reset our current location to the end of that link
		iCurrentLocation = iLinkEnd
	Loop

	' Tack on the end of the string.  I need to do this so we
	' don't miss any trailing non-linked text
	strOutput = strOutput & Mid(strInput, iCurrentLocation)

        strOutput = Replace(strOutput,"(Y)","<img title=""Approve"" src=""" & SiteURL & "Images/Emoticons/Approve.gif"">")
        strOutput = Replace(strOutput,":$","<img title=""Blush"" src=""" & SiteURL & "Images/Emoticons/Blush.gif"">")
        strOutput = Replace(strOutput,"(H)","<img title=""Cool"" src=""" & SiteURL & "Images/Emoticons/Cool.gif"">")
        strOutput = Replace(strOutput,"(Clown)","<img title=""Clown"" src=""" & SiteURL & "Images/Emoticons/Clown.gif"">")
        strOutput = Replace(strOutput,"(X)","<img title=""Dead"" src=""" & SiteURL & "Images/Emoticons/Dead.gif"">")
        strOutput = Replace(strOutput,"(D)","<img title=""Depressed"" src=""" & SiteURL & "Images/Emoticons/Depressed.gif"">")
        strOutput = Replace(strOutput,"(6)","<img title=""Evil"" src=""" & SiteURL & "Images/Emoticons/Evil.gif"">")
        strOutput = Replace(strOutput,"(8)","<img title=""Note"" src=""" & SiteURL & "Images/Emoticons/Note.gif"">")
        strOutput = Replace(strOutput,":D","<img title=""Grin"" src=""" & SiteURL & "Images/Emoticons/Grin.gif"">")
        strOutput = Replace(strOutput,"(Hurt)","<img title=""Hurt"" src=""" & SiteURL & "Images/Emoticons/Hurt.gif"">")
        strOutput = Replace(strOutput,"(K)","<img title=""Kiss"" src=""" & SiteURL & "Images/Emoticons/Kiss.gif"">")
        strOutput = Replace(strOutput,":@","<img title=""Mad"" src=""" & SiteURL & "Images/Emoticons/Mad.gif"">")
        strOutput = Replace(strOutput,"(Mail)","<img title=""Mail"" src=""" & SiteURL & "Images/Emoticons/Mail.gif"">")
        strOutput = Replace(strOutput,"(Entry)","<img title=""Post"" src=""" & SiteURL & "Images/Emoticons/Post.gif"">")
        strOutput = Replace(strOutput,"(User)","<img title=""Profile"" src=""" & SiteURL & "Images/Emoticons/Profile.gif"">")
        strOutput = Replace(strOutput,"(?)","<img title=""Question"" src=""" & SiteURL & "Images/Emoticons/Question.gif"">")
        strOutput = Replace(strOutput,":(","<img title=""Sad"" src=""" & SiteURL & "Images/Emoticons/Sad.gif"">")
        strOutput = Replace(strOutput,":)","<img title=""Smile"" src=""" & SiteURL & "Images/Emoticons/Smile.gif"">")
        strOutput = Replace(strOutput,":-O","<img title=""Shock"" src=""" & SiteURL & "Images/Emoticons/Shock.gif"">")
        strOutput = Replace(strOutput,"(Shy)","<img title=""Shy"" src=""" & SiteURL & "Images/Emoticons/Shy.gif"">")
        strOutput = Replace(strOutput,"^_^","<img title=""Sleepy"" src=""" & SiteURL & "Images/Emoticons/Sleepy.gif"">")
        strOutput = Replace(strOutput,"(*)","<img title=""Star"" src=""" & SiteURL & "Images/Emoticons/Star.gif"">")
        strOutput = Replace(strOutput,":P","<img title=""Tongue"" src=""" & SiteURL & "Images/Emoticons/Tongue.gif"">")
        strOutput = Replace(strOutput,"(URL)","<img title=""URL"" src=""" & SiteURL & "Images/Emoticons/URL.gif"">")
        strOutput = Replace(strOutput,";-)","<img title=""Wink"" src=""" & SiteURL & "Images/Emoticons/Wink.gif"">")

	' Set the return value
	LinkURLs = strOutput
End Function 'LinkURLs
'***** END FUNCTIONS *****
%>
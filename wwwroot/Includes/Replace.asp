<%
' --------------------------------------------------------------------------
'¦Introduction : Replace Functions.                                         ¦
'¦Purpose      : Provides functionality to convert line feeds to HTML, strip¦
'¦               HTML from users' comments, turn URLs into clickable links, ¦
'¦               handle emoticons and truncates entries.                    ¦
'¦Requires     : Nothing.                                                   ¦
'¦Used By      : Most pages.                                                ¦
'---------------------------------------------------------------------------

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************

Function HTML2Text(strInput)
 strInput  = Replace(strInput,"&","&amp;")
 strInput  = Replace(strInput,"<","&lt;")
 strInput  = Replace(strInput,">","&gt;")
 strInput  = Replace(strInput,"%20"," ")
 HTML2Text = Replace(strInput,"""","&quot;")
End Function

'-- Source: http://www.devx.com/vb2themax/Tip/19160 --'
' Convert a string so that it can be used on a URL query string
' Server.URLEncode is too over zealous.
'-----------------------------------------------------'

Function StandardURL(Text)
 Dim i
 Dim acode
 Dim char	
    
 StandardURL = Text
    
 For i = Len(StandardURL) To 1 Step -1
  acode = Asc(Mid(StandardURL, i, 1))
  '-- Do not modify alphanumeric characters, or /, or :
  If (acode < 46 OR acode > 58) AND (acode < 65 or acode > 90) AND (acode < 97 or acode > 122) AND acode <> 61 AND acode <> 63 AND acode <> 26 Then
   '-- Replace punctuation characters with "%hex" --'
   StandardURL = Left(StandardURL, i - 1) & "%" & Hex(acode) & Mid(StandardURL, i + 1)
  End If
 Next
   
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

		' Set the position of the beginning of the line
		SpacePos = InStr(iCurrentLocation, strInput, " http://", 1)

		EnterPos = InStr(iCurrentLocation, strInput, VbCrlf & "http://", 1)
	        BracketPos = InStr(iCurrentLocation, strInput, Chr(40) & "http://", 1)

                If ((SpacePos < EnterPos) OR (EnterPos = 0)) AND (SpacePos <> 0) Then
                    If ((SpacePos < BracketPos) OR (BracketPos = 0)) AND (SpacePos <> 0) Then iLinkStart = SpacePos + 1 Else iLinkStart = BracketPos + 1
                Else
                    If ((EnterPos < BracketPos) OR (BracketPos = 0)) AND (EnterPos <> 0) Then iLinkStart = EnterPos + 2 Else iLinkStart = BracketPos + 1
                End If

		If (Instr(iCurrentLocation, strInput, " irc://", 1) > iLinkStart) AND (iLinkStart <> 0) Then iLinkStart = Instr(iCurrentLocation, strInput, " irc://", 1) + 1

		' Set the position of the end of the link.  I use the
		' first space as the determining factor.
                BracketPos = InStr(iLinkStart + 1, strInput, ")", 1)
                SpacePos = InStr(iLinkStart + 1, strInput, " ", 1)
                EnterPos = InStr(iLinkStart + 1, strInput, VbCrlf, 1)

		If ((SpacePos < EnterPos) OR (EnterPos = 0)) AND (SpacePos > 0) Then
		  iLinkEnd = SpacePos
              Else
                '-- This takes into account that a VbCrlf has a "<br/>" before it.
                iLinkEnd = EnterPos - 5
              End If

		'-- It's always possible there's a bracket and then a full stop.. so fix this! --'
		If (BracketPos < iLinkEnd) AND (BracketPos > 0) Then iLinkEnd = BracketPos

		' If we didn't find a space then we link to the
		' end of the string
		If iLinkEnd <= 0 Then iLinkEnd = Len(strInput) + 1

		' Take care of any punctuation we picked up
		Select Case Mid(strInput, iLinkEnd - 1, 1)
			Case ".", "!", "?", ",", VbCrlf
				iLinkEnd = iLinkEnd - 1
		End Select

        '-- Take care of a "<" if we ended up linking it in --'
		If Instr(Mid(strInput, iLinkStart, iLinkEnd - iLinkStart),"<") Then iLinkEnd = InStr(iLinkStart + 1, strInput, "<", 1)

		' This adds to the output string all the non linked stuff
		' up to the link we're curently processing.
		strOutput = strOutput & Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation)
		'Response.Write iCurrentLocation & " " & iLinkEnd - iLinkStart

		' Get the text we're linking and store it in a variable
		strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)

		' Build our link and append it to the output string
		strOutput = strOutput & "<a href=""" & strLinkText & """>" & strLinkText & "</a>"

		' Some good old debugging
		'Response.Write iLinkStart & "," & iLinkEnd & "<br/>" & vbCrLf

		' Reset our current location to the end of that link
		iCurrentLocation = iLinkEnd
	Loop

	' Tack on the end of the string.  I need to do this so we
	' don't miss any trailing non-linked text
	strOutput = strOutput & Mid(strInput, iCurrentLocation)

        strOutput = Replace(strOutput," (Y)"," <img alt=""Approve"" src=""" & SiteURL & "Images/Emoticons/Approve.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," :$"," <img alt=""Blush"" src=""" & SiteURL & "Images/Emoticons/Blush.gif""/>")
        strOutput = Replace(strOutput," (H)"," <img alt=""Cool"" src=""" & SiteURL & "Images/Emoticons/Cool.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (Clown)"," <img alt=""Clown"" src=""" & SiteURL & "Images/Emoticons/Clown.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (X)"," <img alt=""Dead"" src=""" & SiteURL & "Images/Emoticons/Dead.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (D)"," <img alt=""Depressed"" src=""" & SiteURL & "Images/Emoticons/Depressed.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (6)"," <img alt=""Evil"" src=""" & SiteURL & "Images/Emoticons/Evil.gif""/>")
        strOutput = Replace(strOutput," (8)"," <img alt=""Note"" src=""" & SiteURL & "Images/Emoticons/Note.gif""/>")
        strOutput = Replace(strOutput," :D"," <img alt=""Grin"" src=""" & SiteURL & "Images/Emoticons/Grin.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (Hurt)"," <img alt=""Hurt"" src=""" & SiteURL & "Images/Emoticons/Hurt.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (K)"," <img alt=""Kiss"" src=""" & SiteURL & "Images/Emoticons/Kiss.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," :@"," <img alt=""Mad"" src=""" & SiteURL & "Images/Emoticons/Mad.gif""/>")
        strOutput = Replace(strOutput," (Mail)"," <img alt=""Mail"" src=""" & SiteURL & "Images/Emoticons/Mail.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (Entry)"," <img alt=""Post"" src=""" & SiteURL & "Images/Emoticons/Post.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (User)"," <img alt=""Profile"" src=""" & SiteURL & "Images/Emoticons/Profile.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (?)"," <img alt=""Question"" src=""" & SiteURL & "Images/Emoticons/Question.gif""/>")
        strOutput = Replace(strOutput," :("," <img alt=""Sad"" src=""" & SiteURL & "Images/Emoticons/Sad.gif""/>")
        strOutput = Replace(strOutput," :)"," <img alt=""Smile"" src=""" & SiteURL & "Images/Emoticons/Smile.gif""/>")
        strOutput = Replace(strOutput," :-O"," <img alt=""Shock"" src=""" & SiteURL & "Images/Emoticons/Shock.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (Shy)"," <img alt=""Shy"" src=""" & SiteURL & "Images/Emoticons/Shy.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," ^_^"," <img alt=""Sleepy"" src=""" & SiteURL & "Images/Emoticons/Sleepy.gif""/>")
        strOutput = Replace(strOutput," (*)"," <img alt=""Star"" src=""" & SiteURL & "Images/Emoticons/Star.gif""/>")
        strOutput = Replace(strOutput," :P"," <img alt=""Tongue"" src=""" & SiteURL & "Images/Emoticons/Tongue.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," :-P"," <img alt=""Tongue"" src=""" & SiteURL & "Images/Emoticons/Tongue.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," (URL)"," <img alt=""URL"" src=""" & SiteURL & "Images/Emoticons/URL.gif""/>",1,-1,1)
        strOutput = Replace(strOutput," ;-)"," <img alt=""Wink"" src=""" & SiteURL & "Images/Emoticons/Wink.gif""/>")

	'strOutput = Replace(strOutput, "><br/>", ">",1,-1,1)

	' Set the return value
	LinkURLs = strOutput
End Function 'LinkURLs

'-- This Truncates Long Entries To Bandwith Friendly Feeds --'
Function ShortenEntry(Variable,RecordID)

 If Len(Variable) > 256 Then 

  '-- 1. Check if the small excerpt has an open bracket --'
  Dim OpenTagStart
  OpenTagStart = Instr(Left(Variable,256),"<")

  '-- This isn't perfect as for e.g an IMG could throw it out to include the next tag and its close tag (this is true with any other non closing tag) --'
  '-- Note : Broken with nested tags --'
  If (OpenTagStart <> 0) Then

   '-- 2. Check there's a start of a close bracket in the rest --'
   Dim CloseTagStart
   CloseTagStart = Instr(OpenTagStart,Variable,"</")

	'-- Loop looking for next open tag.. until we run out of characters, there are none, or it exceeds our current end --'
	'-- If there's a new one, find the next end (past the end we thought) --'
	Do While ((Instr(OpenTagStart+1,Variable,"<") > 0) AND (Instr(OpenTagStart+1,Variable,"<") < CloseTagStart))

	 OpenTagStart = Instr(OpenTagStart+1,Variable,"<")

	 '-- It is possible our end tag was on track --'
	 If Instr(CloseTagStart+1,Variable,"</") <> 0 Then CloseTagStart = Instr(CloseTagStart+1,Variable,"</")

	Loop

   '-- 3. Check that there's an end of a close bracket (we rely on this being true later) --'
   Dim CloseTagEnd
   If (CloseTagStart <> 0) Then CloseTagEnd = Instr(CloseTagStart,Variable,">")

   If (CloseTagEnd <> 0) Then
    '-- 4. Append --'
    ShortenEntry = Left(Variable,CloseTagEnd)
    If (Len(Variable) - CloseTagEnd) > 0 Then 
     ShortenEntry = ShortenEntry & "..." & vbcrlf 
     If RecordID <> 0 Then ShortenEntry = ShortenEntry & "<a href=""" & SiteURL & "ViewItem.asp?Entry=" & RecordID & """>"

     ShortenEntry = ShortenEntry & "Read More (" & (Len(Variable) - CloseTagEnd) & " Characters)"
     If RecordID <> 0 Then ShortenEntry = ShortenEntry & "</a>"
    End If
   Else
    '-- Though there's a start tag, it never closes! so it's probably a self closer! --'
    ShortenEntry = Left(Variable,256) & "..." & vbcrlf 
    If RecordID <> 0 Then ShortenEntry = ShortenEntry & "<a href=""" & URLEncode(SiteURL) & "ViewItem.asp?Entry=" & RecordID & """>"

    ShortenEntry = ShortenEntry & "Read More (" & (Len(Variable) - 256) & " Characters)"
    If RecordID <> 0 Then ShortenEntry = ShortenEntry & "</a>"
   End If

  Else
   ShortenEntry = Left(Variable,256) & "..." & vbcrlf
   If RecordID <> 0 Then ShortenEntry = ShortenEntry & "<a href=""" & SiteURL & "ViewItem.asp?Entry=" & RecordID & """>"
   ShortenEntry = ShortenEntry & "Read More (" & (Len(Variable) - 256) & " Characters)"
   If RecordID <> 0 Then ShortenEntry = ShortenEntry & "</a>"
  End If
 Else
  ShortenEntry = Variable
 End If
End Function
'***** END FUNCTIONS *****
%>
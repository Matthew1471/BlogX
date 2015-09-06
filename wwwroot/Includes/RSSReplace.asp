<%
' --------------------------------------------------------------------------
'¦Introduction : RSS Replace Functions.                                     ¦
'¦Purpose      : Provides functionality to convert characters to RSS and    ¦
'¦               truncates entries.                                         ¦
'¦Requires     : Nothing.                                                   ¦
'¦Used By      : RSS/, RSS/Cat/, RSS/Comments.                              ¦
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

'-- This Truncates Long Entries To Bandwith Friendly Feeds --'
Function ShortEncode(Variable)

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
    ShortEncode = Encode(Left(Variable,CloseTagEnd))
    If (Len(Variable) - CloseTagEnd) > 0 Then 
     ShortEncode = ShortEncode & "..." & vbcrlf 
     If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("<a href=""" & URLEncode(SiteURL) & "ViewItem.asp?Entry=" & RecordID & """>")
     ShortEncode = ShortEncode & "Read More (" & (Len(Variable) - CloseTagEnd) & " Characters)"
     If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("</a>")
    End If
   Else
    '-- Though there's a start tag, it never closes! so it's probably a self closer! --'
    ShortEncode = Encode(Left(Variable,256)) & "..." & vbcrlf 
    If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("<a href=""" & URLEncode(SiteURL) & "ViewItem.asp?Entry=" & RecordID & """>")
    ShortEncode = ShortEncode & "Read More (" & (Len(Variable) - 256) & " Characters)"
    If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("</a>")
   End If

  Else
   ShortEncode = Encode(Left(Variable,256)) & "..." & vbcrlf
   If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("<a href=""" & URLEncode(SiteURL) & "ViewItem.asp?Entry=" & RecordID & """>")
   ShortEncode = ShortEncode & "Read More (" & (Len(Variable) - 256) & " Characters)"
   If RecordID <> 0 Then ShortEncode = ShortEncode & Encode("</a>")
  End If
 Else
  ShortEncode = Encode(Variable)
 End If
End Function

Function Encode(Variable)

If Variable <> "" Then

   Dim i

   Encode = Replace(Variable, "Images/Articles/",SiteURL & "/Images/Articles/")
   Encode = Replace(Encode, vbcrlf,"<br>")
   Encode = Replace(Encode, "&","&amp;")
   Encode = Replace(Encode, "'","&#39;")
   Encode = Replace(Encode, "…","...")

   For i = 0 To 31
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 33 To 34
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 37 To 37
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 39 To 47
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 58 To 58
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 60 To 64
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 91 To 96
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next

   For i = 123 To 255
   Encode = Replace(Encode, Chr(i), "&#" & i & ";")
   Next
End If

End function

'--- Let's URLEncode certain characters (Server.URLEncode converts too much) ---'
Function URLEncode(Variable)

If Variable <> "" Then
   URLEncode = Replace(Variable, " ","%20")
   URLEncode = Replace(URLEncode, "'","%27")
   URLEncode = Replace(URLEncode, "’","%92")
End If

End function
'--- End ---'
%>
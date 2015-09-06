<%
' --------------------------------------------------------------------------
'¦Introduction : Pingback Popup Page.                                       ¦
'¦Purpose      : Gathers link, pingback server URL and lets server know that¦
'¦               we linked to them.                                         ¦
'¦Used By      : Admin/AddEntry.asp.                                        ¦
'¦Requires     : Includes/Config.asp, Templates/Config.asp, Admin.asp,      ¦
'¦               Includes/XMLRPC.asp.                                       ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

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
 
'-- Changing this value to True will prevent errors being ignored --'
Dim Debug
Debug = False

'-- If your host does not support parent paths specify the full path here --'
Dim ServerPathToInstalledDirectory
ServerPathToInstalledDirectory = Server.MapPath("..\")
'ServerPathToInstalledDirectory = "C:\inetpub\wwwroot"
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/xmlrpc.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - PingBack!</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 <% If Request.Querystring("Step") < 3 Then Response.Write "<meta http-equiv=""Refresh"" content=""0; URL=PingBack.asp?Step=" & Int(Request.Querystring("Step")) + 1 & """/>" %>
 <!--
 //= - - - - - - - 
 // Copyright 2004, Matthew Roberts
 // Copyright 2003, Chris Anderson
 //= - - - - - - -
 -->
<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
 <link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet"/>
 <!-- #INCLUDE FILE="../Templates/Config.asp" -->
</head>
<body style="background-color:<%=BackgroundColor%>">
<p style="text-align:center">
<%
Select Case Request.Querystring("Step")

 Case 0
  Response.Write "<b>Gathering Links..</b><br/><br/>" & VbCrlf
 Case 1
  Response.Write "<b>Gathering Content Providers..</b><br/><br/>" & VbCrlf

  '-- Read entry text --'
  If (Request.Querystring("Special") = "") Then
   Records.Open "SELECT RecordID, Text FROM Data ORDER BY RecordID DESC",Database, 0, 1
  Else
   Dim Requested
   Requested = Replace(Request.Querystring("Special"),"'","")
   If (NOT IsNumeric(Requested)) Then Requested = 0
   Records.Open "SELECT RecordID, Text FROM Data WHERE RecordID=" & Requested & " ORDER BY RecordID DESC",Database, 0, 1
  End If

  Dim strInput, RecordID

  If NOT Records.EOF Then 
   strInput = Records("Text")
   RecordID = Records("RecordID")
  End If

  Records.Close

  '-- End entry reading --'
  Dim iCurrentLocation  ' Our current position in the input string
  Dim iLinkStart        ' Beginning position of the current link
  Dim iLinkEnd          ' Ending position of the current link
  Dim strLinkText       ' Text we're converting to a link
  Dim StrLinkTest
  
  Dim SpacePos, EnterPos, HrefPos
  
  '-- Start at the first character in the string --'
  iCurrentLocation = 1

  ' Look for http:// in the text from the current position to
  ' the end of the string.  If we find it then we start the
  ' linking process otherwise we're done because there are no
  ' more http://'s in the string.
  Do While (InStr(iCurrentLocation, strInput, "http://", 1) <> 0)

   '-- Set the position of the beginning of the link --'
   iLinkStart = InStr(iCurrentLocation, strInput, "http://", 1)

   'Set the position of the end of the link.  I use a
   'variety of characters as the determining factor.
   SpacePos = InStr(iLinkStart + 1, strInput, " ", 1)
   EnterPos = InStr(iLinkStart + 1, strInput, VbCrlf, 1)
   HrefPos  = InStr(iLinkStart + 1, strInput, """", 1)
   
   If (SpacePos < EnterPos) OR (EnterPos = 0) Then
    If (SpacePos < HrefPos) OR (HrefPos = 0) Then iLinkEnd = SpacePos Else iLinkEnd = HrefPos
   Else
    If (EnterPos < HrefPos) OR (HrefPos = 0) Then iLinkEnd = EnterPos Else iLinkEnd = HrefPos
   End If

   '-- If we did not find a space then we link to the end of the string --'
   If iLinkEnd = 0 Then iLinkEnd = Len(strInput) + 1

   '-- Take care of any punctuation we picked up --'
    Select Case Mid(strInput, iLinkEnd - 1, 1)
     Case ".", "!", "?", ")", "(", "," : iLinkEnd = iLinkEnd - 1
    End Select

    '-- Get the text we are linking and store it in a variable --'
    strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)

    '-- Build our link and append it to the output string --'
    If StrLinkTest <> "" Then StrLinkTest = StrLinkTest & VbCrlf
    strLinkTest = strLinkTest & strLinkText

    '-- Some good old debugging --'
    'Response.Write iLinkStart & "," & iLinkEnd & "<br/>" & vbCrLf

    '-- Reset our current location to the end of that link --'
    iCurrentLocation = iLinkEnd
  Loop

  '-- Set the return value --'
  If strLinkTest <> "" Then Response.Write Replace(strLinkTest,VbCrlf,"<br/>" & VbCrlf) & "<br/><br/>"

  On Error Resume Next
   Set FSO = Server.CreateObject("Scripting.FileSystemObject")

    '-- FSO is NOT enabled --'
    If Err <> 0 Then
     Set FSO = Nothing
     Database.Close
     Set Records  = Nothing
     Set Database = Nothing
     Response.Redirect "PingBack.asp?Step=4"
    End If

  On Error Goto 0

  On Error Resume Next
   Set File = FSO.CreateTextFile(ServerPathToInstalledDirectory & "\Images\Articles\Temp\LastPingBack.txt",true)

   ' -- File Permissions --'
   If Err <> 0 Then
    Set FSO = Nothing
    Database.Close
    Set Records  = Nothing
    Set Database = Nothing
    Response.Redirect "PingBack.asp?Step=5"
   End If

   File.Write(strLinkTest)
   File.Close
  On Error GoTo 0

  Set File = Nothing
  Set FSO = Nothing
 Case 2
 
  Response.Write "<b>Pinging Back Resources..</b><br/><br/>" 
  Dim objXMLHTTP, PingBackServer, CurrentLine, CurrentLine2

  On Error Resume Next

  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
  Set File = FSO.OpenTextFile(ServerPathToInstalledDirectory & "\Images\Articles\Temp\LastPingBack.txt", 1)

  '-- Open Input Text --'
  Records.Open "SELECT RecordID, Text FROM Data ORDER BY RecordID DESC",Database, 1, 3

  If NOT Records.EOF Then 
   strInput = Replace(Records("Text"),VbCrlf,"<br/>")
   RecordID = Records("RecordID")
  End If

  Records.Close
  '--- End Input Gathering ---'

  If CStr(Err.Description) = "" Then

   Do While (NOT File.AtEndOfStream)
    PingBackServer = ""
    CurrentLine = File.ReadLine

    '-- Verify the site is PingBack'd! --'
    '-- If you don't have MSXML3 installed you can revert to the old line: --'
    Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
    'Set objXMLHTTP=Server.CreateObject("MICROSOFT.XMLHTTP")

    objXMLHTTP.open "GET", CurrentLine, true
    objXMLHTTP.SetRequestHeader "User-Agent", "Matthew1471 BlogX"

    On Error Resume Next 
     objXMLhttp.send()

     'Wait for up to 5 seconds if we've not gotten *all* the data yet
     If objXMLHTTP.readyState <> 4 Then objXMLHTTP.waitForResponse 5

    If Err.Number <> 0 Then 
      Response.Write "XMLhttp error " & Hex(Err.Number) & " " & Err.Description 
     ElseIf objXMLhttp.status <> 200 Then 
      Response.Write "http error " & CStr(objXMLhttp.status) & " " & objXMLhttp.statusText
     Else 
      '-- Abort the XMLHttp request --'
      If (objXMLhttp.readyState <> 4) Or (objXMLhttp.Status <> 200) Then objXMLhttp.Abort

      '-- Debugging --'
      'Response.Write objXMLhttp.ResponseText

      If objXMLHTTP.getResponseHeader("X-Pingback") <> "" Then 
       PingBackServer = objXMLHTTP.getResponseHeader("X-Pingback")
      '--- QuickCheck Tm ---'
      ElseIf Instr(1, objXMLhttp.responseText,"pingback", 1) <> 0 Then

       '--- Save the downloaded page ---'
       Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objTS = objFSO.CreateTextFile(ServerPathToInstalledDirectory & "\Images\Articles\Temp\LastPingBackSite.txt", True)
         objTS.Write Replace(objXMLhttp.ResponseText, vbLf, vbNewLine) 
         objTS.Close 
        Set objTS = Nothing
       Set objFSO = Nothing

       Set objXMLHTTP = Nothing
       '--- End of saving ---'

       '-- It is page parsing time! ---'
       '<link rel="pingback" href="pingback server"> |
       '<link rel="pingback" href="pingback server" /> |
       '<link rel="pingback" href="pingback server"/>

       Set File2 = FSO.OpenTextFile(ServerPathToInstalledDirectory & "\Images\Articles\Temp\LastPingBackSite.txt", 1)

        Do While (NOT File2.AtEndOfStream) AND (PingbackServer = "")
         CurrentLine2 = File2.ReadLine
         If Instr(1, CurrentLine2,"pingback",1) <> 0 Then
          TempPingBackServer = Replace(CurrentLine2,"<link rel=""pingback"" href=""","",1,-1,1)
          TempPingBackServer = Replace(TempPingBackServer,""" />","")
          TempPingBackServer = Replace(TempPingBackServer,"""/>","")
          PingBackServer = Replace(TempPingBackServer,""">","")
         End If
        Loop

       File2.Close
       Set File2=Nothing

      End If
     End If
     '-- End of Parsing Time! --'

    If Len(PingBackServer) > 0 Then
     ReDim paramList(2)
     paramList(0)= SiteURL & "ViewItem.asp?Entry=" & RecordID
     paramList(1)= CurrentLine

     Call(xmlRPC (PingBackServer, "pingback.ping", paramList))

     'Dim myresp
     'myresp = xmlRPC (PingBackServer, "pingback.ping", paramList)
     'Response.Write serverResponseText
     'Response.End

     Response.Write CurrentLine & VbCrlf
     Response.Write "<br/>"
    End If
   Loop

  Else
   Response.Write "Error Occured : " & CStr(Err.Description) & "<br/>" & VbCrlf
  End If

  Response.Write "<br/>"

  File.Close
  Set File=Nothing
  Set FSO=Nothing

  '-- More Debugging --'
  'Response.write PingBackServer
  'Response.write "<pre>" & Replace(serverResponseText, "<", "&lt;", 1, -1, 1) & "</pre>"
 Case 3
  Response.Write "<b>Done...</b><br/><br/>" & VbCrlf
  Response.Write "<script type=""text/javascript"">JavaScript:self.close();</script>" & VbCrlf
 Case 4
  Response.Write "An Error Occured:<br/>FSO is not enabled on your webhost.<br/><br/>" & VbCrlf
  If Debug = False Then Response.Write "<script type=""text/javascript"">JavaScript:self.close();</script>" & VbCrlf
 Case 5
  Response.Write "An Error Occured:<br/>IIS needs WRITE permissions on Articles/Temp/.<br/><br/>" & VbCrlf
  If Debug = False Then Response.Write "<script type=""text/javascript"">JavaScript:self.close();</script>" & VbCrlf
End Select

Database.Close
Set Records = Nothing
Set Database = Nothing
%>
 <a href="JavaScript:self.close();">Close</a>
 </p>
</body>
</html>
<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : BlogX Registration Page.                                   ¦
'¦Purpose      : Adds the user to the BlogX list of users and counts hits.  ¦
'¦Used By      : BlogX distributions.                                       ¦
'¦Requires     : None.                                                      ¦
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

'-- Path to official BlogX database --'
Dim DataFile
DataFile = "D:\Inetpub\database\BlogX.mdb"

'-- Find Out Refer --'
If Len(Request.ServerVariables("HTTP_REFERER")) > 0 Then

 Dim ReferURL

 '-- Limit the URL to 100 characters and replace any apostrophes with HTML friendly ones for security --'
 ReferURL = Replace(Left(Request.ServerVariables("HTTP_REFERER"),100),"'", "&#39;")

 '-- Trim off any admin subfolders we are in --'
 ReferURL = Replace(ReferURL,"Admin/", "")

 '-- Trim off the filename as we are only interested in the BlogX directory --'
 ReferURL = Left(ReferURL,InStrRev(ReferURL,"/"))

 '-- The user is from our own network and so does not count --'
 If Instr(Request.ServerVariables("REMOTE_ADDR"),"192.168") > 0 Then
  ReferURL = "Local Address"
 '-- The order is important as once matched IIS will not check other scenarios. --'
 ElseIf Instr(1, Request.ServerVariables("HTTP_REFERER"),"cache:",1) > 0 Then
  '-- The user is viewing a BlogX blog via the Google cache --'
  ReferURL = "Google Cache"
 ElseIf Instr(1, Request.ServerVariables("HTTP_REFERER"),"search/cache.html?",1) > 0 Then
  '-- Same but with Yahoo! instead --'
  ReferURL = "Yahoo! Cache"
 ElseIf Instr(1, Request.ServerVariables("HTTP_REFERER"),"imgurl=",1) > 0 Then
  '-- The user found us on the Google Image Search page --'
  ReferURL = "Google Image Search"
 ElseIf Instr(1, Request.ServerVariables("HTTP_REFERER"),"translate_c",1) > 0 Then
  '-- The URL would link us to a Google translation page instead of the blog --'
  ReferURL = "Google Translation"
 '-- This attempts to remove/fix some of the referers that are not BlogX blogs --'
 ElseIf Instr(1, Request.ServerVariables("HTTP_REFERER"),"localhost",1) > 0 Then
  '-- The user requested this page at the webserver, so try using their WAN IP address --'
  ReferURL = Replace(ReferURL,"localhost",Request.ServerVariables("REMOTE_ADDR"),1,-1,1)
 ElseIf Instr(Request.ServerVariables("HTTP_REFERER"),"127.0.0.1") > 0 Then
  '-- The same as above --'
  ReferURL = Replace(ReferURL,"127.0.0.1",Request.ServerVariables("REMOTE_ADDR"))
 End If

Else
 '-- The user either operates a firewall that blocks referers or they manually requested this page --'
 If Instr(Request.ServerVariables("REMOTE_ADDR"),"192.168") > 0 Then
  ReferURL = "(None)"
 Else
  ReferURL = "Local Address"
 End If
End If

'-- Update list of BlogX users --'
Dim Database
Set Database = Server.CreateObject("ADODB.Connection")
 Database.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataFile

  Dim Records
  Set Records = Server.CreateObject("ADODB.Recordset")
  Records.Open "SELECT ReferURL, ReferHits, IP FROM ScriptRefer WHERE ReferURL='" & ReferURL & "';", Database, 0, 3

   '-- Have we seen this blog before? --'
   If Not Records.EOF Then
    Records("ReferHits") = Int(Records("ReferHits")) + 1
   Else
    Records.AddNew
    Records("ReferURL") = ReferURL
    Records("ReferHits") = 1
   End If

   Records("IP") = Request.ServerVariables("REMOTE_ADDR")

   On Error Resume Next
    Records.Update

    '-- Record locking problems --'
    If Err.Number = -2147217887 Then
      
     '-- Keep trying for 3 seconds --'
     Dim EndTime
     EndTime = DateAdd("s", 3, Now())
     Do While (Now() < EndTime)
      Err.Clear
      Records.Update
       If Err.Number = 0 Then Exit Do
     Loop
    End If
  
    Dim WasError  
    If Err.Number <> 0 Then WasError = True 
   On Error GoTo 0

  '-- Force it again so we get our server error page if needs be --'
  If WasError Then Records.Update

  Records.Close
  Set Records = Nothing

 Database.Close
 Set Database = Nothing

'-- If they have ANY cached copies accept them as valid. --'
If Len(Request.ServerVariables("HTTP_IF_MODIFIED_SINCE")) > 0 Then
 Response.Status = "304 Not Modified"
 Response.End()
Else
 Response.ContentType = "image/GIF"
 Server.Transfer "\Images\Blank.gif"
End If
%>

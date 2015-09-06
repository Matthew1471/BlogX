<%
' --------------------------------------------------------------------------
'¦Introduction : Poll Vote Handler                                          ¦
'¦Purpose      : Processes The Vote Request checking the user has not       ¦
'¦               tried to vote twice, redirecting to PageName once complete.¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Config.asp.                                       ¦
'¦Notes        : This page intentionally ignores proxies, thus a user behind¦
'¦               a proxy will only be able to vote once.                    ¦
'¦Standards    : N/A.                                                       ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
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
%>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
Dim PollID

If Request.Form("Vote") <> "" Then

 '-- Learn current PollID --'
 Records.Open "SELECT PollID FROM Poll ORDER BY PollID DESC",Database, 0, 1
  If Records.EOF = False Then PollID = Records("PollID") Else PollID = 0
 Records.Close

 '-- Have they already voted? --'
 Records.Open "SELECT VoteID FROM Votes WHERE PollID="& PollID & "AND IP='" & Request.ServerVariables("REMOTE_ADDR") & "'", Database, 0, 1

  If Records.EOF = False Then
   Records.Close
   Database.Close
   Set Records = Nothing
   Set Database = Nothing
   Response.Redirect(PageName)
  End If
 
 Records.Close

 '-- Write vote in log --'
 Records.Open "SELECT PollID, IP, Option FROM Votes", Database, 0, 3
 Records.AddNew

  Records("PollID") = PollID
  Records("IP") = Request.ServerVariables("REMOTE_ADDR")

   '-- Filter & Clean --'
   Dim OptionNumber
   OptionNumber = Request.Form("Vote")
   If IsNumeric(OptionNumber) Then OptionNumber = Int(OptionNumber) Else OptionNumber = 0

   Records("Option") = OptionNumber

  Records.Update
 Records.Close

 '-- Update actual poll vote count --'
 Records.Open "SELECT PollID, Op1, Op2, Op3, Op4, Total FROM Poll WHERE PollID=" & PollID, Database

 Select Case OptionNumber
  Case 1: Records("Op1") = Records("Op1") + 1
  Case 2: Records("Op2") = Records("Op2") + 1
  Case 3: Records("Op3") = Records("Op3") + 1
  Case 4: Records("Op4") = Records("Op4") + 1
 End Select

 If OptionNumber <= 4 Then Records("Total") = Records("Total") + 1

 Records.Update
 Records.Close

End If

Database.Close
Set Records = Nothing
Set Database = Nothing
Response.Redirect(PageName)
%>
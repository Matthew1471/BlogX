<%
' -------------------------------------------------
'¦Introduction : Application API                   ¦
'¦Purpose      : Allow programs to upload entries. ¦
'¦Used By      : WinBlogX (VB client).             ¦
'¦Requires     : Includes/Config.asp               ¦
'¦Standards    : BlogX API 0.1.                    ¦
'--------------------------------------------------

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
'-- Check for ban --'
Records.Open "SELECT BannedIP, LoginFailCount, LastLoginFail FROM BannedLoginIP WHERE BannedIP='" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","") & "'",Database, 0, 2
 If Records.EOF = False Then
  If DateDiff("n",Records("LastLoginFail"),Now()) < 15 AND (Records("LoginFailCount") MOD 3 = 0) Then Blacklisted = True
 End If

If BlackListed Then
 Response.Write "You have been banned for 15 minutes"

ElseIf (Ucase(Request.Form("Username")) = Ucase(AdminUsername)) AND (Ucase(Request.Form("Password")) = UCase(AdminPassword)) Then

 '-- Dimension variables --'
 Dim EntryCat
 EntryCat = Request.Form("Category")

 '-- Filter SQL & clean --'
 EntryCat = Replace(EntryCat,"'","&#39;")
 EntryCat = Replace(EntryCat," ","%20")

 '-- Did we type in text? --'
 If Request.Form("Content") = "" Then

  Records.Close
  Set Records = Nothing
  Database.Close
  Set Database = Nothing

  Response.Write "No Text Entered"
  Response.End
 End If

 '-- Open the records ready to write --'
 Records.Close

 Records.CursorType = 2
 Records.LockType = 3
 Records.Open "SELECT Title, Text, Category, Day, Month, Year, Time, UTCTimeZoneOffset, EntryPUK FROM Data", Database

 Records.AddNew

  Records("Title") = Request.Form("Title")
  Records("Text") = Request.Form("Content")
  Records("Category") = EntryCat

  Records("Day") = Day(DateAdd("h",ServerTimeOffset,Now()))
  Records("Month") = Month(DateAdd("h",ServerTimeOffset,Now()))
  Records("Year") = Year(DateAdd("h",ServerTimeOffset,Now()))
  Records("Time") = TimeValue(DateAdd("h",ServerTimeOffset,Time()))

   '-- Work out time offset --'
   Dim Hours
   Hours = Abs(Int(UTCTimeZoneOffset / 60))
   If Hours < 10 Then Hours = "0" & Hours

   Dim Minutes
   Minutes = Abs(UTCTimeZoneOffset Mod 60)
   If Minutes < 10 Then Minutes = "0" & Minutes

   Dim OffsetTime
   If UTCTimeZoneOffset < 0 Then OffsetTime = "+" Else OffsetTime = "-"
   OffsetTime = OffsetTime & Hours & Minutes

  Records("UTCTimeZoneOffset") = OffsetTime

  Randomize Timer
  Records("EntryPUK") = Int(Rnd()*99999999)

 Records.Update

 '-- The entry was saved, no errors occured --'
 Response.Write "Entry Submission Successful"

Else

 If Records.EOF Then 
  Records.AddNew
  Records("BannedIP") = Request.ServerVariables("REMOTE_ADDR")
 End If

 Records("LastLoginFail") = Now()
 Records("LoginFailCount") = Records("LoginFailCount") + 1
 Records.Update

 '-- Form input did not match the admin credentials --'
 Response.Write "User/Password Error"

End If

'-- Close objects --'
Records.Close
Set Records = Nothing

Database.Close
Set Database = Nothing
%>
<%
' --------------------------------------------------------------------------
'¦Introduction : Database Functions.                                        ¦
'¦Purpose      : Handles the connection to the database and reads in the    ¦
'¦               config values to global variables. This also determines and¦
'¦               sets as a variable the server's time offset. When run with ¦
'¦               a default database the CookieName is randomised. The       ¦
'¦               database is left open for other pages to use and must be   ¦
'¦               closed when no longer used.                                ¦
'¦Requires     : All variables to be defined.                               ¦
'¦Used By      : Most pages.                                                ¦
'---------------------------------------------------------------------------

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

'--- Open Database ---'
Set Database = Server.CreateObject("ADODB.connection")
 'Database.Open  "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=" & DataPassword & "; DBQ=" & DataFile
 'Database.Open  "DSN=BlogX;"
 Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Datafile & ";"

 '--- Open Recordset ---'
 Set Records = Server.CreateObject("ADODB.recordset")
  Records.Open "SELECT AdminUsername, AdminPassword, CookieName, Copyright, EnableEmail, EmailAddress, EmailServer, EmailComponent, EnableComments, EntriesPerPage, Polls, ShowCategories, SiteName, SiteDescription, SiteSubTitle, SortByDay, BackgroundColor, ShortTimeFormat, Logging, EnableMainPage, ReaderPassword, Template FROM Config",Database, 0, 1

  If NOT Records.EOF Then
   '-- Read in the configuration details from the recordset --'
   AdminUsername   = Records("AdminUsername")
   AdminPassword   = Records("AdminPassword")
   CookieName      = Records("CookieName")

    If CookieName = "BlogXDefault" Then
     Randomize
     Database.Execute("UPDATE Config SET CookieName='BlogX" & Int(Rnd()*99999999) & "'")
    End If

   Copyright       = Records("Copyright")

   EnableEmail     = Records("EnableEmail")
   EmailAddress    = Records("EmailAddress")
   EmailServer     = Records("EmailServer")
   EmailComponent  = Records("EmailComponent")

   EnableComments  = Records("EnableComments")
   EntriesPerPage  = Records("EntriesPerPage")
   Polls           = Records("Polls")
   ShowCategories  = Records("ShowCategories")
   ShowMonth       = True
   SiteName        = Records("SiteName")
   SiteDescription = Records("SiteDescription")
   SiteSubTitle    = Records("SiteSubTitle")
   SortByDay       = Records("SortByDay")
   BackgroundColor = Records("BackgroundColor")
   TimeFormat      = Records("ShortTimeFormat")
   Logging         = Records("Logging")

   EnableMainPage  = Records("EnableMainPage")
   ReaderPassword  = Records("ReaderPassword")
   Template        = Records("Template")

  End If

 Records.Close

 Records.Open "SELECT Count(*) As SpamAttempts FROM Comments_Unvalidated",Database, 0, 1
  SpamAttempts = Records("SpamAttempts")
 Records.Close

 If EnableMainPage <> True Then PageName = "Default.asp" Else PageName = "Main.asp"
 If Right(SiteURL,1) <> "/" Then SiteURL = SiteURL & "/"

 '-- Auto-Detect how far away the time is from GMT/UTC (including any DST). Only curently used for RSS time stamps. --'
 On Error Resume Next
  Dim oShell
  Set oShell = CreateObject("WScript.Shell") 
   UTCTimeZoneOffset = oShell.RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias") 
  Set oShell = Nothing
 On Error GoTo 0

'UTCTimeZoneOffset = -60 ' This should be automatically detected however if either... :
                         ' - Time zone detection FAILED OR
                         ' - You've altered the time to a different time zone using the option in the config.

                         ' Then Uncomment the line and change this to how many MINUTES you have to add to get UTC/GMT time
                         'e.g. "-60" states you subtract 60 minutes to get GMT/UTC time. This is ONLY shown on RSS times to help the client.

                         ' BOTH are stored in the individual entries (so as to work with DST).
%>
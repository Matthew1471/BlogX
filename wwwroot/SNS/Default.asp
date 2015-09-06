<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : SNS Feed.                                                  ¦
'¦Purpose      : This provides the main SNS feed with all categories.       ¦
'¦Used By      : Includes/Header.asp, Includes/Footer.asp.                  ¦
'¦               Users' SNS readers.                                        ¦
'¦Requires     : Includes/Config.asp, Includes/RSSReplace.asp,              ¦
'¦               Includes/Cache.asp.                                        ¦
'¦Ensures      : SNS feed is generated of all categories.                   ¦
'¦Standards    : SNS 0.1.                                                   ¦
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
Response.AddHeader "WinBlogX_SNS_Version", "0.1"
Response.ContentType = "text/plain"
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<!-- #INCLUDE FILE="../Includes/RSSReplace.asp" -->
<%
'-- Grab all our data and convert it to something more friendly --'
Dim LastDate, LastTime
LastDate = Request.Querystring("Date")
LastTime = Request.Querystring("Time")

If IsDate(LastDate & " " & LastTime) Then

 '--- Open set ---'
 Records.Open "SELECT RecordID, Title, Day, Month, Year, Time, UTCTimeZoneOffset FROM Data ORDER BY RecordID DESC",Database, 1, 1
  Dim RecordID, Title, DayPosted, MonthPosted, YearPosted, TimePosted, EntryUTCTimeZoneOffset
  Dim PubDate, NewCount

  '-- Tell recordset to only check last 3 --'
  Records.PageSize = 4

  '-- Resetting Counter --'
  NewCount = 0

  '-- Loop through records until it is a next page or the end of the records --'
  Do Until (Records.EOF OR Records.AbsolutePage <> 1)

   '--- Setup Variables ---'
   Set RecordID = Records("RecordID")
   Set Title = Records("Title")

   Set DayPosted =  Records("Day")
   Set MonthPosted =  Records("Month")
   Set YearPosted =  Records("Year")
   Set TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

   If CDate(LastDate & " " & LastTime) <= CDate(DayPosted & " " & MonthPosted & " " & YearPosted & " " & TimePosted) Then
    NewCount = NewCount + 1
    PubDate = Left(WeekDayName(WeekDay(DayPosted)),3) & ", " & DayPosted & " " & Left(MonthName(MonthPosted),3) & " " & YearPosted & " " & FormatDateTime(TimePosted,4)

    If NewCount < 4 Then
     Dim ResponseString
     ResponseString = ResponseString & " <item date=""" & PubDate & " " & EntryUTCTimeZoneOffset & """ title="""
     If Title <> "" Then ResponseString = ResponseString & Replace(Title,"""","&quot;") Else ResponseString = ResponseString & PubDate
     ResponseString = ResponseString & """ link=""" & RecordID & """>" & VbCrlf
    Else
     ResponseString = ResponseString & " <item date=""" & PubDate & " " & EntryUTCTimeZoneOffset & """ title=""More NEW Entries Available at " & SiteURL & """ link=""" & SiteURL & """>" & VbCrlf
    End If

   Else

    '-- No point checking EARLIER records than our specified date --'
    Exit Do

   End If

   Records.MoveNext
  Loop

 Records.Close

 Response.Write "<entries count=""" & NewCount & """"
 If NewCount > 0 Then Response.Write " basehref=""" & SiteURL & "ViewItem.asp?Entry=" & """"
 Response.Write ">" & VbCrlf & ResponseString & "</entries>"

Else
 Response.Write "Invalid Syntax!"
End If

Set Records = Nothing
Database.Close
Set Database = Nothing
%>
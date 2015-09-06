<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : Main RSS Feed.                                             ¦
'¦Purpose      : This provides the main RSS feed with all categories.       ¦
'¦Used By      : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp,¦
'¦               Users' RSS readers.                                        ¦
'¦Requires     : Includes/Config.asp, Includes/RSSReplace.asp,              ¦
'¦               Includes/Cache.asp.                                        ¦
'¦Ensures      : RSS feed is generated of all categories.                   ¦
'¦Standards    : RSS 0.92.                                                  ¦
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
Response.ContentType = "text/xml"
%><?xml version="1.0" encoding="UTF-8"?>
<rss version="0.92">
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/RSSReplace.asp" -->
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<channel>
<title><%=Encode(SiteDescription)%></title>
<link><%=URLEncode(SiteURL)%></link>
<description><%=Encode(SiteDescription)%></description>
<language>en-us</language>
<generator>Matthew1471's BlogX / BlogX.co.uk</generator>
<copyright><%=Encode(Copyright)%></copyright>
<docs>http://blogs.law.harvard.edu/tech/rss</docs>
<% If RSSImage <> 0 Then%>
  <image>
    <url><%=URLEncode(SiteURL)%>RSS/Image.jpg</url>
    <title><%=Encode(SiteDescription)%></title>
    <link><%=URLEncode(SiteURL)%></link>
    <width>100</width>
    <height>100</height>
  </image>
<%
End If

'--- Open set ---'
Records.PageSize = EntriesPerPage
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Enclosure, LastModified FROM Data ORDER BY RecordID DESC",Database, 1, 1

 Dim RecordID, Title, Text, Category, Password, DayPosted, MonthPosted, YearPosted, TimePosted, EntryUTCTimeZoneOffset
 Dim PubDate, Enclosure, LastModified

 '-- Loop through records until it's a next page or the end of the records. --'
 Do Until (Records.EOF or Records.AbsolutePage <> 1)

  If (ReaderPassword = "") OR (Ucase(Request.Querystring()) = Ucase(ReaderPassword)) OR (IsNull(ReaderPassword) = True) Then

   '--- Setup Variables ---'
   Set RecordID = Records("RecordID")
   Set Title = Records("Title")
   Set Text = Records("Text")
   Set Category =  Records("Category")
   Set Password = Records("Password")

   Set DayPosted =  Records("Day")
   Set MonthPosted =  Records("Month")
   Set YearPosted =  Records("Year")
   Set TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

   Set Enclosure = Records("Enclosure")

   Set LastModified = Records("LastModified")

   If Len(Password) > 0 Then
    Text = "<b>-- Entry Password  ---</b><p> You need to enter a password to view this entry <p>See " & SiteURL & " for more details"
    Category = "Error"
    Enclosure = ""
   End If

  Else

   RecordID = Records("RecordID")
   Title = "Viewer Password Enabled"
   Text = "<b>-- Reader Password Is Enabled ---</b><p> Please tag ""?<i>Password</i>"" (Replacing <i>password</i> with the reader password) on to the end of the link..<p>See " & SiteURL & " for more details --"
   Category = "Error"

   DayPosted =  Records("Day")
   MonthPosted =  Records("Month")
   YearPosted =  Records("Year")
   TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

  End If

  PubDate = WeekDayName(WeekDay(DayPosted & "/" & MonthPosted & "/" & YearPosted),True,1) & ", " & DayPosted & " " & MonthName(MonthPosted,True) & " " & YearPosted & " " & FormatDateTime(TimePosted,4)
  If Len(EntryUTCTimeZoneOffset) > 0 Then PubDate = PubDate & " " & EntryUTCTimeZoneOffset Else PubDate = PubDate & " UTC"

  '-- Have we already set the LastModified header? --'
  Dim SetLastModifiedHeader
  If (NOT SetLastModifiedHeader) Then

   '-- Not every post has been modified --'
   If IsNull(LastModified) Then LastModified = CDate(DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & TimePosted)

   Dim BuildDate
   BuildDate = WeekDayName(WeekDay(LastModified),True,1) & ", " & Day(LastModified) & " " & MonthName(Month(LastModified),True) & " " & Year(LastModified) & " " & FormatDateTime(LastModified,4)
   If Len(EntryUTCTimeZoneOffset) > 0 Then BuildDate = BuildDate & " " & EntryUTCTimeZoneOffset Else BuildDate = BuildDate & " UTC"

   '-- Proxy Handler --'
   CacheHandle(LastModified)

   'Sun, 12 Aug 2007 09:58:50 GMT
   Response.Write "<lastBuildDate>" & BuildDate & "</lastBuildDate>" & VbCrlf
 
   '-- We do not want to set it twice.. only once, records are descending remember! --'
   SetLastModifiedHeader = True

  End If
%>
    <item>
      <pubDate><%=PubDate%></pubDate>
      <title><%If Title <> "" Then Response.Write Encode(Title) Else Response.Write PubDate%></title>
      <category><%=Encode(Category)%></category>
      <link><%=URLEncode(SiteURL)%>ViewItem.asp?Entry=<%=RecordID%></link>
      <comments><%=URLEncode(SiteURL)%>Comments.asp?Entry=<%=RecordID%></comments>
      <description><%=ShortEncode(Text)%></description>
      <%
      If Enclosure <> "" Then
	   If Left(Enclosure,7) <> "http://" Then Enclosure = SiteURL & "Sounds/" & Enclosure
	   Response.Write "<enclosure url=""" & URLEncode(Enclosure) & """ type=""audio/mpeg"" length=""1""/>"
	  End If
      %>
    </item>

<%
 Records.MoveNext
Loop

Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
  </channel>
</rss>
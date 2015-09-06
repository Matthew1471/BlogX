<% OPTION EXPLICIT
Response.ContentType = "text/xml"
%><?xml version="1.0" encoding="utf-8"?>
<rss version="0.92">
<!-- #INCLUDE FILE="../../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../../Includes/RSSReplace.asp" -->
<!-- #INCLUDE FILE="../../Includes/Cache.asp" -->
<%
Dim Requested
Dim CommentID, Email, Homepage, Content, LastModified
Dim CommentedDate, NewDate, Enclosure

Requested = Replace(Request.Querystring("Entry"),"'","")
Requested = Replace(Requested,"-","")
Requested = Replace(Requested,",","")
If (Requested = "") OR NOT IsNumeric(Requested) Then Requested = 0
RecordID = Encode(Requested)
%>
<channel>
 <title><%=Encode(SiteDescription)%> (Only Comments For Entry #<%=Requested%>)</title>
 <link><%=SiteURL%></link>
 <description><%=Encode(SiteDescription)%></description>
 <language>en-us</language>
 <generator>Matthew1471's BlogX / BlogX.co.uk</generator>
 <copyright><%=Encode(Copyright)%></copyright>
 <docs>http://blogs.law.harvard.edu/tech/rss</docs>
 <% If RSSImage <> 0 Then%>
 <image>
   <url><%=SiteURL%>RSS/Image.jpg</url>
   <title><%=Encode(SiteDescription)%></title>
   <link><%=URLEncode(SiteURL)%></link>
   <width>100</width>
   <height>100</height>
 </image>

 <%
 End If

'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Enclosure, LastModified FROM Data WHERE RecordID=" & RecordID,Database, 1, 3

' Loop through records until it's a next page or End of Records
If Records.EOF = False Then

If (ReaderPassword = "") OR (UCase(Request.Querystring()) = UCase(ReaderPassword)) OR (IsNull(ReaderPassword) = True) Then

'--- Setup Variables ---'
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

   Title = "Viewer Password Enabled"
   Text = "<b>-- Reader Password Is Enabled ---</b><p> Please tag ""?<i>Password</i>"" (Replacing <i>password</i> with the reader password) on to the end of the link..<p>See " & SiteURL & " for more details --"
   Category = "Error"

   DayPosted =  Records("Day")
   MonthPosted =  Records("Month")
   YearPosted =  Records("Year")
   TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")
End If

'-- <pubDate>Wed, 02 Oct 2002 13:00:00 GMT</pubDate> --
PubDate = WeekDayName(WeekDay(DayPosted & "/" & MonthPosted & "/" & YearPosted),True,1) & ", " & DayPosted & " " & MonthName(MonthPosted,True) & " " & YearPosted & " " & FormatDateTime(TimePosted,4)
If Len(EntryUTCTimeZoneOffset) > 0 Then PubDate = PubDate & " " & EntryUTCTimeZoneOffset

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
 
 '-- We don't want to set it twice.. only once, records are descending remember! --'
 SetLastModifiedHeader = True

End If
%>
 <item>
  <pubDate><%=PubDate%></pubDate>
  <title>! <%If Title <> "" Then Response.Write Encode(Title) & " (Original Post)" Else Response.Write PubDate & " (Original Post)"%></title>
  <category><%=Encode(Category)%></category>
  <link><%=URLEncode(SiteURL)%>ViewItem.asp?Entry=<%=RecordID%></link>
  <comments><%=URLEncode(SiteURL)%>Comments.asp?Entry=<%=RecordID%></comments>
  <description><%=ShortEncode(Text)%></description>
<% If Enclosure <> "" Then
      If Left(Enclosure,7) <> "http://" Then Enclosure = SiteURL & "Sounds/" & Enclosure
      Response.Write "<enclosure url=""" & URLEncode(Enclosure) & """ type=""audio/mpeg"" length=""1""/>"
     End If
%> </item>
<%
End If
'--- Close The Records ---
Records.Close

'--- Open set ---'
Records.Open "SELECT CommentID, EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK FROM Comments WHERE EntryID=" & RecordID,Database, 1, 3

 Dim RecordID, Title, Text, Category, Password, DayPosted, MonthPosted, YearPosted, TimePosted, EntryUTCTimeZoneOffset
 Dim PubDate

 '--- Setup Variables ---'
 Set CommentID = Records("CommentID")
 Set Name = Records("Name")
 Set Email = Records("Email")
 Set Homepage =  Records("Homepage")
 Set Content =  Records("Content")

 Set CommentedDate = Records("CommentedDate")
 Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")
                                 
 '--- Loop through records until it's a next page or End of Records --'
 Do Until (Records.EOF)

  '-- We use this in the comment title --'
  Count = Count + 1

  '-- <pubDate>Wed, 02 Oct 2002 13:00:00 GMT</pubDate> --
  PubDate = WeekDayName(WeekDay(CommentedDate),True,1) & ", " & Day(CommentedDate) & " " & MonthName(Month(CommentedDate),True) & " " & Year(CommentedDate) & " " & FormatDateTime(CommentedDate,4)
  If Len(EntryUTCTimeZoneOffset) > 0 Then PubDate = PubDate & " " & EntryUTCTimeZoneOffset
  %>
  <item>
   <pubDate><%=PubDate%></pubDate>
   <title><%= "(" & Count & ") By : " & Encode(Name) %></title>
   <link><%=URLEncode(SiteURL)%>ViewItem.asp?Entry=<%=Encode(Requested)%></link>
   <comments><%=URLEncode(SiteURL)%>Comments.asp?Entry=<%=Encode(Requested)%></comments>
   <description>
   <%=ShortEncode(Content)%>
   <% If Homepage <> "" Then Response.Write VbCrlf & Encode(VbCrlf & "<hr>" & VbCrlf & Name & "'s Homepage : <a href=""" & Homepage & """>" & Homepage & "</a>") %>
   </description>
  </item>
  <%
  Records.MoveNext
 Loop      

'--- Close The Records & Database ---
Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
  </channel>
</rss>
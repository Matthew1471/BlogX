<% OPTION EXPLICIT
Response.ContentType = "text/xml"
%><?xml version="1.0" encoding="ISO-8859-1"?>
<rss version="0.92">
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<%
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

'--- Lets URLEncode Just The End Bits (Server.URLEncode converts the Domain too) ---'
Function URLEncode(Variable)

If Variable <> "" Then
   URLEncode = Replace(Variable, " ","%20")
   URLEncode = Replace(URLEncode, "'","%27")
   URLEncode = Replace(URLEncode, "’","%92")
End If

End function
'--- End ---'
%>
<channel>
<title><%=Encode(SiteDescription)%></title>
<link><%=URLEncode(SiteURL)%></link>
<description><%=Encode(SiteDescription)%></description>
<docs>http://blogs.law.harvard.edu/tech/rss</docs>
<language>en-us</language>
<generator>Matthew1471's BlogX / BlogX.co.uk</generator>
<copyright><%=Encode(Copyright)%></copyright>
<managingEditor><%=EmailAddress%></managingEditor>
<webMaster><%=EmailAddress%></webMaster>
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
    Records.CursorLocation = 3 ' adUseClient
    Records.PageSize = EntriesPerPage
    Records.Open "SELECT * FROM Data ORDER BY RecordID DESC",Database, 1, 3

    Dim RecordID, Title, Text, Category, Password, DayPosted, MonthPosted, YearPosted, TimePosted
    Dim PubDate

' Loop through records until it's a next page or End of Records
Do Until (Records.EOF or Records.AbsolutePage <> 1 )

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

   If Len(Password) > 0 Then
   Text = "<b>-- Entry Password  ---</b><p> You need to enter a password to view this entry <p>See " & SiteURL & " for more details"
   Category = "Error"
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


End If

PubDate = Left(WeekDayName(WeekDay(DayPosted)),3) & ", " & DayPosted & " " & Left(MonthName(MonthPosted),3) & " " & YearPosted & " " & FormatDateTime(TimePosted,4)
%>
    <item>
      <pubDate><%=PubDate & " " & XMLTimeZone%></pubDate>
      <title><%If Title <> "" Then Response.Write Encode(Title) Else Response.Write PubDate%></title>
      <category><%=Encode(Category)%></category>
      <link><%=URLEncode(SiteURL)%>ViewItem.asp?Entry=<%=RecordID%></link>
      <comments><%=URLEncode(SiteURL)%>Comments.asp?Entry=<%=RecordID%></comments>
      <description><%=Encode(Text)%></description>
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
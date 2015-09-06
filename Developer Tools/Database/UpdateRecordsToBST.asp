<% OPTION EXPLICIT

'WARNING: This script probably doesn't check that the time ends on 01:00AM GMT! As it's comparing the finish time..
'          but seen as i'm not often blogging at 01:00AM on the days the clock changes.. I don't really care.

'To prevent a user from accidentally running this, this script is READ ONLY.. Uncomment Records.Update for it to save changes.

%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/RSSReplace.asp" -->
<%
'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT RecordID, Title, Day, Month, Year, Time, UTCTimeZoneOffset FROM Data ORDER BY RecordID DESC",Database, 1, 3

    Dim RecordID, Title, Text, Category, Password, DayPosted, MonthPosted, YearPosted, TimePosted, EntryUTCTimeZoneOffset
    Dim PubDate, Enclosure

' Loop through records until it's a next page or End of Records
Do Until (Records.EOF)

'--- Setup Variables ---'
   Set RecordID = Records("RecordID")
   Set Title = Records("Title")

   Set DayPosted =  Records("Day")
   Set MonthPosted =  Records("Month")
   Set YearPosted =  Records("Year")
   Set TimePosted =  Records("Time")

   Set EntryUTCTimeZoneOffset = Records("UTCTimeZoneOffset")

PubDate = WeekDayName(WeekDay(DayPosted & "/" & MonthPosted & "/" & YearPosted),True) & ", " & DayPosted & " " & MonthName(MonthPosted,True) & " " & YearPosted & " " & FormatDateTime(TimePosted,4)
If Len(EntryUTCTimeZoneOffset) > 0 Then PubDate = PubDate & " " & EntryUTCTimeZoneOffset

'-- Have we already calculated this?! --'
Dim CurrentSearchYear
If CurrentSearchYear <> YearPosted Then

 '-- Calculate Last Sunday In March Date --'
 Dim CurrentMarchSearchDate, CurrentOctoberSearchDate
 CurrentMarchSearchDate = 31
 CurrentOctoberSearchDate = 31

 '-- Note : Sunday = 1 --'
 Do While WeekDay(CurrentMarchSearchDate & "/03/" & YearPosted,1) <> 1
  CurrentMarchSearchDate = CurrentMarchSearchDate - 1
 Loop

 '-- Note : Sunday = 1 --'
 Do While WeekDay(CurrentOctoberSearchDate & "/10/" & YearPosted,1) <> 1
  CurrentOctoberSearchDate = CurrentOctoberSearchDate - 1
 Loop

 '-- We've Calculated The Following Year --'
 CurrentSearchYear = YearPosted

 Response.Write "<b>" & CurrentMarchSearchDate & " was the last Sunday in March " & YearPosted & "!<br>"
 Response.Write CurrentOctoberSearchDate & " was the last Sunday in October " & YearPosted & "!</b><br><br>"
End If
%>
<%=RecordID%>:<%=Title%><br><small>@<%=PubDate%></small><br>

<%
Response.Write "<font color="

If DateDiff("m",CurrentMarchSearchDate & "/03/" & YearPosted & " 1:00am",DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & FormatDateTime(TimePosted,4)) > 0 AND DateDiff("m",CurrentOctoberSearchDate & "/10/" & YearPosted & " 1:00am",DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & FormatDateTime(TimePosted,4)) < 0 Then
 Response.Write """red"">+0100"
 Records("UTCTimeZoneOffset") = "+0100"
Else
 Response.Write """blue"">+0000"
 Records("UTCTimeZoneOffset") = "+0000"
End If

Response.Write "</font><br><br>" & VbCrlf

'--- SCRIPT:READONLY!!!! ---'
'Records.Update

Records.MoveNext
Loop

'--- Close The Records & Database ---
Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
<%
Dim nLastDay, n, nn, nnn, nDS, PostToday

If szPos <> "" Then
If szPos = "NEXT" Then
nDS = 1
Else
nDS = -1
End If
nDS = DateSerial(nYear, nMonth + nDS, 1)
nYear = Year(nDS)
nMonth = Month(nDS)
End If
nLastDay = Day(DateSerial(nYear, nMonth + 1, 1 - 1))
nDay = 1 - Weekday(DateSerial(nYear, nMonth, 1)) + 1
%>
<table class="navCalendar" cellspacing="0" cellpadding="4" border="0" style="border-width:1px;border-style:solid;border-collapse:collapse;">
<tr>
<td colspan="7" style="background-color:<%=CalendarBackground%>;">
<table class="navTitleStyle" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">
<tr>
<td class="navNextPrevStyle" style="width:15%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>&POS=LAST" style="color:Black">&lt;</a></td>
<td align="Center" style="width:70%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>"><%=MonthName(nMonth)%></a> (<%=Right(nYear,2)%>)</td>
<td class="navNextPrevStyle" align="Right" style="width:15%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>&POS=NEXT" style="color:Black">&gt;</a></td>
</tr>
</table>

</td>
</tr>

<%
'### Write Out The Weekdays ###'

Response.Write "<tr>"

For n = 0 To 6
Response.Write "<td class=""navDayHeader"" align=""Center"">" & Left(WeekdayName(n + 1, True),1) & "</TD>" & CHR(13)
Next

Response.Write "</tr>"

'### Write Out Days ###'

For nn = 0 To 5
Response.Write"<TR>" & CHR(13)
For nnn = 0 To 6
If nDay > 0 And nDay <= nLastDay Then

Response.Write "<td class="""

'### Highlight CurrentDay/Weekend ###'
If nDay = Int(Request("Day")) Then
Response.Write "navSelectedDayStyle" 
ElseIf nnn = 0 or nnn = 6 Then Response.Write "navWeekendDayStyle"
Else Response.Write "navDayStyle"
End If
'### End Of Current Day Check ###'

Response.Write """ align=""Center"""

'### Highlight CurrentDay/Weekend ###'
If nDay = Int(Request("Day")) Then
Response.Write " style=""color:White;background-color:" & CalendarBackground & ";width:14%;"">"
Else 
Response.Write " style=""width:14%;"">"
End If
'### End Of Current Day Check ###' 

'### Lets Strip Out That Existing Day From Our Clicky ###'
If CalendarCheck <> 1 Then
If SortByDay = True Then Response.Write "<a href=""" & SiteURL & PageName & "?"
If SortByDay = True Then Response.Write "YearMonth=" & nYear & Right("00" & nMonth, 2) & "&Day=" & nDay & """>"
ElseIf SortByDay <> False Then

    '-- Check if there was something posted on each day --'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT * FROM Data WHERE Day=" & nDay & " AND Month=" & Right("00" & nMonth, 2) & " AND Year=" & nYear & ";",Database, 1, 3
    If Records.EOF = False Then
    PostToday = True
    Response.Write "<a href=""" & SiteURL & PageName & "?"
    Response.Write "YearMonth=" & nYear & Right("00" & nMonth, 2) & "&Day=" & nDay & """>"
    Else
    PostToday = False
    End If
    Records.Close

End If

If (Day(DateAdd("h",TimeOffset,Now())) = nDay) AND (Month(DateAdd("h",TimeOffset,Now())) = nMonth) Then Response.Write "<font color=""red"">"
Response.Write nDay
If (Day(DateAdd("h",TimeOffset,Now())) = nDay) AND (Month(DateAdd("h",TimeOffset,Now())) = nMonth) Then Response.Write "</font>"

If (SortByDay = True) AND ((CalendarCheck <> 1) OR (PostToday = True)) Then Response.Write "</a>"
'### Finished Day Stripping ###'

Response.Write "</TD>" & CHR(13)
Else
Response.Write "<Td class=""navOtherMonthDayStyle"" align=""Center"" style=""width:14%;"">-</TD>" & CHR(13)
End If
nDay = nDay + 1
Next
Response.Write "</TR>" & CHR(13)
Next
%>
<tr><td colspan="7" class="navCalendar" cellspacing="0" cellpadding="4" border="0" style="background-color:<%=CalendarBackground%>;" align="center"><A HREF="<%=SiteURL & PageName%>">This Month!</A></td></tr>
</TABLE>

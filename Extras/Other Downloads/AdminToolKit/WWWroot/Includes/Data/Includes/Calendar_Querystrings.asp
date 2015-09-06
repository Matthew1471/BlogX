<%
Dim szYearMonth, szPos
Dim nYear, nMonth, nDay, SpecificRequest

szYearMonth = Request("YearMonth")
szPos = Request("POS")
nDay = Request("Day")

If szYearMonth = "" Then
SpecificRequest = False
nYear = Year(Now())
nMonth = Month(Now())
Else
SpecificRequest = True
nYear = Left(szYearMonth,4)
nMonth = Right(szYearMonth,2)
End If

'### SQL Attacker Exploit Management ###'
If (IsNumeric(nYear) <> True) OR (IsNumeric(nMonth) <> True) OR (IsNumeric(nDay) <> True) Then Response.Redirect("Hacker.asp")
%>
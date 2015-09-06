<%
' --------------------------------------------------------------------------
'¦Introduction : Cache Functions.                                           ¦
'¦Purpose      : Handles the If-Modified-Since requests and adds a          ¦
'¦               last-modified header.                                      ¦
'¦Requires     : Database, Records variables to be defined.                 ¦
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

Function IsNothing(Value)

 If IsEmpty(Value) Then
  IsNothing = True
  Exit Function
 End If

 If IsObject(Value) Then
  If Value Is Nothing Then
   IsNothing = True
   Exit Function
  End If
 End If

IsNothing = False

End Function

Function PadDigits(n, totalDigits) 
 PadDigits = Right(String(totalDigits,"0") & n, totalDigits) 
End Function 

Function DateFromHTTP(HTTPDate)
 Dim strWeekday, strDay, strMonth
 Dim strYear, strHour, strMinute, strSecond
 Dim ConvertedDate, strHDate

 strHDate = Trim(LCase(HTTPDate)) '-- Case Independent --'

 If Right("---" & strHDate,3) = "gmt" Then
  '-- Split Parts, Assume Correct (Otherwise Error) --'
  strWeekday = Mid(strHDate, 1, 3)
  strDay = Mid(strHDate, 6, 2)
  strMonth = Mid(strHDate, 9, 3)
  strYear = Mid(strHDate, 13, 4)
  strHour = Mid(strHDate, 18, 2)
  strMinute = Mid(strHDate, 21, 2)
  strSecond = Mid(strHDate, 24, 2)

  '-- Try To Build Date --'
  On Error Resume Next
   Err.Clear
   ConvertedDate = DateSerial(strYear, MonthFromString(strMonth), strDay) + TimeSerial(strHour, strMinute, strSecond)
   If Err <> 0 Then ConvertedDate = DateSerial(1970,1,1) '-- Invalid Date --'

   If Err <> 0 Then Response.Write "<!-- HTTPDate: " & HTTPDate & "-->" & VbCrlf

  On Error GoTo 0

 Else
  ConvertedDate = DateSerial(1971,1,1) '-- Invalid, Time doesn't have 'gmt' --'
 End If

 DateFromHTTP = ConvertedDate
End Function

Function MonthFromString(strMonth)

        Dim intMonthNr 
        Select Case strMonth '-- Assume Lower Case --'
            Case "jan" : intMonthNr = 1
            Case "feb" : intMonthNr = 2
            Case "mar" : intMonthNr = 3
            Case "apr" : intMonthNr = 4
            Case "may" : intMonthNr = 5
            Case "jun" : intMonthNr = 6
            Case "jul" : intMonthNr = 7
            Case "aug" : intMonthNr = 8
            Case "sep" : intMonthNr = 9
            Case "oct" : intMonthNr = 10
            Case "nov" : intMonthNr = 11
            Case "dec" : intMonthNr = 12
        End Select

        MonthFromString = intMonthNr

End Function

Function CacheHandle(Posted)

 If IsNull(Posted) Then Posted = Now()

 '-- For Firefox 3 (this basically will make it perform conditional requests every time) --'
 If Response.Buffer Then Response.CacheControl = "no-cache"

 Dim PubDate
 PubDate = Left(WeekDayName(WeekDay(Posted),True),3) & ", " 
 PubDate = PubDate & PadDigits(Day(Posted),2) & " " & MonthName(Month(Posted),True) & " " & Year(Posted) & " "
 PubDate = PubDate & FormatDateTime(Posted,4) & ":" & PadDigits(Second(Posted),2)

'Dates : http://www.w3.org/Protocols/rfc2616/rfc2616-sec3.html
'Response.Write "<!-- Page Last Modified.. " & PubDate & " GMT " & "-->"
 Response.AddHeader "Last-Modified", PubDate & " GMT"

 Dim GivenDate
 GivenDate = Request.ServerVariables("HTTP_IF_MODIFIED_SINCE")
  If GivenDate <> "" Then
   GivenDate = DateFromHTTP(GivenDate)

   '-- Round to 5 minutes for proxy clock drift --'
   If (Posted <= GivenDate + 0.0035) Then

    '-- Only because some of the pages don't get chance to kill records before calling us --'
    If NOT IsNothing(Records) Then
     Const adStateOpen = 1
     If Records.State = adStateOpen Then Records.Close
     Set Records = Nothing
    End If

    If NOT IsNothing(Database) Then
     If Database.State = adStateOpen Then Database.Close
     Set Database = Nothing
    End If

    Response.Clear()
    Response.Status = "304 Not Modified"
    Response.End()
   Else
    Response.Write "<!-- No Proxy Match : DocDate-""" & Posted & """ / ProxyDate-""" & GivenDate & """ -->" & VbCrlf
   End If
  End If
End Function
%>
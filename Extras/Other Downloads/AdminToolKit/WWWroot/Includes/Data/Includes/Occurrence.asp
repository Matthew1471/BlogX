<%
'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string into VisualBasic

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2

'VB Function - FreeVBCode.com
'Converted to ASP By - Matthew1471
'IIF Function For ASP - http://www.developerfusion.com/show/1606
''**********************************************

'------------------------------------------------------------
Public Function IIf(blnExpression, vTrueResult, vFalseResult)
  If blnExpression Then
    IIf = vTrueResult
  Else
    IIf = vFalseResult
  End If
End Function

'------------------------------------------------------------
Function CharCount(OrigString, Chars, CaseSensitive)
Dim lLen, lCharLen, lAns, sInput, sChar, lCtr
Dim lEndOfLoop, bytCompareType

sInput = OrigString

If sInput <> "" Then

lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then lAns = lAns + 1
    Next

CharCount = lAns

End If

End Function

'Response.Write CharCount("hpelp","p",False)
%>
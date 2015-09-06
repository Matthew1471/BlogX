<%
' --------------------------------------------------------------------------
'¦Introduction : Spell Check Functions.                                     ¦
'¦Purpose      : Provides functionality to spell check form input.          ¦
'¦Requires     : ..\Includes\Dictionary\dict-large.txt, Error_Spell.asp.    ¦
'¦Used By      : Most editor pages (when viewed in IE).                     ¦
'---------------------------------------------------------------------------

' spell.asp
' 5/2/2002
'
' By Sam Kirchmeier
' Follow along at http://www.kirchmeier.org/code/pmsc/
'
' This code is released free of charge.  You may copy it and use it
' as you wish, but please include proper credit to the author when
' you incorporate it into your own projects.  Thanks!

Const cstRelativeDictPath = "..\Includes\Dictionary\dict-large.txt"
Dim strDictArray


sub LoadDictArray
    Dim objFSO
    Dim objDictFile
    Dim intDictSize
    Dim intForReading
    Dim objDictStream

On Error Resume Next
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

    If Err <> 0 Then
    Set objFSO = Nothing
    Database.Close
    Set Records  = Nothing
    Set Database = Nothing
    Response.Redirect "Error_Spell.asp?Error=2"
    End If

On Error GoTo 0

If objFSO.FileExists(Server.MapPath(cstRelativeDictPath)) Then

    Set objDictFile = objFSO.GetFile(Server.MapPath(cstRelativeDictPath))
    intDictSize = objDictFile.Size

    intForReading = 1
    Set objDictStream = objDictFile.OpenAsTextStream(intForReading)
    strDictArray = Split(objDictStream.Read(intDictSize), vbNewLine)
    objDictStream.Close

    Set objDictStream = Nothing

    Set objDictFile = Nothing

Else

Set objFSO = Nothing
Database.Close
Set Records  = Nothing
Set Database = Nothing
Response.Redirect "Error_Spell.asp?Error=1"

End If

    Set objFSO = Nothing

end sub


function PrepForSpellCheck(strWord)
    Dim strValidChars
    Dim i
    Dim strLetter

    strValidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'-"

    for i = 1 to Len(strWord)
        strLetter = Mid(strWord, i, 1)
        if InStr(strValidChars, strLetter) > 0 then
            PrepForSpellCheck = PrepForSpellCheck & strLetter
        elseif i < Len(strWord) then
            PrepForSpellCheck = ""
            exit for
        end if
    next
end function


function SpellCheck(strWord)
    Dim intFirst
    Dim intLast
    Dim intMiddle

    if Len(strWord) > 0 then
        SpellCheck = False
        intFirst = LBound(strDictArray)
        intLast = UBound(strDictArray)

        do while intFirst <= intLast
            intMiddle = (intFirst + intLast) \ 2

            if LCase(strDictArray(intMiddle)) = LCase(strWord) then
                SpellCheck = True
                exit do
            elseif LCase(strDictArray(intMiddle)) < LCase(strWord) then
                intFirst = intMiddle + 1
            else
                intLast = intMiddle - 1
            end if
        loop
    else
        SpellCheck = True
    end if
end function


function Soundex(strString)
    Dim i
    Dim strLetter
    Dim strCode

    Soundex = UCase(Left(strString, 1))

    for i = 2 to Len(strString)
        strLetter = UCase(Mid(strString, i, 1))
        select case strLetter
            case "B", "P"
                strCode = "1"
            case "F", "V"
                strCode = "2"
            case "C", "K", "S"
                strCode = "3"
            case "G", "J"
                strCode = "4"
            case "Q", "X", "Z"
                strCode = "5"
            case "D", "T"
                strCode = "6"
            case "L"
                strCode = "7"
            case "M", "N"
                strCode = "8"
            case "R"
                strCode = "9"
            case else
                strCode = ""
        end select
        if Right(Soundex, 1) <> strCode then
            Soundex = Soundex & strCode
        end if
    next
end function


function WordSimilarity(strWord, strSimilarWord)
    Dim intWordLen
    Dim intSimilarWordLen
    Dim intMaxBonus
    Dim intPerfectValue
    Dim intSimilarity
    Dim i

    intWordLen = Len(strWord)
    intSimilarWordLen = Len(strSimilarWord)

    intMaxBonus = 3
    intPerfectValue = intWordLen + intWordLen + intMaxBonus
    intSimilarity = intMaxBonus - Abs(intWordLen - intSimilarWordLen)

    for i = 1 to intWordLen
        if i <= intSimilarWordLen then
            if LCase(Mid(strWord, i, 1)) = LCase(Mid(strSimilarWord, i, 1)) then
                intSimilarity = intSimilarity + 1
            end if

            if LCase(Mid(strWord, intWordLen - i + 1, 1)) = LCase(Mid(strSimilarWord, intSimilarWordLen - i + 1, 1)) then
                intSimilarity = intSimilarity + 1
            end if
        end if
    next

    WordSimilarity = intSimilarity / intPerfectValue
end function


function Suggest(strWord)
    Dim strSoundex
    Dim i
    Dim strSuggestions
    Dim intMaxSuggestions
    Dim intSuggestionCount
    Dim strSuggestion
    Dim strSuggestionArray
    Dim dblSimilarityArray
    Dim dblSimilarity

    intMaxSuggestions = 10
    strSoundex = Soundex(strWord)

    i = 0
    do while i <= UBound(strDictArray)
        if LCase(Left(strDictArray(i), 1)) <> LCase(Left(strWord, 1)) then
            i = i + 1
        else
            exit do
        end if
    loop

    do while i <= UBound(strDictArray)
        if LCase(Left(strDictArray(i), 1)) = LCase(Left(strWord, 1)) then
            if Soundex(strDictArray(i)) = strSoundex then
                if strSuggestions & "" = "" then
                    strSuggestions = strDictArray(i)
                else
                    strSuggestions = strSuggestions & "|" & strDictArray(i)
                end if
            end if
            i = i + 1
        else
            exit do
        end if
    loop

    Suggest = Split(strSuggestions, "|")

    if UBound(Suggest) < intMaxSuggestions then
        intSuggestionCount = UBound(Suggest)
    else
        intSuggestionCount = intMaxSuggestions - 1
    end if
    ReDim strSuggestionArray(intSuggestionCount)
    ReDim dblSimilarityArray(intSuggestionCount)

    for each strSuggestion in Suggest
        dblSimilarity = WordSimilarity(strWord, strSuggestion)
        i = intSuggestionCount
        do while dblSimilarity > dblSimilarityArray(i)
            if i < intSuggestionCount then
                strSuggestionArray(i + 1) = strSuggestionArray(i)
                dblSimilarityArray(i + 1) = dblSimilarityArray(i)
            end if
            strSuggestionArray(i) = strSuggestion
            dblSimilarityArray(i) = dblSimilarity
            i = i - 1
            if i = -1 then
                exit do
            end if
        loop
    next

    Suggest = strSuggestionArray
end function
%>
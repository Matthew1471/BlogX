Attribute VB_Name = "mdlBase64"
Option Explicit

' Reference: "Base 64 Simplified" - http://www.jti.net/brad/base64.htm
' I repeatedly and thoroughly read the information in this site in the making of this code!
' It was EXTREMELY helpful and I wouldn't have succeeded without it!

' Here are the Base 64 characters. In this string, A maps to 1, and / maps to 64.
' In the coding, it's zero-based - A maps to 0 and / maps to 63.
' So when we use InStr or Mid on Base64Chars, remember to decrease the return value by 1
' (for InStr) or to increase the 2nd parameter, Start (for Mid).
Const Base64Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
' Note: This Const is the ONLY thing taken from the C source code on the site mentioned.

' P.S. For the binary numbers, we'll use Strings -
' we aren't applying math to the binary numbers, only
' to the base 10 numbers, and Strings are easier to
' manipulate in this case.

' P.S.S. I know I could have used arrays in many places,
' but they would have been arrays of 3 or 4 members,
' and I'd rather type a line 4 times than use arrays,
' which would add to the amount of code anyway.
' So don't tell me. Thanks.

' P.P.S.S. (Is that right?)
' Change this code all you want, I don't care!
' Write me any question/complaint/praise/comment/bug-notice/glitch-notice/etc.,
' I will appreciate absolutely anything (even hate letters).

Function Base10ToBinary(ByVal Base10 As Long) As String
    ' This function is for internal use only, so there is no
    ' error-trapping in case the variable Base10 is negative.
    ' (Actually, I don't think it would matter, but oh well.)
    '
    ' I really don't feel like explaining this function...
    ' But it works!
    ' This is NOT what everybody usually uses to convert base 10 to 2.
    ' I used logarithms here, which work faster.
    ' (Not that speed matters nowadays... I still like optimizing, though!)
    Dim PrevResult As Integer, CurResult As Integer
    If Base10 = 0 Then
        ' In the case of 0, the method below fails.
        ' Since that is the only case, it is easy to manually deal with it.
        Base10ToBinary = "0"
        Exit Function
    End If
    Do
        CurResult = Int(Log(Base10) / Log(2)) ' Oh no! Logarithms! RUUUUUUN!
        If PrevResult = 0 Then PrevResult = CurResult + 1
        Base10ToBinary = Base10ToBinary & String(PrevResult - CurResult - 1, "0") & "1"
        Base10 = Base10 - 2 ^ CurResult
        PrevResult = CurResult
    Loop Until Base10 = 0
    Base10ToBinary = Base10ToBinary & String(CurResult, "0")
End Function

Function BinaryToBase10(ByVal Binary As String) As Long
    ' This function is for internal use only, so there is no
    ' error-trapping in case the string Binary is not a number,
    ' or a number with digits other than 0 and 1.
    Dim I As Integer
    For I = Len(Binary) To 1 Step -1 ' Stepping through the digits starting at the right
        BinaryToBase10 = BinaryToBase10 + Val(Mid(Binary, I, 1)) * 2 ^ (Len(Binary) - I)
    Next
    ' Converting TO Base 10 is as simple as that!
    ' The above really isn't an advanced mathematical formula,
    ' once you understand what each variable stands for...
End Function

' This sub's name is, "binary three times eight to four times six".
' It translates three 8-digit binary numbers to four 6-digit binary numbers.
' (The opposite of what the below sub, Bin4x6To3x8, does.)
' This sub is necessary for ENCODING.
Sub Bin3x8To4x6(ByVal Bin1Len8 As String, ByVal Bin2Len8 As String, ByVal Bin3Len8 As String, ByRef Bin1Len6 As String, ByRef Bin2Len6 As String, ByRef Bin3Len6 As String, ByRef Bin4Len6 As String)
    ' First, we make sure all three numbers are EXACTLY 8 digits long,
    ' by adding zeros to the beginning - as many as necessary (at most 7).
    Bin1Len8 = Right("0000000" & Bin1Len8, 8)
    Bin2Len8 = Right("0000000" & Bin2Len8, 8)
    Bin3Len8 = Right("0000000" & Bin3Len8, 8)
    ' From the first 8-digit number,
    ' we need these 6 digits: ######XX-XXXXXXXX-XXXXXXXX
    Bin1Len6 = Left(Bin1Len8, 6)
    ' From the first and second 8-digit numbers,
    ' we need these 6 digits: XXXXXX##-####XXXX-XXXXXXXX
    Bin2Len6 = Right(Bin1Len8, 2) & Left(Bin2Len8, 4)
    ' From the second and third 8-digit numbers,
    ' we need these 6 digits: XXXXXXXX-XXXX####-##XXXXXX
    Bin3Len6 = Right(Bin2Len8, 4) & Left(Bin3Len8, 2)
    ' From the third 8-digit number,
    ' we need these 6 digits: XXXXXXXX-XXXXXXXX-XX######
    Bin4Len6 = Right(Bin3Len8, 6)
    ' Now, the four 6-digit numbers may have some zeros in the beginning,
    ' but that doesn't matter.
End Sub

' This sub's name is, "binary four times six to three times eight".
' It translate four 6-digit binary numbers to three 8-digit binary numbers.
' (The opposite of what the above sub, Bin3x8To4x6, does.)
' This sub is necessary for DECODING.
Sub Bin4x6To3x8(ByVal Bin1Len6 As String, ByVal Bin2Len6 As String, ByVal Bin3Len6 As String, ByVal Bin4Len6 As String, ByRef Bin1Len8 As String, ByRef Bin2Len8 As String, ByRef Bin3Len8 As String)
    ' First, we make sure all three numbers are EXACTLY 6 digits long,
    ' by adding zeros to the beginning - as many as necessary (at most 5).
    Bin1Len6 = Right("00000" & Bin1Len6, 6)
    Bin2Len6 = Right("00000" & Bin2Len6, 6)
    Bin3Len6 = Right("00000" & Bin3Len6, 6)
    Bin4Len6 = Right("00000" & Bin4Len6, 6)
    ' From the first and second 6-digit numbers,
    ' we need these 8 digits: ######-##XXXX-XXXXXX-XXXXXX
    Bin1Len8 = Bin1Len6 & Left(Bin2Len6, 2)
    ' From the second and third 6-digit numbers,
    ' we need these 8 digits: XXXXXX-XX####-####XX-XXXXXX
    Bin2Len8 = Right(Bin2Len6, 4) & Left(Bin3Len6, 4)
    ' From the third and fourth 6-digit numbers,
    ' we need these 8 digits: XXXXXX-XXXXXX-XXXX##-######
    Bin3Len8 = Right(Bin3Len6, 2) & Bin4Len6
    ' Now, the three 8-digit numbers may have some zeros in the beginning,
    ' but that doesn't matter.
End Sub

' This sub takes TheString, and removes all instances of WhatToRemove from it.
Function RemoveFromString(ByVal TheString As String, ByVal WhatToRemove As String) As String
    Dim lPos As Long
    If Len(WhatToRemove) = 0 Then Exit Function ' Make sure we are removing something!
    lPos = InStr(TheString, WhatToRemove)
    While lPos > 0
        ' This could be changed a little to become the VB6 Replace function,
        ' which I don't have access to, because I have VB5!
        TheString = Left(TheString, lPos - 1) & Mid(TheString, lPos + Len(WhatToRemove))
        lPos = InStr(TheString, WhatToRemove)
    Wend
    RemoveFromString = TheString
End Function

' This function takes a normal string and encodes it using the Base 64 method.
' It returns an empty string if NormalString is empty.
' Update! You can use the new Break parameter to break the Base 64 string in every line!
' My E-mail server breaks it so that there are 72 chars in every line of the encoded string.
Function Base64Encode(ByVal NormalString As String, Optional ByVal Break As Integer = 0) As String
    Dim I As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    ' Quick error trapping:
    If NormalString = vbNullString Then Exit Function
    ' Go through the string, looking at 3 chars at a time, except the last few.
    ' [Not necessarily the last three, because let's say "I" becomes Len(NormalString) - 4
    ' and we inspect I, I + 1 and I + 2... All we have left are the last TWO chars.
    ' If you don't understand this, never mind! It works.]
    For I = 1 To Len(NormalString) - 3 Step 3
        ' The three chars' ASCII values get translated to binary. The first goes
        ' to Bin1Len8, the second goes to Bin2Len8 and the third goes to Bin3Len8.
        Bin1Len8 = Base10ToBinary(Asc(Mid(NormalString, I, 1)))
        Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, I + 1, 1)))
        Bin3Len8 = Base10ToBinary(Asc(Mid(NormalString, I + 2, 1)))
        ' Now, convert the 3 binary values to 4 binary values.
        Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
        ' We don't care about the Len8's anymore, now we have the Len6's and they feel special.
        ' Now, we have four binary numbers, each ranging between 0 and 63.
        ' Let's put the necessary char from Base64Chars!
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1)
    Next
    ' Now we check how many characters we have left, by removing a multiple of three
    ' from the beginning of NormalString, but do not even try to understand this formula,
    ' you'll get a headache. (It works.) It leaves the last 1, 2 or 3 characters of
    ' NormalString - the ones we haven't touched yet.
    NormalString = Right(NormalString, Len(NormalString) - IIf(Len(NormalString) / 3 = Int(Len(NormalString) / 3), Len(NormalString) - 3, Int(Len(NormalString) / 3) * 3))
    ' We definitely have at least 1 character!
    Bin1Len8 = Base10ToBinary(Asc(Left(NormalString, 1)))
    ' Not sure about the 2nd or 3rd! (If we don't have them, we'll fill in 0's in their place.)
    If Len(NormalString) >= 2 Then Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, 2, 1))) Else Bin2Len8 = "0"
    If Len(NormalString) = 3 Then Bin3Len8 = Base10ToBinary(Asc(Right(NormalString, 1))) Else Bin3Len8 = "0"
    ' Now, we can convert the 3 binary values to the 4 binary values.
    Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
    ' Again, we must check if we need to break the string when adding chars.
    ' We must have Bin1Len8 to have Bin1Len6 and Bin2Len6 - but we have it for sure, so no check is necessary.
    Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
    Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
    ' We must have Bin2Len8 to have Bin3Len6, otherwise we put in a "=" mark.
    Base64Encode = Base64Encode & IIf(Len(NormalString) >= 2, Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1), "=")
    ' We must have Bin3Len8 to have Bin4Len6, otherwise we put in a "=" mark.
    Base64Encode = Base64Encode & IIf(Len(NormalString) = 3, Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1), "=")
    ' That is it! Finally! WE HAVE AN ENCODED BASE 64 STRING!!!
    ' Oh, wait, now we have to break it, if Break > 0.
    If Break > 0 Then
        I = Break + 1 ' The string begins at 1, not 0, so we have to add 1.
        While I < Len(Base64Encode)
            Base64Encode = Left(Base64Encode, I - 1) & vbCrLf & Mid(Base64Encode, I)
            I = I + Break + 2 ' To the next break, and 2 because of the new vbCrLf.
        Wend
    End If
    ' WE'RE DONE encoding, for REAL this time!
End Function

' This function takes a hideously ugly string encoded using the Base 64 method and turns it into a normal string.
' It returns an empty string if Base64String is empty, or if it is not a string encoded using the Base 64 method.
' Update! Some programs give you Base 64 strings with spaces, carriage returns and linefeed characters in them.
' Those characters are now ignored, instead of returning an empty string.
Function Base64Decode(ByVal Base64String As String) As String
    Dim I As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    ' This function is NOT used internally, so we MUST do full error trapping and check that ALL
    ' the characters in Base64String can be found in Base64Chars. Also allowed is "=".
    ' Of course, we must get rid of spaces, carriage returns and linefeed characters.
    Base64String = RemoveFromString(Base64String, " ")
    Base64String = RemoveFromString(Base64String, vbCr)
    Base64String = RemoveFromString(Base64String, vbLf)
    ' Remember that now Base64String may be empty!
    If Base64String = vbNullString Then Exit Function
    ' Almost done with error trapping... Now we can check all the other characters!
    For I = 0 To 255 ' Loop through all the existant characters...
        If InStr(Base64String, Chr(I)) > 0 And Not _
            ((InStr(Base64Chars, Chr(I)) > 0) Or (I = Asc("="))) Then Exit Function
    Next
    ' Also important: The lengths of all Base 64 strings are perfectly divisible by 4.
    ' P.S. Have you ever heard of the \ operator? This is how it works: A \ B <==> Int(A / B)
    ' Very useful! But it's not popular and many people have never heard of it! I had to write
    ' this comment, sorry...
    If Not Len(Base64String) / 4 = Len(Base64String) \ 4 Then Exit Function
    ' Go through the string, looking at 4 chars at a time. Since the length of Base64String is
    ' perfectly divisible by four (or we wouldn't be here), we can safely do a Step 4 without
    ' worrying about missing any character.
    For I = 1 To Len(Base64String) Step 4
        ' The four chars are looked up in Base64Chars. Their ID numbers get translated
        ' to binary. The first ID# goes to Bin1Len6, the second to Bin2Len6, the third to
        ' Bin3Len6 and the fourth to Bin4Len6.
        Bin1Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I, 1)) - 1)
        Bin2Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 1, 1)) - 1)
        ' Last two may be "=" marks... If they are, let's just leave a zero.
        If Mid(Base64String, I + 2, 1) = "=" Then Bin3Len6 = "0" Else Bin3Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 2, 1)) - 1)
        If Mid(Base64String, I + 3, 1) = "=" Then Bin4Len6 = "0" Else Bin4Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 3, 1)) - 1)
        ' Now, convert the 4 binary values to 3 binary values.
        Call Bin4x6To3x8(Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6, Bin1Len8, Bin2Len8, Bin3Len8)
        ' We don't care about the Len8's anymore, now we have the Len6's and they feel special.
        ' Now, we have three binary numbers, each ranging between 0 and 255.
        ' They are ASCII numbers! Let's convert them to chars!
        ' The first one is definitely there.
        Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin1Len8))
        ' If we have a "=" at I + 2, then the second one is disabled.
        If Not Mid(Base64String, I + 2, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin2Len8))
        ' If we have a "=" at I + 3, then the third one is disabled.
        If Not Mid(Base64String, I + 3, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin3Len8))
    Next
    ' That is it! Finally! WE HAVE AN DECODED NORMAL STRING!!!
End Function

Sub Main()
    ' Using the Base 64 Methods, the string "QWxhZGRpbjpvcGVuIHNlc2FtZQ=="
    ' is Base-64-monkey-talk for "Aladdin:open sesame".
    'Dim S As String
    ' Testing ENCODING:
    'S = "The next two lines should be identical:" & vbNewLine & vbNewLine
    'S = S & "QWxhZGRpbjpvcGVuIHNlc2FtZQ==" & vbNewLine
    'S = S & Base64Encode("Aladdin:open sesame") & vbNewLine & vbNewLine
    ' Testing DECODING:
    'S = S & "The next two lines should be identical:" & vbNewLine & vbNewLine
    'S = S & "Aladdin:open sesame" & vbNewLine
    'S = S & Base64Decode("QWxhZGRpbjpvcGVuIHNlc2FtZQ==")
    'Call MsgBox(S, vbExclamation, "Base 64 Method for Encoding and Decoding")
    'Debug.Print Base64Decode(Base64Encode("blah blah blah blah" & String(500, "Z") & "woo hoo" & vbCrLf & "this is working!!", 72))
End Sub

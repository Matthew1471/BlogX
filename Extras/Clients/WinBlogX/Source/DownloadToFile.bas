Attribute VB_Name = "DownloadToFile"
'****************************************************************
'
' Live Program Update Code
'
' Written by:  Blake B. Pell
'              blakepell@hotmail.com
'              bpell@indiana.edu
'              http://www.blakepell.com
'              December 7, 2000
'
' This code is open source, I would appreciate that anybody using
' this is a released application to e-mail or get in contact with
' me.  I hope this makes someone's day easier or helps them learn
' a bit.
'
'
'****************************************************************

Global myVer As String
Global status$
Global UpdateTime As Integer

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    ' Written by: Blake Pell
    
    On Local Error GoTo 100
    
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    
    On Local Error Resume Next
    Kill myFile$
    On Local Error GoTo 100
    
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

' error handler
100 X = MsgBox("An error has occured in the file transfer or write.  Please try again later.", vbInformation)
    GetInternetFile = False
    Resume 105
105 End Function

Attribute VB_Name = "Inifile"
Option Explicit

'Variables
Public Username, Password, Folder, Server As String
Public POP3Username, POP3Password, POP3Server As String

Public Title, Category, Content As String

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

' Read INI
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
ReadIni = Left(RetVal, v)
End Function

' Load All Variables
Public Function LoadVariables()

' Filename
Dim m_File As String
m_File = App.Path & "\Mail2BlogX.ini"

'Read from the file
Username = ReadIni(m_File, "Login", "Username")
Password = ReadIni(m_File, "Login", "Password")
Folder = ReadIni(m_File, "Login", "Folder")
Server = ReadIni(m_File, "Login", "Server")

POP3Username = ReadIni(m_File, "Login", "POP3Username")
POP3Password = ReadIni(m_File, "Login", "POP3Password")
POP3Server = ReadIni(m_File, "Login", "POP3Server")

End Function

' Encode Text To Post Data
Public Function Encode(Value As String) As String
   Dim i As Integer

   Encode = Replace(Value, vbCrLf, "%0D%0A")

   For i = 0 To 31
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 33 To 36
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 38 To 47
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 58 To 64
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 91 To 96
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   For i = 123 To 255
   Encode = Replace(Encode, Chr(i), "%" & Hex$(i))
   Next

   Encode = Replace(Encode, " ", "+")
End Function

Public Function UUDecodeToFile(strUUCodeData As String, strFilePath As String)
    On Error Resume Next
    Dim vDataLine   As Variant      'some variables needed for decoding
    Dim vDataLines  As Variant
    Dim strDataLine As String
    Dim intSymbols  As Integer
    Dim intFile     As Integer
    Dim strTemp     As String
    Dim i
    
    If Left$(strUUCodeData, 6) = "begin " Then  'check if it is a encoded file
        strUUCodeData = Mid$(strUUCodeData, InStr(1, strUUCodeData, vbLf) + 1)
    End If
    If Right$(strUUCodeData, 4) = "end" + vbLf Then 'check if "end" is available
        strUUCodeData = Left$(strUUCodeData, Len(strUUCodeData) - 7)
    End If
    intFile = FreeFile
    Open strFilePath For Binary As intFile  'open output file
        vDataLines = Split(strUUCodeData, vbLf)
        For Each vDataLine In vDataLines    'get every line
                strDataLine = CStr(vDataLine)
                intSymbols = Asc(Left$(strDataLine, 1)) 'get number of chars in
                                                        'one line. This is important
                                                        'for decoding
                strDataLine = Mid$(strDataLine, 2, intSymbols)
                For i = 1 To Len(strDataLine) Step 4
                    'now some decoding
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i, 1)) - 32) * 4 + _
                              (Asc(Mid(strDataLine, i + 1, 1)) - 32) \ 16)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 1, 1)) Mod 16) * 16 + _
                              (Asc(Mid(strDataLine, i + 2, 1)) - 32) \ 4)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 2, 1)) Mod 4) * 64 + _
                              Asc(Mid(strDataLine, i + 3, 1)) - 32)
                Next i
                'put the decoded data in the file
                Put intFile, , strTemp
                strTemp = ""
        Next
    'close the file
    Close intFile
End Function

Public Function PostToBlog()

Dim objXMLHTTP As Object
Dim Response

Dim URL As String
URL = "http://" & Server & "/" & Folder & "/" & "Application.asp"

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
objXMLHTTP.Open "POST", URL, False
objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLHTTP.setRequestHeader "Referer", "VisualBasic"
objXMLHTTP.setRequestHeader "User-Agent", "Matthew1471 Blogging Utility"

'--- Encode The Text ---
Content = Encode(Replace(CheckMail.Content, "=" & vbCrLf, ""))
Title = Encode(CheckMail.Title)
Category = Encode(CheckMail.Category)

' Construct the message body first before we send, it is a name/value pair,separated by ampersands
' which looks like "username=admin&password=letmein"
Dim strbody
strbody = "Username=" & Username & "&Password=" & Password & "&Content=" & Content & "&Title=" & Title & "&Category=" & Replace(Category, " ", "+")

' Send It Baby!
objXMLHTTP.send strbody
Response = objXMLHTTP.ResponseText

' Let Them Know We Failed
If Error <> "" Then MsgBox "Error: " & Error

' So Did The Submission Go Ok?
If Response = "Entry Submission Successfull" Then
Debug.Print "Sent"
ElseIf Response = "No Text Entered" Then
MsgBox "No Text Entered"
ElseIf Response = "" Then
MsgBox "Server Not Found"
Kill (App.Path & "\Mail2BlogX.ini")
ElseIf Response = "User/Password Error" Then
MsgBox "Invalid Username/Password"
Kill (App.Path & "\Mail2BlogX.ini")
ElseIf (InStr(Response, "404") <> 0) Or (InStr(Response, "Not Found") <> 0) Then
MsgBox "Blog Not Found"
Kill (App.Path & "\Mail2BlogX.ini")
Else: MsgBox "Server Error"
MsgBox Response
Kill (App.Path & "\Mail2BlogX.ini")
End If

End Function

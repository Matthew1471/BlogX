<!--#include file="../../Includes/xmlrpc.asp" -->
<!--#include file="../../Includes/Config.asp" -->
<%
function pingbackping(sourceURI, targetURI)

'--- DEBUGGING ---'
'sourceURI = "http://matthew1471.co.uk/ExternalSite.asp"
'targetURI = "http://matthew1471.co.uk/Blog/ViewItem.asp?Entry=167"

'---- Verify The Site Is Not Having Us On ----
  Response.Buffer = True
  Dim objXMLHTTP
  Dim Verified, ResponseCase

  ' Create an xmlhttp object:
  Set objXMLHTTP= Server.CreateObject("Microsoft.XMLHTTP")
  'Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")

  ' Opens the connection to the remote server.
  On Error Resume Next
  objXMLHTTP.Open "GET", sourceURI, False
  objXMLHTTP.Send

  If Instr(objXMLHTTP.responseText,targetURI) > 0 Then 
  Verified = True
  ElseIf Err = 0 Then
  ResponseCase = 17
  Else
  ResponseCase = 16  
  End If
  On Error Goto 0

Set objXMLHTTP = Nothing

'--- End of verification ----

Dim Length, Last, Entry


        If (InstrRev(targetURI,"Comments.asp") - 1) > 0 Then 
        Last = InstrRev(targetURI,"Comments.asp") + 12
	ElseIf (InstrRev(targetURI,"ViewItem.asp") - 1) > 0 Then
	Last = InstrRev(targetURI,"ViewItem.asp") + 12
        Else
        Last = 0
        End If

        Length = Len(targetURI)

        If Last > 0 Then targetURI = Replace(Right(targetURI,Length-Last),"Entry=","")
        If IsNumeric(targetURI) Then Entry = targetURI Else Entry = 0

	'Response.Write "Entry : " & Entry & "<br>" & vbCrlf
	'Response.Write "Verified : " & Verified

'--- Databasing Time! ---'
If (Verified = True) AND (AlreadyPinged = False) Then

'Dimension variables
Dim ReferURL                    'The URL

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3

    '-- Check If We Are Banned --'
    Records.Open "SELECT * FROM BannedIP WHERE IP='" & Request.ServerVariables("REMOTE_ADDR") & "';",Database, 1, 3
    If Records.EOF = False Then Banned = True
    Records.Close

Records.Open "SELECT * FROM PingBack WHERE (SourceURI='" & Replace(SourceURI,"'","''") & "' OR IP='" & Request.ServerVariables("REMOTE_HOST") & "' OR Error='" & Replace(targetURI,"'","''") & "') AND (EntryID=" & Entry & ");", Database

If (Not Records.EOF = True) OR (Banned = True) Then

 ResponseCase = 48  

Else

 Records.AddNew
 Records("EntryID") = Entry
 Records("SourceURI") = Left(SourceURI,255)

 If Entry = 0 Then
  Records("Error") = Left(targetURI,80)
  ResponseCase = 33
 End If

 Records("IP") = Request.ServerVariables("REMOTE_HOST")

End If

Records.Update


'#### Close Objects ###
Records.Close
Set Records = Nothing

End If

Select Case ResponseCase
  case 16
    pingbackping = writeFaultXML("16", "Invalid Source", "The source URI does not exist" )
  case 17
    pingbackping = writeFaultXML("17", "No link", "The source URI does not contain a link to the target URI, and so cannot be used as a source." )
  case 33
    pingbackping = writeFaultXML("33", "Not A Pingback Resource", "The specified target URI cannot be used as a target. It is not a pingback-enabled resource.")
  case 48
    pingbackping = writeFaultXML("48", "Already Registered", "The pingback has already been registered." )
  Case Else
    pingbackping = "The pingback has been registered for entry number " & Entry & "."
End select

End Function

call addServerFunction("pingbackping", "pingbackping")
rpcserver

Database.Close
Set Database = Nothing
%>
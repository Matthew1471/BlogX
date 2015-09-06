<% 
' --------------------------------------------------------------------------
'¦Introduction : Mail The Author Validation Page                            ¦
'¦Purpose      : Validates and sends visitor e-mails.                       ¦
'¦Used By      : Mail.asp.                                                  ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦Notes        : This page is for contacting the blog author but requires   ¦
'¦               e-mail settings to be properly configured in the config.   ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
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

'-- Check For A Proxy --'
Dim MyIPAddress
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" Then 
 MyIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
Else
 MyIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If

'-- Filter & Clean --'
Dim RecordID
RecordID = Request.Querystring("RecordID")
If (IsNumeric(RecordID) = False) OR (RecordID = "") Then Response.Redirect "Mail.asp" Else RecordID = Int(RecordID)

Dim PUK
PUK = Request.Querystring("PUK")
If (IsNumeric(PUK) = False) OR (PUK = "") Then Response.Redirect "Mail.asp" Else PUK = Int(PUK)

PageTitle = "Mail Validation"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<div id="content">
<%

'On Error Resume Next

Dim SQL
SQL = "SELECT RecordID, FromEmail, Subject, Body, IP, PUK FROM Mail_Unvalidated WHERE (RecordID=" & RecordID & ") "
 If Session(CookieName) = False Then SQL = SQL & "AND (IP='" & Replace(MyIPAddress,"'","") & "')"
SQL = SQL & "AND (PUK=" & PUK & ");"

'--- Open Database ---'
Records.Open SQL, Database, 1, 3

If Not (Records.EOF = True) Then

  Dim ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo
  ToName = "Webmaster"
  ToEmail = EmailAddress
  From = EmailAddress
  Name = SiteDescription
  ReplyTo = Records("FromEmail")
  Subject = Records("Subject")
  Body = Records("Body")

  '-- Spammers don't often stay connected to pages very long.. are they still here? ok, send --'
  Response.Flush
  If Response.IsClientConnected Then
   %><!--#INCLUDE FILE="Includes/Mail.asp" --><%
  End If

  '-- Purge This --'
  If Err = 0 Then Records.Delete

  'On Error GoTo 0

 ' ---------- Mail Error Checking ----------
 If Err_Msg <> "" Then 
  Response.Write "<p align=""Center"">" & Err_Msg & "</p>"
 Else
  Response.Write "<p align=""Center"">Thank you " & Request.Querystring("Name") & ",<br/>"
  Response.Write " your e-mail has been successfully queued.</p>"
 End If

 Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"

Else

 Response.Write "<p align=""Center"">Sorry " & Request.Querystring("Name") & ", you have specified an invalid numeric, changed your IP address<br/>" 
 Response.Write "or this e-mail has already been sent.</p>"
 Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"

End If

Records.Close
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
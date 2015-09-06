<%
' --------------------------------------------------------------------------
'¦Introduction : Comment Validation Page.                                   ¦
'¦Purpose      : Performs a few additional anti-spam checks on comments     ¦
'¦               before accepting the comment as genuine.                   ¦
'¦Used By      : Comments.asp.                                              ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Mail.asp.                                         ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

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

PageTitle = "Comment Validation"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<div id="content">
<%
'On Error Resume Next

'-- Check For A Proxy --'
Dim MyIPAddress
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" Then 
 MyIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
Else
 MyIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If

'-- Filter & Clean --'
Dim CommentID
CommentID = Request.Querystring("CommentID")
If (IsNumeric(CommentID) = False) OR (CommentID = "") Then CommentID = 0 Else CommentID = Int(CommentID)

Dim PUK
PUK = Request.Querystring("PUK")
If (IsNumeric(PUK) = False) OR (PUK = "") Then PUK = 0 Else PUK = Int(PUK)

'-- Open The Record Ready To Write --'
Records.CursorType = 2
Records.LockType = 3

Dim SQL
SQL = "SELECT EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK FROM Comments_Unvalidated WHERE (CommentID=" & CommentID & ") "
 If Session(CookieName) = False Then SQL = SQL & "AND (IP='" & Replace(MyIPAddress,"'","") & "')"
SQL = SQL & "AND (PUK=" & PUK & ");"

Records.Open SQL, Database

If Not (Records.EOF = True) Then

 '-- Move to new table --'
 Database.Execute "INSERT INTO Comments (EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK) SELECT EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK FROM Comments_Unvalidated WHERE (CommentID=" & CommentID & ");"

 Dim EntryID
 EntryID = Records("EntryID")

 Dim Email
 Name = Records("Name")
 Email = Records("Email")

 '-- And destroy *all* old ones that matched --'
 Records.Delete
 
 '-- Close Now Anyway --'
 Records.Close

 '-- Spammers don't often stay connected to pages very long.. have they gone? ok, kill page --'
 Response.Flush
 If NOT Response.IsClientConnected Then
  Database.Close
  Set Records = Nothing
  Set Database = Nothing
  Response.End
 End If

 '-- Update Comment Count --'
 Records.Open "SELECT RecordID, Comments, LastModified FROM Data WHERE RecordID=" & EntryID, Database
  Records("Comments") = Records("Comments")+1
  Records("LastModified") = Now()
 Records.Update
 Records.Close

	'-- Honor Subscriptions --'
	If EnableEmail <> False Then

	Dim MailBody, ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo

        '-- Grab the new CommentID --'
        Records.Open "SELECT CommentID, EntryID FROM Comments WHERE EntryID=" & EntryID & " ORDER BY CommentID DESC",Database
         CommentID = Records("CommentID")	
        Records.Close

	Records.Open "SELECT CommentID, EntryID, Name, Email, Subscribe, PUK FROM Comments WHERE Subscribe=True AND EntryID=" & EntryID & " AND Email <> '" & Replace(Email,"'","") & "' AND Email <> '" & EmailAddress & "' AND Email <> '' ORDER BY Email",Database, 1, 3
	Do Until (Records.EOF)

	MailBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf
	MailBody = MailBody & "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en""> " & VbCrlf
	MailBody = MailBody & "<head>" & VbCrlf

	MailBody = MailBody & "<title>Comment submitted on " & SiteDescription & "</title>"

	If TemplateURL = "" Then
	 MailBody = MailBody & "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
	Else
	 MailBody = MailBody & "<link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
	End If

        MailBody = MailBody & "</head>" & VbCrlf
        MailBody = MailBody & "<body style=""background-color:" & BackgroundColor & """>" & VbCrlf

        MailBody = MailBody & "<div class=""content"" style=""text-align:center"">" & VbCrlf
        MailBody = MailBody & "<br/>" & VbCrlf

        MailBody = MailBody & "<div class=""entry"" style=""width: 50%; margin:0 auto"">" & VbCrlf
        MailBody = MailBody & "<h3 class=""entryTitle"">Notification Of Comment Added</h3>" & VbCrlf
        MailBody = MailBody & "<div class=""entryBody"">" & VbCrlf

        MailBody = MailBody & "<p>You are receiving this e-mail as a user (" & Name & ") has submitted a <a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & "#Comment" & CommentID & """>comment</a> on " & SiteDescription & ".</p>" & VbCrlf

        MailBody = MailBody & "</div>" & VbCrlf
        MailBody = MailBody & "</div>" & VbCrlf

        MailBody = MailBody & "<p>To stop receiving update notification for this entry, click <a class=""standardsButton"" href=""" & SiteURL & "CommentNotify.asp?Entry=" & EntryID & "&amp;Email=" & Records("Email") & "&amp;PUK=" & Records("PUK") & """>Unsubscribe</a></p>" & VbCrlf

        MailBody = MailBody & "<p>BlogX V" & Version & "</p>" & VbCrlf

        MailBody = MailBody & "</div>" & VbCrlf
        MailBody = MailBody & "</body>" & VbCrlf
        MailBody = MailBody & "</html>" & VbCrlf

			ToName = Records("Name")
			ToEmail = Records("Email")
			From = NoEmailAddress
			Name = SiteDescription

			Subject = "Blog : Comment Added (Entry #" & EntryID & ")"
			Body = MailBody
	        %><!--#INCLUDE FILE="Includes/Mail.asp" --><%
	Records.MoveNext
	Loop
	End If
	'-- End Of User Subscriptions --'

	If (CommentNotify <> 0) AND (Session(CookieName) <> True) AND (EnableEmail <> False) Then

	MailBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf
	MailBody = MailBody & "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en""> " & VbCrlf
	MailBody = MailBody & "<head>" & VbCrlf

	MailBody = MailBody & "<title>Comment submitted on " & SiteDescription & "</title>"

	 If TemplateURL = "" Then
	  MailBody = MailBody & "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
	 Else
	  MailBody = MailBody & "<link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
	 End If

	MailBody = MailBody & "</head>" & VbCrlf
        MailBody = MailBody & "<body style=""background-color:" & BackgroundColor & """>" & VbCrlf

        MailBody = MailBody & "<div class=""content"" style=""text-align:center"">" & VbCrlf
        MailBody = MailBody & "<br/>" & VbCrlf

	MailBody = MailBody & "<div class=""entry"" style=""width: 50%; margin:0 auto"">" & VbCrlf
	MailBody = MailBody & "<h3 class=""entryTitle"">Notification Of Comment Added</h3>" & VbCrlf
	MailBody = MailBody & "<div class=""entryBody"">" & VbCrlf

	MailBody = MailBody & "<p>You are receiving this e-mail as a user has submitted a <a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & "#Comment" & CommentID & """>comment</a> on " & SiteDescription & ".</p>" & VbCrlf
	MailBody = MailBody & "</div>" & VbCrlf
	MailBody = MailBody & "</div>" & VbCrlf

	MailBody = MailBody & "<p>BlogX V" & Version & "</p>" & VbCrlf

	MailBody = MailBody & "</div>" & VbCrlf
	MailBody = MailBody & "</body>" & VbCrlf
	MailBody = MailBody & "</html>" & VbCrlf

			ToName = "Webmaster"
			ToEmail = EmailAddress
			From = EmailAddress
			If Email <> "" Then ReplyTo = Name & "<" & Email & ">"
			Name = SiteDescription

			Subject = "Blog : Comment Added (Entry #" & EntryID & ")"
			Body = MailBody
%>
                        <!--#INCLUDE FILE="Includes/Mail.asp" -->
<%
	End If

 Response.Write "<p align=""Center"">Comment Submission (#" & Request.Querystring("CommentID") & ") successful.</p>"
 Response.Write "<p align=""Center""><a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & """>Back</font></a></p>"

ElseIf Len(Request.Querystring("CommentID")) > 0 Then
 Response.Write "<p align=""Center"">Sorry but comment " & Request.Querystring("CommentID") & " was not found, you have specified an invalid security code,<br/>"
 Response.Write "changed your IP address or possibly already validated this comment.</p>" & VbCrlf
 Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"
Else
 Response.Write "<p align=""Center"">Sorry but you did not specify a security code.</p>" & VbCrlf
 Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"
End If

Records.Close
%>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
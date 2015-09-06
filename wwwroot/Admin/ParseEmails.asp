<%
' --------------------------------------------------------------------------
'¦Introduction : ArgoSoft E-mail Server Blog Via Mail Import Page.          ¦
'¦Purpose      : Provides a way to queue up blog entries via e-mail.        ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Folder with EML files in.             ¦
'¦Standards    : XHTML Strict.                                              ¦
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
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<%
Dim MailMessage, FilesExist
Set MailMessage = Server.CreateObject("Scripting.FileSystemObject")

'-- Change this to the inbox of the blog e-mail account user --'
Dim Path
Path = "C:\Program Files\Argo Software Design\Mail Server\_users\_nodomain\Blog\Inbox\"

On Error Resume Next
 Dim Folder, Files
 Set Folder = MailMessage.GetFolder(Path)  
 Set Files = Folder.files

 If Err <> 0 Then 
  Set MailMessage = Nothing
  Database.Close
  Set Database = Nothing
  Set Records = Nothing
  Response.Write "<p style=""text-align: center"">You cannot import entries via e-mail for the following reason:<br/>" & VbCrlf

  If Err.Number = 424 Then
   Response.Write "The folder could not be opened, check the Path variable in Admin\ParseEmails.asp." & VbCrlf
  Else
   Response.Write Err.Description & " (" & Err.Number & ")" & VbCrlf
  End If

  Response.Write "</p>" & VbCrlf 

  Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

  Response.Write "</div>" & VbCrlf
  %><!-- #INCLUDE FILE="../Includes/Footer.asp" --><%
  Response.End
 End If

For Each File in Files 

 If MailMessage.GetExtensionName(Path & File.Name) = "eml" Then
  FilesExist = True

  '-- Map the physical system path---'
  FilePath = Path & File.Name

  '-- Get a handle to the file --'
  Dim Email	
  Set Email = MailMessage.GetFile(FilePath)

  '-- Open the file --'
  Dim EmailStream

  '-- Read the file line by line --'
  Set EmailStream = Email.OpenAsTextStream(1, -2)

  Dim ReadText, SentDate, SentTime   

  Response.Write "<li> FileName : " & File.Name & "</li><br/>" & VbCrlf

   Do While Not EmailStream.AtEndOfStream
    ReadText = EmailStream.Readline

     If InStr(ReadText,"Subject: [BlogX] ") <> 0 Then
      Authorised = True
      Subject = Replace(ReadText,"Subject: [BlogX] ","")
      Response.Write "<li> Subject : " & Subject & "</li><br/>" & VbCrlf
     End If

     If Instr(ReadText,"Date:") <> 0 Then 
      SentDate = Replace(ReadText,"Date: ","")
      Length = Len(SentDate)

      '-- Take Off Day Name --'
      SentDate = Right(SentDate,Length-5)

      '-- Take Off GMT Markup --'
      SentDate = Left(SentDate,Length-11)

      SpacePos = InStrRev(SentDate," ")

      Length = Len(SentDate)
      SentTime = Right(SentDate,Length-SpacePos)
      SentDate = Left(SentDate,SpacePos)

      Response.Write "<li> Date : " & SentDate & "</li><br/>" & VbCrlf
      Response.Write "<li> Time : " & SentTime & "</li><br/>" & VbCrlf
     End If

     If InStr(ReadText,"Category: ") <> 0 Then
      EntryCat = Replace(ReadText,"Category: ","")
      Response.Write "<li> Category : " & EntryCat & "</li><br/>"
     End If

     If (InStr(ReadText,":") = 0) AND (InStr(ReadText,"=") = 0) AND (Subject <> "") Then
      Body = Body & VbCrlf & ReadText
      Response.Write VbCrlf & "<br/>" & ReadText
     End If

     '-- Ignore the HTML e-mails --'
     If Instr(ReadText,"This is a multi-part message in MIME format.") <> 0 Then Authorised = False

   Loop   

   EmailStream.Close
   Set EmailStream = nothing
   
   If Authorised = True Then
   
    '-- Filter & Clean --'
    EntryCat = Replace(EntryCat,"'","&#39;")
    EntryCat = Replace(EntryCat," ","%20")

    '-- Open The Records Ready To Write --'
    Records.Open "SELECT Title, Text, Category, Day, Month, Year, Time FROM Data", Database, 2, 3
     Records.AddNew
     Records("Title") = Subject
     Records("Text") = Body
     Records("Category") = EntryCat

     Records("Day") = Day(SentDate)
     Records("Month") = Month(SentDate)
     Records("Year") = Year(SentDate)
     Records("Time") = SentTime
     Records.Update
    Records.Close

    Response.Write "<p style=""text-align:center"">Entry Submission Successful</p>"
    MailMessage.DeleteFile(FilePath)
   Else
    Response.Write "<p style=""text-align:center"">Invalid Auth Details (Or <b>NOT</b> in ""Plain Text"" Format)</p>"
   End If
   
 End If
Next  

Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

Set MailMessage = nothing
Set Folder = Nothing
Set Files = Nothing

If NOT FilesExist Then
 Response.Write "<p style=""text-align:center"">No New E-mails</p>"
 Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"
End If
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
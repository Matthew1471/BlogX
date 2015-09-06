<%
' --------------------------------------------------------------------------
'¦Introduction : Add File Upload Page.                                      ¦
'¦Purpose      : Performs the file upload for blog administrator.           ¦
'¦Used By      : Admin/AddFile.asp.                                         ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/Footer.asp,       ¦
'¦               Templates/Config.asp.                                      ¦
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
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - Add Shared File</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

 <!--
 //= - - - - - - - 
 // Copyright 2004-08, Matthew Roberts
 // Copyright 2003, Chris Anderson
 // 
 // Usage Of This Software Is Subject To The Terms Of The License
 //= - - - - - - -
 -->
 <% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme") %>
 <!-- #INCLUDE FILE="../Templates/Config.asp" -->
 <%
 If TemplateURL = "" Then
  Response.Write "<link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>"
 Else
  Response.Write "<link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>"
 End If

On Error Resume Next

Dim Upload
Set Upload = Server.CreateObject("Persits.Upload")
 Upload.ProgressID = Request.QueryString("PID")
 Upload.IgnoreNoPost = True
 Upload.OverwriteFiles = True

 '-- New line added for any security restrictions that might have been imposed --'
 '-- See http://www.aspupload.com/object_upload.html#SaveVirtual --'
 Count = Upload.Save(SharedFilesPath)
 'Count = Upload.SaveVirtual("..\Images\Articles\Temp")

 If Err.Number = 424 Then
  Message = "Error: This Page Requires<br/><a href=""http://aspupload.com/"">Persits ASP Upload</a>."
 ElseIf Err.Number = 6 Then
  Message = "Error: You are out of disk space on the server."

  '-- Undo the upload to prevent deadlock --'
  For Each File in Upload.Files
   File.Delete
  Next

 ElseIf Err.Number <> 0 Then
  Message = "Error: " & CStr(Err.Description) & "(" & Err.Number & ")"
 Else

  For Each File in Upload.Files

   '-- Check allowed file types --'
   Records.Open "SELECT AllowedExtension FROM FileExtensions WHERE UCase(AllowedExtension) ='" & UCase(Right(Replace(File.filename,"'",""), 3)) & "';", Database, 0, 1

    If (Records.EOF) Then
     Dim Message
     If Len(Message) > 0 Then Message = Message & "<br/>" & VbCrlf
     Message = Message & "Error: The file &quot;" & File.filename & "&quot; was not an allowed file type and has been deleted."
     File.Delete
    Else
     If Len(Message) > 0 Then Message = Message & "<br/>" & VbCrlf
     Message = Message & "The file &quot;" & File.filename & "&quot; was saved."
    End If 

   Records.Close

  Next

  If Len(Message) = 0 AND Len(CStr(Err.Description)) > 0 Then Message = "<p align=""center"">Error : " & Err.Description & "</p>"

 End If
 
Set Upload = Nothing
On Error GoTo 0

Response.Write "</head>" & VbCrlf
Response.Write "<body style=""background-color: " & BackgroundColor & """>" & VbCrlf

 Response.Write " <p style=""text-align:center; font-family:Verdana, Arial, Helvetica;font-size:large"">"

  If (Len(Message) <> 0) Then
   Response.Write "  " & Message & VbCrlf
  ElseIf (Count <= 0) Then
   Response.Write "  Nothing was uploaded,<br/>Check you can write to both folders." & VbCrlf
  End If

 Response.Write " </p>" & VbCrlf

 Response.Write " <p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"
%>
</body>
</html>
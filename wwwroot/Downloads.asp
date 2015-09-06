<%
' --------------------------------------------------------------------------
'¦Introduction : Downloads Page.                                            ¦
'¦Purpose      : This displays any shared files that the owner has shared.  ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Includes/ViewerPass.asp,              ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
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

Function NoTraverse(Path)
 NoTraverse = Path
 NoTraverse = Replace(NoTraverse, "..", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "%2E","", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "INCLUDES", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "http:", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "www.", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, ".com", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "/", "", 1, -1, 1)
 NoTraverse = Replace(NoTraverse, "\", "", 1, -1, 1)
End Function

PageTitle = "Downloads"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->

<div id="content">

 <div class="entry">
 <h3 class="entryTitle">Download Shared Files</h3><br/>

 <div class="entryBody">
 <%
 Dim Folder
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")

 On Error Resume Next

  Set Folder = FSO.GetFolder(SharedFilesPath)

  If (Err.Number = 76) Then 
   Response.Write "<p style=""text-align: center"">The Site Administrator has not set this up to use a valid path.<br/> "
   Response.Write "To allow shared folders edit the SharedFolders path in Config.asp</p>"
  ElseIf Err.Number <> 0 Then 
   Response.Write "Err : " & Err.Decscription
  End If

  Dim NumberOfFiles
  NumberOfFiles = Folder.Files.Count

 On Error GoTo 0

 '-- We Would Like To Delete Something --'
 If (Session(CookieName) = True) AND (Request.Querystring("Delete") <> "") Then 
 %><!-- #INCLUDE FILE="Admin/Admin.asp" --><%
  On Error Resume Next

   Dim DelFile
   Set DelFile = FSO.GetFile(SharedFilesPath & "\" & NoTraverse(Request.Querystring("Delete")))
   If Err.Number <> 0 Then Response.Write "<p style=""text-align:center"">Error While Processing - " & Err.Description & "</p>"
   DelFile.Delete

   If Err.Number = 0 Then 
    Response.Write "<p style=""text-align:center"">Deleted - " & Request.Querystring("Delete") & "</p>" 
   Else
    Response.Write "<p style=""text-align:center"">Error Deleting - " & Err.Description & "</p>"
   End If

   Set DelFile = Nothing

  On Error GoTo 0

 End If

If NumberOfFiles > 0 Then
 Response.Write "  <div class=""MainMargin"">" & VbCrlf
 Response.Write "   <table border=""1"" cellspacing=""0"" style=""align:center; margin: 0 auto;"">" & VbCrlf
 Response.Write "    <tr>" & VbCrlf

 Response.Write "     <th style=""background-color: #0066cc"">Download"
  If (Session(CookieName) = True) Then Response.Write "/Delete"
 Response.Write "</th>" & VbCrlf

 Response.Write "     <th style=""background-color: #0066cc"">Name</th>"
 Response.Write "     <th style=""background-color: #0066cc"">Size</th>"
 Response.Write "     <th style=""background-color: #0066cc"">Type/Editor</th>"
 Response.Write "     <th style=""background-color: #0066cc"">Modified</th>"
 Response.Write "    </tr>" & VbCrlf

  On Error Resume Next
    For Each File in Folder.Files
     If File.Type <> "ASP File" Then

      Dim TotalSize
      TotalSize = File.Size
      If TotalSize >= 1073741824 Then
       TotalSize = TotalSize / 1024 / 1024 / 1024
       TotalSize = FormatNumber(TotalSize, 2) & "&nbsp;GB"
      ElseIf TotalSize >= 1048576 Then
       TotalSize = TotalSize / 1024 / 1024
       TotalSize = FormatNumber(TotalSize, 2) & "&nbsp;MB"
      ElseIf TotalSize >= 1024 Then
       TotalSize = TotalSize / 1024
       TotalSize = FormatNumber(TotalSize, 2) & "&nbsp;KB"
      ElseIf TotalSize < 1024 Then
       TotalSize = TotalSize & " Bytes"
      End If

      Dim EncodedFileName
      EncodedFileName = Replace(File.Name," ","%20")
      EncodedFileName = Replace(EncodedFileName,"&","&amp;")

      Response.Write "    <tr>" & VbCrlf
      Response.Write "     <td align=""center"">" & VbCrlf
      Response.Write "      <a href=""" & SiteURL & "Download/Files/" & EncodedFileName & """><img alt=""Download"" src=""" & SiteURL & "Images/Download.gif""/></a>" & VbCrlf
      If (Session(CookieName) = True) Then Response.Write "      &nbsp;<a href=""?Delete=" & EncodedFileName & """><img alt=""Delete"" src=""" & SiteURL & "Images/Delete.gif""/></a>" & VbCrlf
      Response.Write "     </td>" & VbCrlf

      Response.Write "     <td><a href=""" & SiteURL & "Download/Files/" & EncodedFileName & """><span style=""color:blue"">" & Replace(File.Name,"&","&amp;") & "</span></a></td>" & VbCrlf
      Response.Write "     <td align=""right"">" & TotalSize & "</td>" & VbCrlf
      Response.Write "     <td>" & File.Type & "</td>" & VbCrlf
      Response.Write "     <td>" & File.DateLastModified & "</td>" & VbCrlf
      Response.Write "    </tr>" & VbCrlf
     End If
    Next
	
    Set File   = Nothing
    Set FSO    = Nothing
    Set Folder = Nothing

    On Error GoTo 0
    Response.Write "</table>" & VbCrlf
    ResponSe.Write "<br/>" & VbCrlf
    Response.Write "</div>" & VbCrlf
End If

Response.Write " </div>" & VbCrlf
Response.Write " </div>" & VbCrlf
Response.Write "</div>" & VbCrlf
%>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->					
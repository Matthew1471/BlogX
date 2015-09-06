<%
' --------------------------------------------------------------------------
'¦Introduction : Not Found Page                                             ¦
'¦Purpose      : Shows the error message following the consistent theme     ¦
'¦               and on my server offers some suggestions.                  ¦
'¦Used By      : Your webserver if you have configured it to.               ¦
'¦Requires     : Includes/Header.asp, Includes/ViewerPass.asp,              ¦
'¦               Includes/Nav.asp, Includes/Footer.asp                      ¦
'¦Notes        : This page is important to maintain a consistent user       ¦
'¦               experience but it requires configuration on your part.     ¦
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

 'On Error Resume Next
 Dim BadLink, Refer
 BadLink = Replace(Request.QueryString,":80","")
 BadLink = Replace(BadLink, "404;", "")
 BadLink = HTML2Text(BadLink)
 Refer = HTML2Text(Request.ServerVariables("HTTP_REFERER"))

 '-- I've moved this to the "Official" folder --'
 If Right(LCase(BadLink),10) = "/count.asp" Then
  Response.Clear
  Server.Transfer "Official/Count.asp"
 End If

 '-- This is an old link --'
 If Instr(1, BadLink, "/Blog/", 1) <> 0 Then Response.Redirect Replace(BadLink,"Blog/","",1,1,1)

Response.Status = "404 Not Found"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<div id="content">
 <div class="entry">
  <h3 class="entryTitle">Page/File Not Found</h3><br/>
  <div class="entryBody">
   <% 
   If InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0 Then
    Response.Write "<p>Were you trying to <a href=""/Download/WinBlogX%20Setup.exe"">Download WinBlogx</a> or <a href=""/Download.asp"">Download WebBlogx</a>?</p>" & VbCrlf
   End If
   %>

   <p>It appears that you have stumbled upon a page that is not present on this web site.<br/>
   It could have been moved, spelt incorrectly or it may even be in our plans to expand the site and develop this page.</p>

   <p><b>Error : </b> File "<%=BadLink%>" not found.<br/>
   <b>Referrer : </b> <%
If (Refer <> "") AND (Refer <> BadLink) Then 
 Response.Write "<a href=""" & Refer & """>" & Refer & "</a>"

 If Refer <> SiteURL & "Admin/NotFound.asp" Then
  '-- Open The Records Ready To Write --'
  Records.Open "SELECT URL, ReferringPage, ErrorCount FROM NotFound WHERE URL='" & Left(Replace(BadLink,"'","''"),255) & "';", Database, 0, 3

   If NOT Records.EOF Then
    Records("ErrorCount") = Int(Records("ErrorCount")) + 1
   Else
    Records.AddNew
    Records("URL") = Left(BadLink,255)
    Records("ReferringPage") = Left(Refer,255)
    Records("ErrorCount") = 1
   End If

   Records.Update
   Records.Close
 End If

Else
 Response.Write "You typed in the address manually."
End If
%></p>
  </div>
 </div>

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
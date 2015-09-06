<%
' --------------------------------------------------------------------------
'¦Introduction : About This Blog Page.                                      ¦
'¦Purpose      : If the user has some "About Me" text then this page shows  ¦
'¦               it with a link to viewing the blog, if not it will load it.¦
'¦Used By      : IIS.                                                       ¦
'¦Requires     : Includes/Header.asp, Includes/ViewerPass.asp,              ¦
'¦               Includes/Cache.asp, Includes/NAV.asp, Includes/Footer.asp  ¦
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

Response.Buffer = True

PageTitle = "Home Page"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<% If EnableMainPage <> True Then
Database.Close
Set Records = Nothing
Set Database = Nothing

'*** WHY IS THIS PAGE ERORRING ON MY IIS 3/4????
'	Answer : ASP 2.0 does not know "Server.Transfer"
'	Solution : Replace Server.Transfer with Response.Redirect
'***********************************************

Response.Clear

On Error Resume Next
Server.Transfer("Main.asp")
Response.Redirect("Main.asp")
On Error Goto 0
Else
%>
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<%
'--- Open RecordSet ---'
Records.Open "SELECT MainID, MainText, LastModified FROM Main",Database, 1, 3

If NOT Records.EOF Then

Dim MainID, MainText, LastModified

'--- Setup Variables ---'
   MainID = Records("MainID")
   MainText = Records("MainText")
   LastModified = Records("LastModified")

End If

Records.Close

If (NOT DontSetModified) AND (Session(CookieName) = False) AND (Request.Cookies(CookieName) <> "True") Then

 '-- Not every post has been modified --'
 If IsNull(LastModified) Then LastModified = GeneralModifiedDate

 '-- Proxy Handler --'
 CacheHandle(LastModified)

 'Sun, 12 Aug 2007 09:58:50 GMT
 'Response.Write "<!-- Page Last Modified.. " & PubDate & "-->"

End If

%>
<div id="content">

<!-- Start Header -->
<div class="date" id="Main">
<h2 class="dateHeader"><%=SiteSubTitle%></h2>
<!-- End Header -->
</div>

<!-- Start Content -->
<div class="entry">
<h3 class="entryTitle"><%=SiteDescription%><%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your About Page""><a href=""Admin/EditMainPage.asp""><img alt=""Edit This Text"" src=""Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3>
<div class="entryBody"><%=LinkURLs(Replace(MainText, vbcrlf, "<br/>" & vbcrlf)) %>

<p style="text-align:center"><a href="Main.asp">View The Blog</a></p>
</div>
<p class="entryFooter">
<% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""E-mail Me"" src=""Images/Email.gif"" style=""border-style: none""/></a></acronym>"%></p></div>
<!-- End Content -->

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%End If%>
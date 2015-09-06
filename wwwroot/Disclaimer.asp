<%
' --------------------------------------------------------------------------
'¦Introduction : Disclaimer Page.                                           ¦
'¦Purpose      : This will display any legal or general site disclaimers.   ¦
'¦Used By      : Includes/Header.asp.                                       ¦
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

PageTitle = "Disclaimer"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<%
'--- Open set ---'
Set Records = Server.CreateObject("ADODB.recordset")
    Records.Open "SELECT DisclaimerText, LastModified FROM Disclaimer",Database, 0, 1

 If NOT Records.EOF Then
  '--- Setup Variables ---'
  Dim DisclaimerText, LastModified
  DisclaimerText = Records("DisclaimerText")
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

 <!--- Start Header -->
 <div class="date" id="Disclaimer">
  <h2 class="dateHeader">Disclaimer</h2>
 </div>
 <!--- End Header -->

 <!--- Start Disclaimer -->
 <div class="entry">
  <h3 class="entryTitle">Disclaimer <%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your Disclaimer Page""><a href=""Admin/EditDisclaimer.asp""><img alt=""Edit Disclaimer"" src=""Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3>
  <div class="entryBody"><%=Replace(DisclaimerText, vbcrlf, "<br/>" & vbcrlf)%>
  </div>

  <p class="entryFooter">
  <% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><img alt=""E-mail Me"" src=""Images/Email.gif"" style=""border-style: none""/></a></acronym>"%></p></div>
  <!--- End Disclaimer -->

  <p style="text-align:center"><a href="<%=PageName%>">Back To The Main Page</a></p>

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
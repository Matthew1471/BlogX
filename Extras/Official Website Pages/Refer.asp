<%
' --------------------------------------------------------------------------
'¦Introduction : Version Update Hijack Page.                                ¦
'¦Purpose      : Display new version information before redirecting.        ¦
'¦Used By      : WinBlogX.                                                  ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp.¦
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

PageTitle = "Update Information"
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<div id="content">

 <!--- Start Information -->
 <div class="entry">
  <h3 class="entryTitle">Important Information About Your WinBlogX</h3>
  <div class="entryBody">
  <p><% If Request.Querystring("Refer")="WinBlogX" AND Request.Querystring("Version")< "1.04.14" Then%>
  You Are Using An <b>OLD</b> Version Of WinBlogX.<br/>
  There is a newer version of <a href="/About.asp"><%=Request.Querystring("Refer")%></a> than V<%=Request.Querystring("Version")%>
  <% Else Response.Redirect "../Default.asp"
  End If %></p>
  </div>
 </div>
 <!--- End Information -->

 <p style="text-align: center"><a href="<%=PageName%>">Back To The Main Page</a></p>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
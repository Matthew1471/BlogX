<%
' --------------------------------------------------------------------------
'¦Introduction : No SNS Sites Added Help.                                   ¦
'¦Purpose      : This provides information on adding SNS feeds.             ¦
'¦Used By      : WinBlogX SNS.                                              ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Cache.asp.                                        ¦
'¦Ensures      : WinBlogX SNS users understand how to add SNS feeds.        ¦
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

PageTitle = "SNS Feed Help"
%>
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<%
'-- Proxy Handler --'
CacheHandle(GeneralModifiedDate)
%>
<div id="content">

 <div class="entry">
 <h3 class="entryTitle">How do I add an SNS feed?</h3><br/>
  <div class="entryBody">
   <p>SNS (<b>S</b>imple <b>N</b>otification <b>S</b>ervice) is a new and exciting feature currently in development.</p>
   <p>It allows you to know INSTANTLY if a blog has been updated on your readers list.</p>
   <p>It is similar to RSS yet is guaranteed to save bandwith and can be used as an extension to RSS.</p>
   <p>To add a site click a link to an SNS feed.
   <br/> Like this : <a class="standardsButton" onclick="makeSNSRequest('http://blogx.co.uk/SNS/')" href="#">SNS</a></p>
  </div>
 </div>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
<%
' --------------------------------------------------------------------------
'¦Introduction : SNS Information Page.                                      ¦
'¦Purpose      : Explains to the user what SNS is.                          ¦
'¦Used By      : BlogX SNS icon error (Includes/Header.asp).                ¦
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

PageTitle = "SNS Information"
%>
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<%
'-- Proxy Handler --'
CacheHandle(GeneralModifiedDate)
%>
<div id="content">

 <div class="entry">
 <h3 class="entryTitle">What is SNS?</h3><br/>
  <div class="entryBody">
   <p>SNS (<b>S</b>imple <b>N</b>otification <b>S</b>ervice) is a new and exciting feature currently in development.</p>
   <p>It allows you to know INSTANTLY if a blog has been updated on your readers list.</p>
   <p>It is similar to RSS yet is guaranteed to save bandwith and can be used as an extension to RSS.</p>
   <p>Unfortunatly the program has not been finished yet, but this is a place holder for when it has.</p>
   <p>SNS was originially going to be called RTS (<b>R</b>eally <b>T</b>iny <b>S</b>yndication) 
   but the name was dropped to prevent confusion with other definitions using this acronym.</p>
  </div>
 </div>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
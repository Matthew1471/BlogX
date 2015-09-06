<%
' --------------------------------------------------------------------------
'¦Introduction : Hacker Page                                                ¦
'¦Purpose      : Warns user that we detected their SQL exploit attempt.     ¦
'¦Used By      : Any page that relies on Includes/Calendar_Querystrings.asp.¦
'¦Requires     : Includes/Header.asp, Includes/Nav.asp, Include/Cache.asp   ¦
'¦               Includes/Footer.asp                                        ¦
'¦Notes        : This page issues a pro-active warning to those attempting  ¦
'¦               to hack the database.                                      ¦
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

PageTitle = "Hack Attempt"
%>
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<%
'-- Proxy Handler --'
If (NOT DontSetModified) AND (Session(CookieName) = False) AND (Request.Cookies(CookieName) <> "True") Then CacheHandle(GeneralModifiedDate)
%>
<div id="content">
        <p align="Center">I'm sorry but hacking is not currently supported by this server.</p>
        <p align="Center">Please do not make any further attempts</p>
        <p align="Center"><b>Your IP address has been logged and you will be reported.</b></p>
        <p align="Center"><a href="<%=PageName%>">Back</font></a></p>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
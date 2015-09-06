<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : Templates Redirect.                                        ¦
'¦Purpose      : If a user is looking to browse our templates folder send   ¦
'¦               them back to the themes page.                              ¦
'¦Used By      : User.                                                      ¦
'¦Requires     : /Themes.asp.                                               ¦
'¦Ensures      : No sites where the user could browse the templates.        ¦
'¦Standards    : N/A.                                                       ¦
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
Response.Redirect "../Themes.asp" %>
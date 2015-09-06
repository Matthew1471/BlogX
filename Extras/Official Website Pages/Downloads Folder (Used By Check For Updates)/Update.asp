<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : Old Legacy WinBlogX Update Information.                    ¦
'¦Purpose      : Provides version update information for WinBlogX.          ¦
'¦Used By      : WinBlogX.                                                  ¦
'¦Requires     : Nothing.                                                   ¦
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
Response.ContentType = "text/plain"
Response.Write "1.4.15"
%>
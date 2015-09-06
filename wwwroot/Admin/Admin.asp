<%
' --------------------------------------------------------------------------
'¦Introduction : Admin Login Check.                                         ¦
'¦Purpose      : Checks that the user is logged in as the blog administrator¦
'¦               and thus is allowed access to that administrative page.    ¦
'¦Used By      : All administrative pages.                                  ¦
'¦Requires     : /Default.asp                                               ¦
'¦Standards    : N/A.                                                       ¦
'---------------------------------------------------------------------------

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

If (Session(CookieName) = False or IsNull(Session(CookieName)) = True) AND (Request.Cookies(CookieName) <> "True") Then
 Dim QuerystringURL
 If Len(Request.Querystring) > 0 Then QueryStringURL = "&Query=" & Server.URLEncode(Request.Querystring)
 Response.Redirect"../Default.asp?Return=" & Server.URLEncode(Right(Request.ServerVariables("URL"),Len(Request.ServerVariables("URL"))-1)) & QueryStringURL 
End If%>
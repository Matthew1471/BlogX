<%
' --------------------------------------------------------------------------
'¦Introduction : Spell Check Error Page.                                    ¦
'¦Purpose      : Explains why the spell check cannot continue.              ¦
'¦Used By      : Admin/Spell.asp.                                           ¦
'¦Requires     : Includes/Config.asp, Admin.asp.                            ¦
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
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<%
Dim Title, Text

Select Case Request.Querystring("Error")
 Case 1
  Title = "The dictionary file could not be found"
  Text  = "The file <br/><span style="color:red">Includes/Dictionary/dict-large.txt</span><br/> (or its specified alternative) is missing."
 Case 2
  Title = "FSO Not Enabled On Your Website"
  Text  = "The call to Server.CreateObject failed,<br/>Ask your Webhost for FSO support."
 Case Else
  Title = "Unknown Error"
  Text  = "An unknown error number was passed."
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - SpellCheck Error</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 <!--
 //= - - - - - - - 
 // Copyright 2004, Matthew Roberts
 // Copyright 2003, Chris Anderson
 //= - - - - - - -
 -->
 <link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet"/>
</head>
<body>
 <p style="text-align: center"><%=Title%><br/>SpellCheck is now disabled.</p>
 <p style="text-align: center"><b>Technical :</b> <%=Text%></p>
 <%
 Database.Close
 Set Database = Nothing
 Set Records = Nothing
 %>
 <p style="text-align: center"><a href="JavaScript:self.close()">Close Window</a></p>
</body>
</html>
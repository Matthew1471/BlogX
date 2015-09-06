<%@EnableSessionState = False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : Who Uses BlogX OPML Feed.                                  ¦
'¦Purpose      : This provides an export of the ScriptRefer table in the    ¦
'¦               OPML format, useful for importing into RSS readers.        ¦
'¦Used By      : RSS Readers.                                               ¦
'¦Requires     : Includes/Config.asp, Includes/RSSReplace.asp.              ¦
'¦Ensures      : Users are able to easily add BlogX blogs to their RSS feed.¦
'¦Standards    : OPML 1.1.                                                  ¦
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
Response.ContentType = "text/xml"
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Includes/RSSReplace.asp" --><?xml version="1.0" encoding="UTF-8"?>
 <opml version="1.1">
 <head>
  <title>Users of Matthew1471's BlogX</title>
  <dateCreated>Sat, 6 Sep 2008 15:46:30 +0100</dateCreated>
 </head>
 <body>
  <outline text="BlogX">
  <%
  Records.Open "SELECT ReferURL, Approved, ReferHits FROM ScriptRefer WHERE Approved ORDER BY ReferHits DESC",Database, 0, 1

   Do Until (Records.EOF) 
    Response.Write "   <outline text=""" & Encode(Records("ReferURL")) & """ title=""" & Encode(Records("ReferURL")) & """ xmlUrl=""" & Encode(Records("ReferURL")) & "RSS/"" htmlUrl=""" & Encode(Records("ReferURL")) & """ type=""rss""></outline>" & VbCrlf
    Records.MoveNext
   Loop

  Records.Close
  Database.Close
  Set Records  = Nothing
  Set Database = Nothing
 %>
  </outline>
 </body>
</opml>
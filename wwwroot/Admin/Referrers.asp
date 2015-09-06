<%
' --------------------------------------------------------------------------
'¦Introduction : Referrers Page.                                            ¦
'¦Purpose      : Provides a list of HTTP referrers to this blog.            ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/Replace.asp.                 ¦
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
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<table border="0">
 <%
 '-- Open The Records Ready To Write --'
 Records.Open "SELECT TOP 150 ReferURL, ReferHits FROM Refer ORDER BY ReferHits DESC;", Database

  Count = 0

  Do Until (Records.EOF)
   URL = HTML2Text(Records("ReferURL"))
   Count = Count + Records("ReferHits")
   Response.Write " <tr>" & VbCrlf
   Response.Write "  <td>"
   If InStr(URL,"http://") <> 0 Then Response.Write "<a href=""" & URL & """>"
   Response.Write URL
   If InStr(URL,"http://") <> 0 Then Response.Write "</a>"
   Response.Write "  </td>" & VbCrlf
   Response.Write "  <td>" & Records("ReferHits") & "</td>" & VbCrlf
   Response.Write " </tr>" & VbCrlf
   Records.MoveNext
  Loop

 '-- Close Objects --'
 Records.Close
 %>
 <tr>
  <td>Total</td>
  <td><%=Count%></td>
 </tr>
</table>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
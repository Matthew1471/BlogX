<%
' --------------------------------------------------------------------------
'¦Introduction : Not Found Admin Page.                                      ¦
'¦Purpose      : Provides a list of all the non-present linked pages.       ¦
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

PageTitle = "Pages Not Found"
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<table border="0">
 <%
 '-- Open The Records Ready To Write --'
 Records.Open "SELECT TOP 150 URL, ReferringPage, ErrorCount FROM NotFound ORDER BY ErrorCount DESC;", Database, 0, 1

  Count = 0

  Do Until (Records.EOF)
   URL = HTML2Text(Records("URL"))
   Count = Count + Records("ErrorCount")
   Response.Write " <tr>" & VbCrlf
   Response.Write "  <td style=""padding-bottom: 10px"">"

   '-- Show the URL --'
   If InStr(URL,"http://") <> 0 Then Response.Write "<a href=""" & StandardURL(Records("URL")) & """>"
   Response.Write URL
   If InStr(URL,"http://") <> 0 Then Response.Write "</a>"

   Response.Write "<br/>" & VbCrlf

   '-- Now Referrer --'
   URL = HTML2Text(Records("ReferringPage"))
   Response.Write "<span style=""font-size:xx-small"">(From:"
   If InStr(URL,"http://") <> 0 Then Response.Write "<a href=""" & StandardURL(Records("ReferringPage")) & """>"
   Response.Write URL
   If InStr(URL,"http://") <> 0 Then Response.Write "</a>"
   Response.Write ")</span>" & VbCrlf

   Response.Write "  </td>" & VbCrlf
   Response.Write "  <td>" & Records("ErrorCount") & "</td>" & VbCrlf
   Response.Write " </tr>" & VbCrlf
   Records.MoveNext
  Loop

 '-- Close Objects --'
 Records.Close
 %>
 <tr>
  <td colspan="2">Total</td>
  <td><%=Count%></td>
 </tr>
</table>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
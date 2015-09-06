<%
' --------------------------------------------------------------------------
'¦Introduction : Who Uses BlogX Page.                                       ¦
'¦Purpose      : Displays all known public verified BlogX blogs.            ¦
'¦Used By      : Links table.                                               ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
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

PageTitle = "Other BlogX Users"
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<div id="content">

 <div class="entry">
  <h3 class="entryTitle" style="text-align:center">Who Else Uses BlogX</h3>

  <p style="text-align:center">When you install BlogX, there is an option in the "Includes\Config.asp" file called "Register".<br/>
  <br/>By setting "Register" to "= True" your site appears on this page (this page should update every 2 weeks).<br/>
  In other words, you get <b>free traffic to your blog</b>.</p>

  <p style="text-align:center; color:red">An automated computer program (called a "spider") maintains this list.<br/>
  If you notice the User-Agent "Matthew1471 BlogX" visiting from this IP; that is the spider in action.</p>

  <p style="text-align:center; color:Green">I do not officially endorse any of these pages,<br/>
  the views and comments they contain are those of their respective holder(s) only.</p>

  <table border="0" style="align:center; margin: 0 auto">
   <%
   Records.Open "SELECT ReferURL, ReferHits FROM ScriptRefer WHERE Approved=True ORDER BY ReferHits DESC;", Database, 0, 1

   Count = 0

   Do Until (Records.EOF)
    Set ReferURL = Records("ReferURL")
    Count = Count + Records("ReferHits")

    If InStr(ReferURL,"http://") <> 0 Then 
     Response.Write "   <tr>" & VbCrlf
     Response.Write "    <td>" & "<a href=""" & ReferURL & """>" & ReferURL & "</a>" & "</td>" & VbCrlf
     Response.Write "    <td>" & Records("ReferHits") & "</td>" & VbCrlf
     Response.Write "   </tr>" & VbCrlf
    End If

    Records.MoveNext
   Loop

   Records.Close
   %>
   <tr>
    <td>Total</td>
    <td><%=Count%> Visits</td>
   </tr>
  </table>

  <div style="text-align: center; font-size: x-small">Current Spider Version : 2.0</div>
 </div>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
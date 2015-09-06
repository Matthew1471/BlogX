<%
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
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->

<div id="content">

 <table border="0">
 <%
 '-- Open The Records Ready To Write --'
 Records.Open "SELECT CommentID, CommentedDate, CommentedTime FROM Comments ORDER BY CommentID;", Database

 '-- Did we actually find any records? --'
 If Records.EOF Then

  Response.Write "<p align=""Center"">No Comments Found</P>" & VbCrlf & "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

 Else

  Do Until (Records.EOF) 

   Dim ConvertedDate
   ConvertedDate = Records("CommentedDate") & " " & Records("CommentedTime")
   %>
   <tr>
    <td style="background-color: #FF0000; color:#FFFFFF"><b><%=Records("CommentID")%></b></td>
   <%
   If IsDate(ConvertedDate) Then
    Records("CommentedDate") = ConvertedDate
    Response.Write "<td style=""background-color: #0000FF; color:#FFFFFF""><b>" & ConvertedDate & "</b></td>"
   Else
    Response.Write "<td style=""background-color: #FF0000; color:#FFFFFF""><b>Ignored - " & ConvertedDate & "</b></td>"
   End If
   
   Response.Write "</tr>"

   Records.MoveNext
  Loop
  
 End If

 '-- Close Objects --'
 Records.Close
 %>
 </table>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
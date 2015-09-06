<%
' --------------------------------------------------------------------------
'¦Introduction : Banned IP Address Page.                                    ¦
'¦Purpose      : Provides the blog administrator with an option to unban    ¦
'¦               previously banned IP addresses as well as view them.       ¦
'¦Used By      : Includes/Nav.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp,¦
'¦               Includes/Cache.asp.                                        ¦
'¦Notes        : This page is for unbanning IP addresses and listing them   ¦
'¦               but not always banning, that is often done by Comments.asp.¦
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
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->

<div id="content">

 <%
 '-- Did the user request to unban an IP address? --'
 Dim DelIP
 DelIP = Request.Querystring("Delete")
 DelIP = Replace(DelIP,"'","")
 If DelIP <> "" Then Database.Execute "DELETE FROM BannedIP WHERE IP='" & DelIP & "'"

 '-- We are welcome to ban people :D --'
 Dim BanIP
 BanIP = Request.Querystring("Ban")
 BanIP = Replace(BanIP,"'","")

 On Error Resume Next
  If BanIP <> "" Then Database.Execute "INSERT INTO BannedIP (IP) VALUES ('" & BanIP & "')" & ";"
 On Error Goto 0

 '-- Open The Records Ready To Write --'
 Records.Open "SELECT IP, BannedDate FROM BannedIP ORDER BY IP;", Database

 '-- Did we actually find any records? --'
 If Records.EOF Then

  Response.Write "<p style=""text-align:Center"">No banned users found.</p>" & VbCrlf
  Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

 Else

  Response.Write "<table border=""0"">" & VbCrlf

  Dim IP, BannedDate, BannedTime
  Set IP = Records("IP")
  Set BannedDate= Records("BannedDate")

  Do Until (Records.EOF) 
   Dim LatestDate
   If BannedDate > LatestDate Then LatestDate = BannedDate
   %>
   <tr>
    <td style="background-color: #FF0000; color:#FFFFFF"><b><%=IP%></b></td>
    <td style="background-color: #0000FF; color:#FFFFFF"><b><%=FormatDateTime(BannedDate,vblongdate) & " (" & FormatDateTime(BannedDate,vbLongTime) & ")" %></b></td>
    <td><acronym title="Unban User"><a href="?Delete=<%=IP%>"><img alt="Unban User" style="border:none" src="../Images/Key.gif"/></a></acronym></td>
   </tr>
   <%
   Records.MoveNext
  Loop
  
  Response.Write "</table>" & VbCrlf

  '-- Proxy Handler --'
  CacheHandle(LatestDate)

 End If

 '-- Close Objects --'
 Records.Close
 %>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
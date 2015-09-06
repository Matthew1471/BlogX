<%
' --------------------------------------------------------------------------
'¦Introduction : Check For Update Page.                                     ¦
'¦Purpose      : Check for BlogX engine updates.                            ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp,                            ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
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
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
 <p style="text-align:center">
 <%
  On Error Resume Next 
  Dim objXMLHTTP, VersionResponse

  '-- If you do not have MSXML3 installed you can revert to the old line:
  Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
  'Set objXMLHTTP=Server.CreateObject("MICROSOFT.XMLHTTP")

  objXMLHTTP.open "GET", "http://BlogX.co.uk/Download/UpdateBlogX.asp", true
   objXMLHTTP.setRequestHeader "Content-Type", "text/xml"
   objXMLHTTP.SetRequestHeader "User-Agent", "Matthew1471 BlogX"
  objXMLhttp.send()

  '-- Wait for up to 5 seconds (if we've not gotten the data yet) --'
  If objXMLHTTP.readyState <> 4 Then objXMLHTTP.waitForResponse 5

  '-- Abort the request --'
  If (objXMLhttp.readyState <> 4) Or (objXMLhttp.Status <> 200) Then objXMLhttp.Abort

  VersionResponse = ObjXMLHTTP.ResponseText

  '-- Write it out --'
  If (VersionResponse = Version) AND (Len(VersionResponse) > 0) Then 
   Response.Write "You currently have the <b>LATEST</b> BlogX engine (V" & VersionResponse & ")"
  ElseIf (Len(VersionResponse) > 0) AND IsNumeric(Replace(VersionResponse,".","")) = True Then
   Response.write "<a href=""http://freewebs.com/matthew1471/"">BlogX V" & VersionResponse & "</a> Is Now Available!"
  Else
   Response.Write "BlogX Update server is currently down, you may <a href=""http://freewebs.com/matthew1471/"">manually download the latest version</a>."
  End If

  Set objXMLHTTP = Nothing
 On Error Goto 0
 %>
  <br/><br/><a href="<%=SiteURL & PageName%>">Back</a>
 </p>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
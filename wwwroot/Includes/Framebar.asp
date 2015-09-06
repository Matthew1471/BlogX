<%@EnableSessionState=False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : File Upload Framebar.                                      ¦
'¦Purpose      : This is used to show the Persits ASP Upload progress bar.  ¦
'¦Used By      : Admin/AddFile.asp, Admin/AddPicture.asp                    ¦
'¦Requires     : Includes/Bar.asp, Includes/Note.htm                        ¦
'¦Standards    : XHTML Transitional & Frameset (Browser Defined).           ¦
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
Response.Expires = -1 %>
<% If Request("b") = "IE" Then %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
 <title>Uploading files</title>
 <style type="text/css">td {font-family:arial; font-size: 9pt }</style>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% Else %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title>Uploading files</title>
 <style type="text/css">td {font-family:arial; font-size: 9pt }</style>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<% End If %>
</head>

<% If Request("b") = "IE" Then %>
 <body bgcolor="#E0E0F6" text="midnightblue" link="darkblue" alink="red" vlink="red">
 <!-- Internet Explorer -->
  <iframe src="Bar.asp?PID=<%=Request("PID") & "&amp;to=" & Request("to")%>" title="Upload Progress" scrolling="no" frameborder="0" width="369" height="65"></iframe>
    <table border="0" width="100%" cellpadding="2" cellspacing="0">
     <tr>
      <td align="center">To cancel uploading, press your browser's <b>STOP</b> button.</td>
     </tr>
    </table>
 </body>
<%Else%>
 <!-- Netscape Navigator etc ... -->
 <frameset rows="60%, *" cols="100%">
  <frame src="Bar.asp?PID=<%= Request("PID") & "&amp;to=" & Request("to") %>" scrolling="no" frameborder="0" name="sp_body" noresize="noresize"/>
  <frame src="Note.htm" noresize="noresize" scrolling="no" frameborder="0" name="sp_note"/>
 </frameset>
<%End If%>
</html>

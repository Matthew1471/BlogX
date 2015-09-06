<%@EnableSessionState=False%>
<%
' --------------------------------------------------------------------------
'¦Introduction : File Upload Progress Bar.                                  ¦
'¦Purpose      : This is used to show the Persits ASP Upload progress bar.  ¦
'¦Used By      : Includes/Framebar.asp.                                     ¦
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
Response.Expires = -1

Dim PID, TimeO, UploadProgress 
PID = Request("PID")
TimeO = Request("to")

Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
 Dim BarContent
 BarContent = UploadProgress.FormatProgress(PID, TimeO, "#00007F", "%TUploading files...%t%B3%T%R left (at %S/sec) %r%U/%V(%P)%l%t")
Set UploadProgress = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en"> 
 <head>
<% If ("" = BarContent) Then %>
  <title>Upload Finished</title>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
  <script type="text/javascript">
  function CloseMe()
   {
    window.parent.close();
    return true;
   }
  </script>
 </head>
 <body onload="CloseMe()">
<% Else %>
  <title>Uploading Files...</title>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
  <meta http-equiv="Refresh" content="2;URL=<%=Request.ServerVariables("URL") & "?to=" & TimeO & "&amp;PID=" & PID %>"/>

  <style type="text/css">
   a:link{color:#191970}
   a:active{color:red}
   a:visited {color:red }
   td {font-family:arial; font-size: 9pt }
   td.spread {font-size: 6pt; line-height:6pt }
   td.brick {font-size:6pt; height:12px}
  </style>
 </head>
 <body style="background-color:#E0E0F6;color:#191970;margin-top:0">
  <%=BarContent%>
<% End If %>
 </body>
</html>
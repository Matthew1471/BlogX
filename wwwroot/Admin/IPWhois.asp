<%
' --------------------------------------------------------------------------
'¦Introduction : IP Information Page.                                       ¦
'¦Purpose      : Cross references an IP address with all logged comments.   ¦
'¦Used By      : Comments.asp.                                              ¦
'¦Requires     : Includes/Config.asp, Admin.asp.                            ¦
'¦Notes        : This page is useful for rapid comment deletion.            ¦
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - IPWhois</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

 <!--
 //= - - - - - - - 
 // Copyright 2004-09, Matthew Roberts
 // Copyright 2003, Chris Anderson
 // 
 // Usage Of This Software Is Subject To The Terms Of The License
 //= - - - - - - -
 -->
<%
Dim RequestedIP, Returned
RequestedIP = Replace(Request.Querystring("IP"),"'","")

Records.Open "SELECT EntryID, IP, Name, Email, Content FROM Comments WHERE IP='"& RequestedIP & "' OR Content LIKE '%" & RequestedIP & "%' ORDER BY Name ASC",Database, 1, 3

If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
 <link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet"/>
</head>
<body style="background-color: <%=BackgroundColor%>; text-align:center">
<p>
 <b>Registered Names for IP (<%=RequestedIP%>)</b>
</p>
 <%
 If NOT Records.EOF Then 

  Dim EntryID, Email, IP
  Set EntryID = Records("EntryID")
  Set Email = Records("Email")
  Set Name = Records("Name")
  Set IP = Records("IP")

  Do Until Records.EOF

  Response.Write "<p>" & VbCrlf

  '-- Are we deleting them? --'
  If Request.Querystring("DeleteAllByIP") = "True" Then

   Dim EntryRecords
   Set EntryRecords = Server.CreateObject("ADODB.recordset")
    EntryRecords.Open "SELECT RecordID, Comments, LastModified FROM Data WHERE RecordID=" & EntryID,Database,1,3
    EntryRecords("Comments") = EntryRecords("Comments") - 1
    EntryRecords("LastModified") = Now()
    EntryRecords.Update
    EntryRecords.Close
   Set EntryRecords = Nothing

   Name = Name
   EntryID = EntryID

   Records.Delete
   Records.MoveNext

   Response.Write "Deleted " & Name & " from Entry " & EntryID & "<br/>" & VbCrlf

 Else

  Dim LastName

  If LCase(Name) <> LastName Then

  Response.Write "<a href=""IPWhois.asp?IP=" & RequestedIP & "&amp;DeleteAllByIP=True""><img alt=""Delete all by this IP address"" src=""../Images/Key.gif"" style=""border: none""/></a>" & VbCrlf

  Response.Write "<a href=""" & SiteURL  & "Comments.asp?Entry=" & EntryID & """><img alt=""View this comment"" src=""../Images/Emoticons/Profile.gif"" style=""border: none""/></a> " & Name & " (" & IP 
   If Email <> "" Then Response.Write " / " & Email
  Response.Write ")<br/>" & VbCrlf

On Error Resume Next

  Dim DNSLook
  Set DNSLook = Server.CreateObject("AspDNS.Lookup")

  If (Err.Number <> 0) Then 
   If Err.Number = -2147221005 Then
    Response.Write "(<b><span style=""color:red"">ASPDNS not installed..<br/>No reverse lookup performed</span></b>)" & VbCrlf
   Else
    Response.Write "(<b><span style=""color:red"">An Error was encountered while trying to create a refrence to ASPDNS :<br/>" & Err.Number & " - " & Err.Description & "</span></b>)" & VbCrlf
   End If
  Else
   Response.Write "(<b><span style=""color:red"">" & DNSLook.ReverseDNSLookup(IP) & "</span></b>)" & VbCrlf
  End If

  Set DNSLook = Nothing

On Error GoTo 0

  Response.Write "</p>" & VbCrlf

  LastName = LCase(Name)
  End If

  Response.Flush
  Records.MoveNext

 End If

 Loop

Else
 Response.Write "<p align=""center"">IP Address Not Found</p>" & VbCrlf
End If 

 Response.write "<p><a href=""Javascript:window.close()"">Close Window</a></p>" & VbCrlf

Records.Close
%>
</body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing
%>
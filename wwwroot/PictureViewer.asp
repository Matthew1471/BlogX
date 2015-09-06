<% Option Explicit 
Dim PHPEnabled

'-- This Host ALSO supports PHP --'
PHPEnabled = True

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<html>
<head>
<title><%=SiteDescription%> - Picture Viewer</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")

        Dim Requested	'Holds the name of the requested file
        Requested = Request("Folder")
        Requested = Replace(Requested, ".", "", 1, -1, vbTextCompare)
        Requested = Replace(Requested, "\", "", 1, -1, vbTextCompare)
        Requested = Replace(Requested, "/", "", 1, -1, vbTextCompare)
%>
<link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet">
<script language="JavaScript" type="text/javascript">
<!--
function MyPicturesWindow(url) {
 window.open(url,'Pictures','height=510,width=680,resize=yes,status=yes,toolbar=no,menubar=no,location=no,directories=no,scrollbars=yes,left=300,top=200');
}	
//-->
</script>
</head>
<body bgcolor="<%=BackgroundColor%>">
<%
Dim DIR, EXT, Something

On Error Resume Next
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set DIR = FSO.GetFolder(Server.MapPath("Images\Articles\") & "\" & Requested & "\")
%>

<br>

<TABLE width="95%" cellpadding="5" align="center" border=1 cellspacing=0><TR>
<% For Each File in Dir.Files 

  Ext = UCase(Right(File.Name, 3)) 
  If Ext = "JPG" or Ext = "GIF" or Ext = "PNG" or Ext = "BMP" Then

  Dim Thumbnails
  If FSO.FileExists(Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails\tn" & File.Name) Then 
   Thumbnails = Requested & "/Thumbnails/tn" & File.Name
  ElseIf PHPEnabled = True Then
   '--Lets Generate A Thumbnail --'
   FSO.CreateFolder Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails"
   Thumbnails = "Thumbnail.php?f=" & Requested & "/" & Server.URLEncode(File.Name)
  Else
   Thumbnails = Requested & "/" & Server.URLEncode(File.Name)
  End If

   Something = True
   Count = Count + 1

   If Count > 7 Then 
    Response.Write VbCrlf & "</TR><TR>"
    Count = 1
   End If
%>
   <TD valign="top"><img width="50" height="50" src="Images/Articles/<%=Thumbnails%>" alt="<%=File.Name%>"><br>
   <A HREF="#" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'") & "/" & Replace(File.Name,"'","\'")%>')"><%=File.Name%></A><br>
   <%=FormatNumber(File.Size/1000)%> Kb<br>
   <input type="button" value="View" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'") & "/" & Replace(File.Name,"'","\'")%>')">
</TD>
<%  
   End If

Next

On Error GoTo 0

Do While Count > 0 and count < 7
 Count = Count + 1
 Response.Write "<TD>--Empty--</TD>"
Loop
%>
</tr>
</table>

<% If Something = False then Response.Write "<center><b>No Pictures Available To View<BR></b>" & "(" & Server.MapPath("Images\Articles\") & "\" & Requested & "\" & "*.*" & ")</center>"%>

</Body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing

Set FSO = Nothing
Set DIR    = Nothing
%>
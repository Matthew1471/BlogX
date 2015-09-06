<% Option Explicit 
Dim PHPEnabled

'-- This Host ALSO supports PHP --'
PHPEnabled = True
%>
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
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
<script language="JavaScript">
<!--
function MyPicturesWindow(url) {

window.open(url,'Pictures','height=510,width=680,resize=yes,status=yes,toolbar=no,menubar=no,location=no,directories=no,scrollbars=yes,left=300,top=200');
}	
//-->
</script>
</head>
<body bgcolor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<Center>
<%
Dim Upload, DIR, EXT, Something

On Error Resume Next
Set Upload = Server.CreateObject("Persits.Upload.1")
Set Dir = Upload.Directory(Server.MapPath("Images\Articles\") & "\" & Requested & "\" & "*.*", Request("sortby"))
'On Error GoTo 0
%>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td valign=top>
<center>

<TABLE BORDER=1 CELLSPACING=0><TR>
<% For Each File in Dir %>

<% If Not File.IsSubdirectory Then
Ext = UCase(Right(File.FileName, 3)) 
If Ext = "JPG" or Ext = "GIF" or Ext = "PNG" or Ext = "BMP" Then

  Dim Thumbnails
  If Upload.FileExists(Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails\tn" & File.FileName) Then 
   Thumbnails = Requested & "/Thumbnails/tn" & File.FileName
  ElseIf PHPEnabled = True Then
   '--Lets Generate A Thumbnail --'
   Upload.CreateDirectory Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails", True
   Thumbnails = "Thumbnail.php?f=" & Requested & "/" & Server.URLEncode(File.FileName)
  Else
   Thumbnails = Requested & "/" & Server.URLEncode(File.FileName)
  End If

Something = True
Count = Count + 1
Do While Count > 7
Response.Write "<TR>"
Count = 1
Loop
%>
   <TD><img width="50" height="50" src="Images/Articles/<%=Thumbnails%>" alt="<%=File.FileName%>"><br>
   <A HREF="#" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'") & "/" & Replace(File.FileName,"'","\'")%>')"><%=File.FileName%></A><br>
   <% = formatnumber(File.Size/1000)%> Kb<br>
   <input type="button" value="View" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'") & "/" & Replace(File.FileName,"'","\'")%>')">
</TD>
<%  
End If
End If
Next
Do While Count > 0 and count < 7
Count = Count + 1
Response.Write "<TD>--Empty--</TD>"
Loop
%>
</table>
</table>
</table>
</Form>
<% If Something = False then Response.Write "<b>No Pictures Available To View<BR></b>" & "(" & Server.MapPath("Images\Articles\") & "\" & Requested & "\" & "*.*" & ")"
%>

</Center>
</Body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing

Set Upload = Nothing
Set DIR    = Nothing
%>
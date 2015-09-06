<%Option Explicit
PhotoMode = True

'-- This Host ALSO supports PHP --'
Dim PHPEnabled
PHPEnabled = True

PageTitle = "Photo Album"
%><!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<div id="content">
<%
 Dim Requested	'Holds the name of the requested file
 Requested = Request("Folder")
 Requested = Replace(Requested, ".", "", 1, -1, vbTextCompare)
 Requested = Replace(Requested, "\", "", 1, -1, vbTextCompare)
 If Left(Requested,1) = "/" Then Requested = Replace(Requested, "/", "", 1, -1, vbTextCompare)

Dim DIR, EXT, Something

On Error Resume Next
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")
On Error GoTo 0

If (Requested <> "" AND FSO.FolderExists(Server.MapPath("Images\Articles\") & "\" & Requested & "\")) OR (Requested = "" AND FSO.FolderExists(Server.MapPath("Images\Articles\") & "\")) Then

Set DIR = FSO.GetFolder(Server.MapPath("Images\Articles\") & "\" & Requested & "\")
%>
<script type="text/javascript">
<!-- Hide javascript so W3C doesn't choke on it
function MyPicturesWindow(url) {

var TheNewWin = window.open('','Pictures','height=' + (screen.height/2) +',width=' + (screen.width/2) + ',resizable=yes,status=yes,toolbar=no,menubar=no,location=no,directories=no,scrollbars=auto,left=' + (screen.height/4)+',top='+ (screen.height/4));

TheNewWin.document.open;
TheNewWin.document.write('<html>\r');
TheNewWin.document.write('<head>\r');
TheNewWin.document.write('<title><%=Replace(SiteDescription,"'","\'")%> - Image Display<\/title>\r');
TheNewWin.document.write('<META HTTP-EQUIV="Content-Type" CONTENT="text\/html; charset=utf-8">\r');

TheNewWin.document.write('<!-' + '-\r');
TheNewWin.document.write('\/\/= - - - - - - - \r');
TheNewWin.document.write('\/\/ Copyright 2006, Matthew Roberts\r');
TheNewWin.document.write('\/\/ Copyright 2003, Chris Anderson\r');
TheNewWin.document.write('\/\/ \r');
TheNewWin.document.write('\/\/ Usage Of This Software Is Subject To The Terms Of The License\r');
TheNewWin.document.write('\/\/= - - - - - - -\r');
TheNewWin.document.write('-' + '->\r');

TheNewWin.document.write('<\/head>\r');
TheNewWin.document.write('<body bgcolor="<%=BackgroundColor%>" onload="fitPic();">\r');
TheNewWin.document.write('<!DOCTYPE HTML PUBLIC "-\/\/W3C\/\/DTD HTML 4.0 Transitional\/\/EN">\r');

<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")
If TemplateURL = "" Then
 Response.Write "TheNewWin.document.write('<link href=""" & SiteURL & "Templates\/" & Template & "\/Blogx.css"" type=text\/css rel=stylesheet>\r\r');"
Else
 Response.Write "TheNewWin.document.write('<link href=""" & TemplateURL & Template & "\/Blogx.css"" type=text/css rel=stylesheet>\r\r');"
End If %>

TheNewWin.document.write('<sc' + 'ript language=\'javascript\'>\r'); 

TheNewWin.document.write(' function fitPic() {\r\r');
TheNewWin.document.write(' var winW = 630, winH = 460;\r');
TheNewWin.document.write(' var NS = (navigator.appName=="Netscape")?true:false;\r\r');

TheNewWin.document.write('  iWidth = (NS)?window.innerWidth:document.body.clientWidth;\r');
TheNewWin.document.write('  iHeight = (NS)?window.innerHeight:document.body.clientHeight;\r');
TheNewWin.document.write('  iWidth = (document.images[0].width - iWidth) + 50;\r');
TheNewWin.document.write('  iHeight = (document.images[0].height - iHeight) + 80;\r');

// If we have a large image, do not bother resizing
TheNewWin.document.write('  if ((screen.width > document.images[0].width) & screen.height > document.images[0].height) { window.resizeBy(iWidth, iHeight); } else { window.location = \'' + url.replace("'","\\'") + '\';}\r\r');

TheNewWin.document.write(' if (parseInt(navigator.appVersion)>3) {\r');
TheNewWin.document.write('  if (NS) {\r');
TheNewWin.document.write('   winW = window.innerWidth;\r');
TheNewWin.document.write('   winH = window.innerHeight;\r');
TheNewWin.document.write('  }\r');
TheNewWin.document.write('  if (navigator.appName.indexOf("Microsoft")!=-1) {\r');
TheNewWin.document.write('   winW = document.body.offsetWidth;\r');
TheNewWin.document.write('   winH = document.body.offsetHeight;\r');
TheNewWin.document.write('  }\r');
TheNewWin.document.write(' }\r\r');

TheNewWin.document.write('  window.moveTo((screen.width/2)-(winW/2),(screen.height/2)-(winH/2));');

TheNewWin.document.write('  self.focus();\r');
TheNewWin.document.write(' };\r');
TheNewWin.document.write('<\/sc' + 'ript>\r\r'); 

var arrTemp=url.split("/"); 
var fileName = (arrTemp.length>0)?arrTemp[arrTemp.length-1]:url;

TheNewWin.document.write('<center>\r');
TheNewWin.document.write(' <span style="color: red">File : ' + fileName + '<\/span><br/>\r');
TheNewWin.document.write(' <hr\/><img name="actualImage" src="' + url.replace("'","\'") + '" border=""0""><hr\/>\r')
TheNewWin.document.write(' <a href="#" onclick="self.close();return false;">Close Window<\/a>\r');
TheNewWin.document.write('<\/center>\r\r');

TheNewWin.document.write('<\/body>\r');
TheNewWin.document.write('<\/html>');
TheNewWin.document.close();
}	
//-->
</script>

<%
If (Requested <> "") Then
 Dim LastFolder
 If InstrRev(Requested,"/") <> 0 Then LastFolder = Left(Requested, InstrRev(Requested,"/") - 1)

    Response.Write "<p>Browsing /<span style=""color: red"">" & Requested & "</span>/</p>" & VbCrlf

    Response.Write "<table width=""95%"" cellpadding=""5"" style=""align: center"" border=""1"" cellspacing=""0"">" & VbCrlf
    Response.Write " <tr>" & VbCrlf
    Response.Write "  <td valign=""top"">" & VbCrlf
     If LastFolder <> "" Then Response.Write "    <a href=""?Folder=" & LastFolder & """>" Else Response.Write "<a href=""PhotoAlbum.asp"">"
    Response.Write "   <img style=""border-style: none"" width=""50"" height=""50"" src=""Images/Editor/SpellCheck.gif"" alt=""Go Back To The "
     If LastFolder <> "" Then Response.Write LastFolder Else Response.Write "Main"
    Response.Write " Folder""/><br/>" & VbCrlf
    Response.Write "Back To The "
     If LastFolder <> "" Then Response.Write LastFolder Else Response.Write "Main" 
    Response.Write " Folder</a><br/>" & VbCrlf
    Response.Write "</td>" & VbCrlf

   Count = Count + 1
 Else
   Response.Write "<table width=""95%"" cellpadding=""5"" border=""1"" cellspacing=""0"" style=""align: center; margin: 0 auto;"">" & VbCrlf
   Response.Write "<tr>" & VbCrlf
 End If

Dim SubFolder
For Each SubFolder in Dir.SubFolders
 If (SubFolder.Name <> "Thumbnails") AND (SubFolder.Name <> "Temp") Then
%>
    <td valign="top"><a href="?Folder=<%=Requested%><% If Requested <> "" Then Response.Write "/" %><%=SubFolder.Name%>">
     <img border="0" width="50" height="50" src="Images/Editor/Image.gif" alt="<%=SubFolder.Name%> Folder"><br/>
     <%=SubFolder.Name%></a><br/>
     <%=FormatNumber(SubFolder.Size/1000)%> Kb<br/>
    </td>
<%
   Count = Count + 1

   If Count => 4 Then 
    Response.Write VbCrlf & "</tr>" & VbCrlf
    Response.Write "<tr>" & VbCrlf
    Count = 0
   End If
 End If
Next

For Each File in Dir.Files 

  Ext = UCase(Right(File.Name, 3)) 
  If Ext = "JPG" or Ext = "GIF" or Ext = "PNG" or Ext = "BMP" Then

  Dim Thumbnails
  If FSO.FileExists(Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails\tn" & File.Name) Then 
   Thumbnails = Server.URLEncode(Requested & "/Thumbnails/tn" & File.Name)
   Thumbnails = Replace(Thumbnails,"+","%20")
  ElseIf PHPEnabled = True Then
   '--Lets Generate A Thumbnail --'
   On Error Resume Next
    FSO.CreateFolder Server.MapPath("Images\Articles\") & "\" & Requested & "\Thumbnails"
   On Error GoTo 0

   Thumbnails = "Thumbnail.php?f=" 
   If Requested <> "" Then Thumbnails = ThumbNails & Server.URLEncode(Requested) & "/" 
   Thumbnails = Thumbnails & Server.URLEncode(File.Name)
  Else
   Thumbnails = ""
   If Requested <> "" Then Thumbnails = Server.URLEncode(Requested) & "/"
   Thumbnails = Thumbnails & Server.URLEncode(File.Name)
  End If

   Something = True
   Count = Count + 1

   If Count > 4 Then 
    Response.Write VbCrlf & "</tr><tr>"
    Count = 1
   End If
%>
   <td valign="top"><a href="Images/Articles/<%=Server.URLEncode(Requested)%><%If Requested <> "" Then Response.Write "/" %><%=Server.URLEncode(File.Name)%>" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'")%><%If Requested <> "" Then Response.Write "/" %><%=Replace(Replace(File.Name,"&","&amp;"),"'","\'")%>'); return false;">
   <img style="border-style: none;" width="50" height="50" src="Images/Articles/<%=Thumbnails%>" alt="<%=Replace(File.Name,"&","&amp;")%>"/><br/>
   <%=Replace(File.Name,"&","&amp;")%></a><br/>
   <%=FormatNumber(File.Size/1000)%> Kb<br/>
   <input type="button" value="View" onclick="javascript:MyPicturesWindow('Images/Articles/<%=Replace(Requested,"'","\'")%><%If Requested <> "" Then Response.Write "/" %><%=Replace(Replace(File.Name,"&","&amp;"),"'","\'")%>')"/>
</td>
<%  
   End If

Next

On Error GoTo 0

Do While Count > 0 and count < 4
 Count = Count + 1
 Response.Write "<td>--Empty--</td>"
Loop

If Something = False Then 
 Response.Write "<td style=""text-align: center""><p><b>No Pictures Available To View<br/></b>" & "(" & Server.MapPath("Images\Articles\") & "\" 
 If Len(Requested) > 0 Then Response.Write Requested & "\" 
 Response.Write "*.*" & ")</p></td>" & VbCrlf
End If
%>
</tr>
</table>

<%End If%>

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
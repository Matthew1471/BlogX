<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Documentation</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
      <center>Useful information to Bloggers, Blog Viewers and the Administrative powers within:</center>

<%
        Dim FSO, Folder, Folders, Files

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(AppPath & "Includes\Documentation")
        Set Folders = Folder.SubFolders

        For Each Folder in Folders
        Response.Write "    <H3>" & Folder.Name & "</h3>" & VbCrlf

        Set SubFolder = FSO.GetFolder(AppPath & "Includes\Documentation\" & Folder.Name)
        Set Files = SubFolder.Files

        For Each File in Files
        Response.Write "    <img src=""Includes/Images/eBlog.gif"">" & VbCrlf
        Response.Write "    <a href=""Includes/Documentation/" & Folder.Name & "/" & File.Name & """>" & File.Name & "</a><br>" & VbCrlf
        Next

        Next

        Set FSO = Nothing
	Set Folder = Nothing
	Set Folders = Nothing
	Set SubFolder = Nothing
	Set Files = Nothing
%>
	<br>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
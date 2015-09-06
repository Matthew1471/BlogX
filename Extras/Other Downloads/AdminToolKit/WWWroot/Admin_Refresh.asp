<!-- #INCLUDE FILE="Includes/Admin.asp" -->
<% Server.ScriptTimeout = 6000 %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Refresh Blogs</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">

<h2 align="center">Welcome To The BlogX Administrative Toolkit</h2>

<p align="center">Processed!</p>
<p align="center">
<% 

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

'FSO.CreateFolder(AppPath & Path)

Dim DatabaseUpdate, ConfigFile, Path
Dim Folder, Folders, Count, UpdatedBlogs

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(AppPath)
        Set Folders = Folder.SubFolders

        For Each Folder in Folders

	Path = Folder.Name

		If (Folder.Name <> "Includes") AND (FSO.FileExists(AppPath & Path & "\Default.bak") <> True) Then 

		If NOT FSO.FileExists(DatabasePath & Path & ".mdb") Then
		DatabaseUpdate = True
		FSO.CopyFile AppPath & "Includes\DataSource\BlogX.mdb", DatabasePath & Path & ".mdb"
		Else
		DatabaseUpdate = False
		End If

		FSO.CopyFolder AppPath & "\Includes\Data", AppPath & Path

		If FSO.FileExists(AppPath & Path & "\Default.bak") Then FSO.DeleteFile(AppPath & Path & "\Default.bak")
		If FSO.FileExists(AppPath & Path & "\Includes\Config.bak") Then FSO.DeleteFile(AppPath & Path & "\Includes\Config.bak")

		Set ConfigFile = FSO.CreateTextFile(AppPath & Path & "\Includes\Datafile.asp", True)
		ConfigFile.WriteLine(Chr(60) & "% DataFile = """ & DatabasePath & Path & ".mdb""")
		ConfigFile.WriteLine("SiteURL = ""http://" & Request.ServerVariables("SERVER_NAME") & Root & Path & "/"" %" & Chr(62)) & VbCrlf
		ConfigFile.Close
		Set ConfigFile = Nothing

		Response.Write Path & " was <b>UPDATED</b><br>" & VbCrlf
		Response.Flush

		ElseIf (Folder.Name <> "Includes") Then
		Response.Write Path & " is <b><font color=""red"">DISABLED</font></b><br>" & VbCrlf

		End If

	Next

        Set FSO = Nothing
	Set Folder = Nothing
	Set Folders = Nothing

%>
</p>

<p align="Center"><Input Type="button" Value="Finish->" onClick="document.location.href='Admin_Main.asp'"></p>

<p align="center"><font color="#FF0000">Note : </font>
This will write to <font color="Red"><%=Root%></font> & <font color="Red"><%=DatabasePath%></font> make sure IIS can write to that folder (and that it exists)</p>

      </td>
      <!--- End Of Content -->
<% WriteFooter %>
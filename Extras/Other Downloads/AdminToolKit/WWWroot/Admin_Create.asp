<!-- #INCLUDE FILE="Includes/Admin.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<% Dim CreateBox

If DropDownBoxMode <> True Then
CreateBox = "<Input Name=""Path"" Type=""Text"">"
Else

'<option value="">-- New --</Option>

        Dim Folder, Folders, Count

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(AppPath)
        Set Folders = Folder.SubFolders

	CreateBox = "<select name=""Path"">"

        For Each Folder in Folders
        If Folder.Name <> "Includes" Then CreateBox = CreateBox & "<option>" & Folder.Name & "</Option>" & VbCrlf
        Next

	CreateBox = CreateBox & "</select>" & VbCrlf

        Set FSO = Nothing
	Set Folder = Nothing
	Set Folders = Nothing

End If
%>
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Create/Enable Blog</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">

<h2 align="center">Welcome To The BlogX Administrative Toolkit</h2>

<% If Request.Form("Path") = "" Then %>
<p align="center">
Please Type In The Desired Folder Name :<br>
<Form Name="NewBlog" Method="Post">
http://<%=Request.ServerVariables("SERVER_NAME") & Root & CreateBox%>/</P>

<p align="Center">
<Input Name="Submit" Type="Button" Value="<-Back" onclick="javascript:history.back()">
<Input Name="Submit" Type="Submit" Value="Next->">
</Form>
</p>
<% Else
Dim Path

Path = Ucase(Request.Form("Path"))
Path = Replace(Path, "..", "")
Path = Replace(Path, "%2E", "")
Path = Replace(Path, "INCLUDES", "")

Path = LCase(Path)
Path = UCase(Left(Path,1)) & Right(Path,Len(Path)-1)

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

'FSO.CreateFolder(AppPath & Path)

Dim DatabaseUpdate, ConfigFile

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

Set FSO = Nothing
%>
<p align="center">Processed!</p>
<p align="center"><% If DatabaseUpdate = False Then Response.Write "Blog was <b>UPDATED</b>" Else Response.Write "New Blog Created"%></p>


<p align="center"><a href="http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/">http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/</a></p>

<p align="Center"><Input Type="button" Value="Finish->" onClick="document.location.href='Admin_Main.asp'"></p>
<% End If %>

<p align="center"><font color="#FF0000">Note : </font>
This will write to <font color="Red"><%=Root%></font> & <font color="Red"><%=DatabasePath%></font> make sure IIS can write to that folder (and that it exists)</p>

      </td>
      <!--- End Of Content -->
<% WriteFooter %>
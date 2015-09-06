<!-- #INCLUDE FILE="Includes/Admin.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Disable Blog</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
<h2 align="center">Welcome To The BlogX Administrative Toolkit</h2>

<% If Request.Form("Path") = "" Then %>
<p align="center">
Please Type In The Desired Disabled Blog :<br>
<Form Name="NewBlog" Method="Post">
http://<%=Request.ServerVariables("SERVER_NAME") & Root%><Input Name="Path" Type="Text">/</P>

<p align="Center">
<Input Name="Submit" Type="Button" Value="<-Back" onclick="javascript:history.back()">
<Input Name="Submit" Type="Submit" Value="Next->">
</Form>
</p>
<% Else
Dim Path
Path = Request.Form("Path")
Path = Replace(Path, "..", "")
Path = Replace(Path, "%2E", "")

Dim FSO, Message
Set FSO = CreateObject("Scripting.FileSystemObject")

If (FSO.FolderExists(AppPath & Path)) AND NOT (FSO.FileExists(AppPath & Path & "\Default.bak")) then

  FSO.MoveFile AppPath & Path & "\Default.asp", AppPath & Path & "\Default.bak"

  '-- Lets stuff up the Config file so there is no back door entry --'
  FSO.MoveFile AppPath & Path & "\Includes\Config.asp", AppPath & Path & "\Includes\Config.bak"

	Dim ConfigFile
	Set ConfigFile = FSO.CreateTextFile(AppPath & Path & "\Includes\Config.asp", True)
	ConfigFile.WriteLine(Chr(60) & "% Response.Write ""Disabled""")
	ConfigFile.WriteLine("Response.End %" & Chr(62)) & VbCrlf
	ConfigFile.Close
	Set ConfigFile = Nothing

  FSO.CopyFile AppPath & "Includes\DataSource\Disabled.asp", AppPath & Path & "\Default.asp"

  Message = "Processed!"

ElseIf FSO.FileExists(AppPath & Path & "\Default.bak") Then
  Message = "Blog Already Disabled!"
Else
  Message = "Blog Does NOT Exist!"
End If

Set FSO = Nothing
%>
<p align="center"><%=Message%></p>

<p align="center"><a href="http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/">http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/</a></p>

<p align="Center"><Input Type="button" Value="Finish->" onClick="document.location.href='Admin_Main.asp'"></p>

<% End If %>

<p align="center"><font color="#FF0000">Note : </font>
This will write to <font color="Red"><%=Root%></font> & <font color="Red"><%=AppPath%></font> make sure IIS can write to that folder (and that it exists)</p>
</td>
<% WriteFooter %>
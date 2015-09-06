<% Option Explicit %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Sign Up For A Free Blog</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
<% 
If EnableGuestSignups <> 0 Then
If Request.Form("Path") = "" Then %>

<Form Name="NewBlog" Method="Post">
<OL>
<LI><p align="center"><font color="red">Please Type In The Desired Folder Name :</font><br><br>
http://<%=Request.ServerVariables("SERVER_NAME") & Root %><Input Name="Path" Type="Text">/
</p></LI>

<LI><p align="center"><font color="red">Please Type In Your E-mail Address<br><small>(Used <b>ONLY</b> To Notify You About Your Account)</small> :</font><br><br>
E-mail Address : <Input Name="Email" Type="Text"></P></LI>
</OL>

<p align="Center">
<Input Name="Submit" Type="Button" Value="<-Back" onclick="javascript:history.back()">
<Input Name="Submit" Type="Submit" Value="Next->">
</Form>
</p>
<% Else

         '---- First CHECK IP -----'
	 Dim Database, Records, AlreadyIP

         '### Create a connection odject ###
         Set Database = Server.CreateObject("ADODB.Connection")

         '### Set an active connection to the Connection object ###
         Database.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

         '### Create a recordset object ###
         Set Records = Server.CreateObject("ADODB.Recordset")

	 '### Open The Records Ready To Write ###
	 Records.CursorType = 2
	 Records.LockType = 3
	 Records.Open "SELECT * FROM IP WHERE IP='" & Request.ServerVariables("REMOTE_ADDR") & "';", Database

	 If Not Records.EOF = True Then AlreadyIP = True

	 Records.Close 

Dim Path

Path = Request.Form("Path")
Path = Replace(Path, "..", "", 1, -1, 1)
Path = Replace(Path, "%2E","", 1, -1, 1)
Path = Replace(Path, "INCLUDES", "", 1, -1, 1)

'--- Create Or Error ---
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim AlreadyExists, ConfigFile, NotValidEmail

If (Instr(Request.Form("Email"),"@") <> 0) AND NOT (FSO.FolderExists(AppPath & Path)) AND (AlreadyIP = False) Then 


	 Records.Open "SELECT * FROM IP WHERE IP='" & Request.ServerVariables("REMOTE_ADDR") & "';", Database
	 If Not Records.EOF = False Then
	 Records.AddNew
	 Records("IP") = Request.ServerVariables("REMOTE_ADDR")
	 Records.Update
	 End If
	 Records.Close 

	 Records.Open "SELECT UserName, EmailAddress FROM Alerts;", Database
	 Records.AddNew
	 Records("UserName")        = Path
	 Records("EmailAddress")    = Request.Form("Email")
	 Records.Update
	 Records.Close

FSO.CopyFolder AppPath & "\Includes\Data", AppPath & Path

Set ConfigFile = FSO.CreateTextFile(AppPath & Path & "\Includes\Datafile.asp", True)
ConfigFile.WriteLine(Chr(60) & "% DataFile = """ & DatabasePath & Path & ".mdb""")
ConfigFile.WriteLine("SiteURL = ""http://" & Request.ServerVariables("SERVER_NAME") & Root & Path & "/"" %" & Chr(62)) & VbCrlf
ConfigFile.Close

FSO.CopyFile AppPath & "Includes\DataSource\BlogX.mdb", DatabasePath & Path & ".mdb"

Set FSO = Nothing
Set ConfigFile = Nothing

ElseIf FSO.FolderExists(AppPath & Path) Then
AlreadyExists = True
Else
NotValidEmail = True
End If 
'--- Create Or Error --- 
%>
<p align="center">
<% If AlreadyIP = True Then 
Response.Write "Sorry, But you already have ONE blog!" 
ElseIf AlreadyExists = True Then 
Response.Write "Sorry, But That Username Already Exists"
ElseIf NotValidEmail = True Then
Response.Write "You specified an invalid e-mail address"
Else
Response.Write "New Blog Created"
End If%></p>

<% If (AlreadyExists = False) AND (AlreadyIP = False) AND (NotValidEmail = False) Then %>
<p align="center"><b>Username :</b> admin<br>
<b>Password :</b> letmein</p>
<% End If %>

<% If (AlreadyExists = False) AND (NotValidEmail = False) Then %>
<p align="center"><a href="http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/">http://<%=Request.ServerVariables("SERVER_NAME") & Root & Path%>/</a></p>
<% End If

	 Database.Close

	 '#### Close Objects ###	
	 Set Database = Nothing
	 Set Records = Nothing

%>

<p align="Center"><Input Type="button" Value="Finish->" onClick="document.location.href='Default.asp'"></p> 
<%
End If
Else %>
      <center>Not Allowed</center>
      </td>
      <!--- End Of Content -->
<% End If %>
</td>
<%WriteFooter%>
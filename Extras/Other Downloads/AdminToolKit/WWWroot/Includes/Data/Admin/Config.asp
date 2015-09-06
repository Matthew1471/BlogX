<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>

<script language="JavaScript">
function Colors(url) { popupWin = window.open('SelectColor.asp?' + url,'new_page','width=400,height=450,scrollbars=yes') }
</script>

<Form Name="Form1" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P class="config"><span id="Label1">SiteName<font color="Red">*</Font> : </span><input Name="SiteName" type="text" style="width:90%;" Value="<%If SiteName <> "" Then Response.Write Replace(SiteName,"""","&quot;")%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">CookieName<font color="Red">*</Font> : </span><input Name="CookieName" type="text" style="width:20%;" Value="<%If CookieName <> "" Then Response.Write CookieName%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">Copyright<font color="Red">*</Font> : </span><input Name="Copyright" type="text" style="width:30%;" Value="<%If Copyright <> "" Then Response.Write Replace(Copyright,"""","&quot;")%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">Description : </span><input Name="SiteDescription" type="text" style="width:30%;" Value="<%If SiteDescription <> "" Then Response.Write Replace(SiteDescription,"""","&quot;")%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">Comments<font color="Red">*</Font> : </span><input Name="EnableComments" type="Checkbox" Value="True" onChange="return setVarChange()" <%If EnableComments = True Then Response.Write "CHECKED"%>></P>
            <P class="config"><span id="Label1">EntriesPerPage<font color="Red">*</Font> : </span><input Name="EntriesPerPage" type="text" style="width:5%;" Value="<%=EntriesPerPage%>"></P>
            <P class="config"><span id="Label1">ReaderPassword : </span><input Name="ReaderPassword" type="text" style="width:40%;" Value="<%If ReaderPassword <> "" Then Response.Write Replace(ReaderPassword,"""","&quot;")%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">SiteSubTitle : </span><input Name="SiteSubTitle" type="text" style="width:40%;" Value="<%If SiteSubTitle <> "" Then Response.Write Replace(SiteSubTitle,"""","&quot;")%>" onChange="return setVarChange()"></P>
            <P class="config"><span id="Label1">Polls<font color="Red">*</Font> : </span><input Name="Polls" type="Checkbox" Value="True" onChange="return setVarChange()" <%If Polls = True Then Response.Write "CHECKED"%>></P>
            <P class="config"><span id="Label1">ShowCategories<font color="Red">*</Font> : </span><input Name="ShowCategories" type="Checkbox" Value="True" onChange="return setVarChange()" <%If ShowCat = True Then Response.Write "CHECKED"%>></P>
            <P class="config"><span id="Label1">SortByDay<font color="Red">*</Font> : </span><input Name="SortByDay" type="Checkbox" Value="True" onChange="return setVarChange()" <%If SortByDay = True Then Response.Write "CHECKED"%>></P>
            <P class="config"><span id="Label1"><a href="../Themes.asp">Theme</a> Template<font color="Red">*</Font> : </span>
<Select name="Template" onChange="return setVarChange()">
<%
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(Server.MapPath("../Templates/"))
        Set Folders = Folder.SubFolders

        For Each Folder in Folders 
        Response.Write "<Option"
        If Template = Folder.Name Then Response.Write " SELECTED"
        Response.Write ">" & Folder.Name & "</Option>"
        Next

	Set Folders = Nothing
	Set Folder = Nothing
        Set FSO = Nothing
%>
</select>
</P>
            <P class="config"><span id="Label1">BackgroundColor<font color="Red">*</Font> : </span><input Name="BackgroundColor" type="text" style="width:20%;" Value="<%=BackgroundColor%>"> <a href="JavaScript:Colors('Box=BackgroundColor')"><img src="../Images/Color.gif" border="0"></a></P>
            <P class="config"><span id="Label1">12Hour Times<font color="Red">*</Font> : </span><input Name="TimeFormat" type="Checkbox" Value="True" onChange="return setVarChange()" <%If TimeFormat = True Then Response.Write "CHECKED"%>></P>
            <P class="config"><span id="Label1">Logging<font color="Red">*</Font> : </span><input Name="Logging" type="Checkbox" Value="True" <%If Logging = True Then Response.Write "CHECKED"%>></P>
            <P></P>
            <Input Type="submit" Value="Save">
            <P class="config" align="Center"><font color="Red">*</Font> - Indicates a required field.</P>
            <P class="config" align="Center"><B>Note</B>: To change CommentNotification, PingOMatic, ArgosoftMail support, CalendarCheck, MailingList, OtherLinks, Register, RSS, RSSImage, UseImagesInEditor and the TimeOffset...Edit Includes/Config.asp using notepad</P>
        </form>
<% Else

'### CheckBox Check! ###'
EnableComments = Request.Form("EnableComments")
Logging        = Request.Form("Logging")
Polls 	       = Request.Form("Polls")
ShowCategories = Request.Form("ShowCategories")
SortByDay      = Request.Form("SortByDay")
TimeFormat     = Request.Form("TimeFormat")

If EnableComments = "" Then EnableComments = False
If Logging        = "" Then Logging = False
If Polls          = "" Then Polls = False
If ShowCategories = "" Then ShowCategories = False
If SortByDay      = "" Then SortByDay = False
If TimeFormat     = "" Then TimeFormat = False

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Config", Database
Records("SiteName") = Request.Form("SiteName")
Records("CookieName") = Request.Form("CookieName")
Records("Copyright") = Request.Form("Copyright")
Records("EntriesPerPage") = Request.Form("EntriesPerPage")
Records("EnableComments") = EnableComments
Records("ReaderPassword") = Request.Form("ReaderPassword")
Records("Polls") = Polls
Records("ShowCategories") = ShowCategories
Records("SiteDescription") = Request.Form("SiteDescription")
Records("SiteSubTitle") = Request.Form("SiteSubTitle")
Records("Template") = Request.Form("Template")
Records("SortByDay") = SortByDay
Records("BackgroundColor") = Request.Form("BackgroundColor")
Records("12HourTimeFormat") = TimeFormat
Records("Logging") = Logging
Records.Update

'#### Close Objects ###	
Records.Close

Response.Write "<p align=""Center"">Config Update Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
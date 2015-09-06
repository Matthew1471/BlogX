<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script language="JavaScript" type="text/javascript" src="../Includes/RTF.js"></script>
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<%
'--- Open set ---'
Records.Open "SELECT * FROM Main",Database, 1, 3

If NOT Records.EOF Then

'--- Setup Variables ---'
   MainID = Records("MainID")
   MainText = Records("MainText")

End If

Records.Close
%>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Enable The Main Page<font color="Red">*</Font> : </span><input Name="EnableMainPage" type="Checkbox" Value="True" onChange="return setVarChange()" <%If EnableMainPage = True Then Response.Write "CHECKED"%>></P>
            <P><span id="Label1">Main Page : </span><br>
            <table border="0" cellpadding="1" cellspacing="0" width="100%">
			<tr>
			<td bgcolor="<%=CalendarBackground%>" align="left">

		<% If UseImagesInEditor <> 0 Then %>
			<img src="<%=SiteURL%>Images/Editor/Bold.gif" title="Bold" onclick="boldThis()">
			<img src="<%=SiteURL%>Images/Editor/Italicize.gif" title="Italics" onclick="italicsThis()">
			<img src="<%=SiteURL%>Images/Editor/Underline.gif"  title="Underline" onclick="underlineThis()">
			<img src="<%=SiteURL%>Images/Editor/Strike.gif"title="CrossOut" onclick="crossThis()">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<img src="<%=SiteURL%>Images/Editor/Left.gif" title="Left" onclick="leftThis()">
			<img src="<%=SiteURL%>Images/Editor/Center.gif" title="Center" onclick="centerThis()">
			<img src="<%=SiteURL%>Images/Editor/Right.gif" title="Left" onclick="rightThis()">
			<img src="<%=SiteURL%>Images/Editor/Photo.gif" title="Style the image as a photo" onclick="photoThis()">
			</td>

			<td bgcolor="<%=CalendarBackground%>" align="right">
			<img src="<%=SiteURL%>Images/Editor/SpellCheck.gif" title="Spell Check" onclick="SpellThis()">
			<img src="<%=SiteURL%>Images/Editor/URL.gif" title="Link" onclick="linkThis()">
			<img src="<%=SiteURL%>Images/Editor/Image.gif" title="Image" onclick="imageThis('')">
			&nbsp;
			<img src="<%=SiteURL%>Images/Editor/Line.gif" title="Line" onclick="lineThis()">
		<% Else %>
			<input type="button" value="Bold" onclick="boldThis()">
			<input type="button" value="Italics" onclick="italicsThis()">
			<input type="button" value="Underline" onclick="underlineThis()">
			<input type="button" value="CrossOut" onclick="crossThis()">
			</td>

			<td bgcolor="<%=CalendarBackground%>" align="right">
			<input type="button" value="Link" onclick="linkThis()">
			<input type="button" value="Image" onclick="imageThis('')">
			&nbsp;
			<input type="button" value="Line" onclick="lineThis()">
		<% End If %>

			</td>
			</tr>

            <tr>
            <td colspan="2">
            <Textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:15em;width:100%;" onChange="return setVarChange()"><%=MainText%></textarea>
            </tr>
	    </table>
            </P>
            <P></P>
            <Input Type="submit" Value="Save">
        </form>
<% Else

EnableMainPageRequest = Request.Form("EnableMainPage")
If EnableMainPageRequest  = "" Then EnableMainPageRequest = False

If EnableMainPageRequest <> EnableMainPage Then

Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Config", Database
Records("EnableMainPage") = EnableMainPageRequest
Records.Update
Records.Close

Response.Write "<p align=""Center"">Config Update Successful</p>"
End If

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Main", Database
If Records.EOF = True Then Records.AddNew
Records("MainText") = Request.Form("Content")
Records.Update

'#### Close Objects ###
Records.Close

Response.Write "<p align=""Center"">MainPage Update Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
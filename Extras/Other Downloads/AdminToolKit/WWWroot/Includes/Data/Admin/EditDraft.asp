<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script language="JavaScript" type="text/javascript" src="../Includes/RTF.js"></script>
<DIV id=content>
<%
If Request.Form("Action") = "Post" Then

	'### Did We Type In Text? ###'
	If Request.Form("Content") = "" Then
	Response.Write "<p align=""Center"">No Text Entered</p>"
	Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
	Response.Write "</Div>"
	%>
	<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
	<%
	Response.End
	End If

	'### Open The Records Ready To Write ###
	Records.CursorType = 2
	Records.LockType = 3

	Records.Open "SELECT Title, Text FROM Draft", Database

	If NOT Records.EOF Then
	Records("Title") = Left(Request.Form("Title"),80)
	Records("Text") = Request.Form("Content")
	Records.Update
	Response.Write "<p align=""center""><b>Saved To Draft NotePad<br>(" & Now() & ")</b></p>"
	Else
	Response.Write "<h1 align=""center""><b>Drafts are unavailable</b></h1>"
	End If

	Records.Close

End If 

'--- Open set ---'
    Records.Open "SELECT Title, Text FROM Draft",Database, 1, 3

If NOT Records.EOF Then

'--- Setup Variables ---'
   Title    = Records("Title")
   Text     = Records("Text")

End If

Records.Close
%>

<!--- Start Content For Draft --->
<DIV class=entry>
<H3 class=entryTitle><%=Title%> (Preview Entry)</H3>
<DIV class=entryBody><P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P></DIV>
<P class=entryFooter>
<% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""../Images/Email.gif""></a></acronym>"%>
<b><%=Now()%></b>
<% If EnableComments <> False Then Response.Write " | <SPAN class=""comments"">Comments [0]</A></SPAN>"%>
</SPAN></P></DIV>
<!--- End Content --->

<hr>

<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Title : </span><input Name="Title" type="text" value="<%=Replace(Title,"""","&quot;")%>" style="width:80%;" onChange="return setVarChange()"></P>

            <P>Content :<br>
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
            <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:30em;width:100%;" onChange="return setVarChange()"><%=Replace(Replace(Text,"&","&amp;"),"<","&lt;")%></textarea>
            </tr>
	    <% If LegacyMode <> True Then %>
            <tr>
             <td bgcolor="<%=CalendarBackground%>" align="right" colspan="2">&nbsp;</td>
            </tr>
	    <% End If %>
			</table>
            </P>

            <P align="center"><Input Type="submit" Value="Save Notes"></P>
            
        </form>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
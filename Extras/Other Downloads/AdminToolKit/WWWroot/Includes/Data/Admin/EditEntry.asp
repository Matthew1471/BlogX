<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script language="JavaScript" type="text/javascript" src="../Includes/RTF.js"></script>
<SCRIPT>
// Show/Hide functions for non-pointer layer/objects
function show(id) {

	        var Advanceditem = document.all.item(id)

		if (Advanceditem != null)
		{
			if (Advanceditem.length != null)
			{
			    for (i=0; i<Advanceditem.length; i++)
			    {
					Advanceditem(i).style.display = "block";
				}                                 
			}                                         
		}


}
</SCRIPT>
<style type="text/css"> 
#idHidden{ 
display : none;
} 
</style>
<DIV id=content>
<%
'--- Querish Querystring ---'
Dim Requested, DelRecNo
Requested = Request.Querystring("Entry")
Delete = Request.Querystring("Delete")

If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then Requested = 0

If Request.Form("Action") <> "Post" Then

'--- Open set ---'
    If (Requested <> 0) AND (Delete <> "True") Then 
    Records.Open "SELECT * FROM Data WHERE RecordID=" & Requested,Database, 1, 3
    ElseIf Requested = 0 Then
    Records.Open "SELECT * FROM Data ORDER By RecordID DESC",Database, 1, 3
    Else
 
    Database.Execute "DELETE FROM Data WHERE RecordID=" & Requested
    Database.Close
    Set Records = Nothing
    Set Database = Nothing
    Response.Redirect(SiteURL & PageName) 

    End If

If NOT Records.EOF Then

'--- Setup Variables ---'
   RecordID = Records("RecordID")
   Title = Records("Title")
   Text = Records("Text")
   Category =  Records("Category")
   Password =  Records("Password")

   sDay = Records("Day")
   sMonth = Records("Month")
   sYear = Records("Year")

   TimePosted = Records("Time")

End If

Records.Close
%>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Title : </span><input Name="Title" type="text" value="<%=Replace(Title,"""","&quot;")%>" style="width:80%;" onChange="return setVarChange()"> <a href="?Entry=<%=Requested%>&Delete=True" title="DELETE this entry" onClick="return confirm('Warning! If You Continue Entry #<%=Requested%> Will Be DELETED.')"><img src="<%=SiteURL%>Images/Delete.gif" width="15" height="15" border="0"></a></P>

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
            <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:40em;width:100%;" onChange="return setVarChange()"><%=Replace(Replace(Text,"&","&amp;"),"<","&lt;")%></textarea>
            </tr>
	    <% If LegacyMode <> True Then %>
            <tr>
             <td bgcolor="<%=CalendarBackground%>" align="right" colspan="2">
             <a href="#" onclick="javascript:show('idHidden'); return false;"><img alt="Turn on Advanced editing features" border="0" src="<%=SiteURL%>Images/Editor/Advanced.gif" width="61" height="16"></a>
             </td>
            </tr>
	    <% End If %>
			</table>
            </P>

            <% If ShowCat <> False Then Response.Write "<P>Category : <input Name=""Category"" type=""text"" value=""" & Replace(Category,"%20"," ") & """ style=""width:10%;"" onChange=""return setVarChange()""></P>"

	    If LegacyMode <> True Then 

			Response.Write "<P id=""idHidden"">Change the entry date, Day : "
			Response.Write "<select name=""nDay"">" & VbCrlf			
			For i = 1 to 31
				if i=sDay then
					Response.Write "<option selected=""true"" value=""" & i & """>" & i & "</option>" & VbCrlf
				else
					Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
				end if
			next
		    Response.Write "</select>"
			Response.Write " Month : "
			Response.Write "<select name=""nMonth"">" & VbCrlf			
			For i = 1 to 12
				if i=sMonth then
					Response.Write "<option selected=""true"" value=""" & i & """>" & i & "</option>" & VbCrlf
				else
					Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
				end if
			next
		    Response.Write "</select>"
			Response.Write " Year : "
			Response.Write "<select name=""nYear"">" & VbCrlf			
			For i = 2000 to 2030
				if i=sYear then
					Response.Write "<option selected=""true"" value=""" & i & """>" & i & "</option>" & VbCrlf
				else
					Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
				end if
			next
		    Response.Write "</select>"

		    Response.Write "<br><br>Change the entry time, Time : <input name=""Time"" value=""" & TimePosted & """ style=""width:10%;"">"

			Response.Write "</p>"
		%>

            <table border="0" cellpadding="0" cellspacing="0" width="30%" id="idHidden">
	    <tr><td bgcolor="<%=CalendarBackground%>" align="left"><font color="White">
	        <acronym title="If you type in a password, your viewers will need to enter it to view the Entry, leaving it blank means everyone can see your entry"><img border=0 src="<%=SiteURL%>Images/Help.gif">Optional<br>Entry Password</acronym></font></td>
		<td bgcolor="<%=CalendarBackground%>" align="center"><input name="password" type="text" value="<%=Password%>" maxlength="10" onChange="return setVarChange()"></td>
            </tr>
	    </table>
	    <% End If %>

            <P></P>
            <Input Type="submit" Value="Save">
        </form>
<% Else
'Dimension variables
Dim EntryCat            'Category

EntryCat = Request.Form("Category")

'### Filter & Clean ###
EntryCat = Replace(EntryCat,"'","&#39;")
EntryCat = Replace(EntryCat," ","%20")

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

    If Requested <> 0 Then
    Records.Open "SELECT * FROM Data WHERE RecordID=" & Requested,Database, 1, 3
    Else
    Records.Open "SELECT * FROM Data ORDER By RecordID DESC", Database
    End If

Records("Title") = Left(Request.Form("Title"),80)
Records("Text") = Request.Form("Content")
Records("Password") = Request.Form("Password")
Records("Category") = EntryCat

If IsNumeric(Request.Form("nDay")) Then Records("Day") = Request.Form("nDay")
If IsNumeric(Request.Form("nMonth")) Then Records("Month") = Request.Form("nMonth")
If IsNumeric(Request.Form("nYear")) Then Records("Year") = Request.Form("nYear")

If IsDate(Request.Form("Time")) Then Records("Time") = Request.Form("Time")

Records.Update
Records.Close

Response.Write "<p align=""Center"">Entry Update Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
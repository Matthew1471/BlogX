<%
   PingbackPage = True
   AlertBack = True
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/xmlrpc.asp" -->
<script language="JavaScript" type="text/javascript" src="../Includes/RTF.js"></script>
<SCRIPT>
// Show/Hide functions for non-pointer layer/objects
function show(id) {

	if(document.all.item) //IE
	{
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
	else
	{
		alert("Your browser is not supported");
	}

}





</SCRIPT>
<style type="text/css"> 
#idHidden{ 
display : none;
} 
</style> 
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P>Title : <input Name="Title" type="text" style="width:80%;" maxlength="80"></P>

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
             <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:15em;width:100%;" onChange="return setVarChange()"></textarea>
             </td>
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

                <% If (ShowCat <> False) AND (LegacyMode <> True) Then

			'--- Open set ---'
    			Records.CursorLocation = 3 ' adUseClient
    			Records.Open "SELECT * FROM Data ORDER BY Category",Database, 1, 3

			'--- Set Category ---'
			Set Category = Records("Category")

			'-- Write Them In ---'
                	Response.Write "Select an existing category : "
			Response.Write "<select name=""SelectCategory"" onChange=""document.AddEntry.Category.value = this[this.selectedIndex].value; "">" & VbCrlf
			Response.Write "<option value="""">-- New --</Option>" & VbCrlf

			Do Until (Records.EOF or Records.BOF)
			If (LastCat <> Category) OR (IsNull(LastCat) = True) AND (Category <> "") Then 
			Response.Write "<option value=""" & Replace(Category, "%20", " ") & """>" & Replace(Category, "%20", " ") & "</option>" & VbCrlf
			LastCat = Category
			End If
			Records.MoveNext
			Loop
                
			Response.Write "</select>"

			'-- Close The Database & Records ---'
			Records.Close

			Response.Write " or create/edit the selected Category : <input Name=""Category"" type=""text"" style=""width:10%;"" maxlength=""50""></P>"

			ElseIf ShowCat <> False Then 
			Response.Write "<P>Category : <input Name=""Category"" type=""text"" style=""width:10%;"" maxlength=""50""></P>"

                  End If %>
	    <% If LegacyMode <> True Then %>
            <table border="0" cellpadding="0" cellspacing="0" width="30%" id="idHidden">
	    <tr><td bgcolor="<%=CalendarBackground%>" align="left"><font color="White">
	        <acronym title="If you type in a password, your viewers will need to enter it to view the Entry, leaving it blank means everyone can see your entry"><img border=0 src="<%=SiteURL%>Images/Help.gif">Optional<br>Entry Password</acronym></font></td>
		<td bgcolor="<%=CalendarBackground%>" align="center"><input name="password" type="text" maxlength="10"></td>
            </tr>
	    </table>
	    <% End If %>

	    <P id="idHidden"><font color="red">Note :</font> You can drag the following link : <a title="BlogIt!" href="javascript:Q='';x=document;y=window;if(x.selection){Q=x.selection.createRange().text;}else if(y.getSelection){Q=y.getSelection();}else if(x.getSelection){Q=x.getSelection();}void(window.open('<%=Replace(SiteURL,"'","\'")%>Admin/Toolbar.asp?t='+escape(Q)+'&u='+escape(location.href)+'&n='+escape(document.title),'bloggerForm','scrollbars=no,width=475,height=300,top=175,left=75,status=yes,resizable=yes'));">BlogIt!</a> to your links bar or add it to your favourites and when you click it, it'll open up a window with information (Including any highlighted text) and the link to the site you’re currently browsing so you can post about it.</P>

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
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time FROM Data", Database
Records.AddNew

Records("Title") = Left(Request.Form("Title"),80)
Records("Text") = Request.Form("Content")
Records("Category") = EntryCat
Records("Password") = Request.Form("Password")

Records("Day") = Day(DateAdd("h",TimeOffset,Now()))
Records("Month") = Month(DateAdd("h",TimeOffset,Now()))
Records("Year") = Year(DateAdd("h",TimeOffset,Now()))
Records("Time") = TimeValue(DateAdd("h",TimeOffset,Time()))
Records.Update

Records.MoveLast

Dim RecordID
RecordID = Records("RecordID")

'#### Close Objects ###
Records.Close

If NotifyPingOMatic <> 0 Then 
        On Error Resume Next
	ReDim paramList(2)
	paramList(0)=SiteName
	paramList(1)=SiteURL & "ViewItem.asp?Entry=" & RecordID
	myresp = xmlRPC ("http://rpc.pingomatic.com/", "weblogUpdates.ping", paramList)

        '-- DEBUG --'
	'Response.write("<pre>" & Replace(serverResponseText, "<", "&lt;", 1, -1, 1) & "</pre>")

On Error GoTo 0
End If

Response.Write "<p align=""Center"">Entry Submission Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
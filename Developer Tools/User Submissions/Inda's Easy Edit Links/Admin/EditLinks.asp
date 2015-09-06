<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->

<%
'save form results
If AllowEditingLinks <> 0 Then

	If Request.ServerVariables("Content_Length") > 0 Then

		Dim Content, Content2
		
		'build content for Links file
		for i = 1 To Request.Form("Titles1").Count
			
			if(Request.Form("Titles1")(i) <> "" AND Request.Form("URLs1")(i) <> "") then
				
				Content = Content & "*" & VbCrlf &_
							Request.Form("Titles1")(i) & VbCrlf &_
							Request.Form("URLs1")(i) & VbCrlf
			end if
		Next
		
		
		'build content for Other Links file
		for i = 1 To Request.Form("Titles2").Count
			
			if(Request.Form("Titles2")(i) <> "" AND Request.Form("URLs2")(i) <> "") then
			
				Content2 = Content2 & "*" & VbCrlf &_
							Request.Form("Titles2")(i) & VbCrlf &_
							Request.Form("URLs2")(i) & VbCrlf
			end if			
		Next
		

		'### Write ###
		Set FSO = CreateObject("Scripting.FileSystemObject")

		Set File = FSO.CreateTextFile(LinksPath, True)
		File.Write Content
		File.Close
		Set File = nothing

		Set File = FSO.CreateTextFile(OtherLinksPath, True)
		File.Write Content2
		File.Close
		Set File = nothing

		Set FSO = nothing
	
	end if
end if
%>

<style>#pLinks1 td img, #pLinks2 td img {cursor:pointer}</style>
	
<script>

	//set for mozilla because F5 (refresh) doesn't reset the forms.
	window.onload = function ResetForm() { if(document.AddEntry) document.AddEntry.reset(); }


	//redrawing the complete table for every change is the only foolproof way that
	//I found of doing this. Messy.
	
	function GetFormData(iType)
	{
		var aData = new Array();
		
		var oTitles = eval("document.AddEntry.Titles" + iType)
		var oURLS = eval("document.AddEntry.URLs" + iType)
		
		if(oTitles.length)
		{
			for(var i=0; i<oTitles.length; i++)
			{
				aData[i] = new Array();
				aData[i][0] = oTitles[i].value;
				aData[i][1] = oURLS[i].value;
			}
		}
		else
		{
			aData[0] = new Array();
			aData[0][0] = oTitles.value;
			aData[0][1] = oURLS.value;
		}

		return aData;
	}
	
	
	function WriteTable(iType, aData)
	{
		var oPara = document.getElementById("pLinks" + iType)
		
		//create the HTML
		var sHTML = '<table width="100%" cellpadding="1" cellspacing="0" border="0" id="Links' + iType + '">';
		sHTML += '<tr><th width="33%">Title</th><th width="67%">URL</th><th></th></tr>'
	
		//loop through each row of data
		for(var i=0; i<aData.length; i++)
		{
			sHTML += '<tr>'
			
			sHTML += '<td>'
			sHTML += '<input autocomplete="off" type="text" name="Titles' + iType + '" style="width:100%" '
			sHTML += 'value="' + aData[i][0] + '" '
			sHTML += 'onchange="TextChanged(' + iType + ', 1, ' + i + ')" '
			sHTML += 'onkeyup="TextChanged(' + iType + ', 1, ' + i + ')">'
			sHTML += '</td>'
			
			sHTML += '<td>'
			sHTML += '<input autocomplete="off" type="text" name="URLs' + iType + '" style="width:100%" '
			sHTML += 'value="' + aData[i][1] + '" '
			sHTML += 'onchange="TextChanged(' + iType + ', 2, ' + i + ')" '
			sHTML += 'onkeyup="TextChanged(' + iType + ', 2, ' + i + ')">'
			sHTML += '</td>'
			
			sHTML += '<td nowrap="true">'
			if(i != aData.length-1) //don't add buttons on the last line
			{
				sHTML += '<img src="../Images/up.gif" title="Move up" border="0" onclick="MoveUp(' + iType + ', ' + i + ')"> '
				sHTML += '<img src="../Images/down.gif" title="Move down" border="0" onclick="MoveDown(' + iType + ', ' + i + ')"> '
				sHTML += '<img src="../Images/cancel.gif" title="Delete" border="0" onclick="DeleteLink(' + iType + ', ' + i + ')">'
			}
			else
			{
				sHTML += '<img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0">'
			}
			sHTML += '</td>'
			
			sHTML += '</tr>'
		}
		
		sHTML += '</table>'
		
		oPara.innerHTML = sHTML
	}
	
	function DeleteLink(iType, iRow)
	{
		//get begining section
		var aData = GetFormData(iType).slice(0,iRow)
		
		//add end section
		aData = aData.concat(GetFormData(iType).slice(iRow+1))
								
		WriteTable(iType, aData)
	}
	
	function MoveUp(iType, iRow)
	{
		var aData = GetFormData(iType)
		
		//check if top row has been clicked
		if(iRow != 0)
		{
			//swap the values
			var sTemp1 = aData[iRow-1][0]
			var sTemp2 = aData[iRow-1][1]
			
			aData[iRow-1][0] = aData[iRow][0]
			aData[iRow-1][1] = aData[iRow][1]
			
			aData[iRow][0] = sTemp1
			aData[iRow][1] = sTemp2
		
			WriteTable(iType, aData)
		}		
	}
	
	function MoveDown(iType, iRow)
	{
		var aData = GetFormData(iType)
		
		//check if bottom row has been clicked
		if(iRow != aData.length-2)
		{
			//swap the values
			var sTemp1 = aData[iRow+1][0]
			var sTemp2 = aData[iRow+1][1]
			
			aData[iRow+1][0] = aData[iRow][0]
			aData[iRow+1][1] = aData[iRow][1]
			
			aData[iRow][0] = sTemp1
			aData[iRow][1] = sTemp2
		
			WriteTable(iType, aData)
		}
	}

	function TextChanged(iType, iField, iRow)
	{
							
		var aData = GetFormData(iType)
		
		//if the last row is changed, add another row.
		//tabbing can cause the wrong textbox to fire this function. check if last row is already blank
		if(aData.length == iRow + 1 && !(aData[aData.length-1][0] == "" && aData[aData.length-1][1] == ""))
		{
			AddNewRow(iType)
			
			//replace focus
			var s = (iField == 1) ? "Titles" : "URLs";
							
			var oTB = eval("document.AddEntry." + s + iType + "[" + iRow + "]")
			
			try //IE
			{
				var oRange = oTB.createTextRange();
				oRange.collapse(false);
				oRange.select();
			}
			catch(er) //Moz
			{
				oTB.focus()
			}		
		}
					
		setVarChange()
	}
		
	function AddNewRow(iType)
	{
		var aData = GetFormData(iType)
				
		aData[aData.length] = ["",""]

		WriteTable(iType, aData)		
	}
	
</script>



<DIV id=content>
<% 
If AllowEditingLinks <> 0 Then


	Dim iCount
	iCount = 0
	
	'write form
	%>
	
	<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">

		<P><b>Links File Location: </b><%=LinksPath%>
		
		<p id="pLinks1">
		<table width="100%" cellpadding="1" cellspacing="0" border="0" id="Links1">
	
			<tr><th width="33%">Title</th><th width="67%">URL</th><th></th></tr>
	<%

	Set FSO = server.CreateObject("Scripting.FileSystemObject")

	' Get a handle to the file	
	Set File = FSO.GetFile(LinksPath)

    ' Read the file line by line
	Set TextStream = File.OpenAsTextStream(1, -2)

	'Populate textboxes
	Do While Not TextStream.AtEndOfStream          
		If TextStream.Readline = "*" Then
			%>
			<tr>
				<td><input autocomplete="off" type="text" name="Titles1" style="width:100%" value="<%=TextStream.Readline%>" onchange="TextChanged(1, 1, <%=iCount%>)" onkeyup="TextChanged(1, 1, <%=iCount%>)"></td>
				<td><input autocomplete="off" type="text" name="URLs1" style="width:100%" value="<%=TextStream.Readline%>" onchange="TextChanged(1, 2, <%=iCount%>)" onkeyup="TextChanged(1, 2, <%=iCount%>)"></td>
				<td nowrap="true">
					<img src="../Images/up.gif" title="Move up" border="0" onclick="MoveUp(1, <%=iCount%>)">
					<img src="../Images/down.gif" title="Move down" border="0" onclick="MoveDown(1, <%=iCount%>)">
					<img src="../Images/cancel.gif" title="Delete" border="0" onclick="DeleteLink(1, <%=iCount%>)">
				</td>
			</tr>
			<%
			iCount = iCount + 1
		end if
	Loop
	
    TextStream.Close
	Set TextStream = nothing
	
	'add blank line
	%>
			<tr>
				<td><input autocomplete="off" type="text" name="Titles1" style="width:100%" value="" onchange="TextChanged(1, 1, <%=iCount%>)" onkeyup="TextChanged(1, 1, <%=iCount%>)"></td>
				<td><input autocomplete="off" type="text" name="URLs1" style="width:100%" value="" onchange="TextChanged(1, 2, <%=iCount%>)" onkeyup="TextChanged(1, 2, <%=iCount%>)"></td>
				<td nowrap="true"><img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0"></td>
			</tr>
		</table>
		
		<hr>
	
		<P><b>Other Links File Location: </b><%=OtherLinksPath%>
		
		<p id="pLinks2">
		<table width="100%" cellpadding="1" cellspacing="0" border="0" id="Links2">
	
			<tr><th width="33%">Title</th><th width="67%">URL</th><th></th></tr>		
	
	<%
	'reset counter
	iCount = 0
	
	' Get a handle to the file	
	Set File = FSO.GetFile(OtherLinksPath)

    ' Read the file line by line
	Set TextStream = File.OpenAsTextStream(1, -2)

	'Populate textboxes
	Do While Not TextStream.AtEndOfStream          
		If TextStream.Readline = "*" Then
			%>
			<tr>
				<td><input autocomplete="off" type="text" name="Titles2" style="width:100%" value="<%=TextStream.Readline%>" onchange="TextChanged(2, 1, <%=iCount%>)" onkeyup="TextChanged(2, 1, <%=iCount%>)"></td>
				<td><input autocomplete="off" type="text" name="URLs2" style="width:100%" value="<%=TextStream.Readline%>" onchange="TextChanged(2, 2, <%=iCount%>)" onkeyup="TextChanged(2, 2, <%=iCount%>)"></td>
				<td nowrap="true">
					<img src="../Images/up.gif" title="Move up" border="0" onclick="MoveUp(2, <%=iCount%>)">
					<img src="../Images/down.gif" title="Move down" border="0" onclick="MoveDown(2, <%=iCount%>)">
					<img src="../Images/cancel.gif" title="Delete" border="0" onclick="DeleteLink(2, <%=iCount%>)">
				</td>
			</tr>
			<%
			iCount = iCount + 1
		end if
	Loop
	
    TextStream.Close
	Set TextStream = nothing
	
	'add blank line
	%>
			<tr>
				<td><input autocomplete="off" type="text" name="Titles2" style="width:100%" value="" onchange="TextChanged(2, 1, <%=iCount%>)" onkeyup="TextChanged(2, 1, <%=iCount%>)"></td>
				<td><input autocomplete="off" type="text" name="URLs2" style="width:100%" value="" onchange="TextChanged(2, 2, <%=iCount%>)" onkeyup="TextChanged(2, 2, <%=iCount%>)"></td>
				<td nowrap="true"><img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0"> <img src="../Images/blank.gif" width="14" height="0"></td>
			</tr>
		</table>
	<%

	Set File = nothing
	Set FSO = nothing

	
	%>
		<p align="center"><Input Type="submit" Value="Save All"></p>
		
	</form>
	
	<% 

Else

Response.Write "<p align=""Center"">You are not allowed to edit links</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If
%>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
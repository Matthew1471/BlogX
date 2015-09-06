<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<% 
If Request.Form("Action") <> "Post" Then

Dim Text, Text2

Set FSO = server.CreateObject("Scripting.FileSystemObject")

Set File = fso.OpenTextFile(LinksPath,1,1,0)
If File.AtEndOfStream <> True Then Text = File.ReadAll
File.close

Set File = fso.OpenTextFile(OtherLinksPath,1,1,0)
If File.AtEndOfStream <> True Then Text2 = File.ReadAll
File.close

Set File = nothing
Set FSO = nothing
%>
<Form Name="AddEntry" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P>Links :<br>
            <table border="0" cellpadding="1" cellspacing="0" width="100%">
			<tr>
			<td bgcolor="<%=CalendarBackground%>" align="left"><%=LinksPath%></td>
			</tr>

            <tr>
            <td colspan="2">
            <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:10em;width:100%;" onChange="return setVarChange()"><%=Text%></textarea>
            </tr>
			</table>
            </P>

            <P>Other Links :<br>
            <table border="0" cellpadding="1" cellspacing="0" width="100%">
			<tr>
			<td bgcolor="<%=CalendarBackground%>" align="left"><%=OtherLinksPath%></td>
			</tr>

            <tr>
            <td colspan="2">
            <textarea Name="Content2" DESIGNTIMEDRAGDROP="96" style="height:10em;width:100%;" onChange="return setVarChange()"><%=Text2%></textarea>
            </tr>
			</table>
            </P>

            <P></P>
            <Input Type="submit" Value="Save">
        </form>
<% Else

Content = Request.Form("Content")
Content2 = Request.Form("Content2")

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
Content = "*" & VbCrlf & VbCrlf & VbCrlf
End If

If Request.Form("Content2") = "" Then
Content2 = "*" & VbCrlf & VbCrlf & VbCrlf
End If

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

Response.Write "<p align=""Center"">Links Update Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If
%>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
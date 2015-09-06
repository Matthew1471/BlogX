<!-- #INCLUDE FILE="Includes/Admin.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Edit News</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
<% If Request.Form("Action") <> "Post" Then

Dim Database, Records
Dim Title, Content, TimePosted, DatePosted

'--- Open Database ---'
Set Database = Server.CreateObject("ADODB.connection")
Database.Open  "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'--- Open set ---'
set Records = Server.CreateObject("ADODB.recordset")
    Records.Open "SELECT * FROM News ORDER By ID DESC",Database, 1, 3

If NOT Records.EOF Then

'--- Setup Variables ---'
Title = Records("Title")
Content = Records("Content")
TimePosted = Records("Time")
DatePosted = Records("Date")

End If

Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
<Form Name="AddEntry" Method="Post">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">Title : </span><input Name="Title" type="text" value="<%=Title%>" maxlength="50" size="50"></P>

                        <P>Content :<br>
            <table border="0" cellpadding="0" cellspacing="0" width="100%">

            <tr>
            <td colspan="2">
            <textarea Name="Content" DESIGNTIMEDRAGDROP="96" style="height:10em;width:100%;"><%=Content%></textarea>
            </tr>
			</table>
            </P>
            <P></P>
            <P align="center"><Input Type="submit" Value="Save"></P>
        </form>
<% Else

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
Response.Write "<p align=""Center"">No Text Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
WriteFooter
Response.End
End If

'### Create a connection odject ###
Set Database = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
Database.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set Records = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM News ORDER By ID DESC", Database
Records("Title") = Request.Form("Title")
Records("Content") = Request.Form("Content")

Records("Time") = Time()
Records("Date") = Date()
Records.Update

'### Close Objects ###
Records.Close
Database.Close

'#### Destroy Objects ###	
Set Database = Nothing
Set Records = Nothing

Response.Write "<p align=""Center"">Entry Update Successfull</p>"
Response.Write "<p align=""Center""><a href=""Admin_Main.asp"">Back</font></a></p>"

End If
%>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
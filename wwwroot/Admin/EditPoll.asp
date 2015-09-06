<%
' --------------------------------------------------------------------------
'¦Introduction : Edit Poll Page.                                            ¦
'¦Purpose      : Allows blog administrator to edit the current poll.        ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/RTF.js.                      ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************

AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script type="text/javascript" src="../Includes/RTF.js"></script>
<div id="content">
<% If Request.Form("Action") <> "Post" Then 

'--- Open set ---'
Records.Open "SELECT Content, Des1, Des2, Des3, Des4 FROM Poll ORDER By PollID DESC",Database, 0, 1

If NOT Records.EOF Then

  If Request.Querystring("Delete") <> "True" Then
	  '--- Setup Variables ---'
	  Dim Content
   	  Content = Records("Content")
   	  Des1 =  Records("Des1")
   	  Des2 =  Records("Des2")
   	  Des3 =  Records("Des3")
   	  Des4 =  Records("Des4")
  Else
    	Records.Delete
    	Records.Update
   	Records.Close

    	Database.Close
    	Set Records = Nothing
    	Set Database = Nothing
    	Response.Redirect(SiteURL & PageName) 
  End If

End If

Records.Close
%>
<form id="AddEntry" method="post" action="EditPoll.asp" onsubmit="return setVar()">

 <p>
  <input name="Action" type="hidden" value="Post"/>
  Question <a href="?Delete=True" title="DELETE this poll" onclick="return confirm('Warning! If You Continue This Poll Will Be DELETED.')"><img style="border-style: none;" alt="Delete Icon" src="<%=SiteURL%>Images/Delete.gif" width="15" height="15"/></a> :
 </p>

 <table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
   <td style="background-color: <%=CalendarBackground%>" align="left">
   <% If UseImagesInEditor <> 0 Then %>
    <img src="<%=SiteURL%>Images/Editor/Bold.gif" title="Bold" alt="Bold" onclick="boldThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Italicize.gif" title="Italics" alt="Italics" onclick="italicsThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Underline.gif"  title="Underline" alt="Underline" onclick="underlineThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Strike.gif"title="CrossOut" alt="CrossOut" onclick="crossThis()"/>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <img src="<%=SiteURL%>Images/Editor/Left.gif" title="Left" alt="Left" onclick="leftThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Center.gif" title="Center" alt="Center" onclick="centerThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Right.gif" title="Right" alt="Right" onclick="rightThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Photo.gif" title="Style the image as a photo" alt="Style the image as a photo" onclick="photoThis()"/>
   </td>
     <td style="background-color: <%=CalendarBackground%>" align="right">
    <img src="<%=SiteURL%>Images/Editor/SpellCheck.gif" title="Spell Check" alt="Spell Check" onclick="SpellThis()"/>
    <img src="<%=SiteURL%>Images/Editor/URL.gif" title="Link" alt="Link" onclick="linkThis()"/>
    <img src="<%=SiteURL%>Images/Editor/Image.gif" title="Image" alt="Image" onclick="imageThis('')"/>
    &nbsp;
    <img src="<%=SiteURL%>Images/Editor/Line.gif" title="Line" alt="Line" onclick="lineThis()"/>
   <% Else %>
    <input type="button" value="Bold" onclick="boldThis()"/>
    <input type="button" value="Italics" onclick="italicsThis()"/>
    <input type="button" value="Underline" onclick="underlineThis()"/>
    <input type="button" value="CrossOut" onclick="crossThis()"/>
   </td>
    <td style="background-color: <%=CalendarBackground%>" align="right"/>
    <input type="button" value="Link" onclick="linkThis()"/>
    <input type="button" value="Image" onclick="imageThis('')"/>
    &nbsp;
    <input type="button" value="Line" onclick="lineThis()"/>
   <% End If %>
   </td>
  </tr>
  <tr>
   <td colspan="2">
    <textarea name="Content" cols="141" rows="3" style="height:5em;width:99%;" onchange="return setVarChange()"><%=Content%></textarea>
   </td>
  </tr>
 </table>

 <p>
  Option1 : <input name="Option1" type="text" style="width:10%;" maxlength="50" value="<%=Des1%>"/><br/>
  Option2 : <input name="Option2" type="text" style="width:10%;" maxlength="50" value="<%=Des2%>"/><br/>
  Option3<span style="color:Red">*</span> : <input name="Option3" type="text" style="width:10%;" maxlength="50" value="<%=Des3%>"/><br/>
  Option4<span style="color:Red">*</span> : <input name="Option4" type="text" style="width:10%;" maxlength="50" value="<%=Des4%>"/>
 </p>

 <p class="config" style="text-align: center">
  <span style="color:Red">*</span> - You do not need to fill in <b>all</b> of these.<br/>
  <input type="submit" value="Save"/>
 </p>

</form>
<% Else

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
 Response.Write "<p style=""text-align: center"">No Question Entered</p>" & VbCrlf
 Response.Write "<p style=""text-align: center""><a href=""javascript:history.back()"">Back</a></p>" & VbCrlf
 Response.Write "</div>" & VbCrlf
 %>
 <!-- #INCLUDE FILE="../Includes/Footer.asp" -->
 <%
 Response.End
End If

'### Open The Records Ready To Write ###
Records.Open "SELECT PollID, Content, Des1, Des2, Des3, Des4 FROM Poll ORDER By PollID DESC", Database, 0, 2

Records("Content") = Left(Request.Form("Content"),80)

Records("Des1") = Request.Form("Option1")
Records("Des2") = Request.Form("Option2")
Records("Des3") = Request.Form("Option3")
Records("Des4") = Request.Form("Option4")

Records.Update

'#### Close Objects ###
Records.Close

Response.Write "<p style=""text-align: center"">Poll Submission Successful</p>"
Response.Write "<p style=""text-align: center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
<%
' --------------------------------------------------------------------------
'¦Introduction : Edit Draft Page.                                           ¦
'¦Purpose      : Allows Blog administrator to edit a draft post.            ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/RTF.js.                      ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
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
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script type="text/javascript" src="../Includes/RTF.js"></script>
<div id="content">
<h1>Preview/Publish</h1>
<%
If Request.Form("Action") = "Post" Then

 '-- Did We Type In Text? --'
 If Request.Form("Content") = "" Then
  Response.Write "<p style=""text-align:Center"">No text entered.</p>"
  Response.Write "<p style=""text-align:Center""><a href=""javascript:history.back()"">Back</a></p>"
  Response.Write "</div>"
 %>
 <!-- #INCLUDE FILE="../Includes/Footer.asp" -->
 <%
  Response.End
 End If

 '-- Open The Records Ready To Write --'
 Records.CursorType = 2
 Records.LockType = 3

 Records.Open "SELECT Title, Text FROM Draft", Database

  If NOT Records.EOF Then
   Records("Title") = Left(Request.Form("Title"),80)
   Records("Text") = Request.Form("Content")
   Records.Update
   Response.Write "<p style=""text-align:Center""><b>Saved To Draft NotePad<br/>(" & Now() & ")</b></p>"
  Else
   Response.Write "<h1 style=""text-align:Center""><b>Drafts are unavailable.</b></h1>"
  End If

 Records.Close

End If 

'-- Open set --'
Records.Open "SELECT Title, Text FROM Draft",Database, 1, 3

 If NOT Records.EOF Then

  '-- Setup Variables --'
  Dim Title, Text
  Title    = Records("Title")
  Text     = Records("Text")

 End If

Records.Close
%>

<!-- Start Content For Draft -->
<div class="entry">
 <h3 class="entryTitle"><%=Title%> (Preview Entry)</h3>
 <div class="entryBody"><%If Len(Text) > 0 Then Response.Write LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></div>
 <p class="entryFooter">
  <% If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""../Mail.asp""><img alt=""Email The Author"" src=""../Images/Email.gif"" style=""border-style: none""/></a></acronym>"%>
  <b><%=Now()%></b>
  <% If EnableComments <> False Then Response.Write " | <span class=""comments"">Comments [0]</span>"%>
 </p>
</div>
<!-- End Content -->

<form method="get" action="AddEntry.asp">
 <p style="text-align:center">
  <input name="Import" type="hidden" value="FromDraft"/>
  <input value="Publish To Blog" type="submit"/>
 </p>
 </form>

<hr/>

<h1>Edit Draft</h1>

<form id="AddEntry" action="EditDraft.asp" method="post" onsubmit="return setVar()">
 <p><input name="Action" type="hidden" value="Post"/>
 Title : <input name="Title" type="text" value="<%=Replace(Title,"""","&quot;")%>" style="width:80%;" onchange="return setVarChange()" maxlength="80"/></p>

 <p>Content :<br/></p>

 <table border="0" cellpadding="1" cellspacing="0" width="100%">
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
    <textarea name="Content" cols="141" rows="10" style="height:15em;width:99%;" onchange="return setVarChange()"><%If Len(Text) > 0 Then Response.Write Replace(Replace(Text,"&","&amp;"),"<","&lt;")%></textarea>
   </td>
  </tr>
  <% If LegacyMode <> True Then %>
  <tr>
   <td style="background-color: <%=CalendarBackground%>" align="right" colspan="2">
    <%
    '-- Write Them In ---'
    Response.Write "&nbsp;<select id=""AdvancedTools"" name=""SelectCategory"" onchange=""AdvancedThis(this[this.selectedIndex].value);this.selectedIndex = 0;"">" & VbCrlf
    Response.Write "    <option value="""">-- Tools --</option>" & VbCrlf
    Response.Write "    <option value=""!-- PictureViewer Snippit 1.0 (Edit the FOLDERNAME to the folder inside Images\Articles) -->" & VbCrlf & "<iframe src=&quot;PictureViewer.asp?Folder=FOLDERNAME&quot; width=&quot;99%&quot;></iframe>" & VbCrlf & "<!-- End Of BlogX Snippit --"">Photo Album</option>" & VbCrlf
    Response.Write "    <option value=""!-- Quote Snippit 1.0 (Edit the TEXT to what your quoting) -->" & VbCrlf & "<div style=&quot;background:white; color:blue;&quot;>TEXT</div>" & VbCrlf & "<!-- End Of BlogX Snippit --"">Quote</option>" & VbCrlf
    Response.Write "   </select>"
    %>	
   </td>
  </tr>
  <% End If %>
 </table>

 <p style="text-align:center"><input type="submit" value="Save Draft"/></p>
            
</form>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
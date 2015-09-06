<%
' --------------------------------------------------------------------------
'¦Introduction : Edit Entry Page.                                           ¦
'¦Purpose      : Allows Blog administrator to edit an entry.                ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/RTF.js.                      ¦
'¦Standards    : Almost XHTML (javascript id element admin needs updating). ¦
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
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script type="text/javascript" src="../Includes/RTF.js"></script>

<div id="content">
<%
'--- Querish Querystring ---'
Dim Requested, Delete
Requested = Request.Querystring("Entry")
Delete = Request.Querystring("Delete")

If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then Requested = 0

If Request.Form("Action") <> "Post" Then

 '--- Open set ---'
 If (Requested <> 0) AND (Delete <> "True") Then 
  Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, StopComments, Enclosure FROM Data WHERE RecordID=" & Requested,Database, 0, 1
 ElseIf Requested = 0 Then
  Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, StopComments, Enclosure FROM Data ORDER By RecordID DESC",Database, 0, 1
 Else
  Database.Execute "DELETE FROM Data WHERE RecordID=" & Requested
  Database.Close
  Set Records = Nothing
  Set Database = Nothing
  Response.Redirect(SiteURL & PageName) 
 End If

 If NOT Records.EOF Then

  '--- Setup Variables ---'
  Dim Title, Text, Password, sDay, sMonth, sYear, TimePosted, Enclosure

  Title = Records("Title")
  Text = Records("Text")

  Category =  Records("Category")
  If IsNull(Category) Then Category = ""

  Password =  Records("Password")

  sDay = Records("Day")
  sMonth = Records("Month")
  sYear = Records("Year")

  TimePosted = Records("Time")

  StopComments = Records("StopComments")

  Enclosure = Records("Enclosure")
 End If

Records.Close
%>
<form id="AddEntry" action="EditEntry.asp?Entry=<%=Requested%>" method="post" onsubmit="return setVar()">
 <p><input name="Action" type="hidden" value="Post"/>
 Title : <input name="Title" type="text" style="width:80%;" maxlength="80" value="<%=Replace(Replace(Title,"""","&quot;"),"&","&amp;")%>"/>  <a href="?Entry=<%=Requested%>&amp;Delete=True" title="DELETE this entry" onclick="return confirm('Warning! If You Continue Entry #<%=Requested%> Will Be DELETED.')"><img alt="DELETE This Entry" src="<%=SiteURL%>Images/Delete.gif" width="15" height="15" style="border-style:none"/></a></p>

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
    <textarea name="Content" cols="141" rows="10" style="height:40em;width:99%;" onchange="return setVarChange()"><%=Replace(Replace(Text,"&","&amp;"),"<","&lt;")%></textarea>
   </td>
  </tr>
  <% If LegacyMode <> True Then %>
  <tr>
   <td style="background-color: <%=CalendarBackground%>" align="right" colspan="2">
    <a href="#" onclick="javascript:show('AdvancedTools'); return false;"><img alt="Turn on Advanced editing features" style="border: none" src="<%=SiteURL%>Images/Editor/Advanced.gif" width="61" height="16" id="AdvancedButton"/></a>
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

  <% If ShowCategories <> False Then Response.Write "<p>Category : <input name=""Category"" type=""text"" value=""" & Replace(Category,"%20"," ") & """ maxlength=""50"" style=""width:10%;"" onchange=""return setVarChange()""/></p>" & VbCrlf

  If LegacyMode <> True Then 

   Response.Write "<p id=""AdvancedTools"">Change the entry date, Day : "
   Response.Write "<select name=""nDay"">" & VbCrlf

   Dim i		
   For i = 1 To 31
    If i=sDay Then
     Response.Write "<option selected=""selected"" value=""" & i & """>" & i & "</option>" & VbCrlf
    Else
     Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
    End If
   Next

   Response.Write "</select>"

   Response.Write " Month : "
   Response.Write "<select name=""nMonth"">" & VbCrlf
			
   For i = 1 To 12

    If i=sMonth Then
     Response.Write "<option selected=""selected"" value=""" & i & """>" & i & "</option>" & VbCrlf
    Else
     Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
    End If

   Next

   Response.Write "</select>"
   Response.Write " Year : "
   Response.Write "<select name=""nYear"">" & VbCrlf
			
    For i = 2000 to 2030

     If i=sYear Then
      Response.Write "<option selected=""selected"" value=""" & i & """>" & i & "</option>" & VbCrlf
     Else
      Response.Write "<option value=""" & i & """>" & i & "</option>" & VbCrlf
     End If

    Next

    Response.Write "</select>"

    Response.Write "<br/><br/>Change the entry time, Time : <input name=""Time"" value=""" & TimePosted & """ style=""width:10%;""/>"

    Response.Write "<br/><br/>Stop comments on this entry : <input name=""StopComments"" value=""True"" type=""checkbox"""
    If StopComments = True Then Response.Write "checked=""true"""
    Response.Write "/>"

    Response.Write "<br/><br/>There is an associated audio file and its audio is at : <input name=""Enclosure"" value=""" & Enclosure & """ style=""width:40%;""/><br/><br/>"
  %>
  </p>


   <table border="0" cellpadding="0" cellspacing="0" width="30%" id="AdvancedTools">
    <tr style="background-color:<%=CalendarBackground%>; color:white">
     <td align="left">
      <acronym title="If you type in a password, your viewers will need to enter it to view the entry, leaving it blank means everyone can see your entry."><img style="border: none" src="<%=SiteURL%>Images/Help.gif" alt="Help"/>Optional<br/>Entry Password</acronym>
     </td>
     <td align="center">
      <input name="password" type="text" maxlength="10" value="<%=Password%>"/>
     </td>
    </tr>
   </table>

  <% End If %>

  <p><input type="submit" value="Save"/></p>

  </form>
<% Else

 Dim EntryCat 'Category
 EntryCat = Request.Form("Category")

 '-- Filter & Clean --'
 EntryCat = Replace(EntryCat,"'","&#39;")
 EntryCat = Replace(EntryCat," ","%20")

 Dim StopComments
 StopComments = Request.Form("StopComments")
 If StopComments = "" Then StopComments = False

 '-- Did we type in text? --'
 If Request.Form("Content") = "" Then
  Response.Write "<p align=""Center"">No Text Entered</p>"
  Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
  Response.Write "</div>"
  %>
  <!-- #INCLUDE FILE="../Includes/Footer.asp" -->
  <%
  Response.End
 End If

 '-- Open the records ready to write --'
 If Requested <> 0 Then
  Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, StopComments, Enclosure, LastModified FROM Data WHERE RecordID=" & Requested,Database, 0, 2
 Else
  Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, StopComments, Enclosure, LastModified FROM Data ORDER By RecordID DESC", Database, 0, 2
 End If

  Records("Title") = Request.Form("Title")
  Records("Text") = Request.Form("Content")
  Records("Password") = Request.Form("Password")
  Records("Category") = EntryCat

  '-- Update the last modified date because our cache code relies on this --'
  Records("LastModified") = Now()

  If IsNumeric(Request.Form("nDay")) AND Request.Form("nDay") <> ""  Then Records("Day") = Request.Form("nDay") Else Records("Day") = Records("Day")
  If IsNumeric(Request.Form("nMonth")) AND Request.Form("nMonth") <> "" Then Records("Month") = Request.Form("nMonth") Else Records("Month") = Records("Month")
  If IsNumeric(Request.Form("nYear")) AND Request.Form("nYear") <> "" Then Records("Year") = Request.Form("nYear") Else Records("Year") = Records("Year")

  If IsDate(Request.Form("Time")) AND Request.Form("Time") <> "" Then Records("Time") = Request.Form("Time") Else Records("Time") = Records("Time")

  Records("StopComments") = StopComments

  Records("Enclosure") = Request.Form("Enclosure")

  Records.Update
 Records.Close

Response.Write "<p style=""text-align:Center"">Entry update successful.</p>"
Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & "Comments.asp?Entry=" & Requested & """>Back</a></p>"

End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
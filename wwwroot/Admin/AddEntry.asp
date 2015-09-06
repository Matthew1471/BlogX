<%
' --------------------------------------------------------------------------
'¦Introduction : Add Entry Page.                                            ¦
'¦Purpose      : Allows blog administrator to submit an entry.              ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/XMLRPC.asp,       ¦
'¦               Includes/NAV.asp, Includes/Footer.asp, Includes/RTF.js.    ¦
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

PingbackPage = True
AlertBack = True
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/XMLRPC.asp" -->
<script type="text/javascript" src="../Includes/RTF.js"></script>
 
<div id="content">
<%
 If Request.Form("Action") <> "Post" Then 
   
   '-- Import from drafts if requested --'
   If Request.Querystring("Import") = "FromDraft" Then

    '-- Open set --'
    Records.Open "SELECT Title, Text FROM Draft",Database, 1, 3

    If NOT Records.EOF Then
     '-- Setup Variables --'
     Dim ImportedTitle, ImportedText
     ImportedTitle = " value=""" & Replace(Replace(Records("Title"),"&","&amp;"),"<","&lt;") & """"
     ImportedText  = Replace(Replace(Records("Text"),"&","&amp;"),"<","&lt;")
    End If

    Records.Close

   End If
 %>
<form id="AddEntry" action="AddEntry.asp" method="post" onsubmit="return setVar()">
 <p><input name="Action" type="hidden" value="Post"/>
 Title : <input name="Title" type="text" style="width:80%;" maxlength="80"<%=ImportedTitle%>/></p>

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
    <textarea name="Content" cols="141" rows="10" style="height:15em;width:99%;" onchange="return setVarChange()"><%=ImportedText%></textarea>
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

  <p>

  <%  
  If (ShowCategories <> False) AND (LegacyMode <> True) Then

   '--- Open set ---'
   Records.CursorLocation = 3 ' adUseClient
   Records.Open "SELECT DISTINCT Category FROM Data ORDER BY Category",Database, 1, 3

   '--- Set Category ---'
   Set Category = Records("Category")

   '-- Write Them In ---'
   Response.Write "Select an existing category : " & VbCrlf
   Response.Write "  <select name=""SelectCategory"" onchange=""document.forms['AddEntry'].Category.value = this[this.selectedIndex].value; "">" & VbCrlf
   Response.Write "   <option value="""">-- New --</option>" & VbCrlf

   Do Until (Records.EOF or Records.BOF)
    If (Category <> "") Then Response.Write "   <option value=""" & Replace(Category, "%20", " ") & """>" & Replace(Category, "%20", " ") & "</option>" & VbCrlf
    Records.MoveNext
   Loop
                
   Response.Write "  </select>" & VbCrlf

   '-- Close The Database & Records ---'
   Records.Close

   Response.Write "  or create/edit the selected Category : <input name=""Category"" type=""text"" style=""width:10%;"" maxlength=""50""/></p>"

  ElseIf (ShowCategories <> False) Then 
   Response.Write "<p>Category : <input name=""Category"" type=""text"" style=""width:10%;"" maxlength=""50""/></p>"
  End If 

  If (LegacyMode <> True) Then %>

   <table border="0" cellpadding="0" cellspacing="0" width="30%" id="AdvancedTools">
    <tr style="background-color:<%=CalendarBackground%>; color:white">
     <td align="left">
      <acronym title="If you type in a password, your viewers will need to enter it to view the entry, leaving it blank means everyone can see your entry."><img style="border: none" src="<%=SiteURL%>Images/Help.gif" alt="Help"/>Optional<br/>Entry Password</acronym>
     </td>
     <td align="center">
      <input name="password" type="text" maxlength="10"/>
     </td>
    </tr>
   </table>

   <% Response.Write "<p id=""AdvancedTools""><br/><br/>There is an associated audio file and its audio is at : <input name=""Enclosure"" style=""width:40%;""/></p>"
  End If %>

  <p id="AdvancedTools"><br/><br/><span style="color: red">Note :</span> You can drag the following link : <a title="BlogIt!" href="javascript:Q='';x=document;y=window;if(x.selection){Q=x.selection.createRange().text;}else if(y.getSelection){Q=y.getSelection();}else if(x.getSelection){Q=x.getSelection();}void(window.open('<%=Replace(SiteURL,"'","\'")%>Admin/Toolbar.asp?t='+escape(Q)+'&amp;u='+escape(location.href)+'&amp;n='+escape(document.title),'bloggerForm','scrollbars=no,width=475,height=300,top=175,left=75,status=yes,resizable=yes'));">BlogIt!</a> to your links bar or add it to your favourites and when you click it, it will open up a window with information (including any highlighted text) and the link to the site you are currently browsing so you can post about it.</p>

  <p><input type="submit" value="Save"/></p>
  
  </form>

<% Else
 '-- Dimension variables --'
 Dim EntryCat            'Category
 EntryCat = Request.Form("Category")

 '-- Filter & Clean --'
 EntryCat = Replace(EntryCat,"'","&#39;")
 EntryCat = Replace(EntryCat," ","%20")

 '-- Did We Type In Text? --'
 If Request.Form("Content") = "" Then
  Response.Write "<p style=""text-align:Center"">No Text Entered</p>"
  Response.Write "<p style=""text-align:Center""><a href=""javascript:history.back()"">Back</a></p>"
  Response.Write "</div>"
  %>
  <!-- #INCLUDE FILE="../Includes/Footer.asp" -->
  <%
  Response.End
 End If

 '-- Open The Records Ready To Write --'
 Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Enclosure, EntryPUK FROM Data", Database, 0, 2
  Records.AddNew

  Records("Title") = Left(Request.Form("Title"),80)
  Records("Text") = Request.Form("Content")
  Records("Category") = EntryCat
  Records("Password") = Request.Form("Password")

  Records("Day") = Day(DateAdd("n",ServerTimeOffset,Now()))
  Records("Month") = Month(DateAdd("n",ServerTimeOffset,Now()))
  Records("Year") = Year(DateAdd("n",ServerTimeOffset,Now()))
  Records("Time") = TimeValue(DateAdd("n",ServerTimeOffset,Time()))

   '-- Work out Time Offset --'
   Dim Hours
   Hours = Abs(Int(UTCTimeZoneOffset / 60))
   If Hours < 10 Then Hours = "0" & Hours

   Dim Minutes
   Minutes = Abs(UTCTimeZoneOffset Mod 60)
   If Minutes < 10 Then Minutes = "0" & Minutes

   Dim OffsetTime
   If UTCTimeZoneOffset < 0 Then OffsetTime = "+" Else OffsetTime = "-"
   OffsetTime = OffsetTime & Hours & Minutes

  Records("UTCTimeZoneOffset") = OffsetTime

  Records("Enclosure") = Request.Form("Enclosure")

   Randomize Timer
   Records("EntryPUK") = Int(Rnd()*99999999)

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

 Response.Write "<p style=""text-align:Center"">Entry Submission Successful</p>"
 Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
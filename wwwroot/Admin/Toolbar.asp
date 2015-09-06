<%
' --------------------------------------------------------------------------
'¦Introduction : Toolbar Entry Post Popup Page.                             ¦
'¦Purpose      : A mini entry submitter to be used on a browsers' links bar.¦
'¦Used By      : AddEntry.asp.                                              ¦
'¦Requires     : Includes/Config.asp, Admin.asp, Includes/XMLRPC.asp,       ¦
'¦               Includes/RTF.js, Templates/Config.asp.                     ¦
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
PingbackPage = True
AlertBack = True
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/xmlrpc.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
 <title><%=SiteDescription%> - Toolbar Submission</title>
 <meta http-equiv="Content-Type" content="text/html; CHARSET=utf-8"/>
 <!--
 //= - - - - - - - 
 // Copyright 2008, Matthew Roberts
 // Copyright 2003, Chris Anderson
 //= - - - - - - -
 -->
 <% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
 <link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet"/>
 <!-- #INCLUDE FILE="../Templates/Config.asp" -->
 <script type="text/javascript" src="../Includes/RTF.js"></script>
</head>

<body style="background-color:<%=BackgroundColor%>">
<% If Request.Form("Action") <> "Post" Then %>
<form id="AddEntry" method="post" action="Toolbar.asp" onsubmit="return setVar()">
 <p style="text-align:center">
  <input name="Action" type="hidden" value="Post"/>
  Title : <input name="Title" type="text" style="width:80%;" maxlength="80" value="<%=Request.Querystring("n") & " (" & Request.Querystring("u") & ")"%>"/>
 </p>

 <p style="text-align:center">Content :</p>

 <table border="0" cellpadding="1" cellspacing="0" width="80%" style="align:center; margin: 0 auto">
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
    <textarea name="Content" rows="4" cols="67" style="height:6em;width:99%;" onchange="return setVarChange()"><%If Request.Querystring("t") <> "" Then Response.Write """" & Request.Querystring("t") & """"%></textarea>
   </td>
  </tr>
 </table>

 <%
 If (ShowCategories <> False) AND (LegacyMode <> True) Then
  '-- Open set --'
  Records.Open "SELECT DISTINCT Category FROM Data ORDER BY Category",Database, 0, 1

  '-- Set Category --'
  Dim Category
  Set Category = Records("Category")

  '-- Write them in --'
  Response.Write "<p style=""text-align:center"">Select an existing category : " & VbCrlf
  Response.Write "<select id=""SelectCategory"" onchange=""document.forms['AddEntry'].Category.value = this[this.selectedIndex].value;"">" & VbCrlf
  Response.Write "<option value="""">-- New --</option>" & VbCrlf

  Do Until (Records.EOF or Records.BOF)
   If Category <> "" Then Response.Write "<option value=""" & Replace(Category, "%20", " ") & """>" & Replace(Category, "%20", " ") & "</option>" & VbCrlf
   Records.MoveNext
  Loop

  Response.Write "</select>"

  '-- Close The Database & Records --'
  Records.Close

  Response.Write "<br/>or create/edit the selected Category : <input name=""Category"" type=""text"" style=""width:20%;"" maxlength=""50""/></p>"

 ElseIf ShowCategories <> False Then 
  Response.Write "<p style=""text-align:center"">Category : <input name=""Category"" type=""text"" style=""width:20%;"" maxlength=""50""/></p>"
 End If
  %>
  <p style="text-align:center"><input type="submit" value="Save"/></p>
 </form>
<%
Else

 Dim EntryCat
 EntryCat = Request.Form("Category")

 '-- Filter & Clean --'
 EntryCat = Replace(EntryCat,"'","&#39;")
 EntryCat = Replace(EntryCat," ","%20")

 '-- Did we type in text? --'
 If Request.Form("Content") = "" Then
  Response.Write "<p style=""text-align:center"">No Text Entered</p>"
  Response.Write "<p style=""text-align:center""><a href=""javascript:history.back()"">Back</a></p>"
  %><!-- #INCLUDE FILE="../Includes/Footer.asp" --><%
  Response.End
 End If

 '-- Open The Records Ready To Write --'
 Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, EntryPUK FROM Data", Database, 0, 2
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

   Randomize Timer
   Records("EntryPUK") = Int(Rnd()*99999999)

  Records.Update

  Records.MoveLast

  Dim RecordID
  RecordID = Records("RecordID")

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

 Response.Write "<p style=""text-align:center"">Entry Submission Successful</p>"
 Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

End If

Database.Close
Set Database = Nothing
Set Records = Nothing
%>
</body>
</html>
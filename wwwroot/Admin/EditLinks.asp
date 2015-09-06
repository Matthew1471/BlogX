<%
' --------------------------------------------------------------------------
'¦Introduction : Edit Links Page.                                           ¦
'¦Purpose      : Allows the administrator to edit the NAV links and buddies.¦
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
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<% 
If AllowEditingLinks <> 0 Then

 Dim DelLink
 DelLink = Request.Querystring("Delete")
 DelLink = Replace(DelLink,"'","")
 If IsNumeric(DelLink) Then DelLink = Int(DelLink) Else DelLink = 0
 If (DelLink <> "") Then Database.Execute "DELETE FROM Links WHERE LinkID=" & DelLink

 If (Request.Form("Action") <> "Post") AND (Request.Querystring("Delete") = "") Then

  '-- Links --'

  '-- Write Preview --'
  Response.Write "<!-- Links -->" & VbCrlf
  Response.Write "<div class=""sidebar"" style=""width:50%; float:left;"">" & VbCrlf
  Response.Write " <div class=""section"">" & VbCrlf
  Response.Write "  <h3 class=""sectionTitle"">Links</h3>" & VbCrlf
  Response.Write "  <ul>" & VbCrlf

  Records.Open "SELECT LinkID, LinkName, LinkURL, LinkRSS, LinkType FROM Links Where LinkType='Main Links' ORDER BY LinkName;", Database

  If Records.EOF Then
   Response.Write "<li>No Links In Database</li>"
  Else

   Dim LinkID, LinkName, LinkURL, LinkRSS, LinkType
   Set LinkID   = Records("LinkID")
   Set LinkName = Records("LinkName")
   Set LinkURL  = Records("LinkURL")
   Set LinkRSS  = Records("LinkRSS")
   Set LinkType = Records("LinkType")

   '-- We are in an Admin folder so relative links will not work --'
   If Left(LinkURL,7) <> "http://" Then LinkURL = SiteURL & LinkURL

   Do Until (Records.EOF)   
    Response.Write "<li><a href=""" & LinkURL & """>" & LinkName & "</a> <acronym title=""Delete Link""><a onclick=""return confirm('Are you *sure* you want to delete this link?');"" href=""?Delete=" & LinkID & """><img alt=""Delete Link"" style=""border:none"" width=""12"" height=""12"" src=""../Images/Delete.gif""/></a></acronym></li>"
    Records.MoveNext
   Loop

  End If

  Records.Close

  Response.Write "  </ul>" & VbCrlf
  Response.Write " </div>" & VbCrlf
  Response.Write "</div>" & VbCrlf

  '-- Other Links --'

  '-- Write Preview --'
  Response.Write "<!-- Other Links -->" & VbCrlf
  Response.Write "<div class=""sidebar"" style=""width:50%; float:right;"">" & VbCrlf
  Response.Write " <div class=""section"">" & VbCrlf
  Response.Write "  <h3 class=""sectionTitle"">Other Links</h3>" & VbCrlf
  Response.Write "  <ul>" & VbCrlf

  Records.Open "SELECT LinkID, LinkName, LinkURL, LinkRSS, LinkType FROM Links Where LinkType='Other Links' ORDER BY LinkName;", Database

  If Records.EOF Then
   Response.Write "<li>No Other Links In Database</li>"
  Else

   Set LinkID   = Records("LinkID")
   Set LinkName = Records("LinkName")
   Set LinkURL  = Records("LinkURL")
   Set LinkRSS  = Records("LinkRSS")
   Set LinkType = Records("LinkType")

   '-- We are in an Admin folder so relative links will not work --'
   If Left(LinkURL,7) <> "http://" Then LinkURL = SiteURL & LinkURL

   Do Until (Records.EOF) 
    Response.Write "<li><a href=""" & LinkURL & """>" & LinkName & "</a> <acronym title=""Delete Link""><a onclick=""return confirm('Are you *sure* you want to delete this link?');"" href=""?Delete=" & LinkID & """><img alt=""Delete Link"" style=""border:none"" width=""12"" height=""12"" src=""../Images/Delete.gif""/></a></acronym></li>" & VbCrlf
    Records.MoveNext
   Loop

  End If

  Records.Close

  Response.Write "  </ul>" & VbCrlf
  Response.Write " </div>" & VbCrlf
  Response.Write "</div>" & VbCrlf

  Response.Write "<p>&nbsp;</p>" & VbCrlf

  '-- RSS Budies --'

  '-- Write Preview --'
  Response.Write "<!-- RSS Buddies Links -->" & VbCrlf
  Response.Write "<div class=""sidebar"" style=""width:100%; float:none;"">" & VbCrlf
  Response.Write "<div class=""section"">" & VbCrlf
  Response.Write "  <h3 class=""sectionTitle"">RSS Buddies</h3>" & VbCrlf
  Response.Write "  <ul>" & VbCrlf

  Records.Open "SELECT LinkID, LinkName, LinkURL, LinkRSS, LinkType FROM Links Where LinkType='RSS' ORDER BY LinkName;", Database

  If Records.EOF Then
   Response.Write "<li>No RSS Buddies In Database</li>"
  Else

   Set LinkID   = Records("LinkID")
   Set LinkName = Records("LinkName")
   Set LinkURL  = Records("LinkURL")
   Set LinkRSS  = Records("LinkRSS")
   Set LinkType = Records("LinkType")

   '-- We are in an Admin folder so relative links will not work --'
   If Left(LinkURL,7) <> "http://" Then LinkURL = SiteURL & LinkURL

   Do Until (Records.EOF) 
    Response.Write "<li><a href=""" & LinkURL & """>" & LinkName & "</a> "
    If LinkRSS <> "" Then Response.Write "(<a href=""" & LinkRSS & """>RSS</a>) "
    Response.Write "<acronym title=""Delete Link""><a onclick=""return confirm('Are you *sure* you want to delete this link?');"" href=""?Delete=" & LinkID & """><img alt=""Delete Link"" style=""border:none"" width=""12"" height=""12"" src=""../Images/Delete.gif""/></a></acronym></li>" & VbCrlf
    Records.MoveNext
   Loop

  End If
  Records.Close

  Response.Write "  </ul>" & VbCrlf
  Response.Write " </div>" & VbCrlf
  Response.Write "</div>" & VbCrlf
 %>

 <!-- Add New Link -->
 <div id="AddNew" class="date">
  <div class="comment">
    <h3 class="commentTitle">Add New Link</h3>
     <div class="commentBody">
      <form method="post" action="EditLinks.asp" onsubmit="return setVar()">
      <p>
       <input name="Action" type="hidden" value="Post"/>
       Link Name : <input name="LinkName" type="text" maxlength="80" onchange="return setVarChange()"/>
      </p>
      <p>
       Link URL : <input name="LinkURL" type="text" onchange="return setVarChange()"/>
      </p>
      <p>
       Link RSS (Optional) : <input name="LinkRSS" type="text" onchange="return setVarChange()"/>
      </p>
      <p>
       Link Type :


      <select name="LinkType" onchange="return setVarChange()">
       <option value="Main Links">Main Links</option>
       <option value="Other Links">Other Links</option>
       <option value="RSS">RSS Buddies</option>
      </select>
      </p>
      <p><input type="submit" value="Add Link" accesskey="a"/></p>
      </form>

     </div>
  </div>
 </div>
<%
ElseIf (Request.Form("Action") = "Post") Then

 '-- Did We Type In Text? --'
 If (Request.Form("LinkName") = "") OR (Request.Form("LinkURL") = "") OR (Request.Form("LinkType") = "")  Then
  Response.Write "<p style=""text-align:Center"">Missing link name or link URL.</p>"
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
 Records.Open "SELECT LinkName, LinkURL, LinkRSS, LinkType FROM Links", Database
 Records.AddNew
  Records("LinkName") = Left(Request.Form("LinkName"),80)
  Records("LinkURL") = Request.Form("LinkURL")
  If Len(Request.Form("LinkRSS")) > 0 Then Records("LinkRSS") = Request.Form("LinkRSS")
  Records("LinkType") = Request.Form("LinkType")
  Records.Update
 Records.Close

 Response.Write "<p style=""text-align:Center"">Link added.</p>"
 Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & "Admin/EditLinks.asp"">Back</a></p>"

 ElseIf (DelLink <> 0) Then
  Response.Write "<p style=""text-align:Center"">Link deleted.</p>"
  Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & "Admin/EditLinks.asp"">Back</a></p>"
 End If 

Else
 Response.Write "<p style=""text-align:Center"">You are not allowed to edit links.</p>"
 Response.Write "<p style=""text-align:Center""><a href=""" & SiteURL & PageName & """>Back</a></p>"
End If
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
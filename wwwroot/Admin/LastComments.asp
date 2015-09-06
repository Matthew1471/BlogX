<%
' --------------------------------------------------------------------------
'¦Introduction : Last Comments Display Page.                                ¦
'¦Purpose      : Provides a quick list of recent validated comments.        ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/Replace.asp.                 ¦
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
%>
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<%
'-- There should be a header STRAIGHT away --'
Response.Flush

 '### Open The Records Ready To Write ###
 Records.Open "SELECT CommentID, EntryID, PUK, Name, Email, Homepage, Content, CommentedDate, IP FROM Comments ORDER BY CommentID DESC",Database, 1, 1

  Dim EntryID, CommentID, Email, Homepage, Content, CommentedDate, IP

  Set Content = Records("Content")
  Set CommentedDate = Records("CommentedDate")
  Set EntryID = Records("EntryID")

  Set CommentID = Records("CommentID")
  Set Homepage = Records("Homepage")
  Set IP = Records("IP")

  '-- Split records in to pages --'
  Records.PageSize = EntriesPerPage

  If NOT Records.EOF Then 
   Response.Write "<p style=""text-align:Center"">The following are a list of the last " & EntriesPerPage & " (where applicable) validated comments:</p>" & VbCrlf
   
    Do Until (Records.EOF) OR (Records.AbsolutePage <> 1)

    '-- These have already been used before.. and VB has now decided that they are strings --'
    Set Name = Records("Name")
    Set Email = Records("Email")

    '--- We're British, Let's 12Hour Clock Ourselves ---'
    Dim NewTime, NewDate, CommentedTime

    CommentedTime = FormatDateTime(CommentedDate,vbLongTime)

    NewTime = ""

    If TimeFormat <> False Then

     If Hour(CommentedTime) > 12 Then 
      NewTime = Hour(CommentedTime) - 12 & ":"
     Else
      NewTime = Hour(CommentedTime) & ":"
     End If
 
    If Minute(CommentedTime) < 10 Then
     NewTime = NewTime & "0" & Minute(CommentedTime)
    Else
     NewTime = NewTime & Minute(CommentedTime)
    End If

    If (Hour(CommentedTime) < 12) AND (Hour(CommentedTime) <> 12) Then
     NewTime = NewTime & " AM"
    Else
     NewTime = NewTime & " PM"
    End If

   Else
    If Hour(CommentedTime) < 10 Then NewTime = "0"
    NewTime = NewTime & Hour(CommentedTime) & ":"
    If Minute(CommentedTime) < 10 Then NewTime = NewTime & "0"
    NewTime = NewTime & Minute(CommentedTime)
   End If

   '-- Turn HTML proxy comment into a variable --'
   Dim ProxyPosition 
   ProxyPosition = Instr(Content,"<!-- Proxy Servers ORIGINAL Address : ")

   If ProxyPosition <> 0 Then
    Dim ProxyAddress
    ProxyAddress = Mid(Content,ProxyPosition+38)
    ProxyAddress = Left(ProxyAddress,Len(ProxyAddress)-3)
    'Response.Write "<!-- Debug " & ProxyPosition & ": """ & ProxyAddress & """ -->"
   Else
    ProxyAddress = ""
   End If

   NewDate = Day(CommentedDate) & "/" & Month(CommentedDate) & "/" & Year(CommentedDate) 
   Response.Write "EntryID : " & EntryID & "<br/>"
 %>
  <!--- Start Content For Comment <%=CommentID%> -->
  <div class="comment">
   <h3 class="commentTitle">
  <%
    Response.Write " <acronym title=""Users Using This IP""><a href=""#"" onclick=""javascript:PrintPopup('" & SiteURL & "IPWhois.asp?IP=" & IP & "');""><img alt=""Users Using This IP"" src=""" & SiteURL & "Images/Emoticons/Profile.gif"" style=""border: none""/></a></acronym> "
    If ProxyAddress <> "" Then Response.Write "<acronym title=""List User's Proxy Information""><a href=""http://whois.domaintools.com/" & IP & """><img alt=""List User's Proxy Information"" src=""" & SiteURL & "/Images/Print.gif"" style=""border: none""/></a></acronym> "
    Response.Write "<acronym title=""Ban User""><a "
    If ProxyAddress <> "" Then Response.Write "onclick=""return confirm('Are you *sure* you want to ban this user?\n\nThis user was behind a proxy. Check the address is creditable before banning.');"" "
    Response.Write "href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & "&amp;Ban=" & IP & "#Comment" & CommentID & """><img title=""Ban User"" alt=""Color Icon"" src=""" & SiteURL & "Images/Color.gif"" style=""border: none""/></a></acronym> "
    If ProxyAddress <> "" Then Response.Write "<acronym title=""Ban User's Proxy""><a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & "&amp;Ban=" & ProxyAddress & "#Comment" & CommentID & """ onclick=""return confirm('Are you *sure* you want to ban all users behind this proxy?\n\nProxies are sometimes used by big internet providers, other times they can be created by spammers.');""><img title=""Ban User's Proxy"" alt=""Color Icon"" src=""" & SiteURL & "Images/Color.gif"" style=""border: none""/></a></acronym> "
    Response.Write "<acronym title=""Delete Comment""><a onclick=""return confirm('Are you *sure* you want to delete this comment?');"" href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & "&amp;Delete=" & CommentID & """><img title=""Delete Comment"" alt=""Key Icon"" src=""" & SiteURL & "Images/Key.gif"" style=""border: none""/></a></acronym>"
    %>
    <%=NewDate%>&nbsp;<%=NewTime%>
   </h3>

   <span class="commentBody"><%=LinkURLs(Replace(Content, vbcrlf, "<br/>" & vbcrlf))%></span>

   <p class="commentFooter"><%If HomePage <> "" Then Response.Write "<a class=""permalink"" rel=""nofollow"" href=""" & HTML2Text(Homepage) & """>"%><%=HTML2Text(Name)%><%If HomePage <> "" Then Response.Write "</a>"%>
   <%If (Email <> "") Then Response.Write " | <span class=""comments""><a href=""mailto:" & HTML2Text(Email) & """>" & Email & "</a></span>"%></p>
  </div>
  <!--- End Content -->
  <%
   Response.Flush

   Records.MoveNext
  Loop
  
 Else
  Response.Write "<p style=""text-align:center"">There are no comments in the database.</p>" & VbCrlf
 End If

Records.Close

Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
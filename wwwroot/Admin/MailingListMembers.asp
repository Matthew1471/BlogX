<%
' --------------------------------------------------------------------------
'¦Introduction : Admin Mailing List Management Page.                        ¦
'¦Purpose      : Provides a way to send messages and manage subscriptions.  ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/RTF.js, Includes/Mail.asp.   ¦
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
AlertBack = True
Server.ScriptTimeout = 6000 %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<script type="text/javascript" src="../Includes/RTF.js"></script>
<div id="content">
<%
 Records.Open "SELECT SubscriberAddress, PUK, SubscriberIP, Active FROM MailingList ORDER BY Active ASC, SubscriberAddress ASC;", Database
 
  Dim SubscriberAddress, PUK, IP, Active
  Set SubscriberAddress = Records("SubscriberAddress")
  Set PUK = Records("PUK")
  Set IP = Records("SubscriberIP")
  Set Active = Records("Active")

  '------------------------------------------------------------------------------------------------------------------------
  If Request.Form("Action") <> "Send" Then %>
 <form id="AddEntry" method="post" action="MailingListMembers.asp" onsubmit="return setVar()">
  
 <p>
  <input name="Action" type="hidden" value="Send"/>
  Subject : <input name="Title" type="text" style="width:80%;" maxlength="80" onchange="return setVarChange()"/>
 </p>

 <p>Content :</p>
 
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
    <textarea name="Content" cols="141" rows="10" style="height:15em;width:99%;" onchange="return setVarChange()"></textarea>
   </td>
  </tr>
  </table>
  <p><input type="submit" value="Send"/></p>
 </form>

<table border="0">
<tr>
 <td>
  <b>Mailing List Members</b>
 </td>
</tr>
<%
 Count = 0

 Do Until (Records.EOF)
  Count = Count + 1
  Response.Write "<tr><td>" & VbCrlf
  If Active = False Then Response.Write "<span style=""text-decoration:line-through"">" & VbCrlf
  Response.Write "<acronym title=""" & IP & """><a href=""mailto:" & SubscriberAddress & """>" & SubscriberAddress & "</a></acronym>" & VbCrlf
  If Active = False Then Response.Write "</span>" & VbCrlf
  If Active = True Then Response.Write "<a class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & SubscriberAddress & "&amp;PUK=" & PUK & """>Unsubscribe</a>"
  Response.Write "</td></tr>" & VbCrlf
  Records.MoveNext
 Loop
%>
<tr>
 <td>
  <b>Total</b> : <%=Count%> Members
 </td>
</tr>
</table>
<% 
'------------------------------------------------------------------------------------------------------------------------
Else

'-- Did we type in text? --'
If Request.Form("Content") = "" Then
 Response.Write "<p align=""Center"">No Text Entered</p>"
 Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
 Response.Write "</div>"
 %><!-- #INCLUDE FILE="../Includes/Footer.asp" --><%
 Response.End
End If

Response.Write "<p align=""center"">"

On Error Resume Next

 If Len(Request.Form("Title")) > 0 Then Subject = "Blog : " & Request.Form("Title") Else Subject = "Blog : " & Now()

 Do Until (Records.EOF)

 Dim MailBody, ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg, ReplyTo
 
 If Active = True Then
  MailBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf
  MailBody = MailBody & "<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">" & VbCrlf
  MailBody = MailBody & "<head>" & VbCrlf
 
  If TemplateURL = "" Then
   MailBody = MailBody & " <link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
  Else
   MailBody = MailBody & " <link href=""" & TemplateURL & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf
  End If
 
  MailBody = MailBody & " <title>" & Subject & "</title>" & VbCrlf
 
  MailBody = MailBody & "</head>" & VbCrlf
  MailBody = MailBody & "<body style=""background-color:" & BackgroundColor & """>" & VbCrlf

  MailBody = MailBody & " <div class=""content"" style=""text-align: center"">" & VbCrlf
  MailBody = MailBody & "  <br/>" & VbCrlf
  MailBody = MailBody & "  <div class=""entry"" style=""width: 50%; margin:0 auto"">" & VbCrlf
  MailBody = MailBody & "  <h3 class=""entryTitle"">" & Subject & "</h3>" & VbCrlf
  MailBody = MailBody & "  <div class=""entryBody"" style=""text-align: left"">" & VbCrlf

  MailBody = MailBody & "  " & Replace(Request.Form("Content"), vbcrlf, "<br/>" & vbcrlf & "  ") & VbCrlf
  MailBody = MailBody & "  </div>" & VbCrlf
  MailBody = MailBody & "  </div>" & VbCrlf

  MailBody = MailBody & "  <p>To Unsubscribe click <a class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & SubscriberAddress & "&amp;PUK=" & PUK & """>Unsubscribe</a></p>" & VbCrlf

  MailBody = MailBody & " </div>" & VbCrlf
  MailBody = MailBody & "</body>" & VbCrlf
  MailBody = MailBody & "</html>"

  ToName = "Member Of " & SiteDescription
  ToEmail = SubscriberAddress
  From = EmailAddress
  Name = SiteDescription
  
  Body = MailBody
  %><!-- #INCLUDE FILE="../Includes/Mail.asp" --><%
  End If
  
  If Err_Msg <> "" Then
   Response.Write Err_Msg & "<br/>"
  Else
   If Active = True Then
    Response.Write "Mail Sent To <b>" & SubscriberAddress & "</b><br/>"
   Else
    Response.Write "Mail NOT Sent To <b>" & SubscriberAddress & "</b><br/>"
   End If
  End If

  Records.MoveNext
 Loop

 Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"
End If

'-- Close Objects --'
Records.Close
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
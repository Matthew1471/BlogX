<%
' --------------------------------------------------------------------------
'¦Introduction : Spellcheck Popup Page.                                     ¦
'¦Purpose      : Checks spelling and provides corrections.                  ¦
'¦Used By      : AddEntry.asp, EditEntry.asp, EditDisclaimer.asp,           ¦
'¦               AddPoll.asp, EditPoll.asp.                                 ¦
'¦Requires     : Includes/Config.asp, Admin.asp, Includes/Spell.asp.        ¦
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
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/Spell.asp" -->
<% If Request.Querystring("AddToDic") <> "" Then
	Records.Open "SELECT Word FROM UserDictionary",Database, 1, 3
	 Records.AddNew
	 Records("Word") = Request.Querystring("AddToDic") 
	 Records.Update
	Records.Close
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - SpellCheck</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 <!--
 //= - - - - - - - 
 // Copyright 2004, Matthew Roberts
 // Copyright 2003, Chris Anderson
 //= - - - - - - -
 -->
<link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type="text/css" rel="stylesheet"/>
<script type="text/javascript">
<%
Dim ReplaceWith, ReplaceText
ReplaceText = Request.Querystring("Replace")
ReplaceWith = Request.Querystring("With")

If Request.Form("Content") = "" Then %>
 function Paste() {
  document.AddEntry.Content.select(); 
  document.AddEntry.Content.focus(); 
  Copied = document.AddEntry.Content.createTextRange();
  Copied.execCommand("Paste");

 <% If ReplaceText <> "" Then %>
 var re = new RegExp (' <%=Replace(ReplaceText,"'","\'")%> ', 'gi');
 document.AddEntry.Content.value = document.AddEntry.Content.value.replace(re, ' <%If ReplaceWith <> "" Then Response.Write Replace(ReplaceWith,"'","\'") %> ');

 var re = new RegExp ('<%=Replace(ReplaceText,"'","\'")%>', 'gi');
 document.AddEntry.Content.value = document.AddEntry.Content.value.replace(re, '<%=Replace(ReplaceWith,"'","\'")%>');

  <% If ReplaceWith <> "" Then %>
  opnform=window.opener.document.forms['AddEntry'];
  opnform['Content'].value = opnform['Content'].value.replace(re, '<%=Replace(ReplaceWith,"'","\'")%>');
  <% End If
 End If%>
 Copy()
 }
 function Copy() {
  document.AddEntry.Content.select(); 
  document.AddEntry.Content.focus();
  Copied = document.AddEntry.Content.createTextRange();
  Copied.execCommand("Copy");
 }
 <%Else%>
 function Suggest(Word) {
  document.Change.With.value = Word;
 }
 <%End If%>
 </script>
</head>
<body style="background-color:<%=BackgroundColor%>" <% If Request.Form("Content") = "" Then Response.Write "onload=""Paste()"""%>>
<%
If Request.Form("Content") = "" Then

	Response.Write "<form id=""AddEntry"" method=""post"" action=""Spell.asp"">" & VbCrlf
    Response.Write "<p style=""text-align:center"">" & VbCrlf
	Response.Write "Remove any HTML formatting or any non dictionary words<br/><br/><b>Words To Be Checked:</b>" & VbCrlf
	Response.Write "<textarea name=""Content"" rows=""8"" cols=""70"" style=""height:10em;width:98%;"" onchange=""javascript:Copy()""></textarea>" & VbCrlf
	Response.Write "<input type=""submit"" value=""Check Spelling""/>" & VbCrlf
    Response.Write "</p>" & VbCrlf
	Response.Write "</form>" & VbCrlf

Else

 LoadDictArray

 Dim Word, sarySearch
 Word = Request.Form("Content")
 sarySearch = Split(Trim(Word), " ")

 For Count = 0 To UBound(sarySearch)

  '--- User submitted dictionary check ---'
  Records.Open "SELECT Word FROM UserDictionary",Database

   Do Until (Records.EOF)

    If PrepForSpellCheck(sarySearch(Count)) = Records("Word") Then 
	 Records.MoveFirst
	 sarySearch(Count) = "ignored"
    End If

    Records.MoveNext
   Loop

  Records.Close
  '--- End of user submitted dictionary check ---'

  Response.Write "<p style=""text-align:center"">" & VbCrlf

  If SpellCheck(PrepForSpellCheck(sarySearch(Count))) <> True Then
   Response.Write "<br/><b>" & PrepForSpellCheck(sarySearch(Count)) & "</b> is misspelled.</p>" & VbCrlf
   Response.Write "<p style=""text-align:center""><b>Suggestions:</b><br/>" & VbCrlf
   Response.Write "-- <a href=""Spell.asp?AddToDic=" & PrepForSpellCheck(sarySearch(Count)) & """>Add To The Custom Dictionary</a> --<br/>" & VbCrlf
   Response.Write "-- <a href=""Spell.asp?Replace=" & PrepForSpellCheck(sarySearch(Count)) & """>Ignore This Word</a> --<br/>" & VbCrlf

   Dim strWord
   For Each strWord In Suggest(sarySearch(Count))
    Response.Write "<a href=""javascript:Suggest('" & Replace(strWord,"'","\'") & "')"">" & StrWord & "</a><br/>" & VbCrlf
   Next
   Response.Write "</p>" & VbCrlf

   Response.Write "<form id=""Change"" method=""get"" action=""Spell.asp"">" & VbCrlf
   Response.Write "<p style=""text-align:center"">" & VbCrlf
   Response.Write "<input type=""hidden"" name=""Replace"" value=""" & PrepForSpellCheck(sarySearch(Count)) & """/>" & VbCrlf
   Response.Write "<b>Change To:</b> <input type=""text"" name=""With""/>" & VbCrlf
   Response.Write "<input type=""submit"" value=""Change""/>" & VbCrlf
   Response.Write "</p>" & VbCrlf
   Response.Write "</form>" & VbCrlf

   Database.Close
   Set Records = Nothing
   Set Database = Nothing
   Response.Write "</body>" & VbCrlf
   Response.Write "</html>"
   Response.End
  End If

 Next

 Response.Write "All Checked Words are now correctly spelt." & VbCrlf
 Response.Write "</p>" & VbCrlf
 Response.Write "<p style=""text-align:center""><a href=""JavaScript:self.close()"">Close Window</a></p>" & VbCrlf

End If

Database.Close
Set Database = Nothing
Set Records = Nothing
%>
</body>
</html>
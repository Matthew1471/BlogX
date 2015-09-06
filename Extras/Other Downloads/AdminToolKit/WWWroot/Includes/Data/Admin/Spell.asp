<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/Spell.asp" -->
<% If Request.Querystring("AddToDic") <> "" Then
		    Records.CursorType = 2
		    Records.LockType = 3
		Records.Open "SELECT Word FROM UserDictionary",Database
		Records.AddNew
		Records("Word") = Request.Querystring("AddToDic") 
		Records.Update
		Records.Close
End If
%>
<html>
<head>
<title><%=SiteDescription%> - SpellCheck</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% 
Dim ReplaceWith, ReplaceText
ReplaceText = Request.Querystring("Replace")
ReplaceWith = Request.Querystring("With")
%>
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>

<script language="JavaScript">
<% If Request.Form("Content") = "" Then %>
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
	<% 	End If
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
<body bgcolor="<%=BackgroundColor%>" <% If Request.Form("Content") = "" Then Response.Write "onload=""Paste()"""%>>
<p align="Center">
<%
If Request.Form("Content") = "" Then

	Response.Write "<Form Name=""AddEntry"" Method=""Post"" Action=""Spell.asp"">"
	Response.Write "Remove any HTML formatting or any non dictionary words<br><br><b>Words To Be Checked:</b> "
	Response.Write "<textarea Name=""Content"" DESIGNTIMEDRAGDROP=""96"" style=""height:10em;width:100%;"" onchange=""javascript:Copy()""></textarea>"
	Response.Write "<input type=""submit"" value=""Check Spelling"">"
	Response.Write "</Form>"

Else

	LoadDictArray

	Dim Word, sarySearch
	Word = Request.Form("Content")
	sarySearch = Split(Trim(Word), " ")


	For Count = 0 To UBound(sarySearch)

		'--- User Submitted Dictionary Check ---'
		Records.Open "SELECT Word FROM UserDictionary",Database
		Do Until (Records.EOF)

		If PrepForSpellCheck(sarySearch(Count)) = Records("Word") Then 
		Records.MoveFirst
		sarySearch(Count) = "ignored"
		End If

		Records.MoveNext
		Loop
		Records.Close

		'--- End of User Submitted Dictionary Check ---'

	If SpellCheck(PrepForSpellCheck(sarySearch(Count))) <> True Then

	Response.Write "<br><b>" & PrepForSpellCheck(sarySearch(Count)) & "</b> is misspelled.</p>" & vbNewLine

	Response.Write "<p align=""center""><b>Suggestions:</b><br>" & vbNewLine

			Response.Write "-- <a href=""Spell.asp?AddToDic=" & PrepForSpellCheck(sarySearch(Count)) & """>Add To The Custom Dictionary</a> --<br>"
			Response.Write "-- <a href=""Spell.asp?Replace=" & PrepForSpellCheck(sarySearch(Count)) & """>Ignore This Word</a> --<br>"

    			For Each strWord In Suggest(sarySearch(Count))
        			Response.Write "<a href=""javascript:Suggest('" & Replace(strWord,"'","\'") & "')"">" & StrWord & "</a><br>" & vbNewLine
    			Next
				Response.Write "<form name=""Change"" method=""get"">" & vbNewLine
				Response.Write "<input type=""hidden"" name=""Replace"" value=""" & PrepForSpellCheck(sarySearch(Count)) & """>" & vbNewLine
				Response.Write "<p align=""center""><b>Change To:</b> <input type=""text"" name=""With"">" & vbNewLine
				Response.Write "<input type=""submit"" value=""Change"">" & vbNewLine
				Response.Write "</form>" & vbNewLine
	Database.Close
	Set Records = Nothing
	Set Database = Nothing
	Response.Write "</p>"
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
	End If

	Next
	Response.Write "<br>"
	Response.Write "All Checked Words are now correctly spelt."
	Response.Write "<p align=""center""><a href=""JavaScript:self.close()"">Close Window</a></p>"

	End If

Database.Close
Set Database = Nothing
Set Records = Nothing
%>
</p>
</Body>
</html>
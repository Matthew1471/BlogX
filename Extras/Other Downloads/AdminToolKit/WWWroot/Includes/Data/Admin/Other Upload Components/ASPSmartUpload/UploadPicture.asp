<% Option Explicit %>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<html>
<head>
<title><%=SiteDescription%> - Upload/Smiley</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% 
If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
<script language="JavaScript">
function Smileyinfo(Smile)
{
    //Inda: Use the functions from AddEntry.asp (RTF.js) to insert smilies
    if(window.opener.document.AddEntry.Content.selectionStart > -1) //Mozilla
	{
		window.opener.changeMozilla(Smile, true, true);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		window.opener.changeIE(Smile, true, true);
	}
	else
	{
		alert("Your browser is not supported");
	}
	
    self.close();
}
</script>
</head>
<body bgcolor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<Center>
<b>Insert Smiley</b><br><br>
<a href="JavaScript:Smileyinfo('\(Y\)')"><img border=0 src="../Images/Emoticons/Approve.gif"></a>
<a href="JavaScript:Smileyinfo(':$')"><img border=0 src="../Images/Emoticons/Blush.gif"></a>
<a href="JavaScript:Smileyinfo('\(H\)')"><img border=0 src="../Images/Emoticons/Cool.gif"></a>
<a href="JavaScript:Smileyinfo('\(Clown\)')"><img border=0 src="../Images/Emoticons/Clown.gif"></a>
<a href="JavaScript:Smileyinfo('\(X\)')"><img border=0 src="../Images/Emoticons/Dead.gif"></a>
<a href="JavaScript:Smileyinfo('\(D\)')"><img border=0 src="../Images/Emoticons/Depressed.gif"></a>
<a href="JavaScript:Smileyinfo('\(6\)')"><img border=0 src="../Images/Emoticons/Evil.gif"></a>
<a href="JavaScript:Smileyinfo('\(8\)')"><img border=0 src="../Images/Emoticons/Note.gif"></a>
<a href="JavaScript:Smileyinfo(':D')"><img border=0 src="../Images/Emoticons/Grin.gif"></a>
<a href="JavaScript:Smileyinfo('\(Hurt\)')"><img border=0 src="../Images/Emoticons/Hurt.gif"></a>
<a href="JavaScript:Smileyinfo('\(K\)')"><img border=0 src="../Images/Emoticons/Kiss.gif"></a><br>
<a href="JavaScript:Smileyinfo(':@')"><img border=0 src="../Images/Emoticons/Mad.gif"></a>
<a href="JavaScript:Smileyinfo('\(Mail\)')"><img border=0 src="../Images/Emoticons/Mail.gif"></a>
<a href="JavaScript:Smileyinfo('\(Entry\)')"><img border=0 src="../Images/Emoticons/Post.gif"></a>
<a href="JavaScript:Smileyinfo('\(User\)')"><img border=0 src="../Images/Emoticons/Profile.gif"></a>
<a href="JavaScript:Smileyinfo('\(?\)')"><img border=0 src="../Images/Emoticons/Question.gif"></a>
<a href="JavaScript:Smileyinfo(':(')"><img border=0 src="../Images/Emoticons/Sad.gif"></a>
<a href="JavaScript:Smileyinfo(':\)')"><img border=0 src="../Images/Emoticons/Smile.gif"></a>
<a href="JavaScript:Smileyinfo(':-O')"><img border=0 src="../Images/Emoticons/Shock.gif"></a>
<a href="JavaScript:Smileyinfo('\(Shy\)')"><img border=0 src="../Images/Emoticons/Shy.gif"></a>
<a href="JavaScript:Smileyinfo('^_^')"><img border=0 src="../Images/Emoticons/Sleepy.gif"></a>
<a href="JavaScript:Smileyinfo('\(*\)')"><img border=0 src="../Images/Emoticons/Star.gif"></a>
<a href="JavaScript:Smileyinfo(':P')"><img border=0 src="../Images/Emoticons/Tongue.gif"></a>
<a href="JavaScript:Smileyinfo('\(URL\)')"><img border=0 src="../Images/Emoticons/URL.gif"></a>
<a href="JavaScript:Smileyinfo(';-\)')"><img border=0 src="../Images/Emoticons/Wink.gif"></a>

<hr>
<b>Upload Photo</b>

<FORM Name="MyForm" METHOD="POST" ENCTYPE="multipart/form-data" ACTION="UploadPicture_Save.asp?<% If Request.Querystring() = "MainPhoto" Then Response.Write "&MainPage=True"%>">
<INPUT TYPE="HIDDEN" Name="Action" Value="Upload">
<INPUT TYPE="FILE" SIZE="30" NAME="FILE1"><br>
<INPUT type=submit value="Upload!">
</FORM>

</Center>

</Body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing
%>
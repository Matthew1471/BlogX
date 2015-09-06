<% Option Explicit 
Dim Returned, UploadProgress, PID, Barref %>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="User.asp" -->
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
<%
On Error Resume Next
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")

If Len(CStr(Err.Description)) > 0 Then

If (Instr(Err.Description,"Object required") = 0) AND (Instr(Err.Description,"Object not a collection") = 0) AND (InStr(Err.Description,"Server.CreateObject Failed") = 0) Then
Returned = "<p align=""center"">Error : " & Err.Description & "</p>"
Else
Returned = "<p align=""center""><font color=""red"">Photo Uploads Require<br> <a href=""http://aspupload.com/"">PERSITS ASP Upload</a></font></p>"
End If

End If

PID = "PID=" & UploadProgress.CreateProgressID()
barref = Replace(SiteURL,"'","\'") & "Includes/Framebar.asp?to=10&" & PID
%>
<SCRIPT LANGUAGE="JavaScript">
function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.MyForm.FILE1.value != "")
  {
    if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
    {
      winstyle = "dialogWidth=375px; dialogHeight:145px; center:yes";
      window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
    }
    else
    {
      window.open('<% = barref %>&b=NN','','width=370,height=115', true);
    }
  }
  return true;
}
</SCRIPT>
<script language="JavaScript">
function Smileyinfo(Smile)
{
    opnform=window.opener.document.forms['AddEntry'];
    opnform['Content'].value+=Smile;
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
<% If Returned = "" Then %>
<FORM Name="MyForm" METHOD="POST" ENCTYPE="multipart/form-data" ACTION="UploadPicture_Save.asp?<% = PID %><% If Request.Querystring() = "MainPhoto" Then Response.Write "&MainPage=True"%>" OnSubmit="return ShowProgress();">
<INPUT TYPE="HIDDEN" Name="Action" Value="Upload">
<INPUT TYPE="FILE" SIZE="30" NAME="FILE1"><br>
<INPUT type=submit value="Upload!">
</FORM>
<% Else
Response.Write Returned
End If %>
</Center>
</Body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing
%>
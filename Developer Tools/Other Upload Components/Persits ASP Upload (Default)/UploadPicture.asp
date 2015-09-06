<%
' --------------------------------------------------------------------------
'¦Introduction : Add Picture Or Smiley Form Popup Page.                     ¦
'¦Purpose      : Allows blog administrator to upload a picture or select a  ¦
'¦               pre-defined smiley face.                                   ¦
'¦Used By      : AddEntry.asp, EditEntry.asp, EditDisclaimer.asp,           ¦
'¦               AddPoll.asp, EditPoll.asp.                                 ¦
'¦Requires     : Includes/Config.asp, Admin.asp, UploadPicture_Save.asp,    ¦
'¦               Includes/Framebar.asp.                                     ¦
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
Dim Returned, UploadProgress, PID, Barref %>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title><%=SiteDescription%> - Add Picture/Smiley</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 <!--
 //= - - - - - - - 
 // Copyright 2004-08, Matthew Roberts
 // Copyright 2003, Chris Anderson
 // 
 // Usage Of This Software Is Subject To The Terms Of The License
 //= - - - - - - -
 -->
<% 
If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")
Response.Write " <link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=""text/css"" rel=""stylesheet""/>" & VbCrlf

'-- Create a Persits progress bar process, this also serves as a good test for if Persits is even installed --'
On Error Resume Next
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")

 If Len(CStr(Err.Description)) > 0 Then

  If (Instr(Err.Description,"Object required") = 0) AND (Instr(Err.Description,"Object not a collection") = 0) AND (InStr(Err.Description,"Server.CreateObject Failed") = 0) Then
   Returned = "<p style=""text-align:center"">Error : " & Err.Description & "</p>"
  Else
   Returned = "<p style=""text-align:center; color:red"">Photo Uploads Require<br/> <a href=""http://aspupload.com/"">PERSITS ASP Upload</a></p>"
  End If
 Else
  PID = "PID=" & UploadProgress.CreateProgressID()
  barref = Replace(SiteURL,"'","\'") & "Includes/Framebar.asp?to=10%26" & PID
 End If

Set UploadProgress = Nothing
On Error GoTo 0
%>
 <script type="text/javascript">
  function ShowProgress() {
   strAppVersion = navigator.appVersion;
  
   if (document.forms['AddFile'].file1.value != "") {
    if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4) {
     winstyle = "dialogWidth=375px; dialogHeight:140px; center:yes";
     window.showModelessDialog(unescape('<%=barref%>%26b=IE'),null,winstyle);
    } else {
     var left = (screen.width-375)/2;
     var top = (screen.height-75)/2;
     window.open(unescape('<%=barref%>%26b=NN'),'','width=375,height=75,top='+ top +',left='+ left, true);
    }
   }
  return true;
  }

 function Smileyinfo(Smile) {
  //Inda: Use the functions from AddEntry.asp (RTF.js) to insert smilies
  if(window.opener.document.forms['AddEntry'].Content.selectionStart > -1) {
   //Mozilla
   window.opener.changeMozilla(Smile, true, true);
  } else if(document.selection && document.selection.createRange) {
   //IE
   window.opener.changeIE(Smile, true, true);
  } else {
   alert("Your browser is not supported");
  }
 self.close();
 }
</script>
</head>
<body style="background-color:<%=BackgroundColor%>; text-align:center">
 <p>
  <b>Insert Smiley</b><br/><br/>
  <a href="JavaScript:Smileyinfo('\(Y\)')"><img alt="Approve" style="border:none" src="../Images/Emoticons/Approve.gif"/></a>
  <a href="JavaScript:Smileyinfo(':$')"><img alt="Blush" style="border:none" src="../Images/Emoticons/Blush.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(H\)')"><img alt="Cool" style="border:none" src="../Images/Emoticons/Cool.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(Clown\)')"><img alt="Clown" style="border:none" src="../Images/Emoticons/Clown.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(X\)')"><img alt="Dead" style="border:none" src="../Images/Emoticons/Dead.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(D\)')"><img alt="Depressed" style="border:none" src="../Images/Emoticons/Depressed.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(6\)')"><img alt="Evil" style="border:none" src="../Images/Emoticons/Evil.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(8\)')"><img alt="Note" style="border:none" src="../Images/Emoticons/Note.gif"/></a>
  <a href="JavaScript:Smileyinfo(':D')"><img alt="Grin" style="border:none" src="../Images/Emoticons/Grin.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(Hurt\)')"><img alt="Hurt" style="border:none" src="../Images/Emoticons/Hurt.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(K\)')"><img alt="Kiss" style="border:none" src="../Images/Emoticons/Kiss.gif"/></a><br/>
  <a href="JavaScript:Smileyinfo(':@')"><img alt="Mad" style="border:none" src="../Images/Emoticons/Mad.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(Mail\)')"><img alt="Mail" style="border:none" src="../Images/Emoticons/Mail.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(Entry\)')"><img alt="Post" style="border:none" src="../Images/Emoticons/Post.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(User\)')"><img alt="Profile" style="border:none" src="../Images/Emoticons/Profile.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(?\)')"><img alt="Question" style="border:none" src="../Images/Emoticons/Question.gif"/></a>
  <a href="JavaScript:Smileyinfo(':(')"><img alt="Sad" style="border:none" src="../Images/Emoticons/Sad.gif"/></a>
  <a href="JavaScript:Smileyinfo(':\)')"><img alt="Smile" style="border:none" src="../Images/Emoticons/Smile.gif"/></a>
  <a href="JavaScript:Smileyinfo(':-O')"><img alt="Shock" style="border:none" src="../Images/Emoticons/Shock.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(Shy\)')"><img alt="Shy" style="border:none" src="../Images/Emoticons/Shy.gif"/></a>
  <a href="JavaScript:Smileyinfo('^_^')"><img alt="Sleepy" style="border:none" src="../Images/Emoticons/Sleepy.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(*\)')"><img alt="Star" style="border:none" src="../Images/Emoticons/Star.gif"/></a>
  <a href="JavaScript:Smileyinfo(':P')"><img alt="Tongue" style="border:none" src="../Images/Emoticons/Tongue.gif"/></a>
  <a href="JavaScript:Smileyinfo('\(URL\)')"><img alt="URL" style="border:none" src="../Images/Emoticons/URL.gif"/></a>
  <a href="JavaScript:Smileyinfo(';-\)')"><img alt="Wink" style="border:none" src="../Images/Emoticons/Wink.gif"/></a>
 </p>
 
 <hr/>
  
 <p><b>Upload Photo</b></p>

 <%
 If Returned = "" Then
  Response.Write " <form id=""AddFile"" method=""post"" enctype=""multipart/form-data"" action=""UploadPicture_Save.asp?" & PID
  If Request.Querystring() = "MainPhoto" Then Response.Write "&MainPage=True"
  Response.Write """ onsubmit=""return ShowProgress();"">" & VbCrlf
 
  Response.Write "  <p>" & VbCrlf
  Response.Write "   <input type=""hidden"" name=""Action"" value=""Upload""/>" & VbCrlf
  Response.Write "   <input type=""file"" size=""30"" name=""file1""/><br/>" & VbCrlf
  Response.Write "   <input type=""submit"" value=""Upload!""/>" & VbCrlf
  Response.Write "  </p>" & VbCrlf
 
  Response.Write " </form>" & VbCrlf
 Else
  Response.Write Returned
 End If
%>
</body>
</html>
<% Database.Close
Set Database = Nothing
Set Records = Nothing
%>
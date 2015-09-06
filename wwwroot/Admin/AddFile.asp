<%
' --------------------------------------------------------------------------
'¦Introduction : Add File Form Page.                                        ¦
'¦Purpose      : Allows blog administrator to add a file to shared files.   ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, AddFile_Save.asp,          ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
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

Dim Returned, UploadProgress, PID, Barref %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<%
On Error Resume Next
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")

If Len(CStr(Err.Description)) > 0 Then

If (Instr(Err.Description,"Object required") = 0) AND (Instr(Err.Description,"Object not a collection") = 0) AND (InStr(Err.Description,"Server.CreateObject Failed") = 0) Then
 Returned = "<p>Error : " & Err.Description & "</p>"
Else
 Returned = "<p style=""color:red"">Photo Uploads Requires<br/><a href=""http://aspupload.com/"">Persits ASP Upload</a></p>"
End If

End If

PID = "PID=" & UploadProgress.CreateProgressID()
barref = Replace(SiteURL,"'","\'") & "Includes/Framebar.asp?to=10%26" & PID

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
     var w = 375;
     var h = 75;
     var winl = (screen.width-w)/2;
     var wint = (screen.height-h)/2;
     if (winl < 0) winl = 0;
     if (wint < 0) wint = 0;
     window.open(unescape('<%=barref%>%26b=NN'),'','width=375,height=75,top='+ wint +',left='+ winl, true);
    }
   }
  return true;
  }
  </script>

<div style="text-align:center">
<b>Upload File</b>
<%
If Returned = "" Then %>

 <form id="AddFile" method="post" enctype="multipart/form-data" action="AddFile_Save.asp?<%=PID%>" onsubmit="return ShowProgress();">
  <p>
   <input type="hidden" name="Action" value="Upload"/>
   <input type="file" size="30" name="file1"/><br/>
   <input type="submit" value="Upload!"/>
  </p>
 </form>

  <% 

  Response.Write "<p class=""config"" style=""text-align:center; font-size: smaller"">You are only allowed to upload the following filetypes:<br/>"

  '### Write Out FileTypes ###
  Records.Open "SELECT AllowedExtension FROM FileExtensions ORDER BY AllowedExtension ASC;", Database

  If Records.EOF Then
    Response.Write "<b>None!</b>"
  Else
   Do Until (Records.EOF)
    Response.Write Records("AllowedExtension")
    Records.MoveNext
    If Records.EOF = False Then Response.Write ", "
   Loop
  End If

  Records.Close

  Response.Write "<br/><br/>" & VbCrlf

  Response.Write "<b>Note</b>: To add more allowed file types, ask your administrator to add it to the ""FileExtensions"" table.</p>" & VbCrlf

Else
  Response.Write Returned
End If %>
</div>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
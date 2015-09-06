<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="User.asp" -->
<%
On Error Resume Next

ArticleNumber = "BlogXProxy"

Set Upload = Server.CreateObject("Persits.Upload")
Upload.ProgressID = Request.QueryString("PID")
Upload.IgnoreNoPost = True
Upload.OverwriteFiles = True
Count = Upload.Save(Server.MapPath("..\Images\Articles\Temp") & "\")
For Each File in Upload.Files
   Success = True
   Randomize
   FileName = "Entry" & ArticleNumber & "_" & Int(Rnd*9999)+1 & ".jpg"
   File.Copy Server.MapPath("..\Images\Articles") & "\" & Filename
   File.Delete
Next
Set Upload = Nothing
%>
<html>
<head>
<title><%=SiteDescription%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>

<script language="JavaScript">
function retinfo()
{
    opnform=window.opener.document.forms['AddEntry'];
    opnform['Content'].value+='<img src=\"<%=Replace(SiteURL,"""","\""")%>Images/Articles/<%=FileName%>\" border=\"0\">';
    self.close();
}
</script>
</head>
<body bgColor="<%=BackgroundColor%>">
<p align="center">
<font face="Verdana, Arial, Helvetica" size="4"><% If (Count > 0) AND (Len(CStr(Err.Description)) = 0) Then
Response.Write "Filename : " & FileName & "<br>"
Response.Write "Picture uploaded for Article " & ArticleNumber & "."
Response.Write "<meta http-equiv=""Refresh"" content=""4; URL=JavaScript:retinfo()"">"
Else
Response.Write "Nothing was uploaded,<br>Check you can write to both folders!</p>"
End If

If Len(CStr(Err.Description)) > 0 Then 
If (Instr(Err.Description,"Object required") = 0) AND (Instr(Err.Description,"Object not a collection") = 0) AND (InStr(Err.Description,"Server.CreateObject Failed") = 0) Then
Response.Write "<p align=""center"">Error : " & Err.Description & "</p>"
Else
Response.Write "<p align=""center"">Error : This Page Requires<br> <a href=""http://aspupload.com/"">PERSITS ASP Upload</a></p>"
End If
End If
%>
</font></p>
<p align="center">
<font face="Verdana, Arial, Helvetica" size="2"><a href="JavaScript:retinfo()">Close Window</font></a></p>
</font>
</Center>
</Body>
</html>

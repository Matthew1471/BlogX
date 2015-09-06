<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<%
On Error Resume Next

If Request.Querystring("MainPage") = "" Then

'--- Open set ---'
    Records.Open "SELECT * FROM Data ORDER BY RecordID DESC",Database, 1, 3
    If NOT Records.EOF Then ArticleNumber = Records("RecordID")+1

'--- Close Database ---'
Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing

Else
ArticleNumber = "Main"
End If

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
    //Inda: Use the functions from AddEntry.asp (RTF.js) to insert image markup
    if(window.opener.document.AddEntry.Content.selectionStart > -1) //Mozilla
	{
		window.opener.changeMozilla("img", true, false, "src", "Images/Articles/<%=FileName%>", "border", "0");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		window.opener.changeIE("img", true, false, "src", "Images/Articles/<%=FileName%>", "border", "0");
	}
	else
	{
		alert("Your browser is not supported");
	}
    
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

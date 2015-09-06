<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<!-- #INCLUDE FILE="../Includes/Spell.asp" -->
<html>
<head>
<title><%=SiteDescription%> - SpellCheck Error</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
<br>
<p align="center">The dictionary file could not be found<br>SpellCheck is now disabled.</p>
<p align="center"><b>Technical :</b> The file <br><font color="Red">Includes/Dictionary/dict-large.txt</font><br> (or its specified alternative) is missing.</p>
<%
Database.Close
Set Database = Nothing
Set Records = Nothing
%>
<p align="center"><a href="JavaScript:self.close()">Close Window</a></p>
</Body>
</html>
<%
  On Error Resume Next
  Dim BadLink, Refer
  BadLink = Request.QueryString
  BadLink = Replace(BadLink, "404;", "")

  Refer = Request.ServerVariables("HTTP_REFERER")
  
  BadLink = Replace(BadLink,"/blogs","",1,1,VbTextCompare)
  If Instr(BadLink,"?") <> 0 Then NewLink = BadLink & "&AUTOCORRECT"
  If Instr(BadLink,"?") =  0 Then NewLink = BadLink & "?AUTOCORRECT"
  If Instr(BadLink,"AUTOCORRECT") = 0 Then Response.Redirect(NewLink)
%>

<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28" align="center"><b><font color="#FFFFFF" face="Verdana" size="2">Not Found</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
      <center>Whoops, It appears that you have stumbled upon a page or blog that is not present on this web site.<br>
      It could have been moved, spelled incorrectly or deleted.</center>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
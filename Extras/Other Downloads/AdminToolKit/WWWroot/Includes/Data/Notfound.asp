<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>
<%
  On Error Resume Next
  Dim BadLink, Refer
  BadLink = Replace(Request.QueryString,":80","")
  BadLink = Replace(BadLink, "404;", "")

  Refer = Request.ServerVariables("HTTP_REFERER")
%>

<DIV class=entry>
<h3 class=entryTitle>Page/File Not Found</h3><br>
<DIV class=entryBody>
<% If InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1 <> 0) Then%>
<P>Were you trying to <a href="/Download/WinBlogX%20Setup.exe">Download WinBlogx</a> or <a href="/Download.asp">Download WebBlogx</a>?</P>

<P><b>All pages have been moved up a directory, I apoligise for the inconvenience.</b></P>
<%
If Instr(1, BadLink, "/Blog/", 1) <> 0 Then Response.Write "<p>Have you tried <a href=""" & Replace(BadLink, "Blog/", "", 1,1,1) & """>" & Replace(BadLink, "Blog/", "", 1,1,1) & "</a>?</P>"

End If %>

<P>It appears that you have stumbled upon a page that is not present on this web site.<br>
It could have been moved, spelled incorrectly, or it may even be in our plans to expand the site and develop this page.</P>

<P><b>Error : </b> File "<%=BadLink%>" Not Found<br>
<b>Referrer : </b> <%If Refer <> "" Then Response.Write "<a href=""" & Refer & """>" & Refer & "</a>" Else Response.Write "You Typed In The Address Manually"%></P>
</Div>
</Div>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<DIV id=content>

<!--- Start Information --->
<DIV class=entry>
<H3 class=entryTitle>Important Information About Your WinBlogX</H3>
<DIV class=entryBody>
<P><% If Request.Querystring("Refer")="WinBlogX" AND Request.Querystring("Version")< "1.04.14" Then%>
You Are Using An <b>OLD</b> Version Of WinBlogX.<br>
There is a newer version of <a href="About.asp"><%=Request.Querystring("Refer")%></a> than V<%=Request.Querystring("Version")%></td>
<% Else Response.Redirect"Default.asp"
End If %></P>
</Div></Div>

<!--- End Information --->

<p align="Center"><a href="<%=PageName%>">Back To The Main Page</a></p>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
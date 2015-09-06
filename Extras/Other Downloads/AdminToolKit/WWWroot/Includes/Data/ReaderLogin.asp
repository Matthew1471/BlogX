<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<%
Dim Page
Page = Request.Querystring()

If (Ucase(Request.Form("ReaderPassword")) = Ucase(ReaderPassword)) OR (Session("Reader") = True) OR (Request.Cookies("Reader") = "True") Then

Session("Reader") = True

If Request.Form("ReaderSave") = "True" then
Response.Cookies("Reader") = "True"
Response.Cookies("Reader").Expires = "July 31, 2008"
End If
End If

%>
<div id=content>
<!--- Login As A Reader --->
<% If Session("Reader") = True Then
If Page <> "" Then Response.Redirect(Page) Else Response.Redirect "Default.asp"
Else
Session("Reader") = False
If Session("CookieTest") = "AOK" Then %>
<Form Name="Login" Method="Post" Action="ReaderLogin.asp?<%If Page <> "" Then Response.Write Page Else Response.Write "Default.asp"%>">
<DIV class="section" id=login>
<center> 
<H3><img border="0" src="images/Key.gif">Reader Sign In</H3>
<P><b>The owner of <%=SiteDescription%>, has decided to set a "Reader Password".</b></p>
<P><br><b>You must enter this password to continue.</b></P>
<P><br>Reader Password: <INPUT name="ReaderPassword" type="password"></P>
<P><INPUT name="ReaderSave" type="checkbox" Value="True">Remember Me</P>
<P><INPUT name="SignIn" type="Submit" value="Sign In"></P>
</Center>
</DIV><BR></FORM>
<%Else%>
<DIV class="section" id=login>   
<H3>Login Error</H3>
<p>Please <b>Enable</b> Cookies to login</p>
</Div>
<%
End If
End If %>
</DIV>
<!-- #INCLUDE FILE="Includes/NAV.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
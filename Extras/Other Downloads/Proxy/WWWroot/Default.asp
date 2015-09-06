<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<% If (Session(UserCookieName) = True) OR (Request.Cookies(UserCookieName) = "True") OR (Session(CookieName) = True) OR (Request.Cookies(CookieName) = "True") Then Response.Redirect"Admin/AddEntry.asp"%>

<DIV id=content>
<!--- Start Content --->
<DIV class=entry>
<H3 class=entryTitle><img border="0" src="<%=SiteURL%>Images/Key.gif">User Sign In</H3>
<DIV class=entryBody>

<Form Name="Login" Method="Post" Action="CheckUser.asp">

<P align="center">Username: <INPUT name="username"></P>
<P align="center">Password: <INPUT name="password" type="password"></P>
<P align="center"><INPUT name="Remember" type="checkbox" Value="True">Remember Login</P>
<P align="center"><INPUT name="SignIn" type="Submit" value="Sign In"></P><BR></FORM>
</Div>

</DIV>
<!--- End Content --->
</DIV></DIV>

<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->


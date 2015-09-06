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
<div id="content">
<!--- Login As A Reader --->
<% If Session("Reader") = True Then
	If Page <> "" Then
	 Database.Close
	 Set Records  = Nothing
	 Set Database = Nothing
	 Response.Redirect(Page) 
	Else
	 Database.Close
	 Set Records  = Nothing
	 Set Database = Nothing
	 Response.Redirect "Default.asp"
	End If
Else
 Session("Reader") = False

If Session("CookieTest") = "AOK" Then %>
<form name="Login" method="Post" action="ReaderLogin.asp?<%If Page <> "" Then Response.Write Page Else Response.Write "Default.asp"%>">
<DIV class="section" id=login>
<center> 
 <h3><img border="0" src="images/Key.gif">Reader Sign In</h3>
 <p><b>The owner of <%=SiteDescription%>, has decided to set a "Reader Password".</b></p>
 <p><br><b>You must enter this password to continue.</b></p>
 <p><br>Reader Password: <input name="ReaderPassword" type="password"></p>
 <p><input name="ReaderSave" type="checkbox" Value="True">Remember Me</p>
 <p><input name="SignIn" type="Submit" value="Sign In"></p>
</center>
</div><br>
</form>
<%Else%>
<div class="section" id="login">   
 <h3>Login Error</h3>
 <p>Please <b>Enable</b> Cookies to login</p>
</div>
<%
End If
End If %>
</DIV>
<!-- #INCLUDE FILE="Includes/NAV.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
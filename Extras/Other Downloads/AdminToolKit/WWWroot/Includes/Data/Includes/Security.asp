<%
If (Ucase(Request.Form("Username")) = Ucase(AdminUsername)) AND (Ucase(Request.Form("Password")) = UCase(AdminPassword)) OR (Session(CookieName) = True) OR (Request.Cookies(CookieName) = "True") Then

Session(CookieName) = True

If Request.Form("Remember") = "True" then
Response.Cookies(CookieName) = "True"
Response.Cookies(CookieName).Expires = "July 31, 2008"
End If
End If
%>
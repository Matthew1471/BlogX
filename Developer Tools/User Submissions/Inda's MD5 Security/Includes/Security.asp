<!-- #INCLUDE FILE="class_md5.asp" -->

<%
'Inda: Create MD5 objects
Dim objMD5
Set objMD5 = New MD5

'Inda: Set MD5 text
objMD5.Text = Request.Form("Password")

If (Ucase(Request.Form("Username")) = Ucase(AdminUsername)) AND (Ucase(objMD5.HEXMD5) = UCase(AdminPassword)) OR (Session(CookieName) = True) OR (Request.Cookies(CookieName) = "True") Then

Session(CookieName) = True

If Request.Form("Remember") = "True" then
Response.Cookies(CookieName) = "True"
Response.Cookies(CookieName).Expires = "July 31, 2008"
End If
End If
%>
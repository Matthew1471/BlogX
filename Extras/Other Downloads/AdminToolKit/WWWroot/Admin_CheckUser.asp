<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
	'Read in the password for the user from the database
	If (Request.Form("Password") = AdminPassword) AND (Request.Form("Username") = AdminUsername)Then 

        Session("ToolKitAdmin") = True
        Response.Redirect("Admin_Main.asp")

        Else

        Response.Redirect"Admin_Unauthorised.asp"

	End If
%>
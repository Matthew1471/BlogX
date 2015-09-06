<%
OPTION EXPLICIT
If Session("ToolKitAdmin") <> True Then Response.Redirect("Admin_Unauthorised.asp") %>
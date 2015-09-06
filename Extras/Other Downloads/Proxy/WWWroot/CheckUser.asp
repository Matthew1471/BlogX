<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
'Dimension variables

'Initalise the strUserName variable
Dim UserName
UserName = Replace(Request.Form("UserName"),"'","''")

Records.Open "SELECT Password FROM tblUsers WHERE tblUsers.UserID ='" & UserName & "'", Database

'If the recordset finds a record for the username entered then read in the password for the user
If NOT Records.EOF Then
	
	'Read in the password for the user from the database
	If (Request.Form("Password")) = Records("Password") Then
		
		'If the password is correct then set the session variable to True
		Session("blnIsUserGood") = True
		Records.Close
	                           
                '-- RESET PUK Count & UPDATE LastIP --   
                Records.CursorType = 2
                Records.LockType = 3
                Records.Open "SELECT UserID, IP FROM tblUsers WHERE UserID ='" & UserName & "'", Database
                Records("IP") = Request.ServerVariables("REMOTE_ADDR")
                Records.Update
                Records.Close
		              
		'--Close Objects before redirecting --
		Set Records = Nothing
		Database.Close
		Set Database = Nothing
		
		'-- Send USER to AUTHORISED page --
		'Redirect to the authorised user page and send the users name
                Session("UserName") = UserName
		Session(UserCookieName) = True

		If Request.Form("Remember") = "True" then
		Response.Cookies(UserCookieName) = "True"
		Response.Cookies("UserName") = Session("UserName")
		Response.Cookies(UserCookieName).Expires = "July 31, 2008"
		End If

                Action = UserName & " <b>Logged In</b>"
                %>
                <!-- #INCLUDE FILE="Includes/Add_Action.asp" -->
                <%
		Response.Redirect "Admin/AddEntry.asp"
Else
		'#### Close Objects ###	
		Records.Close
		Set Records = Nothing
		Database.Close
		Set Database = Nothing
                Response.Redirect"Unauthorised.asp"
	End If
End If
		
'#### Close Objects ###
Records.Close
Set Records = Nothing
Database.Close
Set Database = Nothing
	
'If the script is still running then the user must not be authorised
Session("blnIsUserGood") = False

'Redirect to the unauthorised user page
Response.Redirect"Unauthorised.asp"
%>
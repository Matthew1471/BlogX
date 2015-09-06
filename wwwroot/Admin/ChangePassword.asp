<%
' --------------------------------------------------------------------------
'¦Introduction : Change Password Page.                                      ¦
'¦Purpose      : Allows Blog administrator to change the password.          ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp,                            ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************

AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<% If Request.Form("Action") <> "Post" Then %>
 <form method="post" onsubmit="return setVar()" action="ChangePassword.asp">
  <p>
   <input name="Action" type="hidden" value="Post"/>
  </p>
  <p>
   Username : <input name="AdminUsername" type="text" style="width:20%;" value="<%=AdminUsername%>" onchange="return setVarChange()" maxlength="50"/>
  </p>
  <p>
   New Password : <input name="AdminPassword" type="password" style="width:20%;" onchange="return setVarChange()" maxlength="50"/>
  </p>
  <p>
   Confirm Password : <input name="ConfirmPassword" type="password" style="width:20%;" onchange="return setVarChange()" maxlength="50"/>
  </p>
  <p>
   <input type="submit" value="Save"/>
  </p>
</form>
<% 

If SSLSupported = True Then Response.Write "<p><b>Warning:</b> Though secure logins have been activated in this version of BlogX, this form is not secure.<br/>Do not change this password while using an untrusted internet provider.</p>"

Else

 Dim NewUsername, ConfirmPassword, NewPassword
 NewUsername     = Request.Form("AdminUsername")
 ConfirmPassword = Request.Form("ConfirmPassword")
 NewPassword     = Request.Form("AdminPassword")

 Dim Message

  If (ConfirmPassword <> NewPassword) Then
   Message = "Invalid confirmation password."
  ElseIf (NewPassword = "") Then 
   Message = "Cannot set a blank password."
  Else
   Records.CursorType = 2
   Records.LockType = 3
   Records.Open "SELECT AdminUsername, AdminPassword FROM Config", Database
    Records("AdminUsername") = NewUsername
    Records("AdminPassword") = NewPassword
   Records.Update	
   Records.Close

   Message = "Password update successful."
  End If

 Response.Write "<p style=""text-align:center"">" & Message & "</p>"
 Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
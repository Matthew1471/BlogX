<% AlertBack = True %>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<Form Name="Config" Method="Post" onSubmit="return setVar()">
<input Name="Action" type="hidden" Value="Post">
            <P><span id="Label1">UserName : </span><input Name="AdminUsername" type="text" style="width:20%;" Value="<%=AdminUsername%>" onChange="return setVarChange()"></P>
            <P><span id="Label1">New Password : </span><input Name="AdminPassword" type="password" style="width:20%;" onChange="return setVarChange()"></P>
            <P><span id="Label1">Confirm Password : </span><input Name="ConfirmPassword" type="password" style="width:20%;" onChange="return setVarChange()"></P>
            <P></P>
            <Input Type="submit" Value="Save">
        </form>
<% Else

'### CheckBox Check! ###'
NewUsername     = Request.Form("AdminUsername")
ConfirmPassword = Request.Form("ConfirmPassword")
NewPassword     = Request.Form("AdminPassword")

If ConfirmPassword <> NewPassword Then DoIt = "VerificationFailed"

If DoIt <> "VerificationFailed" Then

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Config", Database
Records("AdminUsername") = NewUsername
Records("AdminPassword") = NewPassword
Records.Update

'#### Close Objects ###	
Records.Close

Message = "Password Update Successfull"

Else
Message = "Invalid Confirmation Password"

End If

Response.Write "<p align=""Center"">" & Message & "</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If %>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
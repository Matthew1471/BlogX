<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<?xml version='1.0'?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<card title="Post To Blog">
<p>
<%
Response.ContentType = "text/vnd.wap.wml"

If (UCase(Request.Form("Username")) = UCase(AdminUsername)) AND (UCase(Request.Form("Password")) = UCase(AdminPassword)) Then
 If Request.Form("Action") <> "DoIt" Then %>
		Title : <input type="text" name="Title"/><br/>
		Content : <input type="text" name="Content"/><br/>
               <% If ShowCategories <> False Then Response.Write "Category : <input type=""text"" name=""Category""/><br/>"%>

		<anchor title="Save">Save
		<go href="AddNew.asp" method="post">
		<postfield name="Action" value="DoIt"/>
		<postfield name="Username" value="$(Username)"/>
		<postfield name="Password" value="$(Password)"/>
		<postfield name="Title" value="$(Title)"/>
		<postfield name="Content" value="$(Content)"/>
                <% If ShowCategories <> False Then Response.Write "<postfield name=""Category"" value=""$(Category)""/>"%>
		</go>
		</anchor>
<%
 Else

  '-- Declare variables --'
  Dim EntryCat
  EntryCat = Request.Form("Category")

  '-- Filter & Clean --'
  EntryCat = Replace(EntryCat,"'","&#39;")
  EntryCat = Replace(EntryCat," ","%20")

  '-- Did we type in text? --'
  If Request.Form("Content") = "" Then
   Response.Write "No Text Entered"
   Response.End
  End If

  '-- Open The Records Ready To Write --'
  Records.CursorType = 2
  Records.LockType = 3
  Records.Open "SELECT Title, Text, Category, Day, Month, Year, Time FROM Data", Database
  Records.AddNew
   Records("Title") = Request.Form("Title")
   Records("Text") = Request.Form("Content")
   Records("Category") = EntryCat

   Records("Day") = Day(Now())
   Records("Month") = Month(Now())
   Records("Year") = Year(Now())
   Records("Time") = Time()
   Records.Update

  '-- Close objects --'
  Records.Close
  Set Records = Nothing

  Response.Write "Entry Submitted<br/>"
  Response.Write "Successfully"

 End If

Else
 Response.Write "Invalid Username/Password"
End If

Database.Close
Set Database = Nothing
%>
</p>
</card>
</wml>
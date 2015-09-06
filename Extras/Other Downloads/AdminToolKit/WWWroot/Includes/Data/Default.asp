<% OPTION EXPLICIT
Response.Buffer = True
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<% If EnableMainPage <> True Then
Database.Close
Set Records = Nothing
Set Database = Nothing

'*** WHY IS THIS PAGE ERORRING ON MY IIS 3/4????
'	Answer : ASP 2.0 does not know "Server.Transfer"
'	Solution : Replace Server.Transfer with Response.Redirect
'***********************************************

Response.Clear

On Error Resume Next
Server.Transfer("Main.asp")
Response.Redirect("Main.asp")
On Error Goto 0
Else
%>
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
'--- Open RecordSet ---'
Records.Open "SELECT MainID, MainText FROM Main",Database, 1, 3

If NOT Records.EOF Then

Dim MainID, MainText

'--- Setup Variables ---'
   MainID = Records("MainID")
   MainText = Records("MainText")

End If

Records.Close
%>
<DIV id=content>

<!--- Start Header --->
<DIV class=date id="Main">
<H2 class=dateHeader><%=SiteSubTitle%></H2>
<!--- End Header --->
</Div>

<!--- Start Content --->
<DIV class=entry>
<H3 class=entryTitle><%=SiteDescription%><%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your About Page""><a href=""Admin/EditMainPage.asp""><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym>"%></H3>
<DIV class=entryBody><P><%=Replace(MainText, vbcrlf, "<p>" & vbcrlf) %></P>
<p align="Center"><a href="Main.asp">View The Blog</a></p>
</DIV>
<P class=entryFooter>
<% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%></P></DIV>
<!--- End Content --->

</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%End If%>
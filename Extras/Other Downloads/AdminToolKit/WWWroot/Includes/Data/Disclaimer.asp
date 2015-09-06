<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
'--- Open set ---'
set Records = Server.CreateObject("ADODB.recordset")
    Records.Open "SELECT * FROM Disclaimer",Database, 1, 3

If NOT Records.EOF Then

Dim DisclaimerID, DisclaimerText

'--- Setup Variables ---'
   DisclaimerID = Records("DisclaimerID")
   DisclaimerText = Records("DisclaimerText")

End If

Records.Close
%>

<DIV id=content>

<!--- Start Header --->
<DIV class=date id="Disclaimer">
<H2 class=dateHeader>Disclaimer</H2>
<!--- End Header --->
</Div>

<!--- Start Disclaimer --->
<DIV class=entry>
<H3 class=entryTitle>Disclaimer <%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your Disclaimer Page""><a href=""Admin/EditDisclaimer.asp""><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym>"%></H3>
<DIV class=entryBody><P><%=Replace(DisclaimerText, vbcrlf, "<br>" & vbcrlf)%></P>

</DIV>
<P class=entryFooter>
<% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%></P></DIV>
<!--- End Disclaimer --->
<p align="Center"><a href="<%=PageName%>">Back To The Main Page</a></p>

</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
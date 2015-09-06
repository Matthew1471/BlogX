<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>

<!--- Start Content --->
<DIV class=entry>
<H3 class=entryTitle><%=SiteDescription%></H3>
<DIV class=entryBody>
<%
'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient

    Records.Open "SELECT PollID FROM Poll ORDER BY PollID DESC",Database, 1, 3
    If Records.EOF = False Then PollID = Records("PollID") Else PollID = 0
    Records.Close

    Records.Open "SELECT VoteID FROM Votes WHERE PollID="& PollID & "AND IP='" & Request.ServerVariables("REMOTE_ADDR") & "'",Database, 1, 3
    If Records.EOF = False Then AlreadyVoted = True
    Records.Close

If (Polls <> False) AND (AlreadyVoted = True) Then 

Records.Open "SELECT Content, Des1, Op1, Des2, Op2, Des3, Op3, Des4, Op4, Total FROM Poll ORDER BY PollID DESC",Database, 1, 3

   If NOT Records.EOF Then 

   PollContent = Records("Content")

   Des1 = Records("Des1")
   Des2 = Records("Des2")
   Des3 = Records("Des3")
   Des4 = Records("Des4")

   Op1 = Records("Op1")
   Op2 = Records("Op2")
   Op3 = Records("Op3")
   Op4 = Records("Op4")

   Total = Records("Total")

   Op1Percent = Cint((Op1 / Total) * 100)
   Op2Percent = Cint((Op2 / Total) * 100)
   Op3Percent = Cint((Op3 / Total) * 100)
   Op4Percent = Cint((Op4 / Total) * 100)

   End If

Records.Close
%>

<center><font size='3' color='Maroon'><strong> <%=PollContent%></strong></font>
<br>
<br>
<table width="52%" height="30" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" class="navCalendar">

<% If Des1 <> "" Then %>
<tr>
<td height='20' width='32%' bgcolor='#CCCCFF'><b><%=Des1%></b></td>
<td height='20' width='48%'><p><img src="<%=SiteURL%>Images/Other Poll Colours/Blue.gif" width="<%=Op1Percent%>%" height="10"></p></td>
<td height='20' width='20%' bgcolor='#CCFFCC'><%=Op1Percent%>%</td>
</tr>
<% End If%>

<% If Des2 <> "" Then %>
<tr>
<td height='20' width='32%' bgcolor='#CCCCFF'><b><%=Des2%></b></td>
<td height='20' width='48%'><p><img src="<%=SiteURL%>Images/Other Poll Colours/Yellow.gif" width="<%=Op2Percent%>%" height="10"></p></td>
<td height='20' width='20%' bgcolor='#CCFFCC'><%=Op2Percent%>%</td>
</tr>
<% End If%>

<% If Des3 <> "" Then %>
<tr>
<td height='20' width='32%' bgcolor='#CCCCFF'><b><%=Des3%></b></td>
<td height='20' width='48%'><p><img src="<%=SiteURL%>Images/Other Poll Colours/Red.gif" width="<%=Op3Percent%>%" height="10"></p></td>
<td height='20' width='20%' bgcolor='#CCFFCC'><%=Op3Percent%>%</td>
</tr>
<% End If%>

<% If Des4 <> "" Then %>
<tr>
<td height='20' width='32%' bgcolor='#CCCCFF'><b><%=Des4%></b></td>
<td height='20' width='48%'><p><img src="<%=SiteURL%>Images/Other Poll Colours/Black.gif" width="<%=Op4Percent%>%" height="10"></p></td>
<td height='20' width='20%' bgcolor='#CCFFCC'><%=Op4Percent%>%</td>
</tr>
<% End If%>

</table>
</center>

<% 
ElseIf Polls = False Then

Response.Write "<p align=""center"">Polls are currently disabled<br><br>" & VbCrlf
Response.Write "<a href=""" & PageName & """>Back To Main</a></p>"

Else

Response.Write "<p align=""center"">You need to vote before viewing the results<br><br>" & VbCrlf
Response.Write "<a href=""" & PageName & """>Back To Main</a></p>"

End If
%>
</DIV>
<P class=entryFooter>
<% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp""><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%></P></DIV>
<!--- End Content --->
</Div>

<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
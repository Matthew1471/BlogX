<DIV class=sidebar id=rightBar>
<BR>

<!-- #INCLUDE FILE="Calendar.asp" -->

<% If (ShowMonth <> False) AND (LegacyMode <> True) Then %>
<!--- Archive --->
<DIV class=section>
<H3>Archive</H3>
<UL>
<%
Dim MonthPost, YearPost, LastMonth

'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT Month, Year FROM Data ORDER BY Month",Database, 1, 3

'--- Set Category ---'
Set MonthPost = Records("Month")
Set YearPost = Records("Year")

'-- Write Them In ---'
Do Until (Records.EOF or Records.BOF)
If (LastMonth <> MonthPost) OR (IsNull(LastMonth) = True) AND (MonthPost <> "") Then 
Response.Write "<A Href=""" & SiteURL & "Main.asp?YearMonth=" & YearPost & Right("00" & MonthPost, 2) & """>" 
Response.Write MonthName(MonthPost) & " " & YearPost & "</A><br>" & VbCrlf
LastMonth = MonthPost
End If
Records.MoveNext
Loop

'-- Close The Records ---'
Records.Close
%></UL>
<BR></DIV><BR>
<%End If%>

<!--- Links --->
<DIV class=section>
<H3>Links <%If (Session(CookieName) = True) AND (AllowEditingLinks <> 0) Then Response.Write " <acronym title=""Edit Your Links""><a href=""" & SiteURL & "Admin/EditLinks.asp""><Img Border=""0"" Src=""" & SiteURL & "Images/Edit.gif""></a></acronym>"%></H3>
<UL>
  <% If EnableMainPage = True Then Response.Write "<LI><A href=""" & SiteURL & """>About Me</A></LI>"%>
  <LI><A href="<%=SiteURL & PageName%>">Blog Home</A></LI>
  <!-- #INCLUDE FILE="Links.asp" -->
</UL></DIV><BR>

<% If OtherLinks <> 0 Then %>
<!--- Other Links --->
<DIV class=section>
<H3>Other Links <%If (Session(CookieName) = True) AND (AllowEditingLinks <> 0) Then Response.Write " <acronym title=""Edit Your Links""><a href=""" & SiteURL & "Admin/EditLinks.asp""><Img Border=""0"" Src=""" & SiteURL & "Images/Edit.gif""></a></acronym>"%></H3>
<UL>
<!-- #INCLUDE FILE="OtherLinks.asp" -->
</UL></DIV><BR>
<% End If

If Polls <> False Then

Dim PollID, AlreadyVoted
Dim Des1, Des2, Des3, Des4
Dim Op1, Op2, Op3, Op4
Dim Total, PollContent

Dim Op1Percent, Op2Percent, Op3Percent, Op4Percent

'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient

    Records.Open "SELECT PollID FROM Poll ORDER BY PollID DESC",Database, 1, 3
    If Records.EOF = False Then PollID = Records("PollID") Else PollID = 0
    Records.Close

    Records.Open "SELECT VoteID FROM Votes WHERE PollID="& PollID & "AND IP='" & Request.ServerVariables("REMOTE_ADDR") & "'",Database, 1, 3
    If Records.EOF = False Then AlreadyVoted = True
    Records.Close

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
   %>

<!--- Poll --->
<DIV class=section>
<H3>Poll <%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your Poll""><a href=""" & SiteURL & "Admin/EditPoll.asp""><Img Border=""0"" Src=""" & SiteURL & "Images/Edit.gif""></a></acronym>"%></H3>
   <%
If AlreadyVoted = False Then

   Response.Write "<form name=""Vote"" method=""post"" action=""" & SiteURL & "Vote.asp"">"  & VbCrlf

   Response.Write "<center>" & Records("Content") & "</center>"
   Response.Write "<br>" & VbCrlf
   Response.Write "<hr width=""30%"">" & VbCrlf

   Response.Write "<INPUT type=""radio"" value=""1"" name=""Vote"">&nbsp;"
   Response.Write "<FONT face=""Verdana, Arial, Helvetica"" size=""1""><B>" & Des1 & "</B></FONT><br>" & VbCrlf

   Response.Write "<INPUT type=""radio"" value=""2"" name=""Vote"">&nbsp;"
   Response.Write "<FONT face=""Verdana, Arial, Helvetica"" size=""1""><B>" & Des2 & "</B></FONT><br>" & VbCrlf

   If Des3 <> "" Then 
   Response.Write "<INPUT type=""radio"" value=""3"" name=""Vote"">&nbsp;"
   Response.Write "<FONT face=""Verdana, Arial, Helvetica"" size=""1""><B>" & Des3 & "</B></FONT><br>" & VbCrlf
   End If

   If Des4 <> "" Then 
   Response.Write "<INPUT type=""radio"" value=""4"" name=""Vote"">&nbsp;"
   Response.Write "<FONT face=""Verdana, Arial, Helvetica"" size=""1""><B>" & Des4 & "</B></FONT>" & VbCrlf
   End If
   
   Response.Write "<center><INPUT type=""image"" src=""" & SiteURL & "Images/vote.gif"" border=""0""></center>" & VbCrlf
   Response.Write "</form>" & VbCrlf

Else
   
   Response.Write "<UL>" & Records("Content")%>
   <br>(Total : <%=Records("Total")%> Vote(s) )
   <hr width="30%">
  <% If Instr(Request.ServerVariables("HTTP_USER_AGENT"),"Firefox") = 0 Then %>
      <TABLE cellSpacing=0 width="30%" border="0">
        <TBODY>
         <TR>
           <TD>
   <%
   Op1Percent = Cint((Op1 / Total) * 100)
   Op2Percent = Cint((Op2 / Total) * 100)
   Op3Percent = Cint((Op3 / Total) * 100)
   Op4Percent = Cint((Op4 / Total) * 100)

   If Des1 <> "" Then Response.Write "&nbsp;" & Des1 & "<br> <img src=""" & SiteURL & "Images/Bar.gif"" width=""" & Op1Percent / 2 & "%"" height=""10""> " & Op1 & " (" & Op1Percent & "%)<br><br>" & VbCrlf

   If Des2 <> "" Then Response.Write "&nbsp;" & Des2 & "<br> <img src=""" & SiteURL & "Images/Bar.gif"" width=""" & Op2Percent / 2 & "%"" height=""10""> " & Op2 & " (" & Op2Percent & "%)<br><br>" & VbCrlf

   If Des3 <> "" Then Response.Write "&nbsp;" & Des3 & "<br> <img src=""" & SiteURL & "Images/Bar.gif"" width=""" & Op3Percent / 2 & "%"" height=""10""> " & Op3 & " (" & Op3Percent & "%)<br><br>" & VbCrlf

   If Des4 <> "" Then Response.Write "&nbsp;" & Des4 & "<br> <img src=""" & SiteURL & "Images/Bar.gif"" width=""" & Op4Percent / 2 & "%"" height=""10""> " & Op4 & " (" & Op4Percent & "%)<br>" & VbCrlf

   Response.Write "</td>"
   Response.Write "</tr>"
   Response.Write "</tbody>"
   Response.Write "</table>"

   Else
   
   Response.Write "</UL>"

   End If

   Response.Write "<center><a href=""" & SiteURL & "Results.asp"">Results</a></center>"
   End If %>
</DIV><BR>
   <% 
    End If
    Records.Close
End If %>

<% If ShowCat <> False Then 
Dim Category, LastCat
%>
<!--- Categories --->
<DIV class=section>
<H3>Categories</H3>
<UL>
<Li><A Href="<%=SiteURL & PageName%>">All</A> (<A Href="<%=SiteURL%>RSS/<%If (ReaderPassword <> "") AND Session("Reader") = True Then Response.Write "?" & ReaderPassword%>">Rss</A>)</Li>
<%

'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT Category FROM Data ORDER BY Category",Database, 1, 3

'--- Set Category ---'
Set Category = Records("Category")

'******************************************************** HACK ************************************
   
   ' ************* AddArray SUB *******'
   Sub AddArray (ByRef aArray, ByRef sString)
	   dim iNewUBound
	   iNewUBound = UBound(aArray) + 1
	   redim preserve aArray(iNewUBound)
	   aArray(iNewUBound) = Ucase(Left(sString,1)) & LCase(Right(sString,Len(sString)-1))
   End Sub
   ' *********************************'
   
   ' ********** ASP BubbleSort *******'
   Sub SingleSorter( byRef arrArray )
    Dim row, j
    Dim StartingKeyValue, NewKeyValue, swap_pos

    For row = 0 To UBound( arrArray ) - 1
    'Take a snapshot of the first element
    'in the array because if there is a 
    'smaller value elsewhere in the array 
    'we'll need to do a swap.
        StartingKeyValue = arrArray ( row )
        NewKeyValue = arrArray ( row )
        swap_pos = row
	    	
        For j = row + 1 to UBound( arrArray )
        'Start inner loop.
            If arrArray ( j ) < NewKeyValue Then
            'This is now the lowest number - 
            'remember it's position.
                swap_pos = j
                NewKeyValue = arrArray ( j )
            End If
        Next
	    
           If swap_pos <> row Then
             'If we get here then we are about to do a swap
             'within the array.		
              arrArray ( swap_pos ) = StartingKeyValue
              arrArray ( row ) = NewKeyValue
           End If	
       Next
    End Sub
    ' ***************************************************'

   
Dim EntryCategoryArray, CategoriesArray()
ReDim CategoriesArray(1)

    Do Until (Records.EOF or Records.BOF)

      EntryCategoryArray = Split(Category,",%20") 

      For Count = 0 To UBound(EntryCategoryArray)
      Call AddArray(CategoriesArray,EntryCategoryArray(Count))
      Next                           
          
    Records.MoveNext        
    Loop

SingleSorter CategoriesArray

Records.Close  
   
  For Count = 0 To Ubound(CategoriesArray)   
  If (LastCat <> CategoriesArray(Count)) OR (IsNull(LastCat) = True) AND (CategoriesArray(Count) <> "") Then 
  Response.Write "<Li><A Href=""" & SiteURL & "ViewCat.asp?Cat=" & Replace(CategoriesArray(Count), " ", "%20") & """>" 
  Response.Write Replace(CategoriesArray(Count), "%20", " ") & "</A> (<A Href=""" & SiteURL & "RSS/Cat/?Category=" & Replace(CategoriesArray(Count), " ", "%20")
  If (ReaderPassword <> "") AND (Session("Reader") = True) Then Response.Write "&Password=" & ReaderPassword
  Response.Write """>Rss</A>)</Li>" & VbCrlf
  LastCat = CategoriesArray(Count)
  End If
  Next

'********************************************** END OF HACK (EOH)  ****************************************************************
%>
</UL></DIV><BR>
<%End If

If (UseExternalPlugin = 1) AND (LegacyMode = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!--- <%=PluginTitle%> --->
<DIV class=section>
<H3><%=PluginTitle%></H3>
<UL><%=PluginText%></UL>
</DIV><BR>
<%
End If

If (LegacyMode <> True) Then %>
<!--- Search Blog --->
<DIV class=section>
<H3>Search</H3>
<BR>
<form name="Search" method="post" action="<%=SiteURL%>Search.asp">
<form name="Mode" type="hidden" value="Normal">
<input Name="Search" Type="text" value="<%=Replace(Request("Search"),"""","&quot;")%>" size="13" maxlength="70"><input Type="submit" value="Search"><br>
<a href="<%=SiteURL%>Search.asp">Advanced Search</a>
</form>
<BR>
</DIV><BR>
<% End If %>

<!--- Login As A Publisher --->
<% If Session(CookieName) = True Then %>
<DIV class="section">
<H3><img border="0" src="<%=SiteURL%>Images/Key.gif">Admin</H3>
    <ul>
        <li><A href="<%=SiteURL%>Admin/EditMainPage.asp">About Me</A></li>
        <li><A href="<%=SiteURL%>Admin/AddEntry.asp">Add Entry</A></li>
        <li><A href="<%=SiteURL%>Admin/AddPoll.asp">Add Poll</A></li>
        <li><A href="<%=SiteURL%>Admin/EditBan.asp">Banned Addresses</A></li>
        <li><A href="<%=SiteURL%>Admin/ChangePassword.asp">Change Password</A></li>
        <li><A href="<%=SiteURL%>Admin/CheckForUpdate.asp">Check For Update</A></li>
        <li><A href="<%=SiteURL%>Admin/Config.asp">Config</A></li>
        <li><A href="<%=SiteURL%>Admin/EditDisclaimer.asp">Disclaimer</A></li>
        <li><A href="<%=SiteURL%>Admin/EmailConfig.asp">Email Settings</A></li>
        <li><A href="<%=SiteURL%>Admin/MailingListMembers.asp">Mailing List</A></li>
        <% If ArgoSoftMailServer = True Then Response.Write "<li><A href=""" & SiteURL & "Admin/ParseEmails.asp"">Parse E-mails</A></li>" %>
        <li><A href="<%=SiteURL%>Admin/Referrers.asp">Referrers</A></li>
        <li><a href="<%=SiteURL & PageName %>?ClearCookie">Logout</a></li>
    </ul>
</DIV><BR>
<%Else
Session(CookieName) = False
If Session("CookieTest") = "AOK" Then %>
<Form Name="Login" Method="Post" Action="Main.asp">
<DIV class="section" id=login>   
<H3><img border="0" src="<%=SiteURL%>Images/Key.gif">Admin Sign In</H3>
<P>Username: <INPUT name="username"></P>
<P>Password: <INPUT name="password" type="password"></P>
<P><INPUT name="Remember" type="checkbox" Value="True">Remember Login</P>
<P><INPUT name="SignIn" type="Submit" value="Sign In"></P></DIV><BR></FORM>
<%Else%>
<DIV class="section" id=login>   
<H3>Login Error</H3>
<p>Please <b>Enable</b> Cookies to login</p>
</Div>
<%
End If
End If %>
</DIV>
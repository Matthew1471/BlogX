<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Occurrence.asp" -->
<%
Dim Requested
Requested = Request("Search")

'### Filter & Clean ###
Requested = Replace(Requested,"'","''")

'-- If anyone has a quick fix for these, E-mail me, I'm bored of playing with it ---'
Requested = Replace(Requested,"%","")
Requested = Replace(Requested,"_","")
Requested = Replace(Requested,"[","")
Requested = Replace(Requested,"]","")

'--- Should We Process The Search? ---'
Response.Write "<DIV id=content>" & VbCrlf

If Len(Requested) <> 0 Then

                Dim sarySearch, strSQL, NewTitle
                sarySearch = Split(Trim(Requested), " ")

'--- Build up the SQL Query on whether it's an "AnyOrder"(Tm) Search or whether it's an EXACT match ---

Dim Spaces
If Request("NoAutoComplete") = "" Then Spaces = "" Else Spaces = " "

If Request("Mode") = "Any" Then
		'Search for the first search word in the URL titles
                strSQL = "SELECT * FROM Data WHERE Text LIKE '%" & Spaces & sarySearch(0) & Spaces & "%'"

		'Loop to search for each search word entered by the user
		For intSQLLoopCounter = 0 To UBound(sarySearch)
			strSQL = strSQL & " AND Text LIKE '%" & Spaces & sarySearch(intSQLLoopCounter) & Spaces & "%'"
		Next
		
		'Order the search results by the RecordID
		strSQL = strSQL & " ORDER BY RecordID DESC;"
Else

                strSQL = "SELECT * FROM Data WHERE (Title LIKE '%" & Spaces & Requested & Spaces & "%') OR (Text LIKE '%" & Spaces & Requested & Spaces & "%')"
                strSQL = strSQL & " ORDER BY RecordID DESC;"

End If

'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open strSQL,Database, 1, 3

'### UnFilter & Scruffisize ###
Requested = Replace(Requested,"''","'")

' Let's see what page are we looking at right now
Dim nPage
nPage = CLng(Request.QueryString("Page"))

'****************************************************************
' Get Records Count
Dim nRecCount
nRecCount = Records.RecordCount

' Tell recordset to split records in the pages of our size
Records.PageSize = 10

' How many pages we've got
Dim nPageCount
nPageCount = Records.PageCount

' Make sure that the Page parameter passed to us is within the range
If nPage < 1 Or nPage > nPageCount Then nPage = 1

	If nRecCount > 0 Then
	' Time to tell user what we've got so far
	Response.Write "<p align=""Right"">Page : " & nPage & "/" & nPageCount & "</p><p>"

	' Give user some navigation

	' First page
	Response.Write "<Center>"
	Response.Write 	"<A HREF=""Search.asp?Search=" & Requested & "&Page=" &  1 & """>First Page</A>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

	' Previous Page
	Response.Write 	"<A HREF=""Search.asp?Search=" & Requested & "&Page=" & nPage - 1 & """>Prev. Page</A>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

	' Next Page
	Response.Write 	"<A HREF=""Search.asp?Search=" & Requested & "&Page=" & nPage + 1 & """>Next Page</A>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"

	' Last Page
	Response.Write 	"<A HREF=""Search.asp?Search=" & Requested & "&Page=" & nPageCount & """>Last Page</A>"
	Response.Write "</center><br>" & VbCrlf
	End If

' Position recordset to the page we want to see
If nRecCount > 0 Then Records.AbsolutePage = nPage

'--- Setup Day Posted ---'
Dim PreviousDay
PreviousDay = "0"

		Dim RecordID, Title, Text, Password

		' Loop through records until it's a next page or End of Records
		Do Until (Records.EOF or Records.AbsolutePage <> nPage)

		'--- Setup Variables ---'
   		Set RecordID = Records("RecordID")
   		Set Title = Records("Title")
   		Set Text = Records("Text")
		Set Password = Records("Password")

		   If Len(Password) > 0 Then
   	           Text = "<form action=""ProtectedEntry.asp"" method=""GET""><center>" & VbCrlf
                   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """>" & VbCrlf 
                   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br>" & VbCrlf
                   Text = Text & "This post is password protected. To view it please enter your password below:"
                   Text = Text & "<br><br>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
                   Text = Text & "</center></form>"
                   End If

   		'********** Replace(TEXT,NEW, START, NUMBER OF OCC, CASE) ********
   		If Request("Mode") = "Any" Then 
   		
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   		       
   		'Loop to search for each search word entered by the user
		Dim intSQLLoopCounter
		For intSQLLoopCounter = 0 To UBound(sarySearch)
                
                If InStr(sarySearch(intSQLLoopCounter),"http://") = 0 Then _
                Text = Replace(Text, " " & sarySearch(intSQLLoopCounter) , _
                " <b><font color=""#800000""><span style=""background-color: #FFFF00"">" & sarySearch(intSQLLoopCounter) & "</span></font></b>",1,-1,1)
		Next                  
		'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  
		
		NewTitle = Title

                Else
                
                NewTitle = Replace(Title, Requested, _
                " <b><font color=""#800000""><span style=""background-color: #FFFF00; color: black"">" & Requested & "</span></font></b>",1,-1,1)
   		
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   		'Loop to search for each search word entered by the user
		For intSQLLoopCounter = 0 To UBound(sarySearch)
                
                If InStr(sarySearch(intSQLLoopCounter),"http://") = 0 Then _
                Text = Replace(Text, " " & Requested , _
                " <b><font color=""#800000""><span style=""background-color: #FFFF00"">" & Requested & "</span></font></b>",1,-1,1)
	       	Next    
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	
   		End If
   		
		Dim CommentsCount, DayPosted, MonthPosted, YearPosted, TimePosted, NewTime, JustDoIt

   		Set Category =  Records("Category")
   		Set CommentsCount = Records("Comments")

   		Set DayPosted =  Records("Day")
   		Set MonthPosted =  Records("Month")
   		Set YearPosted =  Records("Year")
   		Set TimePosted =  Records("Time")

		'--- We're British, Let's 12Hour Clock Ourselves ---'
		NewTime = ""

		If TimeFormat <> False Then
		If Hour(TimePosted) > 12 Then
		NewTime = Hour(TimePosted) - 12 & ":"
		Else
		NewTime = Hour(TimePosted) & ":"
		End If

		If Minute(TimePosted) < 10 Then
		NewTime = NewTime & "0" & Minute(TimePosted)
		Else
		NewTime = NewTime & Minute(TimePosted)
		End If

		If (Hour(TimePosted) < 12) AND (Hour(TimePosted) <> 12) Then
		NewTime = NewTime & " AM"
		Else
		NewTime = NewTime & " PM"
		End If

		Else
		If Hour(TimePosted) < 10 Then NewTime = "0"
		NewTime = NewTime & Hour(TimePosted) & ":"
		If Minute(TimePosted) < 10 Then NewTime = NewTime & "0"
		NewTime = NewTime & Minute(TimePosted)
		End If

		If DayPosted <> PreviousDay Then
		Response.Write vbcrlf & "<!--- Start Date Header --->" & vbcrlf
		Response.Write "<DIV class=date id=2003-11-30>" & vbcrlf
		Response.Write "<H2 class=dateHeader>" & Left(MonthName(MonthPosted),3) & " " & DayPosted & ", " & YearPosted & " (Only containing "
		If Request("Mode") <> "Any" Then Response.Write "<u>EXACTLY</u> (in this order) "
		Response.Write """<b> " & Replace(Requested, "%20", " ") & " </b>"")</H2>" & vbcrlf
		Response.Write "<!--- End Date Header --->" & vbcrlf
		JustDoit = True
		Else
		JustDoIt = False
		End If
		%>
		<!--- Start Content For Search List (<%=DayPosted%>)--->
		<DIV class=entry>
		<H3 class=entryTitle><A href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTitle%></A> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym> "%><% If Request("Mode") <> "Any" Then Response.Write "(" & CharCount(Text,Requested,False) & " Occurrences)"%></H3>
		<DIV class=entryBody><P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P></DIV>
		<P class=entryFooter><% 
If LegacyMode <> True Then Response.Write "<acronym title=""Printer Friendly Version""><a href=""javascript:PrintPopup('Printer_Friendly.asp?Entry=" & RecordID & "')""><Img Border=""0"" Src=""Images/Print.gif""></a></acronym>"
If (EnableEmail = True) AND (LegacyMode <> True) Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%>
		<A class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=NewTime%></A>
		<% 
		If EnableComments <> False Then Response.Write " | <SPAN class=""comments""><A href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</A></SPAN>"
		 If (ShowCat <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write "| <SPAN class=categories>#<A href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</A></SPAN>"%>
		</P></DIV>
		<!--- End Content --->
		<%
		PreviousDay = DayPosted
		Records.MoveNext
		If JustDoIt = True Then Response.Write "</Div>"
		Loop

'--- Close The Records & Database ---
Records.Close

End If

If nRecCount < 1 Then
%>
<!--- Start No|Invalid Text / Default Pageload / EOF Content --->
<DIV class=entry>
<h3 class=entryTitle><% If Request("Search") = "" Then Response.Write "No Text entered" Else Response.Write "No Entries Found With That Criteria"%></h3><br>
<DIV class=entryBody>

<% If (Request("Search") <> "") AND (Requested = "") Then %>
<P align="center">You cannot perform a search on the criteria you selected.</P>
<% ElseIf (Requested = "") AND (Request("Search") = "") Then %>
<P align="center">Welcome, Please enter your query below<br>
<% ElseIf (nRecCount < 1) AND (Requested <> "") Then %>
<P align="center">No Entries found with the criteria "<b><%=Requested%></b>" in either the text's sentences <b>OR</b> <% If Request("Mode") <> "Any" Then Response.Write "the title" ELSE Response.Write "in any order in the text"%>.<br></P>
<% End If
If Request("Search") <> "" Then Response.Write "<P align=""center"">Please try again with a different criteria:" 
%>

<form name="Search" method="post" action="Search.asp">
<input Name="Search" Type="text" value="<%=Replace(Requested,"""","&quot;")%>" size="13" maxlength="70"><input Type="submit" value="Search"><br>
<br>Words Can Be In Any Order In The Text (so long as they <b>ALL</b> appear) : <input Name="Mode" Type="Checkbox" Value="Any" <%If Request("Mode") = "Any" Then Response.Write "CHECKED"%>>
<br>Don't Complete words ("Holl" wont return "Holly") : <input Name="NoAutoComplete" Type="Checkbox" Value="True" <%If Request("NoAutoComplete") = "True" Then Response.Write "CHECKED"%>>
</form>

</P>
</Div>
</Div>

<!--- End No Text Content --->
<% End If%>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
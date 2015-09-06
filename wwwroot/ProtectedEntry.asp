<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>
<%
Dim Requested, NewTime

Requested = Request.Querystring("Entry")
If (IsNumeric(Requested) = False) OR (Len(Requested) = 0) Then Requested = 0

'--- Open set ---'
Records.Open "SELECT RecordID, Title, Text, Category, Password, Day, Month, Year, Time, Comments, Enclosure FROM Data WHERE RecordID=" & Requested,Database, 1, 3

If NOT Records.EOF Then

'--- Setup Variables ---'
   Dim RecordID, Title, Text, CommentsCount, Password
   Dim DayPosted, MonthPosted, YearPosted, TimePosted, Enclosure

   RecordID = Records("RecordID")
   Title = Records("Title")
   Text = Records("Text")
   Category =  Records("Category")
   CommentsCount = Records("Comments")
   Password = Records("Password")

   DayPosted =  Records("Day")
   MonthPosted =  Records("Month")
   YearPosted =  Records("Year")
   TimePosted =  Records("Time")

   If (Len(Password) > 0) AND (Ucase(Password) <> Ucase(Request.Querystring("Password"))) Then
   Text = "<form action=""ProtectedEntry.asp"" method=""GET""><center>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & Requested & """>" & VbCrlf   
   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br/>" & VbCrlf
	If Len(Request.Querystring("Password")) > 0 Then Text = Text & "<b>You have entered an incorrect password</b><br/>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br/><br/>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
   Text = Text & "</center></form>"
   End If

'--- We're British, Let's 12Hour Clock Ourselves ---'
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
%>
<!--- Start ID Header --->
<DIV class=date id=<%=YearPosted%>-<%=MonthPosted%>-<%=DayPosted%>>
<H2 class=dateHeader>Permanant Link For Entry #<%=RecordID%></H2>
<!--- End ID Header --->

<!--- Start Content For --->
<DIV class=entry>
<H3 class=entryTitle><%=Title%> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Last Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym>"%></H3>
<DIV class=entryBody><P><%
        Response.Write LinkURLs(Replace(Text, vbcrlf, "<br/>" & vbcrlf))
        If (Enclosure <> "") AND (Len(Password) = 0) Then
	      If Instr(Enclosure,"http://") = 0 Then Enclosure = "Sounds/" & Enclosure
	      Response.Write "<br/><br/><br/>"
	 %>
	 <!-- Start Podcast Object (http://www.skylab.ws/?p=116) -->
	 <object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="192" height="29" id="mp3player" align="middle">
	 <param name="allowScriptAccess" value="sameDomain" />
	 <param name="movie" value="Includes/mp3player.swf?id=1.2" />
	 <param name="quality" value="high" />
	 <param name="bgcolor" value="#ffffff" />
	 <param name=FlashVars value="zipURL=<%=Enclosure%>&songURL=<%=Enclosure%>">
	 <embed src="Includes/mp3player.swf?id=1.2" FlashVars="zipURL=<%=Enclosure%>&songURL=<%=Enclosure%>" quality="high" bgcolor="#ffffff" width="192" height="29" name="mp3player" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
	</object>
	 <!-- End Podcast Object -->
	<br/><small><a href="<%=enclosure%>">Download this file</a></small>
	<%
	 End If
  %></P></DIV>
<P class=entryFooter>
<acronym title="Printer Friendly Version""><a href="javascript:PrintPopup('Printer_Friendly.asp?Entry=<%=RecordID%>&Password=<%=Request.Querystring("Password")%>')"><Img Border="0" Src="Images/Print.gif"></a></acronym> <% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%>
<b><%=DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime%></b>
<% If EnableComments <> False Then Response.Write " | <SPAN class=""comments""><a href=""Comments.asp?Entry=" & RecordID & """>Comments [" & CommentsCount & "]</a></SPAN>"%>
<% If (ShowCategories <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <SPAN class=""categories"">#<A href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</A>"%></SPAN></P></DIV>
<!--- End Content --->
</Div>
<%Else
'--- We're British, Let's 12Hour Clock Ourselves ---'
If TimeFormat <> False Then
If Hour(Time()) > 12 Then 
NewTime = Hour(Time()) - 12 & ":"
Else
NewTime = Hour(Time()) & ":"
End If
 
If Minute(Time()) < 10 Then
NewTime = NewTime & "0" & Minute(Time())
Else
NewTime = NewTime & Minute(Time())
End If

If (Hour(Time()) < 12) AND (Hour(Time()) <> 12) Then
NewTime = NewTime & " AM"
Else
NewTime = NewTime & " PM"
End If

Else
If Hour(Time()) < 10 Then NewTime = "0"
NewTime = NewTime & Hour(Time()) & ":"
If Minute(Time()) < 10 Then NewTime = NewTime & "0"
NewTime = NewTime & Minute(Time())
End If
%>
<!--- Start EOF Content --->
<DIV class=entry>
<H3 class=entryTitle>Error</H3>
<DIV class=entryBody><p>Sorry, The Record Number You Requested Was Either Invalid Or Has Been Removed.</p>
<p align="Center"><a href="<%=PageName%>">Back To The Main Page</a></p>
</DIV>
<P class=entryFooter><%=NewTime%> 
| <SPAN class=comments><A href="Mail.asp?Whatever happened to record <%=Requested%>?">Report Error</A></SPAN> 
| <SPAN class=categories>#Error</SPAN></P></DIV>
<!--- End EOF Content --->
<%End If
'--- Close The Records ---
Records.Close
%>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
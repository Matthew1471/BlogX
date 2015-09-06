<% 
OPTION EXPLICIT
AlertBack = True 
Response.Buffer = True
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>
<%           
Dim Ban
Dim DelRecNo
Dim EntryID 'EntryID
Dim Requested

	'### Filter & Clean ###
	EntryID = Request.Form("EntryID")
	If (IsNumeric(EntryID) = False) OR (EntryID = "") Then EntryID = 0 Else EntryID = Int(EntryID)

	Requested = Request.Querystring("Entry")
	If (IsNumeric(Requested) = False) OR (Requested = "") Then Requested = 0 Else Requested = Int(Requested)

	Ban = Request.Querystring("Ban")
	Ban = Replace(Ban,"'","")
    
    	DelRecNo = Request.Querystring("Delete")
    	If (IsNumeric(DelRecNo) = False) OR (DelRecNo = "") Then DelRecNo = 0 Else DelRecNo = Int(DelRecNo)

If EntryID <> 0 Then

'--- Open set ---'
    Records.CursorLocation = 3 ' adUseClient
    Dim LastIP
    Records.Open "SELECT CommentID, IP FROM Comments WHERE EntryID="& EntryID & " ORDER BY CommentID DESC",Database, 1, 3
    If Records.EOF = False Then LastIP = Records("IP")
    If EnableComments = False OR LastIP = Request.ServerVariables("REMOTE_ADDR") Then
    Records.Close
    Database.Close
    Set Records = Nothing
    Set Database = Nothing
    Response.Clear
    Response.Redirect(PageName)
    End If
    
    Records.Close

End If

If Request.Form("Action") <> "Post" Then

'--- Open set ---'
Records.Open "SELECT * FROM Data WHERE RecordID=" & Requested,Database, 1, 3

If NOT Records.EOF Then

Dim RecordID, Title, Text, Password, DayPosted, MonthPosted, YearPosted, TimePosted
Dim NewTime

'--- Setup Variables ---'
   RecordID = Records("RecordID")
   Title = Records("Title")
   Text = Records("Text")
   Category = Records("Category")
   Password = Records("Password")

   DayPosted =  Records("Day")
   MonthPosted =  Records("Month")
   YearPosted =  Records("Year")
   TimePosted =  Records("Time")

   If Len(Password) > 0 Then
   Text = "<form action=""ProtectedEntry.asp"" method=""GET""><center>" & VbCrlf
   Text = Text & "<input type=""hidden"" name=""Entry"" value=""" & RecordID & """>" & VbCrlf 
   Text = Text & "<img src=""Images/Key.gif""> Password Protected Entry <br>" & VbCrlf
   Text = Text & "This post is password protected. To view it please enter your password below:"
   Text = Text & "<br><br>Password: <input name=""Password"" type=""text"" size=""20""> <input type=""submit"" name=""Submit"" value=""Submit"">" & VbCrlf
   Text = Text & "</center></form>"
   End If

Records.Close

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
<H2 class=dateHeader>Comments For Entry #<%=RecordID%></H2>
<!--- End ID Header --->

<!--- Start Content For (Entry #<%=RecordID%>) --->
<DIV class=entry>
<H3 class=entryTitle><%=Title%> <%If Session(CookieName) = True Then Response.Write " <acronym title=""Edit Your Entry""><a href=""Admin/EditEntry.asp?Entry=" & RecordID & """><Img Border=""0"" Src=""Images/Edit.gif""></a></acronym> "%>(<a href="RSS/Comments/?Entry=<%=Requested%>">Comments RSS</a>)</H3>
<DIV class=entryBody><P><%=LinkURLs(Replace(Text, vbcrlf, "<br>" & vbcrlf))%></P></DIV>
<P class=entryFooter>
<acronym title="Printer Friendly Version""><a href="javascript:PrintPopup('Printer_Friendly.asp?Entry=<%=RecordID%>')"><Img Border="0" Src="Images/Print.gif"></a></acronym> <% If EnableEmail = True Then Response.Write "<acronym title=""Email The Author""><a href=""Mail.asp?" & HTML2Text(Title) & """><Img Border=""0"" Src=""Images/Email.gif""></a></acronym>"%>
<A class="permalink" href="ViewItem.asp?Entry=<%=RecordID%>"><%=DayPosted & "/" & MonthPosted & "/" & YearPosted & " " & NewTime%></A> 
<% If (ShowCat <> False) AND (Category <> "") AND (IsNull(Category) = False) Then Response.Write " | <SPAN class=""categories"">#<A href=""ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" & Replace(Category, "%20", " ") & "</A>"%></SPAN></P></DIV>
<!--- End Content --->
<%
'--- Open set ---'
Records.CursorLocation = 3 ' adUseClient

    '-- Check If We Are Banned --'
    Records.Open "SELECT * FROM BannedIP WHERE IP='" & Request.ServerVariables("REMOTE_ADDR") & "';",Database, 1, 3
    Dim Banned
    If Records.EOF = False Then Banned = True
    Records.Close
                                                                                                                            
	If IsNumeric(DelRecNo) = False Then
	Database.Close
	Set Records = Nothing
	Set Database = Nothing
	Response.Clear
	Response.Redirect(PageName)
	End If
                                                                                                                        
    If (DelRecNo <> 0) AND (Session(CookieName) = True) Then
    
    	Records.Open "SELECT * FROM Comments WHERE EntryID=" & Requested & " AND CommentID=" & DelRecNo & ";",Database, 1, 3
    	
    	If NOT Records.EOF Then 
    	Database.Execute "DELETE FROM Comments WHERE CommentID=" & DelRecNo

    	'### Write In Comment Count ###'
    	Records.Close
    	Records.Open "SELECT RecordID, Comments FROM Data WHERE RecordID=" & Requested,Database
        Records("Comments") = Records("Comments") - 1
    	Records.Update
    	
    	End If
    	
    	Records.Close
    	Database.Close
	Set Records = Nothing
    	Set Database = Nothing

    Response.Clear
    Response.Redirect("Comments.asp?Entry=" & Requested) 

    End If

    On Error Resume Next
    If (Ban <> "") AND (Session(CookieName) = True) Then Database.Execute "INSERT INTO BannedIP (IP) VALUES ('" & Ban & "')" & ";"
    On Error Goto 0

Records.Open "SELECT * FROM Comments WHERE EntryID=" & Requested & ";",Database, 1, 3

'****************************************************************
' Get Records Count
Dim nRecCount
nRecCount = Records.RecordCount
                         
' Loop through records until it's a next page or End of Records
Dim CommentID, Email, DateCommented, TimeCommented, NewDate
Dim AlreadySubscribed

Do Until (Records.EOF)

'--- Setup Variables ---'
   Set CommentID = Records("CommentID")
   Set Name = Records("Name")
   Set Email = Records("Email")
   Set Homepage =  Records("Homepage")
   Set Content =  Records("Content")
   Set Subscribe = Records("Subscribe")
   
   
   If Subscribe = True AND Records("IP") = Request.ServerVariables("REMOTE_ADDR") Then AlreadySubscribed = True

   Set DateCommented =  Records("Date")
   Set TimeCommented = Records("Time")

'--- We're British, Let's 12Hour Clock Ourselves ---'
NewTime = ""

If TimeFormat <> False Then
If Hour(TimeCommented) > 12 Then 
NewTime = Hour(TimeCommented) - 12 & ":"
Else
NewTime = Hour(TimeCommented) & ":"
End If
 
If Minute(TimeCommented) < 10 Then
NewTime = NewTime & "0" & Minute(TimeCommented)
Else
NewTime = NewTime & Minute(TimeCommented)
End If

If (Hour(TimeCommented) < 12) AND (Hour(TimeCommented) <> 12) Then
NewTime = NewTime & " AM"
Else
NewTime = NewTime & " PM"
End If

Else
If Hour(TimeCommented) < 10 Then NewTime = "0"
NewTime = NewTime & Hour(TimeCommented) & ":"
If Minute(TimeCommented) < 10 Then NewTime = NewTime & "0"
NewTime = NewTime & Minute(TimeCommented)
End If

NewDate = Day(DateCommented) & "/" & Month(DateCommented) & "/" & Year(DateCommented)
%>
<!--- Start Content For Comment <%=CommentID%> --->
<div class="comment">
<h3 class="commentTitle"><%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Ban User""><a href=""Comments.asp?Entry=" & Requested & "&Ban=" & Records("IP") & """><Img Border=""0"" Src=""Images/Color.gif""></a></acronym> <acronym title=""Delete Comment""><a href=""Comments.asp?Entry=" & Requested & "&Delete=" & CommentID & """><Img Border=""0"" Src=""Images/Key.gif""></a></acronym>"%><%=NewDate%>&nbsp;<%=NewTime%></h3>
<span class="commentBody"><%=LinkURLs(Replace(Content, vbcrlf, "<br>" & vbcrlf))%></span>
<p class="commentFooter"><%If HomePage <> "" Then Response.Write "<a class=""permalink"" href=""" & Homepage & """>"%><%=Name%><%If HomePage <> "" Then Response.Write "</a>"%>
<%If (Session(CookieName) = True) AND (Email <> "") Then Response.Write " | <span class=""comments""><a href=""mailto:" & Email & """>" & Email & "</span></a>"%></p>
</div>
<!--- End Content --->
<%
LastIP = Records("IP")
Records.MoveNext
Loop      

'--- Close The Records & Database ---
Records.Close
%>
</div>
<%
'--- Open RecordSet ---'
Records.CursorLocation = 3 ' adUseClient

    '-- Check If We Are Banned --'
    Records.Open "SELECT SourceURI FROM Pingback WHERE EntryID=" & Requested & ";",Database, 1, 3
    If Records.EOF = False Then%>
<!--- Start Content For PingBacks --->
<div class="comment">
<h3 class="commentTitle">Pingbacks For Entry #<%=Requested%></h3>
<span class="commentBody">
<UL>
<%
Do Until (Records.EOF)
Response.Write "<LI><a href=""" & Records("SourceURI") & """>" & Records("SourceURI") & "</a></LI><br>"
Records.MoveNext
Loop
%>
</UL>
</span>
<p></p>
</div>
<!--- End Content --->
<%
    End If

%>
                         
<div id="AddNew" class="date">
                    <div class="comment">
                    <h3 class="commentTitle">Add New Comment</h3>
                    <div class="commentBody">
<% If Banned = True Then%>
<Center><B>You Have Been Banned From Making Comments!</B></Center>
<% ElseIf EnableComments <> True Then %>
<Center><B>Comments have been disabled by the Blog administrator!</B></Center>
<% ElseIf LastIP <> Request.ServerVariables("REMOTE_ADDR") Then %>
                        <Form Name="AddComment" Method="Post" Action="Comments.asp" onSubmit="return setVar()">
                        <input Name="Action" type="hidden" Value="Post">  
                        <input Name="EntryID" type="hidden" Value="<%=RecordID%>">
                            <p><span id="Label1">Name</span><Input Name="Name" Type="text" Value="<%=Request.Cookies("Visitor")("Name")%>" maxlength="50"></p>
                            <p><span id="Label2">E-mail</span><Input Name="Email" Type="text" Value="<%=Request.Cookies("Visitor")("Email")%>" maxlength="50"></p>
                            <p><span id="Label3">Homepage</span><Input name="Homepage" type="text" Value="<%=Request.Cookies("Visitor")("Homepage")%>" maxlength="50"></p>
                            <p><Input Name="RememberMe" Type="checkbox" Checked="True" Value="True">Remember Me
<% If (AlreadySubscribed = False) AND (LegacyMode <> True) AND (Session(CookieName) <> True) Then%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<Input Name="Subscribe" Type="checkbox" Checked="True" Value="True">E-mail me replies</p>
<% End If %>
                            <p><span id="Label4">Content (HTML not allowed)</span></p>
                            <p><Textarea name="Content" rows="12" cols="40" onChange="return setVarChange()"></textarea></p>
                            <p><Input type="submit" Value="Add Comment"></p>
                        </form>
<% Else %>
<Center><B>You Were Already The Last Person To Comment!</B></Center>
<% End If %>
                    </div>
                </div>
              </div>
</Div>

<%
Else
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

Records.Close

Else

'Dimension variables
Dim Subscribe
Subscribe = Request.Form("Subscribe")
If Subscribe = "" Then Subscribe = False

'### Did We Type In Name? ###'
If Request.Form("Name") = "" Then
Response.Write "<p align=""Center"">No Name Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%
Response.End
End If

'### Did We Type In Text? ###'
If Request.Form("Content") = "" Then
Response.Write "<p align=""Center"">No Text Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%
Response.End
End If

Randomize Timer

'## Anti-HTML ###
Dim Content
Content = Request.Form("Content")
Content = Replace(Content, "<","&lt;")
Content = Replace(Content, ">","&gt;")

Name = Request.Form("Name")
Name = Replace(Name, "<","&lt;")
Name = Replace(Name, ">","&gt;")

'## Add a http:// to non http'd links ##
Dim Homepage
Homepage = Request.Form("Homepage")
If (Instr(Homepage,"http://") = 0) AND (Len(Homepage) > 0) Then Homepage = "http://" & Homepage

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3

'### Write In Comments ###'
Records.Open "SELECT * FROM Comments", Database
Records.AddNew
Records("EntryID") = EntryID
Records("Name") = Left(Name,50)
Records("Email") = Left(Request.Form("Email"),50)
Records("Homepage") = Left(Homepage,50)
Records("Content") = Content
Records("Subscribe") = Subscribe
Records("PUK") = Int(Rnd()*99999999)

Records("Date") = DateValue(DateAdd("h",TimeOffset,Now()))
Records("Time") = TimeValue(DateAdd("h",TimeOffset,Time()))

Records("IP") = Request.ServerVariables("REMOTE_ADDR")

Records.Update
Records.Close

'### Write In Comments ###'      
Records.Open "SELECT RecordID, Comments FROM Data WHERE RecordID=" & EntryID, Database
Records("Comments") = Records("Comments")+1
Records.Update
Records.Close

	'#### Honor Subscriptions ####
	If EnableEmail <> False Then

	Dim MailBody, ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg
	Records.Open "SELECT * FROM Comments WHERE Subscribe=True AND EntryID=" & EntryID & " AND Email <> '" & Request.Form("Email") & "'",Database, 1, 3
	Do Until (Records.EOF)

	MailBody = "<html>" & VbCrlf
	MailBody = MailBody & "<head>" & VbCrlf
	MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
	MailBody = MailBody & "</head>" & VbCrlf
	MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

	MailBody = MailBody & "<br>" & VbCrlf
	MailBody = MailBody & "<DIV class=content>" & VbCrlf
	MailBody = MailBody & "<center>" & VbCrlf

	MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
	MailBody = MailBody & "<H3 class=entryTitle>Notification Of Comment Added</H3>" & VbCrlf
	MailBody = MailBody & "<DIV class=entryBody>" & VbCrlf

	MailBody = MailBody & "<p>You are recieving this e-mail as a user has submitted a <a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & """>comment</a> on " & SiteDescription & ".</p>" & VbCrlf
	MailBody = MailBody & "</DIV>" & VbCrlf
	MailBody = MailBody & "</DIV>" & VbCrlf

	MailBody = MailBody & "<p>To stop recieving update notification for this entry, click <a class=""standardsButton"" href=""" & SiteURL & "CommentNotify.asp?Entry=" & EntryID & "&Email=" & Records("Email") & "&PUK=" & Records("PUK") & """>Unsubscribe</a></p>" & VbCrlf

	MailBody = MailBody & "<p>BlogX V" & Version & "</p>" & VbCrlf

	MailBody = MailBody & "</Center>" & VbCrlf
	MailBody = MailBody & "</DIV>" & VbCrlf

	MailBody = MailBody & "</html>" & VbCrlf

			ToName = Records("Name")
			ToEmail = Records("Email")
			From = EmailAddress
			Name = SiteDescription

			Subject = "Blog : Comment Added (Entry #" & EntryID & ")"
			Body = MailBody
	%>
	<!--#INCLUDE FILE="Includes/Mail.asp" -->
	<%
	Records.MoveNext
	Loop
	Records.Close
	End If
	'### End Of Subscriptions ####

If (CommentNotify <> 0) AND (Session(CookieName) <> True) AND (EnableEmail <> False) Then

MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>Notification Of Comment Added</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody>" & VbCrlf

MailBody = MailBody & "<p>You are recieving this e-mail as a user has submitted a <a href=""" & SiteURL & "Comments.asp?Entry=" & EntryID & """>comment</a> on " & SiteDescription & ".</p>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>BlogX V" & Version & "</p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			ToName = "Webmaster"
			ToEmail = EmailAddress
			If Request.Form("Email") <> "" Then From = Request.Form("Email") Else From = EmailAddress
			Name = SiteDescription

			Subject = "Blog : Comment Added"
			Body = MailBody
%>
                        <!--#INCLUDE FILE="Includes/Mail.asp" -->
<%
End If

Response.Write "<p align=""Center"">Comment Submission Successful</p>"
Response.Write "<p align=""Center""><a href=""Comments.asp?Entry=" & EntryID & """>Back</font></a></p>"

If Request.Form("RememberMe") = "True" then
Response.Cookies("Visitor")("Name") = Request.Form("Name")
Response.Cookies("Visitor")("Email") = Request.Form("Email")
Response.Cookies("Visitor")("Homepage") = Request.Form("Homepage")
Response.Cookies("Visitor").Expires = "July 31, 2008"
End If

End If

%>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
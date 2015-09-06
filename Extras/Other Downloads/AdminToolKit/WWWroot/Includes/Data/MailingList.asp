<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>
<% If Request.Form("Action") <> "Post" Then %>
<DIV class=entry>
<h3 class=entryTitle>What Is "<%=SiteName%>" Mailing List</h3><br>
<DIV class=entryBody>
<% If InStr(1,Request.ServerVariables("HTTP_Host"),"blogx.co.uk",1) <> 0 Then%>
<P>This mailing list will notify you the <b><u>MOMENT</u></b> there is a new Matthew1471 BlogX Version available for release.</P>
<p>Click <a href="Download.asp">here</a> to download BlogX.</p>

<p>This will keep up to date on the latest information, news and other information.</p>
<P>Note : You are free to unsubscribe at any time, I will not disclose your e-mail addresses to any third parties.</P>
<% Else %>
<!--- *************** CHANGE ME *********************************************************** --->
<p align="center">If ever there was a time for the site owner to open up "MailingList.asp"<br>
in Notepad and tell people what this mailing list is actually about, now would be it
<img src="images/emoticons/wink.gif">.
</p>
<!--- *************** Leave the rest alone ************************************************ --->
<%End If %>

<form Method="Post">
<Input Name="Action" Type="Hidden" Value="Post">
<center><b>Your Email Address : </b> <Input Name="EmailAddress" type="Text" size="40" maxlength="40">
<Input Type="submit" Value="Subscribe"></center>
</Form>

<Center><B><font Color="Lime">Important : You can only subscribe ONCE from ONE IP address<br>"Test" e-mail subscriptions invalidate your ability to join in future!</font></B></Center>
</Div>
</Div>

<% Else
'Dimension variables
Dim EntryCat            'Category
Dim Email
Dim Returned

Email = Request.Form("EmailAddress")
Email = Replace(Email,"'","")

' ---------- Did We Forget to type our address?  ----------
If Email = "" Then
Response.Write "<p align=""Center"">No E-mail Address Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%
Response.End
End If

' ---------- Test if the form was properly filled in ----------
If InStr(Email,"@") = 0 Then
Response.Write "<p align=""Center"">Invalid E-mail Address Entered</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%
Response.End
End If

' ---------- Test if the user is AOL ----------
If InStr(Email,"aol.com") <> 0 Then
Response.Write "<p align=""Center"">AOL is banning my server, Please use a different address</p>"
Response.Write "<p align=""Center""><a href=""javascript:history.back()"">Back</font></a></p>"
Response.Write "</Div>"
%>
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<%
Response.End
End If

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM MailingList WHERE SubscriberAddress ='" & Email & "' OR SubscriberIP='" & Request.ServerVariables("REMOTE_ADDR") & "'", Database

If Records.EOF Then
Records.AddNew

Records("SubscriberAddress") = Request.Form("EmailAddress")
Records("SubscriberIP") = Request.ServerVariables("REMOTE_ADDR")
Randomize
Records("PUK") = Int(Rnd()*99999999)
Records("Active") = True

Records.Update
Returned = "Mailing List Subscription Successful"

Dim MailBody

MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>Notification Of Subscription</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody align=""left"">" & VbCrlf

If InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0 Then
MailBody = MailBody & "<p>Welcome to the new BlogX mailing list designed to help collaborate BlogX'ers so together we can fix bugs & be kept up to date.</p>" & VbCrlf
MailBody = MailBody & "<p>I guess the first thing you want to do is <a href=""http://blogx.co.uk/Download.asp"">Download BlogX</a>.</p>" & VbCrlf
End If

MailBody = MailBody & "<p>You are recieving this e-mail as you, or someone posing as you, have subscribed to be notified of updates on " & SiteDescription & ".</p>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>To Unsubscribe click <a class=""standardsButton"" href=""" & SiteURL & "Unsubscribe.asp?Email=" & Request.Form("EmailAddress") & "&PUK=" & Records("PUK") & """>Unsubscribe</a></p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			Dim ToName, ToEmail, From, Subject, Body, iConf, Flds, Mail, Err_Msg

			ToName = "Member Of " & SiteDescription
			ToEmail = Request.Form("EmailAddress")
			From = EmailAddress
			Name = SiteDescription

			Subject = "Blog : Subscription"
			Body = MailBody
%>
                        <!--#INCLUDE FILE="Includes/Mail.asp" -->
<%
'---- Notify Webmaster ----'
MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>Notification Of Subscription</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody>" & VbCrlf

MailBody = MailBody & "<p>A new member <b>" & Request.Form("EmailAddress") & "</b> has subscribed to " & SiteDescription & "'s Mailing list.</p>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>To view all subscriptions click <a class=""standardsButton"" href=""" & SiteURL & "Admin/MailingListMembers.asp"">Admin</a></p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

			ToName = "Webmaster"
			ToEmail = EmailAddress
			From = Request.Form("EmailAddress")
			Name = SiteDescription

			Subject = "Blog : New Member"
			Body = MailBody
%>
                        <!--#INCLUDE FILE="Includes/Mail.asp" -->
<%
Else
Returned = "You (Or Someone Using Your IP) Are Already Subscribed"
Returned = Returned & VbCrlf & "<br><br>Please <a href=""Mail.asp"">Contact</a> the Webmaster if you have changed / or need to change your address."
End If

'#### Close Objects ###
Records.Close

Response.Write "<p align=""Center"">" & Returned & "</p>"
Response.Write "<p align=""Center""><a href=""" & PageName & """>Back</font></a></p>"

End If
%>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
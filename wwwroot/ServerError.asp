<%
' --------------------------------------------------------------------------
'¦Introduction : Server Error Page                                          ¦
'¦Purpose      : Shows the error message following the consistent theme     ¦
'¦               and more importantly notifies the webmaster.               ¦
'¦Used By      : Your webserver if you have configured it to.               ¦
'¦Requires     : None (As many include files are included as possible)      ¦
'¦Notes        : This page is one of the most useful pages for blogx.co.uk, ¦
'¦               but it requires configuration on your part. Also this page ¦
'¦               may become out of date quickly due to the amount of        ¦
'¦               replicated unmaintained code.                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT 

  On Error Resume Next
  Response.Clear

  Response.Status = "500 Internal Server Error"

  Dim objError
  Set objError = Server.GetLastError()

Dim SiteURL, Version, RSS
Dim AboutPage, ArgoSoftMailServer, MailingList
Dim OtherLinks, Register, Copyright
Dim ShowCat, ShowMonth, Polls, EmailAddress
Dim SiteName, SiteDescription, SiteSubTitle, SortByDay
Dim BackgroundColor, TimeFormat, EnableMainPage
Dim LegacyMode, Template, PageName, Timeoffset
Dim NoDate
Dim AllowEditingLinks, EnableEmail, EmailServer, EmailComponent

'-- Mailer Specific DIMS --'
Dim SendOk, gMDUser, mbDllLoaded, gMDMessageInfo

'***********************************************'
'Your DatabasePath & Settings
'***********************************************'
SiteURL = "http://blogx.co.uk/"
Version = "1.0.7.02"

AboutPage = True          'Use your About.asp file...or if false, link to the one on BlogX.co.uk instead
AllowEditingLinks = 1	  'Setting to 1 allows a logged in user to edit the links online
ArgoSoftMailServer = 1    'If your server is running "ArgosoftMailServer" you can post from e-mail
MailingList = 1           'Allows you to run a mailinglist and displays the link at the bottom.
LegacyMode = False 	  'This will remove all functionality that wasn't in the ORIGINAL BlogX
NoDate = 0		  'Stops separating entries by day
OtherLinks = 1		  'Display the section "Other Links"
Register = False	  'Notifies BlogX.co.uk That You want Your Site Addded To The "Who Uses" page
RSS = 1                   'Lets people access your small RSS feed
TimeOffset = 0            'Change the time by this many hours e.g. "6" adds 6 hours to the server's time

'***********************************************'
'END OF EDITING
'***********************************************'

	'Read in the configuration details from the recordset
	Copyright       = "Content © 2008 Matthew Roberts"
	
	' This is used in reporting the error and must be valid
	EmailAddress 	= "webmaster@yoursite.co.uk"
	EnableEmail     = True
        EmailServer     = "localhost"
        EmailComponent  = "cdosys"
	
	Polls           = True
	ShowCat         = True
        ShowMonth       = True
	SiteName        = "<Strong>Matthew1471's</Strong> <Font Color=""Black"">BlogX</Font>"
	SiteDescription = "Matthew1471's BlogX"
        SiteSubTitle    = "Please Leave Comments! :-)"
        SortByDay       = True
        BackgroundColor = "darkorange"
        TimeFormat      = True

        EnableMainPage  = True
        Template        = "Swimming Pool"
        
If EnableMainPage <> True Then PageName = "Default.asp" Else PageName = "Main.asp"

If Instr(Request.Querystring(),"LegacyMode") <> 0 Then LegacyMode = True

If LegacyMode = True Then
ArgoSoftMailServer = 0
BackgroundColor = "#FFFFFF"
CommentNotify = 0
NotifyPingOMatic = 0
NoDate = 0
MailingList = 0
Polls = 0
SortByDay = True
Template = "Default"
End If

Session("CookieTest") = "AOK"

Response.AddHeader "X-Pingback",SiteURL & "RSS/PingBack/Default.asp"

Dim szYearMonth, szPos
Dim nYear, nMonth, nDay, SpecificRequest

szYearMonth = Request("YearMonth")
szPos = Request("POS")
nDay = Request("Day")

If szYearMonth = "" Then
SpecificRequest = False
nYear = Year(Now())
nMonth = Month(Now())
Else
SpecificRequest = True
nYear = Left(szYearMonth,4)
nMonth = Int(Right(szYearMonth,2))
End If

'### SQL Attacker Exploit Management ###'
If (IsNumeric(nYear) <> True) OR (IsNumeric(nMonth) <> True) OR (IsNumeric(nDay) <> True) Then Response.Redirect("Hacker.asp")
%>
<HTML>
<HEAD>
<TITLE><%=SiteDescription%> - ASP Error</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
// 
// Usage Of This Software Is Subject To The Terms Of The License
//= - - - - - - -
-->
<SCRIPT LANGUAGE="Javascript">
<!-- begin hiding
  if (parent.frames.length)
  top.location.href= document.location;

function PrintPopup(url) {
  popupWin = window.open(url,'Print','width=450,height=525,scrollbars=yes,toolbar=yes,menubar=yes,resizable=yes')
}

// done hiding -->
</SCRIPT>
<%
If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")

Dim CalendarBackground

Select Case Template
  Case "Black"
    CalendarBackground = "fuchsia"
  case "Default"
    CalendarBackground = "Silver"
  Case "Clouds"
    CalendarBackground = "DodgerBlue"
  Case "Matrix"
    CalendarBackground = "#003300"
  Case "Pebbles"
    CalendarBackground = "RoyalBlue"
  Case "Orange"
    CalendarBackground = "#ff6600"
  Case "Red"
    CalendarBackground = "DarkOrange"
  Case "Stary"
    CalendarBackground = "RoyalBlue"
  Case "Sea"
    CalendarBackground = "DarkBlue"
  Case "SkyBlue"
    CalendarBackground = "darkblue"
  Case "TotallyGreen"
    CalendarBackground = "#363"
  Case "WaterFall"
    CalendarBackground = "#536e55"
  Case Else
  CalendarBackground = ""
End select
%> 
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
<Link rel="Alternate" type="application/rss+xml" title="RSS" href="<%=SiteURL%>Rss/">
<link rel="pingback" href="<%=SiteURL%>RSS/PingBack/Default.asp">
<style><!--
/* image shadow */
.dropshadow {
  clear: both;
  float:left;
  background: url(<%=SiteURL%>Images/shadowAlpha.png) no-repeat bottom right !important;
  background: url(<%=SiteURL%>Images/shadow.gif) no-repeat bottom right;
  margin: 10px 13px 0 6px !important;
  margin: 20px 7px 0 3px;

}

.dropshadow img
{
	display: block;
	position: relative;
	background-color: #fff;
	border: 1px solid #a9a9a9;
	margin: -6px 6px 6px -6px;
	padding: 4px;
}

/* image shadow */
.dropshadowr {
  clear: both;
  float:right;
  background: url(<%=SiteURL%>Images/global/shadowAlpha.png) no-repeat bottom right !important;
  background: url(<%=SiteURL%>Images/global/shadow.gif) no-repeat bottom right;
  margin: 20px 6px 0 30px !important;
  margin: 30px 3px 0 20px;


}

.dropshadowr  img
{
	display: block;
	position: relative;
	background-color: #fff;
	border: 1px solid #a9a9a9;
	margin: -6px 6px 6px -6px;
	padding: 4px;
}
--></style>
</HEAD>
<BODY bgColor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<DIV id=header>
<H1 id=title><A style="TEXT-DECORATION: none" href="<%=SiteURL%>"><%=SiteName%></A></H1>
<P id=byline><%=SiteDescription%></P>
<P id=sideTitle><SPAN class=blogTitleSub><%=SiteSubTitle%></SPAN>
<SPAN class=blogTitleSubDisclaimer>Please read my <A href="<%=SiteURL%>Disclaimer.asp">disclaimer</A>.</SPAN>
</P></DIV>

<DIV id=content>
<%
  Dim ASPCode, COMCode, Refer, Description, FileName, FullDescription, LineNumber, Source
  Dim RequestedURL, RequestedQuerystring

  Refer = Request.ServerVariables("HTTP_REFERER")
  
  ASPCode = objError.ASPCode
  COMCode = objError.Number
  Description = objError.Description
  FileName = objError.File
  FullDescription = objError.ASPDescription
  LineNumber = objError.Line
  Source = objError.Source

  RequestedURL = Request.ServerVariables("URL")
  RequestedQuerystring = Request.ServerVariables("QUERY_STRING")
%>

<DIV class=entry>
<h3 class=entryTitle>Server Error</h3><br>
<DIV class=entryBody>
<embed src="<%=SiteURL & "Includes/"%>Error.mp3" autostart=True hidden=true>

<% If Len(CStr(Description)) > 0 Then %>
<P align="center">It appears that you have stumbled upon an error in the BlogX engine.<br>
It might be that this page is currently being editied</P>

<P align="center">Congratulations <img src="<%=SiteURL & "Images/Emoticons/Wink.gif"%>"></p>

<P align="center">In plain english, The file "<b><%=FileName%></b>" had a problem with Line <b><%=LineNumber%></b> because "<b><%=Description%></b>" happened</p>
<% End If %>

<P align="center">I bet you were doing something you weren't supposed to be doing (shame on you) <img src="<%=SiteURL & "Images/Emoticons/Grin.gif"%>"></p>

<P align="center">
<% If Len(CStr(Description)) > 0 Then Response.Write "<b>Error : </b> " & Description & "<br>" %>
<b>Referrer : </b> <%If Refer <> "" Then Response.Write "<a href=""" & Refer & """>" & Refer & "</a>" Else Response.Write "You Typed In The Address Manually"%><br><br>
<% 
If Len(CStr(ASPCode)) > 0 Then Response.Write "<b>IIS Error Number : </b> " & ASPCode & "<br>"
If COMCode < 0 Then Response.Write "<b>COM Error Number : </b>" & COMCode & " (0x" & Hex(COMCode) & ")" & "<br>"
If Len(CStr(Source)) > 0 Then Response.Write "<b>Error Source : </b> " & Replace(Source,"<","&lt;") & "<br><br>" 
If Len(CStr(FileName)) > 0 Then Response.Write "<b>File Name : </b> " & FileName & "<br>"
If LineNumber > 0 Then Response.Write "<b>Line Number : </b> " & LineNumber & "<br>"
If Len(CStr(FullDescription)) > 0 Then Response.Write "<b>Full Description : </b> " & FullDescription

If Len(CStr(RequestedURL)) > 0 Then Response.Write "<b>URL : </b> " & RequestedURL & "<br/>" & VbCrlf
If Len(CStr(RequestedQuerystring)) > 0 Then Response.Write "<b>Query : </b> " & RequestedQuerystring & "<br/>"

'-- Detect error page errors! --'
If Err.Description <> "" Then Response.Write "<br><b>Even this page errors :</b><br/>" & Err.Description
%>
</P>
<p align="Center"><a href="<%=SiteURL & PageName%>">Back To The Main Page</a></p>
</Div>
</Div>

</Div>

<DIV class=sidebar id=rightBar>
<BR>
<%
Dim nLastDay, n, nn, nnn, nDS, PostToday

If szPos <> "" Then
If szPos = "NEXT" Then
nDS = 1
Else
nDS = -1
End If
nDS = DateSerial(nYear, nMonth + nDS, 1)
nYear = Year(nDS)
nMonth = Month(nDS)
End If
nLastDay = Day(DateSerial(nYear, nMonth + 1, 1 - 1))
nDay = 1 - Weekday(DateSerial(nYear, nMonth, 1)) + 1
%>
<table class="navCalendar" cellspacing="0" cellpadding="4" border="0" style="border-width:1px;border-style:solid;border-collapse:collapse;">
<tr>
<td colspan="7" style="background-color:<%=CalendarBackground%>;">
<table class="navTitleStyle" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">
<tr>
<td class="navNextPrevStyle" style="width:15%;"><a href="?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>&POS=LAST" style="color:Black">&lt;</a></td>
<td align="Center" style="width:70%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>"><%=MonthName(nMonth)%></a> (<%=Right(nYear,2)%>)</td>
<td class="navNextPrevStyle" align="Right" style="width:15%;"><a href="?YearMonth=<%=nYear & Right("00" & nMonth, 2)%>&POS=NEXT" style="color:Black">&gt;</a></td>
</tr>
</table>

</td>
</tr>

<%
'### Write Out The Weekdays ###'

Response.Write "<tr>"

For n = 0 To 6
Response.Write "<td class=""navDayHeader"" align=""Center"">" & Left(WeekdayName(n + 1, True),1) & "</TD>" & CHR(13)
Next

Response.Write "</tr>"

'### Write Out Days ###'

For nn = 0 To 5
Response.Write"<TR>" & CHR(13)
For nnn = 0 To 6
If nDay > 0 And nDay <= nLastDay Then

Response.Write "<td class="""

'### Highlight CurrentDay/Weekend ###'
If nDay = Int(Request("Day")) Then
Response.Write "navSelectedDayStyle" 
ElseIf nnn = 0 or nnn = 6 Then Response.Write "navWeekendDayStyle"
Else Response.Write "navDayStyle"
End If
'### End Of Current Day Check ###'

Response.Write """ align=""Center"""

'### Highlight CurrentDay/Weekend ###'
If nDay = Int(Request("Day")) Then
Response.Write " style=""color:White;background-color:" & CalendarBackground & ";width:14%;"">"
Else 
Response.Write " style=""width:14%;"">"
End If
'### End Of Current Day Check ###' 

'### Lets Strip Out That Existing Day From Our Clicky ###'
If SortByDay = True Then Response.Write "<a href=""" & SiteURL & PageName & "?"
If SortByDay = True Then Response.Write "YearMonth=" & nYear & Right("00" & nMonth, 2) & "&Day=" & nDay & """>"

If (Day(DateAdd("h",TimeOffset,Now())) = nDay) AND (Month(DateAdd("h",TimeOffset,Now())) = nMonth) Then Response.Write "<font color=""red"">"
Response.Write nDay
If (Day(DateAdd("h",TimeOffset,Now())) = nDay) AND (Month(DateAdd("h",TimeOffset,Now())) = nMonth) Then Response.Write "</font>"

If (SortByDay = True) OR (PostToday = True) Then Response.Write "</a>"
'### Finished Day Stripping ###'

Response.Write "</TD>" & CHR(13)
Else
Response.Write "<Td class=""navOtherMonthDayStyle"" align=""Center"" style=""width:14%;"">-</TD>" & CHR(13)
End If
nDay = nDay + 1
Next
Response.Write "</TR>" & CHR(13)
Next
%>
<tr><td colspan="7" class="navCalendar" cellspacing="0" cellpadding="4" border="0" style="background-color:<%=CalendarBackground%>;" align="center"><A HREF="<%=SiteURL & PageName%>">This Month!</A></td></tr>
</TABLE>


<% If (ShowMonth <> False) AND (LegacyMode <> True) Then %>
<!--- Archive --->
<DIV class=section>
<H3>Archive</H3>
<UL>Not Available In "Safe Mode"</UL>
<BR></DIV><BR>
<%End If%>

<!--- Links --->
<DIV class=section>
<H3>Links</H3>
<UL>
  <% If EnableMainPage = True Then Response.Write "<LI><A href=""" & SiteURL & """>About Me</A></LI>"%>
  <LI><A href="<%=SiteURL & PageName%>">Blog Home</A></LI>
</UL></DIV><BR>

<% If OtherLinks <> 0 Then %>
<!--- Other Links --->
<DIV class=section>
<H3>Other Links</H3>
<UL><Li>Not Available In "Safe Mode"</Li></UL>
</DIV><BR>
<% End If %>

<% If Polls <> False Then %>
<!--- Poll --->
<DIV class=section>
<H3>Poll</H3>
 <UL>Not Available In "Safe Mode"</UL>
</DIV><BR>
 <% End If %>

<% If ShowCat <> False Then 
Dim Category, LastCat
%>
<!--- Categories --->
<DIV class=section>
<H3>Categories</H3>
<UL>
<Li>Not Available In "Safe Mode"</Li></UL>
</DIV><BR>
<%End If
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
<DIV class=section>   
<H3>Login</H3>
<UL>
<Li>Disabled In Safe Mode</Li>
</UL>
</Div>
</DIV>
<DIV id=footer>
<%' INSERT BANNERS, ADVERTS, LINKS, Pictures ETC HERE... %>
<A HREF="http://www.aspin.com"><IMG BORDER="0" ALT="Aspin.com" SRC="http://blogx.co.uk/Images/ASPin.gif"></A> 

<P id=copyright><%=Copyright%><% If RSS <> 0 Then 
Response.Write " | Subscribe to my <A class=""standardsButton"" href=""" & SiteURL & "RSS/"">RSS</A> feed"
End If
%>
<% If MailingList <> 0 Then Response.Write VbCrlf & "<br>Subscribe to my <A class=""standardsButton"" href=""" & SiteURL & "MailingList.asp"">Mailing List</A>"%>.</P>

<% '## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE Matthew1471 BlogX LICENSE AGREEMENT & MAY MEAN LEGAL ACTION ##' %>
<P id=poweredBy>Powered by <A href="<% If AboutPage = True Then Response.Write SiteURL & "About.asp" Else Response.Write "http://blogX.co.uk/About.asp"%>"><acronym title="Powered By: Matthew1471 BlogX Version V<%=Version%>">Matthew1471's edition of BlogX</acronym></A></P></DIV>
<% '## END - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE Matthew1471 BlogX LICENSE AGREEMENT & MAY MEAN LEGAL ACTION ##' %>

<%'--- This means you'll get added to "OtherLinks.txt" in future ---'
If Register = True Then Response.Write "<Img Src=""http://blogx.co.uk/Count.asp"" width=""1"" height=""1"">"

Response.Flush

'**********************************************'
'NOTIFY Mission Control
'**********************************************'
If Description <> "" Then

Dim ToName, ToEmail, From, Name, Subject, Body, iConf, Flds, Mail, MailBody, Err_Msg

			ToName = "Webmaster"
			ToEmail = "BlogXError@matthew1471.co.uk"
			From = EmailAddress
			Name = "Matthew1471's BlogX"
			Subject = "Error Report : " & Description & " (" & FileName & ")"

MailBody = "<html>" & VbCrlf
MailBody = MailBody & "<head>" & VbCrlf
MailBody = MailBody & "<Link href=""" & SiteURL & "Templates/" & Template & "/Blogx.css"" type=text/css rel=stylesheet>" & VbCrlf
MailBody = MailBody & "</head>" & VbCrlf
MailBody = MailBody & "<Body bgcolor=""" & BackgroundColor & """>" & VbCrlf

MailBody = MailBody & "<br>" & VbCrlf
MailBody = MailBody & "<DIV class=content>" & VbCrlf
MailBody = MailBody & "<center>" & VbCrlf

MailBody = MailBody & "<DIV class=entry style=""width: 50%"">" & VbCrlf
MailBody = MailBody & "<H3 class=entryTitle>Bug Report From " & SiteURL & "</H3>" & VbCrlf
MailBody = MailBody & "<DIV class=entryBody>" & VbCrlf

MailBody = MailBody & "<p> Hello there Matt, I've got another one of those damn bugs for you here..<BR>" & VbCrlf
MailBody = MailBody & "One of my users accessed the page <b><a href=""http://" & Request.ServerVariables("HTTP_HOST") & RequestedURL & "?" & RequestedQuerystring & """>http://" & Request.ServerVariables("HTTP_HOST") & RequestedURL & "?" & RequestedQuerystring & "</a></b> and the site just died</p>" & VbCrlf 
MailBody = MailBody & "<p> I'd appreciate if you could investigate, we're talking line <b>" & LineNumber & "</b>, with the error ""<b>" & Description & "</b>""."
MailBody = MailBody & "<br><br>TattyBye</p>" & VbCrlf

If Len(CStr(Description)) > 0 Then MailBody = MailBody &  "<P><b>Error : </b> " & Description & "<br>" & VbCrlf

MailBody = MailBody &  "<b>Referrer : </b> "
If Refer <> "" Then MailBody = MailBody & "<a href=""" & Refer & """>" & Refer & "</a>" Else MailBody = MailBody & "They Typed In The Address Manually"
MailBody = MailBody & "<br><br>" & VbCrlf

If Len(CStr(ASPCode)) > 0 Then MailBody = MailBody &  "<b>IIS Error Number : </b> " & ASPCode & "<br>" & VbCrlf
If COMCode < 0 Then MailBody = MailBody &  "<b>COM Error Number : </b> " & COMCode & " (0x" & Hex(COMCode) & ")" & "<br>" & VbCrlf
If Len(CStr(Source)) > 0 Then MailBody = MailBody &  "<b>Error Source : </b> " & Source &  "<br><br>" & VbCrlf
If Len(CStr(FileName)) > 0 Then MailBody = MailBody & "<b>File Name : </b> " & FileName &  "<br>" & VbCrlf
If LineNumber > 0 Then MailBody = MailBody & "<b>Line Number : </b> " & LineNumber &  "<br>" & VbCrlf
If Len(CStr(FullDescription)) > 0 Then MailBody = MailBody &  "<b>Full Description : </b> " & FullDescription & "<br>" & VbCrlf
If Len(CStr(RequestedURL)) > 0 Then MailBody = MailBody &  "<b>URL : </b> " & RequestedURL & "<br>" & VbCrlf
If Len(CStr(RequestedQuerystring)) > 0 Then MailBody = MailBody &  "<b>Query : </b> " & RequestedQuerystring & "<br>" & VbCrlf

 Dim Item

 MailBody = MailBody &  "<b>Request.Form : </b> " & Chr(10)

 For Each Item in Request.Form
     MailBody = MailBody &  "<b>" & Item & " : </b> " & Request.Form(Item).Item & ".<br>" & Chr(10)
 Next

MailBody = MailBody & "</P>" & VbCrlf

MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf

MailBody = MailBody & "<p>From <a class=""standardsButton"" href=""http://ws.arin.net/cgi-bin/whois.pl?queryinput=" & Request.ServerVariables("REMOTE_ADDR") & """>" & Request.ServerVariables("REMOTE_ADDR") & "</a></p>" & VbCrlf

MailBody = MailBody & "</Center>" & VbCrlf
MailBody = MailBody & "</DIV>" & VbCrlf
MailBody = MailBody & "</html>" & VbCrlf

Body = MailBody

select case lcase(EmailComponent) 

	case "abmailer"
		Set Mail = Server.CreateObject("ABMailer.Mailman")
		Mail.ServerAddr = EmailServer
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.SendTo = ToEmail
		Mail.MailSubject = Subject
		Mail.MailMessage = Body
		On Error Resume Next '## Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "aspemail"
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.FromName = Name
		Mail.From = From
		Mail.AddReplyTo From
		Mail.Host = EmailServer
		Mail.AddAddress ToEmail, ToName
		Mail.Subject = Subject
		Mail.Body = Body
		On Error Resume Next '## Ignore Errors
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "aspmail"
		Set Mail = Server.CreateObject("SMTPsvg.Mailer")
                Mail.ContentType = "text/html"
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.ReplyTo = From
		Mail.RemoteHost = EmailServer
		Mail.AddRecipient ToName, ToEmail
		Mail.Subject = Subject
		Mail.BodyText = Body
		On Error Resume Next '## Ignore Errors
		SendOk = Mail.SendMail
		If not(SendOk) <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"
		End if

	case "aspqmail"
		Set Mail = Server.CreateObject("SMTPsvg.Mailer")
                Mail.ContentType = "text/html"
		Mail.QMessage = 1
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.ReplyTo = From
		Mail.RemoteHost = EmailServer
		Mail.AddRecipient ToName, ToEmail
		Mail.Subject = Subject
		Mail.BodyText = Body
		On Error Resume Next '## Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "cdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		Mail.BodyFormat = 0
		Mail.MailFormat = 0
		On Error Resume Next '## Ignore Errors
		Mail.Send From, ToEmail, Subject, Body
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "chilicdonts"
		Set Mail = Server.CreateObject ("CDONTS.NewMail")
		On Error Resume Next '## Ignore Errors
		Mail.Host = EmailServer
		Mail.To = ToName & "<" & ToEmail & ">"
		Mail.From = Name & "<" & From & ">"
		Mail.Subject = Subject
		Mail.Body = Body
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "cdosys"
	        Set iConf = Server.CreateObject ("CDO.Configuration")
        	Set Flds = iConf.Fields 

	        'Set and update fields properties
        	Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	        Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailServer
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
        	Flds.Update

	        Set Mail = Server.CreateObject("CDO.Message")
        	Set Mail.Configuration = iConf

	        'Format and send message
        	Err.Clear 
		Mail.To = ToName & "<" & ToEmail & ">"
		Mail.From = Name & "<" & From & ">"
		Mail.Subject = Subject
		Mail.HTMLBody = Body
        	On Error Resume Next
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	 case "dkqmail"
		Set Mail = Server.CreateObject("dkQmail.Qmail")
		Mail.FromEmail = From
		Mail.ToEmail = ToEmail
		Mail.Subject = Subject
		Mail.Body = Body
		Mail.CC = ""
		Mail.MessageType = "TEXT"
		On Error Resume Next '## Ignore Errors
		Mail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "dundasmailq"
		Set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.QuickSend From, ToEmail, Subject, Body
		On Error Resume Next '##Ignore Errors
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "dundasmails"
		Set Mail = Server.CreateObject("Dundas.Mailer")
		Mail.TOs.Add ToEmail
		Mail.FromAddress = From
		Mail.Subject = Subject
		Mail.Body = Body
		On Error Resume Next '##Ignore Errors
		Mail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "geocel"
		set Mail = Server.CreateObject("Geocel.Mailer")
		Mail.AddServer EmailServer, 25
		Mail.AddRecipient ToEmail, ToName
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.Subject = Subject
		Mail.Body = Body
		On Error Resume Next '##  Ignore Errors
		Mail.Send()
		If Err <> 0 then 
			Response.Write "Your request was not sent due to the following error: " & Err.Description 
		Else
			Response.Write "Your mail has been sent..."
		End If

	case "iismail"
		Set Mail = Server.CreateObject("iismail.iismail.1")
		MailServer = EmailServer
		Mail.Server = EmailServer
		Mail.addRecipient(ToEmail)
		Mail.From = From
		Mail.Subject = Subject
		Mail.body = Body
		On Error Resume Next '## Ignore Errors
		Mail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "jmail"
		Set Mail = Server.CreateObject("Jmail.smtpmail")
		Mail.ServerAddress = EmailServer
		Mail.AddRecipient ToEmail
		Mail.Sender = From
		Mail.Subject = Subject
		Mail.body = Body
		Mail.priority = 3
		On Error Resume Next '## Ignore Errors
		Mail.execute
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "jmail4"
		Set Mail = Server.CreateObject("Jmail.Message")
		'Mail.MailServerUserName = "myUserName"
		'Mail.MailServerPassword = "MyPassword"
		Mail.From = From
		Mail.FromName = Name
		Mail.AddRecipient ToEmail, ToName
		Mail.Subject = Subject
		Mail.Body = Body
		on error resume next '## Ignore Errors
		Mail.Send(EmailServer)
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "mdaemon"
		Set gMDUser = Server.CreateObject("MDUserCom.MDUser")
		mbDllLoaded = gMDUser.LoadUserDll

		If mbDllLoaded = False then
			response.write "Could not load MDUSER.DLL! Program will exit." & "<br />"
		Else
			Set gMDMessageInfo = Server.CreateObject("MDUserCom.MDMessageInfo")
			gMDUser.InitMessageInfo gMDMessageInfo
			gMDMessageInfo.To = ToEmail
			gMDMessageInfo.From = From
			gMDMessageInfo.Subject = Subject
			gMDMessageInfo.MessageBody = Body
			gMDMessageInfo.Priority = 0
			gMDUser.SpoolMessage gMDMessageInfo
			mbDllLoaded = gMDUser.FreeUserDll
		End if

		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End If

	case "ocxmail"
		Set Mail = Server.CreateObject("ASPMail.ASPMailCtrl.1")
		On Error Resume Next '## Ignore Errors
		Result = Mail.SendMail(EmailServer, ToEmail, From, Subject, Body)
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "ocxqmail"
		Set Mail = Server.CreateObject("ocxQmail.ocxQmailCtrl.1")
		On Error Resume Next '## Ignore Errors
		Mail.Q EmailServer,      _
			Name,      _
		        From,      _
		        "",      _
		        "",      _
		        ToEmail,      _
		        "",      _
		        "",      _
		        "",      _
		        Subject,      _
		        Body
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "sasmtpmail"
		Set Mail = Server.CreateObject("SoftArtisans.SMTPMail")
		Mail.FromName = Name
		Mail.FromAddress = From
		Mail.AddRecipient ToName, ToEmail
		'Mail.AddReplyTo From
		Mail.BodyText = Body
		Mail.organization = SiteDescription
		Mail.Subject = Subject
		Mail.RemoteHost = EmailServer
		On Error Resume Next
		SendOk = Mail.SendMail
		If Not(SendOk) <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Mail.Response & "</li>"
		End if

	case "smtp"
		Set Mail = Server.CreateObject("SmtpMail.SmtpMail.1")
		Mail.MailServer = EmailServer
		Mail.Recipients = ToEmail
		Mail.Sender = From
		Mail.Subject = Subject
		Mail.Message = Body
		On Error Resume Next '## Ignore Errors
		Mail.SendMail2
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

	case "vsemail"
		Set Mail = CreateObject("VSEmail.SMTPSendMail")
		Mail.Host = EmailServer
		Mail.From = From
		Mail.SendTo = ToEmail
		Mail.Subject = Subject
		Mail.Body = Body
		On Error Resume Next '## Ignore Errors
		Mail.Connect
		Mail.Send
		Mail.Disconnect
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if

End Select

Set Mail = Nothing

  If Err_Msg <> "" Then Response.Write "<!-- SERIOUS FATAL CRITICAL ERROR : " & Err_Msg & "-->"
  If Err <> 0 Then Response.Write "<!-- Even the error page errors!!! : " & Err.Description & "-->"

On Error Goto 0
End If %>
<!-- Generated by Matthew1471 BlogX v.<%=Version%> --></BODY></HTML>
<%
' --------------------------------------------------------------------------
'¦Introduction : Page Navigation.                                           ¦
'¦Purpose      : Inserts the navigation into the page after the content.    ¦
'¦Requires     : Plugin.asp.                                                ¦
'¦Used By      : Most pages.                                                ¦
'---------------------------------------------------------------------------

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************
%>
<div class="sidebar" id="rightBar">
<br/>
<%
'-- Calendar --'
Dim nLastDay, n, nn, nnn, nDS, PostToday
nLastDay = Day(DateSerial(DataYear, DataMonth + 1, 1 - 1))
DataDay = 1 - Weekday(DateSerial(DataYear, DataMonth, 1)) + 1
%><!-- Begin Calendar -->
<table class="navCalendar" cellspacing="0" cellpadding="4" border="0" style="border-width:1px;border-style:solid;border-collapse:collapse;">
 <tr>
  <td colspan="7" style="background-color:<%=CalendarBackground%>;">
   <table class="navTitleStyle" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">
    <tr>
     <td class="navNextPrevStyle" style="width:15%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=DataYear & Right("00" & DataMonth, 2)%>&amp;POS=LAST" style="color:Black">&lt;</a></td>
     <td align="center" style="width:70%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=DataYear & Right("00" & DataMonth, 2)%>"><%=MonthName(DataMonth)%></a> (<%=Right(DataYear,2)%>)</td>
     <td class="navNextPrevStyle" align="right" style="width:15%;"><a href="<%=SiteURL & PageName%>?YearMonth=<%=DataYear & Right("00" & DataMonth, 2)%>&amp;POS=NEXT" style="color:Black">&gt;</a></td>
    </tr>
   </table>
  </td>
 </tr>
<%
'-- Write out weekdays --'
Response.Write " <tr>" & VbCrlf

For n = 0 To 6
 Response.Write "  <td class=""navDayHeader"" align=""center"">" & Left(WeekdayName(n + 1, True),1) & "</td>" & VbCrlf
Next

Response.Write " </tr>" & VbCrlf

'-- Write out days --'

For nn = 0 To 5

 Response.Write " <tr>" & VbCrlf

  For nnn = 0 To 6
   If DataDay > 0 And DataDay <= nLastDay Then

    Response.Write "  <td class="""

    '-- Highlight CurrentDay/Weekend --' 
    If DataDay = Int(Request("Day")) Then
     Response.Write "navSelectedDayStyle" 
    ElseIf nnn = 0 or nnn = 6 Then 
     Response.Write "navWeekendDayStyle"
    Else 
     Response.Write "navDayStyle"
    End If
    '-- End Of Current Day Check --'

    Response.Write """ align=""center"""

    '-- Highlight CurrentDay/Weekend --'
    If DataDay = Int(Request("Day")) Then
     Response.Write " style=""color:White;background-color:" & CalendarBackground & ";width:14%;"">"
    Else 
     Response.Write " style=""width:14%;"">"
    End If
    '-- End Of Current Day Check --'

    '-- Lets strip out that existing day from our clicky --'
    If (SortByDay = True) AND (CalendarCheck <> 1) Then
     Response.Write "<a href=""" & SiteURL & PageName & "?"
     Response.Write "YearMonth=" & DataYear & Right("00" & DataMonth, 2) & "&amp;Day=" & DataDay & """>"
    ElseIf SortByDay = True Then

     '-- Check if there was something posted on each day --'
     Records.CursorLocation = 3 ' adUseClient
     Records.Open "SELECT Day, Month, Year FROM Data WHERE Day=" & DataDay & " AND Month=" & Right("00" & DataMonth, 2) & " AND Year=" & DataYear & ";",Database, 1, 3

     If Records.EOF = False Then
      PostToday = True
      Response.Write "<a href=""" & SiteURL & PageName & "?"
      Response.Write "YearMonth=" & DataYear & Right("00" & DataMonth, 2) & "&amp;Day=" & DataDay & """>"
     Else
      PostToday = False
     End If

     Records.Close

    End If

    If (Day(DateAdd("h",ServerTimeOffset,Now())) = DataDay) AND (Month(DateAdd("h",ServerTimeOffset,Now())) = DataMonth) Then Response.Write "<span style=""color: red"">"
    Response.Write DataDay
    If (Day(DateAdd("h",ServerTimeOffset,Now())) = DataDay) AND (Month(DateAdd("h",ServerTimeOffset,Now())) = DataMonth) Then Response.Write "</span>"

    If (SortByDay = True AND CalendarCheck <> 1) OR (PostToday = True) Then Response.Write "</a>"
    '-- Finished Day Stripping --'

    Response.Write "</td>" & CHR(13)

   Else
    Response.Write "  <td class=""navOtherMonthDayStyle"" align=""center"" style=""width:14%;"">-</td>" & CHR(13)
   End If

   DataDay = DataDay + 1

  Next

 Response.Write " </tr>"

Next
%>
 <tr>
  <td colspan="7" class="navCalendar" style="background-color:<%=CalendarBackground%>;" align="center"><a href="<%=SiteURL & PageName%>">This Month!</a></td>
 </tr>
</table>
<!-- End Calendar -->

<% If (MultiLanguage = True) Then %>
<!-- Language -->
<div class="section">
<h3 class="sectionTitle">Languages</h3>
<div class="sectionBody">
<p style="text-align: center">Languages!<br/>
<br/></p>
<hr style="width: 15%"/>
<p style="text-align: center">Coming Soon..</p></div></div><br/>
<% End If %>

<% If (Request.ServerVariables("REMOTE_ADDR") <> NoAdvertIP) AND (InStr(1, Request.ServerVariables("HTTP_Host"),"blogx.co.uk", 1) <> 0) AND (LegacyMode <> True) Then %>
<!-- Adverts -->
<div class="section">
<h3 class="sectionTitle">Adverts/Donate</h3>
<div class="sectionBody" style="text-align:center">
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<p>Donating to the BlogX project<br/>
helps improve the product further.<br/>
 <input type="hidden" name="cmd" value="_s-xclick"/>
 <input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but04.gif" name="submit" alt="Make payments with PayPal - it's fast, free and secure!"/>
 <input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----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-----END PKCS7-----
"/></p>
</form>
<br/>
<hr style="width: 30%"/>

<!-- Begin: AdBrite -->
<script type="text/javascript" src="http://ads.adbrite.com/mb/text_group.php?sid=25632&amp;br=1&amp;dk=726567697374657220646f6d61696e5f355f325f776562">
</script>
<div><a class="adHeadline" href="http://www.adbrite.com/mb/commerce/purchase_form.php?opid=25632&amp;afsid=1">Your Ad Here</a></div>
<!-- End: AdBrite -->

<br/>

<!-- Begin: Google AdSense -->
<div style="align: center">
<script type="text/javascript"><!--
google_ad_client = "pub-5730167862640052";
google_ad_width = 200;
google_ad_height = 200;
google_ad_format = "200x200_as";
google_ad_type = "image";
google_ad_channel = "";
google_color_border = ["FFFFFF","336699","000000","6699CC"];
google_color_bg = ["FFFFFF","FFFFFF","F0F0F0","003366"];
google_color_link = ["0000FF","0000FF","0000FF","FFFFFF"];
google_color_text = ["000000","000000","000000","AECCEB"];
google_color_url = ["008000","008000","008000","AECCEB"];
//-->
</script>
<script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js"></script>
</div>
<!-- End: Google Ad Sense -->
<br/>
</div></div><br/>
<% End If %>

<% If (ShowMonth <> False) AND (LegacyMode <> True) Then

'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT DISTINCT Month, Year FROM Data ORDER BY Year, Month",Database, 1, 3

If NOT Records.EOF Then %>
<!-- Archive -->
<div class="section">
<h3 class="sectionTitle">Archive</h3>
<div class="sectionBody">
<%
Dim MonthPost, YearPost

'--- Set Category ---'
Set MonthPost = Records("Month")
Set YearPost = Records("Year")

'-- Write Them In ---'
Do Until (Records.EOF or Records.BOF)
 Response.Write "<a href=""" & SiteURL & "Main.asp?YearMonth=" & YearPost & Right("00" & MonthPost, 2) & """>" 
 Response.Write MonthName(MonthPost) & " " & YearPost & "</a><br/>" & VbCrlf
 Records.MoveNext
Loop

Response.Write "<br/></div></div><br/>" & VbCrlf

End If

'-- Close The Records ---'
Records.Close

End If%>

<!-- Links -->
<div class="section">
<h3 class="sectionTitle">Links <%If (Session(CookieName) = True) AND (AllowEditingLinks <> 0) Then Response.Write " <acronym title=""Edit Your Links""><a href=""" & SiteURL & "Admin/EditLinks.asp""><img alt=""Edit Links Icon"" src=""" & SiteURL & "Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3>
<ul>
  <% If EnableMainPage = True Then Response.Write "<li><a href=""" & SiteURL & """>About Me</a></li>"%>
  <li><a href="<%=SiteURL & PageName%>">Blog Home</a></li>
  <% Records.Open "SELECT LinkName, LinkURL, LinkType FROM Links Where LinkType='Main Links' ORDER BY LinkName Asc;", Database

  Do Until (Records.EOF) 
   Response.Write "<li><a href=""" & Records("LinkURL") & """>" & Records("LinkName") & "</a></li>"
   Records.MoveNext
  Loop

  Records.Close %>
</ul></div><br/>

<%
If (OtherLinks <> 0) AND (FSODisabled = False) Then
 Records.Open "SELECT LinkName, LinkURL, LinkType FROM Links Where LinkType='Other Links' ORDER BY LinkName Asc;", Database
 If NOT Records.EOF Then
%>
<!-- Other Links -->
<div class="section">
<h3 class="sectionTitle">Other Links <%If (Session(CookieName) = True) AND (AllowEditingLinks <> 0) Then Response.Write " <acronym title=""Edit Your Links""><a href=""" & SiteURL & "Admin/EditLinks.asp""><img alt=""Edit Other Links Icon"" src=""" & SiteURL & "Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3>
<ul>
<%
  Do Until (Records.EOF)
   Response.Write "<li><a href=""" & Records("LinkURL") & """>" & Records("LinkName") & "</a></li>"
   Records.MoveNext
  Loop
  Records.Close

  Response.Write "</ul>" & VbCrlf
  Response.Write "</div>" & VbCrlf
  Response.Write "<br/>" & VbCrlf

 End If

End If

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

   	SplitText   = Split(Trim(Records("Content"))," ")
	PollContent = ""

     	  For WordLoopCounter = 0 To UBound(SplitText)
	   PollContent = PollContent & " " & SplitText(WordLoopCounter)
	   If (Int(WordLoopCounter / 4) = (WordLoopCounter / 4)) AND (WordLoopCounter > 0) Then PollContent = PollContent & "<br/>" & VbCrlf
     	  Next

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

<!-- Poll -->
<div class="section">
<h3 class="sectionTitle">Poll<%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your Poll""><a href=""" & SiteURL & "Admin/EditPoll.asp""><img alt=""Edit Poll Icon"" src=""" & SiteURL & "Images/Edit.gif"" style=""border-style: none""/></a></acronym>"%></h3>
   <%
If AlreadyVoted = False Then

   Response.Write "<form method=""post"" action=""" & SiteURL & "Vote.asp"">"  & VbCrlf

   Response.Write "<p style=""text-align: center;"">" & PollContent & "<br/>" & "</p>" & VbCrlf 
   Response.Write "<hr style=""width: 30%""/>" & VbCrlf 

   Response.Write "<p><input type=""radio"" value=""1"" name=""Vote""/>&nbsp;"
   Response.Write "<span style=""font-family: Verdana, Arial, Helvetica; font-size:1""><b>" & Des1 & "</b></span><br/></p>" & VbCrlf

   Response.Write "<p><input type=""radio"" value=""2"" name=""Vote""/>&nbsp;"
   Response.Write "<span style=""font-family: Verdana, Arial, Helvetica; font-size:1""><b>" & Des2 & "</b></span><br/></p>" & VbCrlf

   If Des3 <> "" Then 
   Response.Write "<p><input type=""radio"" value=""3"" name=""Vote""/>&nbsp;"
   Response.Write "<span style=""font-family: Verdana, Arial, Helvetica; font-size:1""><b>" & Des3 & "</b></span><br/></p>" & VbCrlf
   End If

   If Des4 <> "" Then 
   Response.Write "<p><input type=""radio"" value=""4"" name=""Vote""/>&nbsp;"
   Response.Write "<span style=""font-family: Verdana, Arial, Helvetica; font-size:1""><b>" & Des4 & "</b></span></p>" & VbCrlf
   End If
   
   Response.Write "<p><span style=""text-align: center;""><input type=""image"" src=""" & SiteURL & "Images/vote.gif""/></span></p>" & VbCrlf
   Response.Write "</form>" & VbCrlf

Else
   
   Response.Write "<p style=""text-align:center"">" & PollContent %>
   <br/>(Total : <%=Records("Total")%> Vote(s) )</p>
   <hr style="width: 30%"/>
      <table cellspacing="0" width="30%" border="0" style="margin-left: auto; margin-right: auto; text-align: center;">
        <tbody>
         <tr>
           <td>
   <%
   Op1Percent = Cint((Op1 / Total) * 100)
   Op2Percent = Cint((Op2 / Total) * 100)
   Op3Percent = Cint((Op3 / Total) * 100)
   Op4Percent = Cint((Op4 / Total) * 100)

   If Des1 <> "" Then Response.Write "&nbsp;" & Des1 & "<br/> <img alt=""Poll Bar 1"" src=""" & SiteURL & "Images/Bar.gif"" width=""" & Int(Op1Percent / 2) & "%"" height=""10""/> " & Op1 & " (" & Op1Percent & "%)<br/><br/>" & VbCrlf

   If Des2 <> "" Then Response.Write "&nbsp;" & Des2 & "<br/> <img alt=""Poll Bar 2"" src=""" & SiteURL & "Images/Bar.gif"" width=""" & Int(Op2Percent / 2) & "%"" height=""10""/> " & Op2 & " (" & Op2Percent & "%)<br/><br/>" & VbCrlf

   If Des3 <> "" Then Response.Write "&nbsp;" & Des3 & "<br/> <img alt=""Poll Bar 3"" src=""" & SiteURL & "Images/Bar.gif"" width=""" & Int(Op3Percent / 2) & "%"" height=""10""/> " & Op3 & " (" & Op3Percent & "%)<br/><br/>" & VbCrlf

   If Des4 <> "" Then Response.Write "&nbsp;" & Des4 & "<br/> <img alt=""Poll Bar 4"" src=""" & SiteURL & "Images/Bar.gif"" width=""" & Int(Op4Percent / 2) & "%"" height=""10""/> " & Op4 & " (" & Op4Percent & "%)<br/>" & VbCrlf

   Response.Write "</td>"
   Response.Write "</tr>"
   Response.Write "</tbody>"
   Response.Write "</table>"

   Response.Write "<p style=""text-align: center""><a href=""" & SiteURL & "Results.asp"">Results</a></p>"
   End If %>
</div><br/>
   <% 
    End If
    Records.Close
End If %>

<% If ShowCategories <> False Then 
Dim Category, LastCat
%>
<!-- Categories -->
<div class="section">
<h3 class="sectionTitle">Categories</h3>
<ul>
<li><a href="<%=SiteURL & PageName%>">All</a> (<a href="<%=SiteURL%>RSS/<%If (ReaderPassword <> "") AND Session("Reader") = True Then Response.Write "?" & ReaderPassword%>">Rss</a>)</li>
<%

'--- Open Recordset ---'
Records.CursorLocation = 3 ' adUseClient
Records.Open "SELECT DISTINCT Category FROM Data ORDER BY Category",Database, 1, 3

 '--- Set Category ---'
 Set Category = Records("Category")

 '-- Write Them In ---'
 Do Until (Records.EOF or Records.BOF)
  If (Category <> "") Then 
   Response.Write "<li><a href=""" & SiteURL & "ViewCat.asp?Cat=" & Replace(Category, " ", "%20") & """>" 
   Response.Write Replace(Category, "%20", " ") & "</a> (<a href=""" & SiteURL & "RSS/Cat/?Category=" & Replace(Category, " ", "%20")
   If (ReaderPassword <> "") AND (Session("Reader") = True) Then Response.Write "&Password=" & ReaderPassword
   Response.Write """>Rss</a>)</li>" & VbCrlf
  End If
 Records.MoveNext
Loop

'-- Close The Records ---'
Records.Close
%>
</ul></div><br/>
<%End If

If (UseExternalPlugin = 1) AND (LegacyMode = False) AND (FSODisabled = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!-- <%=PluginTitle%> -->
<div class="section">
<h3 class="sectionTitle"><%=PluginTitle%></h3>
<%=PluginText%>
</div><br/>
<%
End If

If (LegacyMode <> True) Then %>
<!-- Search Blog -->
<div class="section">
<h3 class="sectionTitle">Search</h3>

<form method="post" action="<%=SiteURL%>Search.asp">
 <div>
  <input name="Mode" type="hidden" value="Normal"/>
  <input name="Search" type="text" value="<%=Replace(Request("Search"),"""","&quot;")%>" size="30" maxlength="70"/>
  <input type="submit" value="Search"/><br/>
 </div>
</form>

 <a href="<%=SiteURL%>Search.asp">Advanced Search</a>

</div>
<br/>
<% End If %>

<!-- Login As A Publisher -->
<% If Blacklisted = True Then %>
 <div class="section" id="login">   
  <h3 class="sectionTitle">Login Error</h3>
  <p style="text-align:center">Too many failed login attempts,<br/>please wait 15 minutes and try again.</p>
 </div>
<% ElseIf CookiesDisabled = True Then %>
  <div class="section" id="login">   
  <h3 class="sectionTitle">Login Error</h3>
  <p style="text-align:center">Please <b>Enable</b> Cookies to login<br/>
  <a href="Main.asp">Try again?</a></p>
  </div>
<% ElseIf Session(CookieName) = True Then %>
<div class="section">
<h3 class="sectionTitle"><img alt="Login Key" src="<%=SiteURL%>Images/Key.gif" style="border-style: none"/>Admin</h3>
Content<br/>
------<br/>
    <ul>
        <li><a href="<%=SiteURL%>Admin/EditMainPage.asp">About Me</a></li>
        <li><a href="<%=SiteURL%>Admin/AddEntry.asp">Add Entry</a></li>
        <li><a href="<%=SiteURL%>Admin/AddFile.asp">Add Shared File(s)</a></li>
        <li><a href="<%=SiteURL%>Admin/AddPoll.asp">Add Poll</a></li>
        <li><a href="<%=SiteURL%>Admin/MailingListMembers.asp">Mailing List</a></li>
        <% If ArgoSoftMailServer = True Then Response.Write "<li><a href=""" & SiteURL & "Admin/ParseEmails.asp"">Parse E-mails</a></li>" %>
        <li><a href="<%=SiteURL%>Admin/EditDisclaimer.asp">Disclaimer</a></li>
    </ul>

Tools<br/>
------<br/>
    <ul>
        <li><a href="<%=SiteURL%>Admin/CheckForUpdate.asp">Check For Update</a></li>
        <li><a href="<%=SiteURL%>Admin/EditDraft.asp">Drafts</a></li>
        <li><a href="<%=SiteURL%>Admin/LastComments.asp">Last Comments</a></li>
        <li><a href="<%=SiteURL%>Admin/NotFound.asp">Not Found</a></li>
        <li><a href="<%=SiteURL%>Admin/Referrers.asp">Referrers</a></li>
    </ul>

Settings<br/>
------<br/>
    <ul>
        <li><a href="<%=SiteURL%>Admin/EditBan.asp">Banned Addresses</a></li>
        <li><a href="<%=SiteURL%>Admin/ChangePassword.asp">Change Password</a></li>
        <li><a href="<%=SiteURL%>Admin/Config.asp">Config</a></li>
        <li><a href="<%=SiteURL%>Admin/EmailConfig.asp">Email Settings</a></li>
        <li><a href="<%=SiteURL & PageName %>?ClearCookie">Logout</a></li>
    </ul>
</div><br/>
<%Else

  Dim PageURL
  If Len(Request.Querystring("Return")) > 0 Then
   PageURL = "/" & Request.Querystring("Return")
   If Len(Request.Querystring("Query")) > 0 Then PageURL = PageURL & "?" & Request.Querystring("Query")
  Else
   PageURL = Request.ServerVariables("SCRIPT_NAME")
   If Len(Request.Querystring) > 0 Then PageURL = PageURL & "?" & Request.Querystring
  End If

  If SSLSupported = True Then
   Response.Write "  <form method=""post"" action=""https://" & Request.ServerVariables("SERVER_NAME") & PageURL & """>"
  Else
   Response.Write "  <form method=""post"" action=""http://" & Request.ServerVariables("SERVER_NAME") & PageURL & """>"
  End If
 %>
   <div class="section" id="login">   
    <h3 class="sectionTitle"><img alt="Login Key" src="<%=SiteURL%>Images/Key.gif" style="border-style: none"/>Admin Sign In</h3>
    <p>Username: <input name="username"/></p>
    <p>Password: <input name="password" type="password"/></p>
    <p><input name="Remember" type="checkbox" value="True"/>Remember Login</p>
    <p><input name="SignIn" type="submit" value="Sign In"/><br/></p>
   </div>
  </form>
<%End If %>
</div>
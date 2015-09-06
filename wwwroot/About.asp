<%
' --------------------------------------------------------------------------
'¦Introduction : Version Information Page                                   ¦
'¦Purpose      : Provides general information about BlogX and               ¦
'¦               specific information about this particular install.        ¦
'¦Used By      : Includes/Footer.asp                                        ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦Notes        : This page is for downloading BlogX off BlogX.co.uk, but the¦
'¦               options automatically change slightly based on the host.   ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

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

'-- Proxy Handler --'
CacheHandle(CDate("04/12/08 20:18:00"))

PageTitle = "About BlogX"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">

 <!--- Start General Information -->
 <div class="entry">
  <h3 class="entryTitle">About BlogX</h3>

  <div class="entryBody">
   <p>This site is running Matthew1471's version of BlogX V<%=Version%>.</p>
   <p>The site owner can post information about his/her daily or weekly events in a little box and the site presents them for everyone to read.</p>
  </div>
 </div>
 <!--- End General Information -->

 <!--- Start Original BlogX Information -->
 <% If Request.Querystring("ShowOriginal") = "Y" Then Response.Write "<div class=""entry"">"%>
  <h3 class="entryTitle"><% If Request.Querystring("ShowOriginal") = "Y" Then%><acronym title="Hide The Original Information"><a href="?ShowOriginal=N&amp;ShowNew=<%=Request.Querystring("ShowNew")%>&amp;ShowChanges=<%=Request.Querystring("ShowChanges")%>"><img alt="Show Less" src="Images/Less.Gif" style="border-style: none"/></a></acronym><% Else %><acronym title="Show The Original Information"><a href="?ShowOriginal=Y&amp;ShowNew=<%=Request.Querystring("ShowNew")%>&amp;ShowChanges=<%=Request.Querystring("ShowChanges")%>"><img alt="Show More" src="Images/More.Gif" style="border-style: none"/></a></acronym><% End If%> What Was The Original WebBlogX/BlogX</h3><br/>

  <% If Request.Querystring("ShowOriginal") = "Y" Then %><div class="entryBody">
   BlogX is a ASP C# .Net Blogging program, which was originally written for websites by "Chris Anderson" (<a href="http://SimpleGeek.com">http://SimpleGeek.com</a>)
  </div>
 </div><% End If %>
 <!--- End Original BlogX Information -->

 <!--- Start Matthew1471 BlogX Information -->
<% If Request.Querystring("ShowNew") <> "N" Then Response.Write "<div class=""entry"">"%>
<h3 class="entryTitle"><% If Request.Querystring("ShowNew") <> "N" Then%><acronym title="Hide The New Information"><a href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&amp;ShowNew=N&amp;ShowChanges=<%=Request.Querystring("ShowChanges")%>"><img alt="Show Less" src="Images/Less.Gif" style="border-style: none"/></a></acronym><% Else %><acronym title="Show The New Information"><a href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&amp;ShowNew=Y&amp;ShowChanges=<%=Request.Querystring("ShowChanges")%>"><img alt="Show More" src="Images/More.Gif" style="border-style: none"/></a></acronym><% End If%> What Is Matthew1471's BlogX</h3><br/>
<% If Request.Querystring("ShowNew") <> "N" Then %>
<div class="entryBody">
<p>Matthew1471's Edition of BlogX runs "Classic ASP" (Wider supported and simpler) and is programmed in Visual Basic.</p>

<p>What is Matthew1471's BlogX's Features?</p>
<ul>
  <li><b>Firm</b> Advanced Comment Spam Management Control System (including IP validation).</li>
  <li>Full online control panel to <b>edit the configuration</b>.</li>
  <li>Fully supports the innovative <b>Pingback</b></li>
  <li>Allows <b>categories</b>, <b>date linking</b> to be <b>turned on or off</b>.</li>
  <li>Much <b>more responsive</b> than the original WebBlogX.</li>
  <li>Ability to customize the <b>external links</b> online.</li>
  <li>Allows <b>"EntriesPerPage"</b> setting and correctly handles <b>page numbers</b>.</li>
  <li>Allows <b>"Contact Me"</b> setting and supports a range of <b>ASP Mail components</b>.</li>
  <li>Allows <b>"<a href="Themes.asp">Themes</a>"</b>.</li>
  <li>Supports <b>RSS</b> (Really Simple Syndication).</li>
  <li>Can be set to use either <b>12 Hour or 24 Hour times</b>.</li>
  <li>Has an online <b>editable disclaimer</b> and <b>change password</b> utility.</li>
  <li>Has a built in <b>spell check</b> function.</li>
  <li>Dynamically <b>assigns</b> categories upon a new entry added.</li>
  <li>Checks for SQL &amp; HTML exploits/injections.</li>
  <li><b>Anti-bruteforce login code</b> prevents password guesses.</li>
  <li>Supports the Windows client Matthew1471's <b>WinBlogX</b>.</li>
  <li><b>Simple</b> to setup and configure.</li>
  <li>Comments RSS for each post.</li>
  <li>Full customisable <b>search engine</b>.</li>
  <li>Most pages follow the XHTML standard which renders fast on modern browsers.</li>
  <li>Differentiates between a user's address and their proxy and allows independent banning of both.</li>
</ul>
</div></div><% End If %>

<div class="entry">
<h3 class="entryTitle">Download Matthew1471 WebLogX</h3><br/>
<div class="entryBody">
<p>To download the current version of BlogX 
<% Dim Domain
Domain = Request.ServerVariables("HTTP_Host")

If InStr(1, Domain,"blogx.co.uk", 1) <> 0 Then 
 Response.Write "<b>V" & Version & "</b>, "
 Response.Write "click <a href=""Download.asp"">here</a>"
Else
 Response.Write "click <a href=""Share.asp"">here</a>"
End If
%>.<br/>
A list of changes and improvements can be found <a href="http://blogx.co.uk/Official/Changes.asp">here</a>.</p>

<p>To download the current version of WinBlogX (the windows posting client)
<% If InStr(1, Domain,"blogx.co.uk", 1) <> 0 Then Response.Write "<b>V1.04.14</b>,"%>
click <a href="http://blogx.co.uk/Download/WinBlogX%20Setup.exe">here</a>.</p>

<%
If InStr(1, Domain,"blogx.co.uk", 1) = 0 Then
 Response.Write "<p>To view information on the current version of BlogX, click <a href=""http://BlogX.co.uk/About.asp"">here</a>.</p>"
Else
 Response.Write "<p>To subscribe to the BlogX mailing list, click <a href=""http://BlogX.co.uk/MailingList.asp"">here</a>.</p>"
End If
%>
</div></div>

<div class="entry">
<h3 class="entryTitle">Thanks To</h3><br/>
<div class="entryBody">
<p>
Max Web Portal<br/>
Freevbcode.com<br/>
Developerfusion.com<br/>
kirchmeier.org<br/>
ASPxmlrpc SourceForge Team<br/>
<br/>All the beta testers and people who have got in contact with suggestions and reported bugs<!--- Sarah, Dan, Naomi, Russ, Tom your all great guys! Yay you found an easteregg ;-) -->.</p>

<p>Dedicated to Hannah Poulter, a girl who taught me an awful lot about life.</p>

<p>Lee (IndaUK) for the incredible FireFox JS compatibility.</p>

<p>Obviously a huge thanks to Chris Anderson without his site design there would <span style="text-decoration: underline">DEFINITELY</span> not have been a Matthew1471 BlogX today!</p>
</div></div>

<!--- End Information -->

<p style="text-align: center"><a href="<%=PageName%>">Back To The Main Page</a></p>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
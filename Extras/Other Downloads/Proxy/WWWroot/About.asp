<%
OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-04 Matthew Roberts
'**
'** This free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with a link back to
'** http://www.simplegeek.com in the footer of the pages MUST
'** remain visible when the pages are viewed on the internet or intranet.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<DIV id=content>

<!--- Start Information --->
<DIV class=entry>
<H3 class=entryTitle>About BlogXProxy</H3>
<DIV class=entryBody>
<P>This site is running Matthew1471's BlogXProxy V<%=Version%>.</P>
<p>The site owner can delegate access so other users may post information about his/her daily or weekly events to a ASP BlogX compatible blog.</P>
</Div></Div>

<% If Request.Querystring("ShowNew") <> "N" Then Response.Write "<DIV class=entry>"%>
<h3 class=entryTitle><% If Request.Querystring("ShowNew") <> "N" Then%><Acronym title="Hide The New Information"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=N&ShowChanges=<%=Request.Querystring("ShowChanges")%>"><Img Border="0" Src="Images/Less.Gif"></A></Acronym><% Else %><Acronym title="Show The New Information"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=Y&ShowChanges=<%=Request.Querystring("ShowChanges")%>"><Img Border="0" Src="Images/More.Gif"></A></Acronym><% End If%> What Is Matthew1471's BlogX</h3><br>
<% If Request.Querystring("ShowNew") <> "N" Then %>
<DIV class=entryBody>
<p>Matthew1471's Edition of BlogX runs "Classic ASP" (Wider supported and simpler) and is programmed in Visual Basic.</p>

<p>What's Matthew1471's BlogX's Features?</p>
<ul>
  <li><b>Firm</b> Advanced Comment Spam Management Control System.</li>
  <li>Full online control panel to <b>edit the configuration</b>.</li>
  <li>Fully supports the innovative <b>Pingback</b></li>
  <li>Allows <b>categories</b> to be <b>turned on or off</b>.</li>
  <li>Allows <b>date linking</b> to be <b>turned on or off</b>.</li>
  <li>Much <b>more responsive</b> than the original WebBlogX.</li>
  <li>Ability to customize the <b>External links</b> from the "OtherLinks.Txt" file which is parsed.</li>
  <li>Allows <b>"EntriesPerPage"</b> setting and correctly handles <b>page numbers</b>.</li>
  <li>Allows <b>"Contact Me"</b> setting and correctly handles a range of <b>ASP Mail components</b>.</li>
  <li>Allows <b>"Themes"</b>.</li>
  <li>Supports <b>RSS</b> (Really Simple Syndication).</li>
  <li>Can be set to use either <b>12 Hour or 24 Hour times</b>.</li>
  <li>Has an online <b>editable disclaimer</b> and <b>change password</b> utility.</li>
  <li>Has a built in <b>spell check</b> function.</li>
  <li>Dynamically <b>assigns</b> categories upon a new entry added.</li>
  <li>Checks for SQL exploits.</li>
  <li>Uses around <b>112kb</b> (More with dictionaries).</li>
  <li>Supports the Windows client Matthew1471's <b>WinBlogX</b>.</li>
  <li><b>Simple</b> to setup and configure.</li>
  <li>Comments RSS for each post.</li>
  <li>Full customisable search engine.</li>
</ul>
</Div></Div><% End If %>

<DIV class=entry>
<h3 class=entryTitle>Download Matthew1471 WebLogX</h3><br>
<DIV class=entryBody>
<p>To download the current version of BlogX click <a href="http://BlogX.co.uk/Share.asp">here</a>.</p>
<p>To view information on the current version of BlogX, click <a href="http://BlogX.co.uk/About.asp">here</a>.</p>

</Div></Div>

<% If Request.Querystring("ShowChanges") = "Y" Then Response.Write "<DIV class=entry>"%>
<h3 class=entryTitle><% If Request.Querystring("ShowChanges") = "Y" Then%><Acronym title="Hide The Changelog"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=N"><Img Border="0" Src="Images/Less.Gif"></A></Acronym><% Else %><Acronym title="Show The Changelog"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=Y"><Img Border="0" Src="Images/More.Gif"></A></Acronym><% End If%> Recent Changes</h3><br>
<% If Request.Querystring("ShowChanges") = "Y" Then %>
<DIV class=entryBody>

<p><b>17th February 2005</b><br>
Updated code to allow the ability to remove the log in (See Includes/Header.asp).<br>
Updated code to skip login if already logged in.<br>
</p>

<p><b>27th December 2004 (Unreleased Beta)</b><br>
Started programming BlogXProxy.<br>
</p>

</Div></Div><%End If%>

<DIV class=entry>
<h3 class=entryTitle>Thanks To</h3><br>
<DIV class=entryBody>
<p>
Max Web Portal<br>
<br><br>All the beta testers and people who have got in contact with suggestions and reported bugs and also...<br>
special thanks to my perty cheeky monkey Hannah</p>
</Div></Div>

<!--- End Information --->

<p align="Center"><a href="Default.asp">Back To The Main Page</a></p>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
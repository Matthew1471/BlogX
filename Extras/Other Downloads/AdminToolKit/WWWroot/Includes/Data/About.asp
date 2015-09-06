<%
OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-04 Matthew Roberts, Chris Anderson
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
<H3 class=entryTitle>About BlogX</H3>
<DIV class=entryBody>
<P>This site is running Matthew1471's version of BlogX V<%=Version%>.</P>
<p>The site owner can post information about his/her daily or weekly events in a little box and the site presents them for everyone to read.</P>
</Div></Div>

<% If Request.Querystring("ShowOriginal") = "Y" Then Response.Write "<DIV class=entry>"%>
<h3 class=entryTitle><% If Request.Querystring("ShowOriginal") = "Y" Then%><Acronym title="Hide The Original Information"><A Href="?ShowOriginal=N&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=<%=Request.Querystring("ShowChanges")%>"><Img Border="0" Src="Images/Less.Gif"></A></Acronym><% Else %><Acronym title="Show The Original Information"><A Href="?ShowOriginal=Y&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=<%=Request.Querystring("ShowChanges")%>"><Img Border="0" Src="Images/More.Gif"></A></Acronym><% End If%> What Was The Original WebBlogX/BlogX</h3><br>
<% If Request.Querystring("ShowOriginal") = "Y" Then %>
<DIV class=entryBody>
BlogX is a ASP C# .Net Blogging program, which was originally written for websites by "Chris Anderson" (<a href="http://SimpleGeek.com">http://SimpleGeek.com</a>)</P>
</Div></Div><% End If %>

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
  <li>Allows <b>"<a href="Themes.asp">Themes</a>"</b>.</li>
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
<p>To download the current version of BlogX 
<% Dim Domain
Domain = Request.ServerVariables("HTTP_Host")

If InStr(1, Domain,"blogx.co.uk", 1) <> 0 Then 
Response.Write "<b>V" & Version & "</b>, "
Response.Write "click <a href=""Download.asp"">here</a>"
Else
Response.Write "click <a href=""Share.asp"">here</a>"
End If
%>
.</p>

<p>To download the current version of WinBlogX (The Windows Posting Client)
<% If InStr(1, Domain,"BlogX.co.uk", 1) <> 0 Then Response.Write "<b>V1.04.14</b>,"%>
click <a href="http://BlogX.co.uk/Download/WinBlogX%20Setup.exe">here</a>.</p>

<% If InStr(1, Domain,"BlogX.co.uk", 1) = 0 Then Response.Write "<p>To view information on the current version of BlogX, click <a href=""http://BlogX.co.uk/About.asp"">here</a>.</p>"%>

<% If InStr(1, Domain,"BlogX.co.uk", 1) <> 0 Then Response.Write "<p>To subscribe to the BlogX mailinglist, click <a href=""http://BlogX.co.uk/MailingList.asp"">here</a>.</p>"%>

</Div></Div>

<% If Request.Querystring("ShowChanges") = "Y" Then Response.Write "<DIV class=entry>"%>
<h3 class=entryTitle><% If Request.Querystring("ShowChanges") = "Y" Then%><Acronym title="Hide The Changelog"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=N"><Img Border="0" Src="Images/Less.Gif"></A></Acronym><% Else %><Acronym title="Show The Changelog"><A Href="?ShowOriginal=<%=Request.Querystring("ShowOriginal")%>&ShowNew=<%=Request.Querystring("ShowNew")%>&ShowChanges=Y"><Img Border="0" Src="Images/More.Gif"></A></Acronym><% End If%> Recent Changes</h3><br>
<% If Request.Querystring("ShowChanges") = "Y" Then %>
<DIV class=entryBody>

<p><b>28th December 2004 (V1.0.5.10)</b><br>
Added a "Draft NotePad" option.<br>
Added ability to remove uploading images (If you don't have an upload component and arn't likely to get one).<br>
Added support for the free "ASPSmartUpload" upload component (Upon Mikes' request).<br>
Fixed the CheckForUpdate.asp "Back" link.<br>
Splitted a few of the ZIP files (Themes, Dictionaries) into separate zips available <a href="http://matthew1471.co.uk/Downloads.asp">here</a> (to keep the ZIP size low).<br>
</p>

<p><b>27th December 2004 (V1.0.5.10)</b><br>
Fixed FireFox javascript compatibility (All Lee's handywork, I just copied and pasted, Thanks Lee!).<br>
Programmed BlogXProxy (Allow multiple users to post to one blog).<br>
</p>

<p><b>20th December 2004 (V1.0.5.10)</b><br>
Fixed any Commenter field (aside from Comment) > 50 meant loosing the comment.<br>
Optimized "Refer" to not dynamically load in fields, all pages run faster.<br>
</p>

<p><b>18th December 2004 (V1.0.5.10)</b><br>
Added "Sandy" theme.<br>
Optimized "Refer" to not dynamically load in fields, all pages run faster.<br>
</p>

<p><b>15th December 2004 (V1.0.5.10)</b><br>
Added "Advanced" button to EditEntry and AddEntry to unhide "Extra" features.<br>
Added "Extra" features to EditEntry (Change Date, Rons' Coding).<br>
</p>

<p><b>10th December 2004 (V1.0.5.10)</b><br>
Fixed spelling mistake of "Occurrence".<br>
Fixed year roll over bug which meant new years wouldnt be counted in "Archive".<br>
</p>

<p><b>07th December 2004 (V1.0.5.10)</b><br>
Updated PictureViewer.asp to allow thumbnails (something about my 22mb photo collection made me added this).<br>
</p>

<p><b>04th December 2004 (V1.0.5.10)</b><br>
Updated Mail.asp to not center the message text.<br>
Updated Mail.asp to convert character returns to new lines.<br>
Updated Mail.asp's grammer "your mail was sent" to "your message was sent".<br>
Updated Unsubscribe.asp's InStr checking for my site.<br>
Updated Comments.asp so it states the entry ID in the e-mail subject.<br>
Updated mailing list so entries aren't centered.<br>
</p>

<p><b>02nd December 2004 (V1.0.5.10)</b><br>
Updated ViewCat.asp to ignore non-numerical page numbers after a failed exploitation (No security vulnerability, just errors).<br>
</p>

<p><b>28th November 2004 (V1.0.5.10)</b><br>
Added user submitted hacks to the ZIP.<br>
Fixed PingOMatic support having my old site URL hardcoded into it.<br>
Updated EditEntry.asp to replace category "%20"'s with eye friendly spaces.<br>
</p>

<p><b>17th November 2004 (V1.0.5.09*)</b><br>
Fixed "EditEntry.asp" to parse out the "&lt;/textarea>"'s.<br>
Updated "EditEntry.asp" to not convert &amp;nbsp;'s to spaces.<br>
</p>

<p><b>13th November 2004 (V1.0.5.09*)</b><br>
Updated "SwimmingPool" theme so the comments are readable.<br>
Updated RSS so the times are in 24 hour times (Pass Feed Validation).<br>
</p>

<p><b>11th November 2004 (V1.0.5.09*)</b><br>
Programmed "PictureViewer" (Beta).<br>
</p>

<p><b>08th November 2004 (V1.0.5.09*)</b><br>
Added a RSSReader import facility for AdminToolkit.<br>
Updated WinblogX to use new site address.<br>
Updated WinblogX to delete files before downloading new ones (Data corruption).<br>
Updated site to work with the old "Check For Update" on WinBlogX (At Least I hope).<br>
Updated WinBlogX to accept an empty blog folder (Seen as my site now doesnt use one).<br>
</p>

<p><b>06th October 2004 (V1.0.5.09*)</b><br>
Added a little "help" entry to the database to get people started.<br>
Added the rest of the Themes at the expense of a larger ZIP file.<br>
</p>

<p><b>21st October 2004 (V1.0.5.09*)</b><br>
Fixed Mail support for Mdaemon.<br>
Updated Mailing code to DIM error message.<br>
Updated site to not issue e-mails if e-mail is disabled.<br>
</p>

<p><b>11th October 2004 (V1.0.5.09)</b><br>
Updated Readme<br>
</p>

<p><b>10th October 2004 (V1.0.5.09)</b><br>
Added ThemeIT Editor
Added BlogXThemer (ThemeIT) Editor support in Header.asp.<br>
Fixed HTML code in Entry titles messing up E-mail link. (RP4 discovered that)<br>
</p>

<p><b>07th October 2004 (V1.0.5.09)</b><br>
Added "On Error Resume Next" to OtherLinks & Links (So WebHosts not supporting FSO work).<br>
Fixed hacking protection script error (No Vulnerabilities)<br>
Fixed RTF selected text new window linking glitch (Wow what a mouthful)<br>
Fixed Search where an entry category contains NULL.<br>
Fixed Search to not display categories when categories are turned off.<br>
Fixed trying to navigate the calendar when on the comments page.<br>
Updated Mailing List page to actually say what that box was for. (Thanks Ben for that)<br>
Updated Mailing List page to be more user friendly.<br>
Updated all site link checking to ignore Matthew1471.co.uk.<br>
Updated Results page to use the calander CSS.<br>
</p>

<p><b>26th September 2004 (V1.0.5.08)</b><br>
Fixed pingback having HARDCODED into it my website's old address..<br>
</p>

<p><b>29th August 2004 (V1.0.5.07)</b><br>
Fixed rare bug where poll would fail to register a vote.<br>
Fixed time display if times are set to 24hours on some pages.<br>
Fixed CheckForUpdate's link.<br>
Updated CheckForUpdate to verify link exists in future.<br>
</p>

<p><b>28th August 2004 (V1.0.5.06)</b><br>
Fixed links to new site layout.<br>
</p>

<p><b>20th August 2004 (V1.0.5.06)</b><br>
Fixed problem on IIS3/4 where Response.Redirect failed on Comments.asp.<br>
</p>

<p><b>12th August 2004 (V1.0.5.05)</b><br>
Fixed Edit Poll.<br>
</p>


<p><b>09th August 2004 (V1.0.5.05)</b><br>
Added List Plugin (Dan's idea).<br>
Fixed mailing list not showing text on main BlogX domains.<br>
Updated Image Upload Error to still display smileys even if the upload component is not installed.<br>
Updated main.asp to display WeekDayName.<br>
Updated plugin documentation to explain running multiple plugins.<br>
Updated search to change the font on highlighted words (Unreadable yellow highlight on some font colors).<br>
Updated WhoUses.asp to explain things better.<br>
Updated ZIP.<br>
</p>

<p><b>28th July 2004 (V1.0.5.04)</b><br>
Added several new stylesheets (But not to the ZIP, you'll have to manually steal them).<br>
Updated RSS feeds to use new domain.<br>
Updated Comments RSS feed to not confuse feed readers into remarking all comments as new.<br> 
</p>

<p><b>27th July 2004 (V1.0.5.04)</b><br>
Added several new stylesheets (But not to the ZIP, you'll have to manually steal them).<br>
Fixed the link to the Admin EditDisclaimer page (Noticed it this morning).<br>
</p>

<p><b>19th July 2004 (V1.0.5.04)</b><br>
Fixed spell's ignore joining words.<br>
Improved compatibility with FireFox (Poll results).<br>
Validated & fixed all stylesheets.<br>
</p>

<p><b>18th July 2004 (V1.0.5.04)</b><br>
Added ability to remove dictionary (Without causing an ASP 500).<br>
Fixed Config.asp nullifying records.<br>
Updated Share.asp to use new domain and goto FreeWebs instead of PSC.<br>
</p>

<p><b>17th July 2004 (V1.0.5.04)</b><br>
Added "Note.gif" emoticon.<br>
Fixed a non random PUK code on Comments. (Meaning people could unsubscribe others if they know their e-mail address).<br>
Fixed link to a password protected entry from the comments page.<br>
Hopefully fixed a very minor error when the same page is loaded twice at the EXACT same time (and logging is on).<br>
</p>

<p><b>15th July 2004 (V1.0.5.03)</b><br>
Fixed "EditEntry.asp" where Title containted quotes.<br>
Fixed Spell Check were original word contained an appostrophe.<br>
Fixed absolute path in PingBack.asp (C:\Inetpub\wwwroot\Blog\).<br>
</p>

<p><b>13th July 2004 (V1.0.5.03)</b><br>
Added ServerError.asp with Intelligent bug reporting. (You'll need to edit the variables in it)<br>
Fixed Comments.asp when no Entry specified.<br>
Fixed ViewCat.asp nullifying records (Causing an error for Calendar) GoogleBot reported that!<br>
Fixed EditEntry.asp not closing the recordset (Not sure when I created that problem as it worked before).<br>
Fixed Spell Check error when correction contains appostrophe.<br>
Fixed Random Quote Plugin (Down to a 1 in 679 chance it would error) I cannot believe im hitting them all today!.<br>
</p>

<p><b>12th July 2004 (V1.0.5.02)</b><br>
Fixed RSS Feed validation (Thanks Joe for reporting that).<br>
Fixed Spell Check failing to accept user corrections with words with symbols e.g. "BlogX,".<br>
Updated a possible problem with Database collision (Possible I say).<br>
Updated security to accept the cookie as a direct entry to Admin pages.<br>
</p>

<p><b>11th July 2004 (V1.0.5.01)</b><br>
Added "Winamp NowPlaying" plugin.<br>
Added "Random Quotes" plugin.<br>
Added my first user submitted template "Black" (Thanks to Kiz for that).<br>
Fixed link to Poll Results if in an Admin page.<br>
Fixed link on Default.asp that goes to "EditMainPage.asp".<br>
</p>

<p><b>10th July 2004 (V1.0.5.01)</b><br>
Changed Count.asp (Nothing important).<br>
Fixed Comments.asp problem once and for all.<br>
Fixed "Results.asp" not checking if user has voted.<br>
Fixed typo in "CheckForUpdate.asp".<br>
Updated Footer.asp to nullify "Records".<br>
Updated all pages to not nullify and recreate "Records" (it "<i>might</i>" cause a problem says a Microsoft Article).<br>
</p>

<p><b>09th July 2004 (V1.0.5.00)</b><br>
Added words to ZIP's "UserDictionary".<br>
Fixed replace.asp adding a comma to a link.<br>
Fixed "UploadPicture.asp" causing a client side Javascript error when there's an appostrophe in the URL.<br>
Updated WinBlogX installer to install dependencies (and provide a link in the readme to the Visual Basic ones).<br>
</p>

<p><b>08th July 2004 (V1.0.5.00)</b><br>
Fixed demo plugin that shows "Last 5 Entry Titles" showing 6 (Well done to those who can count ;)).<br>
Modified "Credits.txt" to credit PoorMan's SpellCheck.<br>
Removed other dictionaries (so zip file is smaller).<br>
Updated About.asp to credit PMSC, mention beta testers, list new features etc.<br>
</p>

<p><b>07th July 2004 (V1.0.5.00)</b><br>
Added "AllowEditLinks".<br>
Added "CheckForUpdate".<br>
Added Pingback viewer to the comments page.<br>
Fixed "BlogIt" when SiteURL contains an appostrophe.<br>
Mass problems with Comments.Asp (Naomi & Sarah reported this)...think it's fixed.<br>
Notified Mailing List.<br>
Added Spell Check.<br>
Updated Download.asp to not run through a server side component (Since the URL is no longer secret).<br>
Updated Plugin.asp.<br>
Updated pingback to check for the same IP address pinging back for the same entry.<br>
Updated ZIP and Freewebs mirror.<br>
</p>

<p><b>06th July 2004 (V1.0.5.00)</b><br>
Worked on Spell Check.<br>
</p>

<p><b>05th July 2004 (V1.0.5.00)</b><br>
Added a link to SpellCheck in editor (Not yet finished).<br>
Added "NoDate" for CMS.<br>
Deleted PlanetSourceCode zip.<br>
Fixed Count.asp.<br>
Fixed selecting a day messing up the current day highlighting.<br>
Researched WinBlogX RTF/HTML (RTF a waste of time).<br>
Updated Config.asp to include "Dim.asp" and "Database.asp" (Now easier to update).<br>
Updated WinBlogX (New "Check For Update" option).<br>
</p>

<p><b>04th July 2004 (V1.0.5.00)</b><br>
Added moderation to WhoUses.asp.<br>
Added legacy option to Themes.asp.<br>
Fixed Javascript error on AddEntry.asp, triggered when an Entry is saved and the page unloaded.<br>
Updated LegacyMode functions.<br>
</p>

<p><b>03rd July 2004 (V1.0.5.00)</b><br>
Fixed sloppy programming resulting in returning commenters not being able to comment (Unreleased Beta).<br>
Fixed CommentNotify.asp having undefined variables.<br>
Updated comments.asp to remove subscription checkbox for logged in users.<br>
</p>

<p><b>01st July 2004 (V1.0.5.00)</b><br>
Fixed today's day being shown on another month.<br>
</p>

<p><b>30th June 2004 (V1.0.5.00)</b><br>
Added the toolbar bookmarklet.<br>
</p>

<p><b>28th June 2004 (V1.0.5.00)</b><br>
Added the ability to edit any entry.<br>
Added the ability to delete any entry.<br>
Fixed XMLTimeZones being converted by client RSS (Thanks Dan).<br>
Updated Orange stylesheet to underline all links.<br>
</p>

<p><b>24rd June 2004 (V1.0.5.00)</b><br>
Added "UseImagesInEditor" option in "Config.asp"<br>
Added EditLinks.asp.<br>
Added Links.txt file.<br>
Updated text editor to set focus.<br>
Updated text editor to allow highlighting text to link (Thanks Dan).<br>
Updated text editor to ask if you want the link in a new window (Thanks Dan).<br>
Updated calendar to show red (I remember now why I gave up with Bold, it was invisible) for Today (Thanks Dan).<br>
Updated XMLRPC to new version.<br>
</p>

<p><b>23rd June 2004 (V1.0.4.00)</b><br>
Added a new option, "CalendarCheck", which highlights only days with posts (Thanks Dan).<br>
Added "TimeOffset" to offset the time, Surprising really (Thanks Dan).<br>
Added "EditPoll.asp".<br>
Fixed "Any HTTP header modifications must be made before writing page content" on RSS when using IIS 5.0 or less (Thanks Dan).<br>
Fixed calendar on Admin pages.<br>
Fixed Comments.asp dimension.<br>
Fixed ASPUpload not found leaving database open.<br>
Updated comments.asp to disallow new comments if they are disabled, In event of a comment spam attack (Thanks Tom).<br>
Updated Links.asp to open links in a new window (Thanks Dan).<br>
Looked at possible problems with CommentSpam (Thanks Tom).<br>
Updated config.asp to better handle quotes.<br>
Updated config.asp to list manual options only (Thanks Dan).<br>
Updated "Includes/config.asp" to better explain each option (Thanks Dan).<br>
</p>

<p><b>22nd June 2004 (V1.0.4.00)</b><br>
Added another poll results page.<br>
Fixed comments RSS where the querystring was empty.<br>
Fixed dimension of calender.<br>
Fixed dimension of already dimensioned values in EditMainPage.asp.<br>
Fixed image upload link to stylesheet and progressbar.<br>
Fixed LegacyMode.<br>
Fixed logout link.<br>
Fixed problems with ReaderPassword.<br>
Fixed someone else's XMLRPC script. (Now I know i'm a pro ;-) )<br>
Fixed search when on any Admin pages.<br>
Renamed "Blue.gif" to "Bar.gif" (for more colours).<br>
Released code.<br>
Tested new release.<br>
Updated default.asp to clear buffer.<br>
Updated Download.asp to fix script timeout.<br>
Updated documentation.<br>
Updated planet source code.<br>
</p>

<p><b>21th June 2004 (V1.0.4.00)</b><br>
Added dimension of some variables.<br>
Added option explicit to some pages.<br>
Added poll.<br>
Fixed double headers when no main page.<br>
Fixed mailing list.<br>
Fixed NotFound.asp's case sensitive domain check.<br>
Fixed WAP not closing the database.<br>
Removed case sensitivity in usernames and passwords.<br>
Updated Comments.asp Spam protection (Just incase someone did some clever form spoofing)<br>
Updated NotFound.asp's download link.<br>
</p>

<p><b>20th June 2004 (V1.0.4.00)</b><br>
Added Ban lookup to pingback.<br>
Added Password protected entries and modified all pages accordingly.<br>
Added subscription to comments.<br>
Fixed the link to disclaimer from Admin pages.<br>
Fixed photo upload.<br>
Tested pingback fully.<br>
Fixed pingback fully.<br>
Fixed URL linking again.<br>
Updated comments.asp so "Email" reads as "E-Mail" (Thanks RadicalPuppy4).<br>
</p>

<p><b>19th June 2004 (V1.0.4.00)</b><br>
Added confirmation to exiting the mail page without submitting the form.<br>
Tested exploiting the XML HTTP to access local file. (No Security risk)<br>
Tested URL linking fully.<br>
Worked on URL linking again.<br>
</p>

<p><b>18th June 2004 (V1.0.4.00)</b><br>
Added confirmation to exiting the mail page without submitting the form.<br>
Added PingBack Client.<br>
Fixed "main page" link on About.asp<br>
</p>

<p><b>17th June 2004 (V1.0.4.00)</b><br>
Added Printer friendly page.<br>
Worked on PingBack client.<br>
</p>

<p><b>12th June 2004 (V1.0.4.00)</b><br>
Fixed URL linking, if link was hidden behind a "(".<br>
Fixed URL linking, if link had a ")" after it.<br>
</p>

<p><b>06th June 2004 (V1.0.4.00)</b><br>
Added "Archive" to nav.<br>
Fixed page moving on specified dates.<br>
Fixed case sensitive domain check.<br>
Re-Branded BlogX to also use new domain "BlogX.co.uk".<br>
</p>

<p><b>05th June 2004 (V1.0.4.00)</b><br>
Optimised code, Database only opened once.<br>
Optimised code, closed Records reused.<br>
</p>

<p><b>04th June 2004 (V1.0.4.00)</b><br>
Fixed search engine "don't complete words".<br>
Fixed search engine occurance count on "any" mode.<br>
Moved all Admin pages to a seperate folder (Thanks Tom for the suggestion).<br>
</p>

<p><b>03nd June 2004 (V1.0.4.00)</b><br>
Added a link to e-mail on the search results (and that solves the mystery of the missing envelope).<br>
Fixed problems with databases being left open on a few pages (I'm <b>VERY</b> sorry).<br>
</p>

<p><b>02nd June 2004 (V1.0.4.00)</b><br>
Added confirmation to exiting the AddEntry page without submitting the form.<br>
Fixed entries having a title longer than 80 characters.<br>
</p>

<p><b>30th May 2004 (V1.0.4.00)</b><br>
Added "PingBack" for "Ping-O-Matic".<br>
Tidyied up the "Config.asp".<br>
Tested PingBack client (No Problems with "WordPress").<br>
Updated RSS to use the "Entry=" querystring.<br>
</p>

<p><b>29th May 2004 (V1.0.4.00)</b><br>
Added "credits.txt".<br>
Added Error handling to image upload main page (forgot to add it to the file page).<br>
Added "PingBack" table to database.<br>
Added untested PingBack client (works internally).<br>
Fixed category link on "Comments.asp".<br>
Tidyied up the search engine (both apperance and code).<br>
Updated "ViewItem.asp" to use the "Entry=" querystring.<br>
</p>

<p><b>28th May 2004 (V1.0.4.00)</b><br>
Added an occurance count to search.<br>
Added "complete words" to search.<br>
Fixed search where an URL contained a term, breaking up the link.<br>
</p>

<p><b>27th May 2004 (V1.0.4.00)</b><br>
Added Comments RSS.<br>
Fixed Category RSS where an Un-Encoded category is empty.<br>
Fixed Category RSS when an Un-Encoded field is empty.<br>
</p>

<p><b>25th May 2004 (V1.0.4.00)</b><br>
Added a search (both "any order" and exact match).<br>
</p>

<p><b>22nd May 2004 (V1.0.4.00)</b><br>
Added Category select to "Add Entry".<br>
Added Comments.asp to auto add a "http://" if not already included.<br>
Added a "legacy" mode to BlogX, so you can now kill all the fancy stuff i've added and go "classic".<br>
Updated "Email.gif".<br>
Updated WinBlogX to erase WinBlogX.ini if password/server/folder is wrong.<br>
Updated WinBlogX to auto advance login if credentials already found.<br>
Updated WinBlogX to auto save password by default.<br>
</p>

<p><b>24th April 2004 (V1.0.3.06)</b><br>
Added Smiley mode in postings.<br>
Fixed a few pages links to go to "Main.asp" instead of "Default.asp".<br>
Fixed null when "0" is the count of "MailingListMembers" or "Refer".<br>
Updated "Mail.asp" to align center & use site stylesheet.<br>
Updated "ViewItem.asp" to link to report EOF's to the webmasters.<br>
</p>

<p><b>22nd April 2004 (V1.0.3.05)</b><br>
Added greying out of the "Prev Page" if it's already on the first.<br>
Updated comments to hide e-mails from non-admins (What was I thinking before!).<br>
Updated Cat RSS in Zip (been lazy in providing new version).<br>
</p>

<p><b>14th April 2004 (V1.0.3.05)</b><br>
Added BlogX mirroring (Share.asp).<br>
Fixed RSS on null categories/titles.<br>
Updated "Orange" theme template.<br>
</p>

<p><b>13th April 2004 (V1.0.3.05)</b><br>
Added BlogX to PlanetSourceCode again (Any publicity is good publicity).<br>
Updated MailingList.asp & About.asp to reflect this.<br>
</p>

<p><b>08th April 2004</b><br>
Added pagecount to RSS, RSS now shows last 10 entries.<br>
Fixed the documentation in the ZIP, it was showing up the "WinBlogX Readme" instead.<br>
Updated the ZIP file again, replaced all files with my copys just incase I missed a few.<br>
Updated the ZIP file's database to use the orange theme by default.<br>
Updated the orange theme so work with the comments table.<br>
</p>

<p><b>23nd March 2004</b><br>
Updated WinBlogX in the ZIP files "Bin" directory.<br>
</p>

<p><b>21st March 2004</b><br>
Added ability to hide the OtherLinks from the Config.asp.<br>
Added installer to WinBlogX (and finally finished).<br>
Fixed RSS When SiteDescription had an "&amp;" or any other strange symbol.<br>
Dropped the theme "Sky Blue" from the Zip, download is now 300kb less.<br>
Dropped a forgotten debugging line of "Application.asp" which threw off the password checks for WinBlogX.<br>
Updated WinBlogX to fix strange characters.<br>
Updated WinBlogX to fully encode characters.<br>
Updated WinBlogX's error handeling.<br>
Updated License.txt.<br>
Updated "Sea" template.<br>
</p>

<p><b>12th March 2004</b><br>
Updated footer to link more than just "BlogX".<br>
Updated a few themes.<br>
Worked On An Undercover Script.<br>
</p>

<p><b>09th March 2004</b><br>
Fixed comment notification linking with a wrong URL.<br>
Updated comment notification to prevent notification when the admin comments.<br>
</p>

<p><b>05th March 2004</b><br>
Updated WinBlogX to include new RTF functions.<br>
Updated RSS to include optional picture, new information etc.<br>
</p>

<p><b>03rd March 2004</b><br>
Added "Remember Me" to Comments.<br>
Fixed RSS when title was empty (Feedreader looped a "NEW ENTRY", Sarah's doing ;-) ).<br>
</p>

<p><b>02nd March 2004</b><br>
Added Comments Banning System.<br>
Added Comments "You've already posted" System.<br>
Added Comments Delete Function.<br>
</p>

<p><b>01st March 2004</b><br>
Added Comments.<br>
Updaed "Referers" & "Count.asp" to flag LAN addresses.<br>
</p>

<p><b>25th February 2004</b><br>
Added WAP site.<br>
Added E-Mail parser.<br>
Added "SkyBlue" Theme.<br>
</p>

<p><b>24th February 2004</b><br>
Added "Comments".<br>
Added "Comments" to "Data" table to count comments on a post.<br>
Dropped "CommentsURL" and replaced it with "EnableComments" in the DB.<br>
Fixed page recordset to use file names (Cat is now passed in its own querystring).<br>
Fixed "ViewCat.asp" to allow switching through pages.<br>
Updated "MailingList.asp" to warn users of "testing" e-mail addresses.<br>
Updated "MailingList.asp" to display helpful info for already subscribed users.<br>
Updated all pages to show the new Comments link.<br>
Updated "ViewCat.asp" to not link to "ViewCat.asp".<br>
</p>

<p><b>23rd February 2004</b><br>
Fixed a few pages which defaulted to "Default.asp" on the "Go Back" buttons (*Annoying*).<br>
Fixed "Replace.asp".. Strange strange bugs.<br>
</p>

<p><b>22nd February 2004</b><br>
Added MainPage (EditMainPage.asp, Default.asp).<br>
Added "RTF.js" to "EditDisclaimer.asp".<br>
Fixed a minor glitch in a missing "</Span>".<br>
Fixed a minor glitch in "Edit Last Entry" appearing on further pages.<br>
Updated content boxes to be bigger.<br>
Updated database to include "EnableMainPage".<br>
Updated upload picture to include "MainPage" handeling.<br>
Updated "RTF.js" to include an if statement for querystrings.<br>
Updated "Disclaimer.asp" to be held within a box.<br>
Updated MailingList to block AOL. (Can't send mail to them)<br>
</p>


<p><b>15th February 2004</b><br>
Added "NotFound.asp".<br>
Added source for "Count.asp" (To settle any concerns).<br>
Added "WhoUses.asp".<br>
Added "WinBlogX.exe" to the source.<br>
Fixed a minor glitch in image uploading.<br>
Fixed "Default.asp" & "ViewItem.asp" to not error if the Category was set as null.<br>
Fixed "OtherLinks.txt" not working on a few pages where I used "count" before.<br>
Fixed "MailingList" after a messup.<br>
Fixed Images in RSS.<br>
Updated Database to include "ScriptRefer".<br>
Updated Documentation to include "Disallowed Parent Path" information.<br>
Updated RSS.<br>
Updated "MailingList" to convert line breaks into HTML line breaks.<br>
Updated "Includes/Mail.asp" to send the from name and to name with a few components.<br>
</p>

<p><b>13th February 2004</b><br>
Updated RSS.<br>
</p>

<p><b>12th February 2004</b><br>
Added a download license.<br>
Added mailing list Webmaster notification.<br>
Crippled PlanetSource Code's Code to cut down on piracy. (Full version available AFTER accepting License).<br>
Fixed URL replace to handle VbCrlf's better.<br>
Updated About.asp so now source code can only be downloaded by subscribed users (Sick of people abusing the license).<br>
Updated "MailingList" sending.<br>
</p>

<p><b>10th February 2004</b><br>
Added Forum support.<br>
Added & Updated mailing list.<br>
Added "Upload Picture" ability.<br>
Fixed hyperlinking ONLY when there is a space or linebreak before a URL (Image URL problems).<br>
Fixed RSS to convert "Images/Articles/" to the FULL URL.<br>
Fixed image uploading for thoose whoose paths were different (Now uses MapPath for image paths).<br>
Updated formatting tools.<br>
Updated "EditLastEntry" to include formatting tools.<br>
</p>

<p><b>09th February 2004</b><br>
Added "IncludeHTM.txt" (Gets included just before the footer).<br>
Added "UploadPicture.asp".<br>
Updated "AddEntry.asp" to include formatting tools.<br>
Updated clarification of the license agreement a bit (after a few people have removed my ONE line copyright).<br>
</p>

<p><b>07th February 2004</b><br>
Added/Updated "Themes".<br>
Added "BlankTemplate.zip" (Information on how to make your own theme).<br>
Updated "Themes.asp".<br>
</p>

<p><b>06th February 2004</b><br>
Added "Themes".<br>
Added "Themes.asp".<br>
Added a theme preview querystring to "Header.asp".<br>
Created "Themes" & updated a few.<br>
Fixed RSS as somehow setting nothing wasn't the same as nothing?!?. (IsNUll) is now checked.<br>
Updated "AddEntry.asp" to use "MaxLength" attribute.<br>
Updated WinBlogX to use "MaxLength" attribute.<br>
</p>

<p><b>05th February 2004</b><br>
Added "ReaderPassword", Now You Can Restrict Who Reads The Blog (*Database, RSS, ViewerPassword*).<br>
Fixed RSS to use Password attribute...Should a "ReaderPassword" be implemented.<br>
</p>

<p><b>04th February 2004</b><br>
Updated WinBlogX as someone got mixed up with the Hexidecimal for "&" and the ASCII code for it (*Simple mistake ;-) *).<br>
</p>

<p><b>03rd February 2004</b><br>
Updated disclaimer, it conflicted with the "license.txt".<br>
</p>

<p><b>01st February 2004</b><br>
Added a "Show/Hide" function to the "Recent Changes" (Hidden by default).<br>
Added WinBlogX documentation, Source code, Zip...<br>
Fixed RSS handling of "&" after Elin caused a syntax error.<br>
Fixed RSS handling of tags, after new RSS downloads had paragraphs displayed as tags due to "&" converting further down the script.<br>
Updated "About.asp" not to show the link to more information on the more information page.<br>
Updated "About.asp" to hide/allow the ability to hide a few more options.<br>
Updated documentation to have seperate files for additional information. (Smaller mainpage).<br>
Updated WinBlogX to convert a few non Alphanumeric characters to more friendly ones. (Don't know the VB equivallant of ASP's Server.URLEncode() )<br>
</p>

<p><b>31th January 2004</b><br>
Added a "Colour Picker" for background colour.<br>
Added RSS for all (*Really Simple Syndication*).<br>
Added RSS by category (*Really Simple Syndication*).<br>
Added "All" To Categories.<br>
Updated "About.asp" now we have RSS + more.<br>
Updated "Logging.asp" to include more of the refering URL<br>
Updated Database to allow a longer ReferURL<br>
Updated Header to include "Auto RSS Discovery"
Updated Footer to no longer link to SimpleGeek.com.<br>
Updated Documentation to include seperate parts (OtherLinks.txt, Updating, Config Definitions).<br>
</p>

<p><b>30th January 2004</b><br>
Added ability to edit the last entry (Still firmly believe that's as far as you should edit ;-) )<br>
Added "Contact Me" option, EmailServerSettings and the support to use a range of ASP Mail components<br>
Included "Application.asp" for WinBlogX.<br>
Updated security (Yet again) so that it knows after line 2 whether your logged in or not. (No this wasn't a vulnerability).<br>
Updated a possible vulnerability in which ".Inc" files can be read as plain text if your webServer isn't setup to parse them (*oops*).<br>
Updated all the includes to the world defined extension of ".asp" extension in case servers arn't configured to handle ".inc" files (See Above).<br>
</p>

<p><b>28th January 2004</b><br>
Updated About Page.<br>
Updated Pages To Include a "&lt;P&gt;" and updated the post page not to include them<br>
(Now "No Data Entry" is easier to track..Database fields are a few characters shorter...etc etc).<br>
Ready To Release WinBlogX 1.0! (No User Documentation Yet)<br>
</p>

<p><b>27th January 2004</b><br>
Updated a HUGE vulnerability in which session state was being called BEFORE "CookieName" was being defined.<br>
Updated SQL Exploit checking to allow category names with " ' " in them.<br>
Updated "Matthew1471 Homepage" NOT to show.<br>
</p>

<p><b>26th January 2004</b><br>
Distributed on the Internet.<br>
Added A "Register" mode in which your site can get added to future versions of "OtherLinks.txt"<br>
</p>

<p><b>25th January 2004</b><br>
Added documentation & started distributing.<br>
Deleted "ChangeLog.txt" ("About.asp" will replace it).<br>
Researched RSS & Blogging (I just might get into that).<br>
Updated "Calendar.inc" to highlight current day. (Bug Fix)<br>
Updated Windows client to make postings. (Still in Beta State).<br>
Updated "About.asp".<br>
</p>

<p><b>24th January 2004</b><br>
Added a "CheckSessionCookies" login check (Not Sure How Useful This Is).<br>
Updated "Remember Login" to store a cookie (Nav.inc, ?ClearCookie).</p>

<p><b>23rd January 2004</b><br>
Added "ShowCategories" option (Nav.Inc, AddEntry, ViewItem, Default.asp).<br>
Fixed a problem where ReferURL contained an " ' " (Possible SQL Security Problem).<br>
Fixed Broken timings of "011:23" because someone got their "<12" and "<10" muddled on the "add a 0 before the number code".<br>
Allowed user to take advantage of "About.Asp" and Added an acronym for the version number over the "Powered By" link.<br>
Updated database to know what's "Required" and what's not.<br>
Updated copyright messages to "2004" and added a "License.Txt".<br>
Updated "ChangePassword.Asp" to not show the existing password (It was confusing), also simplified the words more.</p>

<p><b>22nd January 2004</b><br>
Updated clarity on what is compulsory in "Config.Asp" and what's not.<br>
Updated CSS to have a darker blue header and a blue entry header.<br>
Added option to hide "Add comment" should "CommentsURL" be blank.<br>
Fixed option to change background color.</p>

<p><b>21st January 2004</b><br>
Fixed problem where "Category" contained a space, then modified all pages to decode and encode it appropriately.<br>
Added "OtherLinks.txt" and implemented format checking and size checking.</p>

* = Not Worthy Of A New Version Number :P

</Div></Div><%End If%>

<DIV class=entry>
<h3 class=entryTitle>Thanks To</h3><br>
<DIV class=entryBody>
<p>
Max Web Portal<br>
Freevbcode.com<br>
Developerfusion.com<br>
kirchmeier.org<br>
ASPxmlrpc SourceForge Team<br>
<br>All the beta testers and people who have got in contact with suggestions and reported bugs<!--- Sarah, Dan, Naomi, Russ, Tom your all great guys! Yay you found an easteregg ;-) ---></p>

<p>Lee (IndaUK) for the incredible FireFox JS compatibility that I could never have possibly done</p>

<p>Obviously a huge thanks to Chris Anderson without his site design there would <u>DEFIANTLY</u> not have been a Matthew1471 BlogX today!</p>
</Div></Div>

<!--- End Information --->

<p align="Center"><a href="<%=PageName%>">Back To The Main Page</a></p>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
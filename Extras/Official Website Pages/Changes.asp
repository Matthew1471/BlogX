<%
' --------------------------------------------------------------------------
'¦Introduction : Change Log Page.                                           ¦
'¦Purpose      : Displays the official change log.                          ¦
'¦Used By      : About.asp.                                                 ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp.¦
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
CacheHandle(CDate("11/04/12 23:32:00"))

PageTitle = "BlogX Change Log"
%>

<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<div id="content">

 <!--- Start Version Information -->
 <div class="entry">
  <h3 class="entryTitle">About BlogX - Change Log / Version History</h3>

<div class="entryBody">
<p style="text-align: center">This site is running Matthew1471's version of BlogX V<%=Version%>.<br/>
You can download a copy <a href="http://blogx.co.uk/Download.asp">here</a>.</p>

<p style="text-align: center; font-size: smaller"><b>Note:</b> This site has new changes/fixes/features which are not available in the main package, this is while they are tested for bugs/errors.
Unreleased changes/fixes/features are listed as having "(Unreleased Beta)" next to them.</p>

<p style="text-align: center">This is the latest version and includes the following changes/additions/fixes:</p>

<p><b>18 August 2012 (Unreleased Beta)</b><br/>
Tidied up the awful non compliant code in Themes.asp.
</p>

<p><b>16 May 2012 (Unreleased Beta)</b><br/>
Fixed posts being made on the same date but on different months not having a new dated heading (ViewCat.asp, Main.asp, Search.asp). 
</p>

<p><b>11 April 2012 (Unreleased Beta)</b><br/>
Changed HTML2Text to ByVal.<br/>
Fixed cross-site scripting vulnerabilities in About.asp, Search.asp, Nav.asp (<a href="http://secunia.com/advisories/48573">Credits</a>).<br/>
Fixed some XHTML issues with Search.asp.<br/>
Fixed poor grammar in Search.asp.<br/>
Optimised About.asp checking for "ShowChanges" that is no longer required on that page.
</p>

<p><b>29 August 2011 (Unreleased Beta)</b><br/>
Fixed EditEntry.asp not checking title length.<br/>
</p>

<p><b>10 April 2011 (Unreleased Beta)</b><br/>
Fixed API including blank categories.<br/>
Optimised Nav.asp category listing so that a check is done in the DB instead of VB.
</p>

<p><b>04 March 2011 (Unreleased Beta)</b><br/>
Updated API to state Blog version in header.<br/>
Added Category fetching method for API (Used in BlogX for Android).
</p>

<p><b>23 November 2010 (Unreleased Beta)</b><br/>
Fixed error message when all links are deleted but enabled (thanks Jo).<br/>
Fixed error message when comment was submitted &amp; validated but e-mailing was not enabled.<br/>
Updated Comment_Validate.asp to use a less intensive locktype.
</p>

<p><b>17 November 2010 (Unreleased Beta)</b><br/>
Fixed bulk comment deletion by IP output and corrupt entry comment counts.<br/>
Fixed LastComments.asp not correctly linking to the IPWhois.asp.
</p>

<p><b>21 February 2010 (Unreleased Beta)</b><br/>
Fixed Comments.asp, ViewItem.asp &amp; RSS/Comments.asp from not stripping out ","'s in querystrings.
</p>

<p><b>14 February 2010 (V2.2*)</b><br/>
Fixed RSS feed validation issues on other servers (thanks Chip).
</p>

<p><b>30 December 2009 (V2.2)</b><br/>
Fixed Comments.asp &amp; ViewItem.asp erroring when a spurious entry number is specified which includes a -<br/>
Optimised Comments.asp when checking if Requested is specified.
</p>

<p><b>10 October 2009 (V2.2)</b><br/>
Updated Replace to be more case insensitive and more efficient.
</p>

<p><b>15 September 2009 (V2.2)</b><<br/>
Fixed spelling mistakes in CommentNotify.asp, Comments_Validate.asp, MailingList.asp, Unsubscribe.asp.
</p>

<p><b>29 July 2009 (V2.2)</b><br/>
Added Last.fm plugin to the ZIP.<br/>
Updated Buddies.asp to use an alternative method of loading in the RSS and provide better errors.<br/>
Updated List Plugin in ZIP to XHTML standards.<br/>
Updated Readme.txt in Plugins folder.
</p>

<p><b>26 July 2009 (V2.2)</b><br/>
Fixed Home screen not automatically linking and adding emoticons (Thanks James Collett).<br/>
Fixed EditLinks.asp not recognising relative links won't work (Thanks James Collett).<br/>
Fixed Delete.gif not loading in Downloads.asp if the blog is not in the root folder (Thanks James Collett).<br/>
Updated Buddies.asp to provide more descriptive errors and handle sites with weird encodings (Thanks James Collett).<br/>
Fixed NotFound.asp redirecting users to the count page when the wrong URL contained "Count.asp".
</p>

<p><b>18 July 2009 (V2.1*)</b><br/>
Added option to send draft to AddEntry page (AddEntry.asp, EditDraft.asp).<br/>
Fixed an error in EditEntry.asp if Categories field was null.<br/>
Fixed a spelling mistake in EditEntry.asp.
</p>

<p><b>13 July 2009 (V2.1*)</b><br/>
Updated "Check Consistency" Query in DB to not fail when there are no comments for an entry.<br/>
Updated "CheckCommentsCount.asp" in the Developer Tools to not miss incorrect records, follow XHTML Strict and fix.
</p>

<p><b>06 July 2009 (V2.1*)</b><br/>
Minor changes to the database properties such as enabling more unicode compression etc.<br/>
Updated Default stylesheet for better Firefox compatibility.
</p>

<p><b>02 June 2009 (V2.1)</b><br/>
Fixed NAV.asp not redirecting correctly on login if BlogX is in a subfolder.
</p>

<p><b>23 May 2009 (V2.1)</b><br/>
Added confirmation warning when deleting links.<br/>
Improved login code so it remembers previous page.<br/>
Optimised Purge.asp.
</p>

<p><b>11 April 2009 (V2.1)</b><br/>
Fixed browsers with aggressive caching not performing conditional GETs.
</p>

<p><b>23 March 2009 (V2.1)</b><br/>
Fixed XHTML issues with UnSubscribe.asp.
</p>

<p><b>02nd February -> 22nd March 2009 (V2.0)</b><br/>
Added further comments and fixed XHTML issues on numerous pages.<br/>
Added additional variable so that the IP address to withhold adverts to can be altered easily.<br/>
Fixed EditEntry.asp losing the password of password protected entries.<br/>
Fixed MailingListMembers.asp repeating the subject.<br/>
Updated MailingListMembers SQL &amp; code to separate unsubscribed members easily.
</p>

<p><b>18th January 2009 (V2.0)</b><br/>
Added comments to Cache.asp, Database.asp, Dim.asp, Header.asp, Footer.asp, ExtraHTM.txt, Languages.asp, Mail.asp, NAV.asp &amp;.<br/>
Optimised ExtraHTM.txt &amp; Footer.asp by not switching contexts as much.
</p>

<p><b>17th January 2009 (V2.0)</b><br/>
Optimised readability of Dim.asp by seperating variables into their source.<br/>
Optimised readability of Bar.asp by converting ' to &quot;.
</p>

<p><b>10th January 2009 (V2.0)</b><br/>
Optimised readability of ViewCat.asp, Dim.asp, Database.asp and included comments.
</p>

<p><b>09th January 2009 (V2.0)</b><br/>
Added comments to ViewCat.asp.<br/>
Optimised ViewCat.asp &amp; Main.asp to not use such an intensive cursor location and locktype.<br/>
Optimised ViewItem.asp &amp; ViewCat.asp to not use such an intensive cursor and locktype.<br/>
Optimised cursortype &amp; locktype on Vote.asp in a few places so as not to use such an intensive locktype.<br/>
Optimised Vote.asp by using a case statement instead of multiple IFs.<br/>
Updated comments in Vote.asp.
</p>

<p><b>03rd January 2009 (V2.0)</b><br/>
Removed old commented out Firefox poll removing code.<br/>
Updated ASPSmartUpload UploadPicture_Save.asp to have the same coding improvements as the Persits.<br/>
Updated UploadPicture_Save.asp to detect errors before saving so an error could not leave a temp file.<br/>
Updated No upload UploadPicture.asp to conform to XHTML and other coding improvements.
</p>

<p><b>02nd January 2009 (V2.0)</b><br/>
Optimised Newspaper.asp to use variable instead of repeatedly calling Server.MapPath.<br/>
Optimised Themes.asp to use variable instead of repeatedly calling Server.MapPath.<br/>
Updated ASPSmartUpload UploadPicture.asp to meet similar coding improvements as Persits UploadPicture.asp.<br/>
Updated UploadPicture.asp to remove some white space and merge better with other files.<br/>
Updated NotFound.asp to handle URLs better (not complete).<br/>
Updated Includes/Replace.asp to include URL encoding functions (not complete).<br/>
Updated Config.asp to include PageTitle and absolute path option for hosts that do not support parent paths.<br/>
Updated NewsPaper.asp to include absolute path option for hosts that do not support parent paths.<br/>
Updated Pingback.asp to include absolute path option for hosts that do not support parent paths.<br/>
Updated Themes.asp to include PageTitle and absolute path option for hosts that do not support parent paths.
</p>

<p><b>01st January 2009 (V2.0)</b><br/>
Optimised locktype on NotFound.asp to not use such an intensive locktype.<br/>
Updated Comments.asp to use less intensive locktype, lock errors should error in right place (where error code is).
</p>

<p><b>30th December 2008 (V2.0)</b><br/>
Added NotFound.asp page for admins to detect broken links (NAV.asp, Admin/NotFound.asp).<br/>
Fixed XHTML issue with pingbacks in Comments.asp.<br/>
Fixed Official/Count.asp incorrectly calling replace function when URL contained localhost.<br/>
Fixed HTML2Text converting &amp; too late (Replace.asp).<br/>
Updated NotFound.asp to not log the admin NotFound.asp page.
</p>

<p><b>25th December 2008 (V2.0)</b><br/>
Fixed Admin/Pingback.asp refering to BackgroundColor but Templates/Config.asp never being included to set it.<br/>
Optimised Admin/Default.asp, Official/Default.asp, Official/FacebookRSS.asp, Official/OPML.asp, RSS/Default.asp, SNS/Default.asp &amp; Templates/Default.asp to skip checking for session variables.<br/>
Optimised readability of Official/FacebookRSS.asp, Official/NoSites.asp, Official/OPML.asp, RSS/Default.asp &amp; SNS/Default.asp.<br/>
Optimised Official/NoSites.asp, Official/SNS.asp by including conditional GET support.<br/>
Optimised Official/FacebookRSS.asp, RSS/Default.asp &amp; RSS/Cat/Default.asp using too intensive cursortype and locktypes.<br/>
Optimised RSS/Default.asp using too intensive cursor location (adUseClient).<br/>
Updated Official/FacebookRSS.asp &amp; RSS/Default.asp to use similar code so it is easier to merge changes.
</p>

<p><b>24th December 2008 (V2.0)</b><br/>
Added popup screen center code for comment preview (Comments.asp).<br/>
Added HTML escape for comment preview so that it represents how it will display on the site better.<br/>
Added disk space error and undo upload method to UploadPicture_Save.asp.<br/>
Added more comments to count.asp to improve readability.<br/>
Fixed content type for UpdateBlogX.asp, Update.asp &amp; Count.asp.<br/>
Fixed XHTML issues (&lt;font) in UploadPicture.asp.<br/>
Fixed XHTML issues in comment preview.<br/>
Renamed BuyCD.asp to Donate.asp.<br/>
Optimised Count.asp (restructured else-if to improve execution).<br/>
Optimised UpdateBlogX.asp, Update.asp, Download/Default.asp &amp; Count.asp to skip checking for session variables.<br/>
Optimised comment preview for readability.<br/>
Optimised UploadPicture_Save.asp &amp; Count.asp using too intensive cursortype and locktypes.<br/>
Optimised UploadPicture.asp code (handles errors better).<br/>
Updated AddFile_Save.asp &amp; UploadPicture_Save.asp to use similar code so it is easier to merge changes.<br/>
Updated Donate.asp to automatically show BlogX's age, be XHTML standard compliant and reduced text.
</p>

<p><b>23th December 2008 (V2.0)</b><br/>
Fixed Comments.asp concurrency issue with attempting to delete already deleted records due to JET caching.<br/>
Optimised Comments.asp using too intensive cursortype and locktypes.<br/>
Optimised Comments.asp by specifying cursortype and locktypes as command params.<br/>
Optimised Comments.asp by removing old close recordset code (inserted to try and remove concurrency issues).<br/>
Optimised Comments.asp by removing deprecated unreferenced recordset count.
</p>

<p><b>22nd December 2008 (V2.0)</b><br/>
Added popup screen center code to RTF.js.<br/>
Optimised RTF.js readability.<br/>
Optimised Database.asp readability.<br/>
Optimised Mail.asp readability.
</p>

<p><b>21st December 2008 (V2.0)</b><br/>
<span style="color:red">Added Picture to FileExtensions table (to distinguish which files work as an img src).</span><br/>
Added clearer error messages to ParseEmails.asp.<br/>
Added popup centering code for UploadPicture.asp (Includes\RTF.asp).<br/>
Added a 150 URL limit to the display of Referrers.<br/>
Added check to UploadPicture_Save.asp to check file extensions are allowed (instead of forcing JPG).<br/>
Fixed Referrers.asp not encoding &amp;.<br/>
Fixed XHTML issues with EditPoll.asp, AddPoll.asp, EmailConfig.asp, MailingListMembers.asp, ParseEmails.asp, Pingback.asp, Spell.asp &amp; UploadPicture.asp<br/>
Fixed MailingListMembers.asp sending out non standard compliant XHTML e-mails.<br/>
Fixed the comment notification e-mails (Comments.asp) not centering when opened up on Firefox/Webkit based browsers.<br/>
Fixed MailingListMembers.asp not DIMensioning mail variables (though the error was masked).<br/>
Fixed ParseEmails.asp not DIMensioning Path.<br/>
Fixed Referrers not being checked for HTML.<br/>
Fixed Referrers code in Header.asp often including the SiteURL as a referrer if it was in a different case.<br/>
Fixed Toolbar.asp having my blog URL hardcoded into the pingback.<br/>
Fixed Toolbar.asp refering to old time offset variables.<br/>
Fixed non Firefox friendly code in Toolbar.asp.<br/>
Fixed user specifying &amp; as their homepage not being converted to &amp;amp; (Replace.asp).<br/>
Fixed Toolbar.asp inserting a rogue &lt;/div&gt;.<br/>
Fixed shared files save file (AddFile_Save.asp) not checking admin credentials.<br/>
Fixed image upload javascript insertion for Mozilla, broke by XHTML update to AddEntry etc (UploadPicture.asp, UploadPicture_Save.asp).<br/>
Fixed a few minor bugs in Mail.asp (i.e. a stray Response.Write).<br/>
Improved readability on admin files and included additional comments (and small fixes to existing ones).<br/>
Optimised EditPoll.asp, EmailConfig.asp, LastComments.asp, Pingback.asp, AddEntry.asp, EditEntry.asp &amp; Toolbar.asp using too intensive cursortype and locktypes.<br/>
Optimised EditPoll.asp, AddPoll.asp, EmailConfig.asp, Spell.asp, Toolbar.asp, EditEntry.asp &amp; AddEntry.asp by specifying cursortype and locktypes as command params.<br/>
Optimised Error_Spell.asp by removing unnecessary SSI.<br/>
Optimised cursor type being specified then overridden by the Open params in LastComments.asp.<br/>
Optimised the SQL in Toolbar.asp for the category lookup to use the optimal DISTINCT keyword (DBMS is faster than interpreted code).<br/>
Optimised the SQL in AddFile_Save.asp by removing the ORDER BY when checking if in extensions table.<br/>
<span style="color:red">Updated Refer table to accept 255 length URLs.</span><br/>
Updated Header.asp to reflect increase in refer URL.
</p>

<p><b>20th December 2008 (V2.0)</b><br/>
Fixed concurrency issue with Comments.asp while under spam attack with spam record being deleted twice.<br/>
Optimised cursor type being specified then overridden by the Open params in Comments.asp.
</p>

<p><b>17th December 2008 (V2.0)</b><br/>
Fixed XHTML &amp;amp; decoding issues in Firefox with AddFile.asp.<br/>
Fixed XHTML issues with AddFile.asp, AddFile_Save.asp.<br/>
Optimised the SQL in AddEntry.asp for the category lookup to use the optimal DISTINCT keyword (DBMS is faster than interpreted code).<br/>
Optimised wording on AddFile.asp.
</p>

<p><b>16th December 2008 (V2.0)</b><br/>
Added e-mail notification icons to Comments.asp.<br/>
Optimised the SQL in AddEntry.asp for the category lookup to use the optimal DISTINCT keyword (DBMS is faster than interpreted code).<br/>
Optimised the SQL in NAV.asp for the category lookup to use the optimal DISTINCT keyword (DBMS is faster than interpreted code).<br/>
</p>

<p><b>15th September 2008 -> 04th December 2008 (V2.0)</b><br/>
Added randomisation of cookie for those who did not manually alter it before.<br/>
Added fix for Firefox's over zealous caching not performing conditional GETs.<br/>
Added attempt to guess default path in Plugin.asp.<br/>
Added auto-lock out to Application.asp.<br/>
Added code to attempt to deal with deadlocks during commenting concurrency.<br/>
Fixed paging issue in unreleased beta.<br/>
Fixed a few more XHTML non-standard pages.<br/>
Fixed &amp; not being properly re-encoded by EditEntry.asp<br/>
Fixed a few pages not appropriately handling EOFs (unlikely to occur, sample DB contains records).<br/>
Fixed very minor things like tabs instead of spaces in ASP comments.<br/>
Fixed ShowCat not being dimensioned and erroring AddEntry.asp when categories are turned off.<br/>
Fixed cookies disabled error not firing.<br/>
Fixed typo in "ErrorDescription" in Downloads.asp.<br/>
Fixed a leftover &lt;/span&gt; being wrongly inserted when categories are turned off.<br/>
Fixed downloads.asp not properly handling empty folders.<br/>
Moved change log to Changes.asp page on BlogX.co.uk.<br/>
Optimised the SQL in NAV.asp to use the optimal DISTINCT keyword (DBMS is faster than interpreted code).<br/>
Optimised a few stylesheets (Matthew1471, Orange) to use CSS margin shortcuts et al.<br/>
Optimised Printer_Friendly.asp to not switch between ASP and HTML as much.<br/>
Removed some quotes from Quotes.txt to remove some of the less interesting/offensive quotes.<br/>
Removed last-modified in ASP comments as it was not practical to maintain.<br/>
Updated About.asp to reflect the blog's features better.
</p>

<p><b>15th September 2008 (V2.0)</b><br/>
Added more descriptive page titles to some pages to improve SEO (user request, thanks Don).<br/>
Optimised a few more files (inappropriate locktype, cursortype, NOT x instead of if statement).
</p>

<p><b>06th September 2008 (V2.0)</b><br/>
Added ShortenEntry function to Replace.asp.<br/>
Fixed Paragraph truncate in Main.asp.<br/>
Merged small include files into calling pages, Header.asp and NAV.asp.<br/>
Optimised Main.asp to not perform not needed checks.
</p>

<p><b>24th August 2008 (V2.0)</b><br/>
Added admin login throttle to prevent brute force attacks.<br/>
<span style="color:red">Added "BannedLoginIP" for invalid admin login throttle.</span>.<br/>
Added SSL support for logins.<br/>
Optimised Header.asp security to not run some pointless functions.<br/>
</p>

<p><b>18th August 2008 (V2.0)</b><br/>
Optimised Comments_Validate.asp to remove additional IF statement.<br/>
Updated Comments_Validate.asp to send standards compliant e-mails.
</p>

<p><b>17th August 2008 (V2.0)</b><br/>
Fixed LastComments.asp not using the correct link for Ban User.<br/>
Updated Mail.asp to use 2 page anti-spam validation given the success it has been for comments.asp.<br/>
Updated Comments.asp to use DELETE FROM instead of iterating through an inefficient recordset for auto-purge.asp.<br/>
Updated Comments.asp purge of multi-spam to use DELETE from to prevent concurrency issues and also to improve speed.<br/>
Updated CommentNotify.asp to not send confirmation if admin requested unsubscribe. (I click unsubscribe links in bounces).
</p>

<p><b>15th August 2008 (V2.0)</b><br/>
Added un-validated deletion to Comments.asp on re-post to prevent user-lockout.<br/>
Added a descriptive error message to Comments.asp instead of blindly re-directing.<br/>
Added option to see the last EntryPageSize number of validated comments.<br/>
Updated Comments.asp to no longer automatically redirect as some (few) spammers can perform that.
</p>

<p><b>14th August 2008 (V2.0)</b><br/>
Added additional spam validation check to refuse to accept Mail and comments if the user disconnected early.<br/>
Added IP spoof checking to Mail.asp.<br/>
Added option to IPWhoIs.asp to delete all comments by a particular IP address.<br/>
Renamed Validate.asp to Comments_Validate.asp as there are now 2 validation files.<br/>
Updated Purge.asp to use the new comments validation link and use a hyperlink instead of javascript onclick.<br/>
Updated Comments.asp to give more accurate feedback as to whether the user passed spam protection and removed the hidden image.<br/>
Updated RTF.js to use XHTML compatible code.<br/>
Updated PrintPopup to center.<br/>
</p>

<p><b>13th August 2008 (V2.0)</b><br/>
Optimised more code.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
Fixed Calendar not correctly closing hyperlinks when SortByDay <> True.<br/>
Fixed some fields not being DIMed in ChangePassword.asp.<br/>
<span style="color:red">Added "LastModified" to Main and Disclaimer.</span>.<br/>
Added LastModified handling for EditMain.asp and EditDisclaimer.asp.<br/>
Removed some smaller include files and merged them into parent files (OtherLinks.asp, Links.asp, Calendar_Querystrings.asp, Security.asp) to reduce confusion and optimise loading.<br/>
</p>

<p><b>12th August 2008 (V2.0)</b><br/>
Optimised more code.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
Rewrote documentation.<br/>
Fixed SNS bar not loading for non logged in users.<br/>
Fixed Search.asp not showing podcast plugin.
</p>

<p><b>10th August 2008 (V2.0)</b><br/>
Optimised more code.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
Changed links to whois.domaintools.com instead of apnic.<br/>
Added ReplyTo to Includes/Mail.asp and others.<br/>
Added more explicit error explanations to Comments.asp.<br/>
Added HTML escape for mail.asp (just for privacy and those who use insecure webmail).<br/>
Fixed Comments.asp not checking Banned IPs for POSTs.<br/>
Updated e-mail scripts to set SPF compatible &quot;from&quot; addresses.<br/>
</p>

<p><b>09th August 2008 (V2.0)</b><br/>
Optimised more code.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
Added a ban proxy button on Comments.asp.<br/>
Added a prompt to delete comment, ban ip, ban proxy.<br/>
Updated Comments.asp to force comments ascending.
</p>

<p><b>08th August 2008 (V2.0)</b><br/>
Optimised more code.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
<span style="color: red">Dropped BannedIP.BannedTime, Comments.CommentedTime, Comments_Unverified.CommentedTime</span><br/>
<span style="color: red">Fixed Data.UTCTimeZoneOffset not storing leading 0.</span><br/>
<span style="color: red">Deleted Deprecated Buddies table.</span><br/>
Fixed Comment count being incremented for comments and then decremented when found to be spam and deleted.<br/>
Added cache handling code to Purge.asp and EditBan.asp.<br/>
Fixed Comments.asp not converting &amp; symbol to &amp;amp;.
</p>

<p><b>04th August 2008 (V2.0)</b><br/>
Fixed RSS giving invalid build dates.<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
</p>

<p><b>03rd August 2008 (V2.0)</b><br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
</p>

<p><b>27th July 2008 (V2.0)</b><br/>
Added cache handler to Download.asp, Mail.asp<br/>
Fixed more XHTML compatibility to files.<br/>
Improved HTML readability on some files.<br/>
</p>

<p><b>15th July 2008 (V2.0)</b><br/>
Fixed Last-Modified code not correctly determining whether records and databases have been terminated (IsEmpty, IsObject and Is Nothing behave very differently).
</p>

<p><b>10th July 2008 (V2.0)</b><br/>
Fixed Last-Modified code not terminating any detected records and databases.
</p>

<p><b>09th July 2008 (V2.0)</b><br/>
Fixed Last-Modified code (and added it to more page) on pages.<br/>
Added more comments and tidied more code.<br/>
Added more XHTML compatibility to files.<br/>
Fixed Main.asp, ViewCat.asp adding a link to Page=1 (just was making google index more).<br/>
Fixed URL linker including tags if they were straight after the URL.<br/>
</p>

<p><b>08th July 2008 (V2.0)</b><br/>
Fixed Last-Modified code (and added it to more page) on pages and made modular now I understand it more.<br/>
Added more comments and tidied more code.<br/>
Added more XHTML compatibility to files.
</p>

<p><b>01st July 2008 (V2.0)</b><br/>
Added declarations to more files to explain what they do and what links to them.
</p>

<p><b>15th June 2008 (V2.0)</b><br/>
Fixed the spam comment display showing in ascending order rather than descending.<br/>
Fixed invalid HTML in About.asp.<br/>
Fixed "EnableMainPageRequest" not declared in EditMainpage.asp.<br/>
Fixed some CSS issues in Matthew1471! stylesheet.<br/>
Improved About.asp, Main.asp, Default.asp, Includes/Header.asp, Includes/Footer.asp, Includes/Calendar.asp, Includes/Nav.asp to follow the XHTML 1.0 standard.<br/>
Improved EditMainPage.asp to not update if the checkbox was not checked (boolean was being compared to a string and ASP treated them differently).<br/>
</p>

<p><b>13th June 2008 (V2.0)</b><br/>
<span style="color: red">Changed "Time" in Banned to "BannedTime", same for "Date".</span><br/>
<span style="color: red">Changed "Time" in Comments to "CommentedTime", same for "Date".</span><br/>
<span style="color: red">Separated "Comments" and "Comments_Unvalidated" to speed up comment handling.</span><br/>
<span style="color: red">PUKValidated field dropped (see above).</span><br/>
Added "noreply" address for user comment notifications.<br/>
Added Mail.asp to check banned addresses after a user reported Mail.asp abuse.<br/>
Improved Mail.asp to not wait until user has entered in a message on Mail.asp before informing them e-mail is disabled.<br/>
Fixed local users being incorrectly logged as "Cache" in referrers.<br/>
Rewrote random photo selector for newspaper.asp.<br/>
Tidied and Optimised Newspaper.asp code.<br/>
Changed Includes/Replace.asp to be more XHTML compliant.<br/>
Changed some tags in pages to start using XHTML syntax.<br/>
Tidied buddies script and improved error handling.</p>

<p><b>07th May 2008 (V2.0)</b><br/>
Fixed SelectColor.asp using a non existent variable.<br/>
Fixed invalid HTML code in SelectColor.asp.<br/>
Improved SelectColor.asp readability.<br/>
Optimised SelectColor.asp.</p>

<p><b>04th May 2008 (V2.0)</b><br/>
Fixed Mail.asp returning blank field.</p>

<p><b>16th April 2008 (V2.0)</b><br/>
Fixed bug in picture upload where 2/3 variables were not defined.<br/>
Removed 1 deprecated variable in picture upload.</p>

<p><b>25th March 2008 (V2.0)</b><br/>
Added declarations to some files to explain what they do and what links to them.<br/>
Indented both source code and output code for easier readability.</p>

<p><b>26th February 2008 (V2.0)</b><br/>
Added HTTP referrer information to the mail.asp page.<br/>
Improved links to the mail page to not use a query string (it was throwing off spiders and web crawlers).<br/>
Fixed Printer_Friendly.asp not specifying an encoding (now specifies UTF8).</p>

<p><b>13th January 2008 (V2.0)</b><br/>
Fixed the WhoUses page creating a "_new" window instead of a blank one.</p>

<p><b>28th December 2007 (V2.0)</b><br/>
Fixed strWord not being dimensioned in Admin/Spell.asp.</p>

<p><b>24th December 2007 (V2.0)</b><br/>
Changed RSS feeds to "utf-8" from my local character set.<br/>
Fixed EditPoll.asp not dimensioning a variable (oops).</p>

<p><b>26th November 2007 (V2.0)</b><br/>
Fixed bug in WAP site referencing obsolete ShowCat variable.<br/>
</p>

<p><b>18th November 2007 (V2.0)</b><br/>
Fixed Comments RSS feed from not checking input was a number.<br/>
</p>

<p><b>12th September 2007 (V2.0)</b><br/>
Fixed Application.asp using a renamed variable (TimeOffset).<br/>
Turned on OPTION EXPLICIT on a few more files (AddEntry.asp, AddPoll.asp etc) to improve execution speed.<br/>
Fixed a large number of variables not being dimensioned on Admin pages (whoops).<br/>
Renamed "ShowCat" to "ShowCategories" (Config.asp x 2, Main.asp, NAV.asp, Comments.asp, DIM.asp, Search.asp).<br/>
</p>

<p><b>23rd July 2007 (V2.0)</b><br/>
Improved time reporting in RSS feeds (Added on AddEntry.asp).<br/>
Moved time detection to separate file and tidied up Config.asp (It's not a really important option.. it can be almost hidden).<br/>
Fixed bug in CommentsRSS that made them crash (shoddy testing).<br/>
<span style="color: red">Added UTCTimeZoneOffset to Comments and Data.</span><br/>
</p>

<p><b>12th August 2007 (V2.0)</b><br/>
Added a simple Last-modified algorithm.. not completed.<br/>
</p>

<p><b>11th August 2007 (V2.0)</b><br/>
Fixed RSS mishandling Enclosure URLs that have a http:// in them.<br/>
</p>

<p><b>21st July 2007 (V2.0)</b><br/>
Improved Pingback to check and truncate for SourceURI's longer than 255.<br/>
<span style="color: red">Updated SourceURI in Pingback to be of length 255.</span><br/>
</p>

<p><b>15th July 2007 (V2.0)</b><br/>
Fixed RSS timestamps being invalid (I just gave the date instead of date and month to the weekday calculator).<br/>
Updated RSS timestamps to truncate without LEFT (it was one of the parameters).<br/>
</p>

<p><b>01st February 2007 (V2.0)</b><br/>
Fixed URL linking including a bracket when there's a full stop after it.<br/>
</p>

<p><b>14th February 2007 (V2.0)</b><br/>
Fixed comment notifications being sent to the spammers too! (ie. unvalidated commenters)<br/>
</p>

<p><b>15th January 2007 (V2.0)</b><br/>
Added Error handling to IPWhois.asp for when ASPDNS is not installed.<br/>
</p>

<p><b>11th January 2007 (V2.0)</b><br/>
Fixed Comments.asp auto-purge purging new comments after 25th December 2006 (Unreleased Beta only).<br/>
</p>

<p><b>26th November 2006 (V2.0)</b><br/>
Updated Header.asp to remove more features on LegacyMode.<br/>
Updated "Default" template CSS to handle overflowing entries.<br/>
Updated WaterFall template CSS.<br/>
</p>

<p><b>25th November 2006 (V2.0)</b><br/>
Fixed RSS trimming to be able to handle nested HTML tags too!<br/>
Fixed variable not declared in updated ServerError.asp (should only effect my site).<br/>
Updated RSS Trimming to properly detect end of HTML tags and to link to more information.<br/>
Moved RSS string functions to a separate include file (RSSReplace.asp).<br/>
</p>

<p><b>08th October 2006 (V2.0)</b><br/>
Updated EditEntry.asp to scrollbar the textarea if there's an overflow (This is to fix an IE bug).<br/>
</p>

<p><b>17th September 2006 (V2.0)</b><br/>
Updated plugin to exit loop once it has found it (didn't realise ASP supported EXIT commands).<br/>
</p>

<p><b>16th September 2006 (V2.0)</b><br/>
Fixed Comments RSS displaying spam comments. (Oops)<br/>
Fixed Comments RSS dynamically loading filed names in. (Oops)<br/>
Fixed Purge.asp not displaying spammers name (not like it matters).<br/>
Fixed Bug in Thumbnail.php which could over-write images if no folder was specified. (Thanks Andrew for your PHP help in fixing it).<br/>
</p>

<p><b>15th September 2006 (V2.0)</b><br/>
Added "Photo Mode".<br/>
Disabled Bug in Thumbnail.php which could over-write images if no folder was specified. (Need Andrew's PHP help to fix).<br/>
</p>

<p><b>07th August 2006 (V2.0)</b><br/>
Added Entry Security (not that it was really needed), it's now almost impossible to randomly POST comments on entries without visiting Comments.asp.<br/>
Added automatic unvalidated comment deletion (after 1 day) to Comments.asp (Can be turned on/off).<br/>
Fixed Typo in AddEntry.asp (2 Lefts).<br/>
Updated AddEntry.asp to generate a EntryPUK (to fit in line with new security).<br/>
<span style="color: red">Added EntryPUK field to Data</span>.<br/>
</p>

<p><b>01st August 2006 (V2.0)</b><br/>
Fixed the encoding on the adbrite adverts.<br/>
Fixed ServerError.asp when error source contained unescaped HTML it could mess up error page layout.<br/>
Fixed capitalisation in ServerError.asp.<br/>
Updated Config.asp to include better descriptions of variables (some were redundant).<br/>
Updated Config.asp to no longer have a Links/Other Links Path.<br/>
Updated OtherLinks.asp and Links.asp to read from database.<br/>
Updated ServerError.asp to take into account the lack of a links file.<br/>
Updated EditLinks.asp to take into account the lack of a links file.<br/>
Updated Buddies.asp to read fields from database.<br/>
<span style="color: red">Added Links table to DB</span>.<br/>
</p>

<p><b>26th July 2006 (V2.0)</b><br/>
Updated documentation regarding updating to newer versions as old information was too cumbersome for dramatically changed databases.<br/>
</p>

<p><b>07th July 2006 (V2.0)</b><br/>
Fixed a bug in the SwimmingPool template in which long worded entries would force the navigation to the bottom of the page (this exists in other themes, they will be updated as I use them.. or if I'm contacted).<br/>
</p>

<p><b>03rd July 2006 (V2.0)</b><br/>
Fixed a bug in comments which a comment ending with a URL while the user is using a proxy made the hyperlink overflow.<br/>
</p>

<p><b>13th June 2006 (V2.0)</b><br/>
Fixed a beta bug in which flush was called on the comments page then headers were modified.<br/>
Fixed Category RSS to include URLEncode methods.<br/>
</p>

<p><b>12th June 2006 (V2.0)</b><br/>
<span style="color: red">Added Enclosure field to DB</span>.<br/>
Added the option to delete polls on the edit polls page.<br/>
Updated AddEntry, EditEntry, Comments, ViewItem, ViewCat, ProtectedEntry to allow enclosures.<br/>
Updated RSS to allow enclosures.<br/>
Updated the RSS SQL to no longer dynamically include fields, now uses FIXED field names (Big speed increase).<br/>
Updated a CSS bug in Swimming Pool and Orange (60% and 30% width).<br/>
</p>

<p><b>11th June 2006 (V2.0)</b><br/>
Fixed users rushing off the "Comment Submitted" page could have their entries marked as spam, so I've flushed the picture before a DB operation.<br/>
Updated "Add Shared Files" with a few minor cosmetic changes (,'s appearing with spaces etc).<br/>
Updated Comment.asp to have one Proxy IP exception for MY site (You can go through and change it to yours).<br/>
Updated "Swimming Pool" CSS &amp; engrish description.<br/>
</p>

<p><b>07th June 2006 (V2.0)</b><br/>
Few changes to update the limit number of spam comments shown to Admin (re-wording etc).<br/>
</p>

<p><b>06th June 2006 (V2.0)</b><br/>
Fixed error in PictureViewer if the file didn't exist.<br/>
Added option to limit number of spam comments shown to Admin (625 records is slowing down page generation).<br/>
</p>

<p><b>30th May 2006 (V1.0.7.2)</b><br/>
Added error handling to Downloads.asp.<br/>
Updated Downloads.asp header to "Shared Files".<br/>
</p>

<p><b>28th May 2006 (V1.0.7.2)</b><br/>
Added another CSS definition for sectionBody.<br/>
Fixed section overflow in firefox making NAV drop to the bottom (Finally).<br/>
Removed FireFox compatibility mode for polls as now fixed.<br/>
Updated 2 plugins to center.<br/>
</p>

<p><b>24th May 2006 (V1.0.7.2)</b><br/>
Started validating the HTML and fixed sloppy HTML to make it more complient (Main.asp, Default.asp, Replace.asp, Footer.asp, Header.asp).<br/>
</p>

<p><b>20th May 2006 (V1.0.7.2)</b><br/>
Updated Comments.asp to assume either false or true for Subscriptions (Spammers are sending malformed posts which are doing nothing more than clogging up my inbox).<br/>
</p>

<p><b>19th May 2006 (V1.0.7.2)</b><br/>
Added code to autotruncate entries after the second paragraph (Requires future revision).<br/>
Investigated CSS bugs with photos' (Possibly down to unstandardized code).<br/>
</p>

<p><b>17th May 2006 (V1.0.7.2)</b><br/>
Fixed wrongly dimensioned variables causing WinBlogX SNS to not post entries.<br/>
Fixed WebServer in SNS client and implemented non-persistant subscriptions.<br/>
Finished WinBlogX SNS' auto-retry code (Requires future revision).<br/>
</p>

<p><b>16th May 2006 (V1.0.7.2)</b><br/>
Rebranded RTS to SNS after unmanageable page ranking expectations.<br/>
</p>

<p><b>17th April 2006 (V1.0.7.2)</b><br/>
<span style="color: red">Added "FileExtensions" table for file extensions.</span><br/>
Added "Shared Files".<br/>
Finished RTS Frontend.<br/>
Updated Photo Uploader code to look better in FireFox.<br/>
</p>

<p><b>15th April 2006 (V1.0.7.2)</b><br/>
Finished RTS Frontend.<br/>
</p>

<p><b>14th April 2006 (V1.0.7.2)</b><br/>
Investigated RTS Frontend.<br/>
</p>

<p><b>12th April 2006 (V1.0.7.2)</b><br/>
Fixed AdminToolkit accepting "http:", "www.", ".com", "/" and "\"'s as sign up addresses (just made it crash or looked stupid, no security risk).<br/>
</p>

<p><b>11th April 2006 (V1.0.7.2)</b><br/>
Fixed BlogX Spider not updating records (database corruption and sloppy error trapped code).<br/>
Fixed BlogX Spider failing with weird encodings (Microsoft bug in XMLHTTP).<br/>
Started programming RTS integration to site (Limitations in AJAX though).<br/>
</p>

<p><b>06th April 2006 (V1.0.7.1*)</b><br/>
Ported BlogX Spider to VB (too many users are using BlogX for an ASP page to be able to handle).<br/>
</p>

<p><b>03rd April 2006 (V1.0.7.1*)</b><br/>
Fixed PictureViewer.asp's awful code (Table inside table, loops instead of if's, it now validates).<br/>
Updated PictureViewer.asp to now use FSO instead of Persits (Thanks Thom for the reminder).<br/>
</p>

<p><b>29th March 2006 (V1.0.7.1)</b><br/>
Fixed Comment notification being sent BEFORE validation.<br/>
Updated admin notification so it's not sent till validation. (There aren't really any false positives)<br/>
</p>

<p><b>24th March 2006 (V1.0.7.0)</b><br/>
Started programming WinBlogX RTS!<br/>
</p>

<p><b>27th February 2006 (V1.0.7.0)</b><br/>
Added Andrews' Thumbnail Generation Script (Requires PHP).<br/>
</p>

<p><b>27th January 2006 (V1.0.7.0)</b><br/>
Fixed Comments CSS for "LighterBlue".<br/>
</p>

<p><b>03rd January 2006 (V1.0.7.0)</b><br/>
Added a new theme "LighterBlue".<br/>
Fixed ThemeIT crashing on empty files.<br/>
Fixed ThemeIT re-diming already an dimmed variable.<br/>
</p>

<p><b>01st October 2005 (V1.0.7.0)</b><br/>
Added a character limit for RSS feeds.<br/>
</p>

<p><b>08th September 2005 (V1.0.7.0)</b><br/>
Fixed Vote.asp not checking for words in form submission, Someone tried exploiting this on my site (No Security Risk, just a page crash).<br/>
</p>

<p><b>20th August 2005 (V1.0.7.0)</b><br/>
Added IP Validation so that spoofed IP addresses cannot be used (Comments.asp, Validate.asp, Includes/Header.asp).<br/>
Fixed intermitent bug with BannedAddresses returning a "Record Already Deleted" error when removing banned entries.<br/>
<span style="color: red">Updated "Comments" table to include "PUKValidated"</span><br/>
</p>

<p><b>15th August 2005 (V1.0.7.0)</b><br/>
Fixed repeated poll title in NAV when viewing results page (Thanks Kevin Whipp).<br/>
</p>

<p><b>30th July 2005 (V1.0.7.0)</b><br/>
Fixed crash on requesting the calendar page directly (No Security Risk).<br/>
</p>

<p><b>27th July 2005 (V1.0.7.0)</b><br/>
Fixed crash on invalid numbers being passed to the Vote.asp page (No Security Risk).<br/>
</p>

<p><b>13th July 2005 (V1.0.7.0)</b><br/>
Fixed crash on months greater than 12 in calander querystring.<br/>
</p>

<p><b>11th July 2005 (V1.0.7.0)</b><br/>
Fixed crash on words in Main.asp querystring.<br/>
</p>

<p><b>03st June 2005 (V1.0.7.0)</b><br/>
Fixed AdminToolkit reporting "/Admin" as URLs' in ranking.<br/>
Finished Adding ability to disable inactive blogs in AdminToolkit.<br/>
</p>

<p><b>01st May 2005 (V1.0.7.0)</b><br/>
Researched ability to disable inactive blogs in AdminToolkit.<br/>
</p>

<p><b>30th May 2005 (V1.0.7.0)</b><br/>
Added "Preview" to comments.<br/>
</p>

<p><b>29th May 2005 (V1.0.7.0)</b><br/>
Fixed emoticon replacing characters without spaces.<br/>
</p>

<p><b>12th April 2005 (V1.0.7.0)</b><br/>
Added internal "IPWHOIS.asp" to Comments system.<br/>
</p>

<p><b>31st March 2005 (V1.0.7.0)</b><br/>
Added Donation links to my main pages. (I know awful :$)<br/>
Updated Admins' Config Template listing to work with multiple hosted blogs.<br/>
</p>

<p><b>28th March 2005 (V1.0.7.0)</b><br/>
Added "Tools" (HTML Snippets) to entry editors (AddEntry, EditEntry)<br/>
</p>

<p><b>16th March 2005 (V1.0.6.00*)</b><br/>
Fixed PingBack not running due to recent OPTION EXPLICIT (Admin/Pingback.asp and Includes/XMLRPC.asp).<br/>
Fixed minor bug with XMLRPC not displaying correct entry number in confirmatin.<br/>
</p>

<p><b>05th March 2005 (V1.0.6.00*)</b><br/>
Fixed StyleSheet loading on Comments, MailingList, Mail e-mails when "TemplatesURL" is used.<br/>
Updated Count.asp to distinguish between your localservers and my LAN requests.<br/>
</p>

<p><b>03rd March 2005 (V1.0.6.00*)</b><br/>
Removed the offensive language from Hacker.asp due to a user complaint (Greg).<br/>
</p>

<p><b>25th February 2005 (V1.0.6.00)</b><br/>
Fixed EditLinks.asp not DIM'ing 2 variables (Thanks Gareth).<br/>
Updated Comments.asp to force a SORT (Incase the sorting is messed up, Thanks Gareth).<br/>
</p>

<p><b>21st February 2005 (V1.0.6.00)</b><br/>
Fixed AdminToolkits' Count.asp.<br/>
</p>

<p><b>20th February 2005 (V1.0.6.00)</b><br/>
Added a security check to Header.asp to ensure the database is not stored in a "WWWroot".<br/>
Added a hidden variable "TemplateURL" for hosting enviroments (32.5mb Saved).<br/>
Fixed Count.asp not correctly identifying "localhost" queries.<br/>
Finished Programming PalmBlogX to a redistributable standard.<br/>
Updated Database.asp to append SiteURL with a "/" if it hasn't already been.<br/>
Updated Count.asp to ignore Google Translation, Google Image Search.<br/>
Updated AdminToolKit to write in a TemplateURL.<br/>
</p>

<p><b>19th February 2005 (V1.0.6.00)</b><br/>
Fixed *SERIOUS* SQL Exploit in Comments.asp (Sorry :$ ).. No Known attacks yet.<br/>
Fixed MailingListMembers.asp, A disabled user would stop the MailingList from sending.<br/>
Updated BETA ZIP (Sorry to those who noticed it was faulty).<br/>
Updated Comments.asp to be more proxy complient (After an untrackable abusive comment on Hannahs' blog).<br/>
Updated Login form to a full URL path, rather than a relative (In case you are running ASPs in external folders).<br/>
</p>

<p><b>15th February 2005 (V1.0.6.00)</b><br/>
Added a hideously eyecatching "NoticeText" function for companies/urgent message.<br/>
Added optional ability in Comments.asp to disable commenting on entries older than a week.<br/>
Changed the character set to UTF-8 for compatibility.<br/>
Updated Comments.asp to prevent a HTML injection attack.<br/>
</p>

<p><b>10th February 2005 (V1.0.6.00)</b><br/>
Fixed bug with BlogX Spider having "Count" already DIMd.<br/>
</p>

<p><b>07th February 2005 (V1.0.6.00)</b><br/>
Fixed bug with Comments.asp sending notification to the same e-mail address (When via a different IP).<br/>
</p>

<p><b>06th February 2005 (V1.0.6.00)</b><br/>
Added rel="nofollow" to Commenters' post to try to stop spam <a href="http://www.google.com/googleblog/2005/01/preventing-comment-spam.html">More Info</a>.<br/>
Created a BETA ZIP to ease distribution of the next release.<br/>
Fixed bug with NAV.asp not declaring "SplitText, WordLoopCounter".<br/>
Fixed bug with Plugin.asp often declaring SplitText and WordLoopCounter.<br/>
Updated all "The Teen Forum" accounts to run the new release.<br/>
Updated AdminToolKit DISABLE to include the DIM.asp to prevent errors.<br/>
</p>

<p><b>05th February 2005 (V1.0.6.00)</b><br/>
Fixed Includes\Spell to handle no FSO (Again, fixing someone elses' code).<br/>
Updated RTF.js code to include hide/show ID instead of defining it twice in AddEntry and EditEntry.<br/>
Updated RTF on EditEntry to match RTF on "AddEntry".<br/>
Updated PingBack.asp to include more error handling (FSOdisabled, Write permissons..) and the ability to show errors.<br/>
Updated PingBack.asp to include "OPTION EXPLICIT" and DIMd 3 variables.<br/>
Updated "Setup.asp" page, to check the databases' tables and fields.<br/>
</p>

<p><b>03rd February 2005 (V1.0.6.00)</b><br/>
Added a "Setup.asp" page, allowing server component information checks (and database checks in future).<br/>
Added an unsupported "MoveRec" page, to help move records into their correct places.<br/>
Fixed an expected end in ViewerPassword.asp (not before noticed due to error trapping)<br/>
Fixed ViewerPassword.asp and ReaderLogin.asp not closing the database before redirecting.<br/>
Fixed Comments RSS, PictureViewer, Pingback DIMing the now already DIMd count.<br/>
Fixed BlogX entering an infinate loop if FSO is disabled on your webhost. (WebHost users should have no problems)<br/>
Fixed Config Template selection if FSO is disabled on your webhost. (WebHost users should have no problems)<br/>
Fixed Themes.asp if FSO is disabled on your webhost. (WebHost users should have no problems)<br/>
Removed ErrorTrapping for database missing tables AboutPage and EmailConfig, you should have added these by now<br/>
Updated the rest of the SQL to no longer dynamically include fields, now uses FIXED field names (Big speed increase).
Updated DIM.asp to include variables from Links.asp (and removed them from Links.asp).<br/>
Updated Pingback.asp to include OPTION EXPLICIT, and had to DIM 2 or 3 variables i'd missed out.<br/>
Updated Footer not to include "Mailing List" if e-mail is turned off.<br/>
</p>

<p><b>02nd February 2005 (V1.0.6.00)</b><br/>
Fixed loads of bugs in which you could Overtype in the EmailConfig fields.<br/>
Updated MailingListMembers.asp, ParseEmails.asp to remove duplication.<br/>
Updated more pages' SQL to no longer dynamically include fields, now uses FIXED field names (big speed increase).
</p>

<p><b>01st February 2005 (V1.0.6.00)</b><br/>
<span style="color: red">Changed Database "12HourTimeFormat" to "ShortTimeFormat" (SQL didn't like the number in the query).</span><br/>
<span style="color: red">Changed Database "BackgroundColor" field size from 50 to 20 (It's a 6 digit hex value most of the time anyway!).</span><br/>
Fixed Error on ChangePassword if a new empty password was tried.<br/>
Fixed Spelling mistake on ChangePassword.asp.<br/>
Fixed loads of bugs in which you could Overtype in the config and Drafts fields.<br/>
Fixed bugs in EditEntry so it now physically restricts over typing in the Category and Title field (Why did I miss that?!).<br/>
Updated more pages' SQl to no longer dynamically include fields, now uses FIXED field names (Big speed increase).
</p>

<p><b>30th January 2005 (V1.0.6.00)</b><br/>
<span style="color: red">Added Database "StopComments" to "Data" (can now stop comments per entry instead of only being able to stop all comments).</span><br/>
Added ability to lock comments on an entry to the EditEntry page (see above, comments.asp also was changed).<br/>
Fixed EditEntry crashing if LegacyMode is turned on (Null is apparantly a date).<br/>
Fixed BlogX newspaper beta, Random image script erroring.<br/>
Updated Most pages' SQl to no longer dynamically include fields, now uses FIXED field names (Big speed increase).<br/>
Updated CommentNotify.asp to display a warning if NO querystrings are passed.
Updated Poll to split up sentences into a bunch of words with line feeds, had to center the poll (Unsure if I like it like that).
</p>

<p><b>29th January 2005 (V1.0.6.00)</b><br/>
Added more error trapping to PingBack (Permission denied, Object required etc...)<br/>
Added a new TitleIcon CSS definition.<br/>
Added comments to the "Default" CSS.<br/>
Fixed ViewCat.asp &amp; Search.asp not taking into account "EntriesPerPage".<br/>
Fixed RSS Comments so the comment times validate.<br/>
Fixed Cat RSS so "EntriesPerPage" is correctly parsed.<br/>
Programmed BlogX newspaper beta.</p>

<p><b>20th January 2005 (V1.0.6.00)</b><br/>
Fixed a minor bug in Drafts in which AlertBack made a minor browser script error after the first save.
</p>

<p><b>16th January 2005 (V1.0.6.00)</b><br/>
Fixed Comments.asp trying to send an e-mail to the same e-mail address (Simple solution I just kept forgetting to fix).<br/>
Updated "Advanced Tools" to improved functionality ("Advanced" Button now disappears when clicked, compatible with Firefox etc).<br/>
Updated Comments.asp to tell the blog owner what entry number was commented on in the subject (Like the changes I made for the visitors).<br/>
Updated Comments.asp to provide a direct link to the new comment.</p>

<p><b>02nd January 2005 (V1.0.6.00)</b><br/>
Fixed Comments.asp trying to send an e-mail to a blank e-mail address.<br/>
</p>

<p><b>01st January 2005 (V1.0.6.00)</b><br/>
Added a few more user Submissions to ZIP.<br/>
Fixed line feeds in HTML code being converted to "&lt;br>"'s.<br/>
Fixed search engine trying to display HTML markup code (it'll now convert it to text).<br/>
</p>

<p><b>28th December 2004 (V1.0.5.10)</b><br/>
<span style="color: red">Added a "Draft NotePad" option.</span><br/>
Added ability to remove uploading images (If you don't have an upload component and arn't likely to get one).<br/>
Added support for the free "ASPSmartUpload" upload component (Upon Mikes' request).<br/>
Fixed the CheckForUpdate.asp "Back" link.<br/>
Splitted a few of the ZIP files (Themes, Dictionaries) into separate zips available <a href="http://matthew1471.co.uk/Downloads.asp">here</a> (to keep the ZIP size low).<br/>
</p>

<p><b>27th December 2004 (V1.0.5.10)</b><br/>
Fixed FireFox javascript compatibility (All Lee's handywork, I just copied and pasted, Thanks Lee!).<br/>
Programmed BlogXProxy (Allow multiple users to post to one blog).<br/>
</p>

<p><b>20th December 2004 (V1.0.5.10)</b><br/>
Fixed any Commenter field (aside from Comment) > 50 meant loosing the comment.<br/>
Optimized "Refer" to not dynamically load in fields, all pages run faster.<br/>
</p>

<p><b>18th December 2004 (V1.0.5.10)</b><br/>
Added "Sandy" theme.<br/>
Optimized "Refer" to not dynamically load in fields, all pages run faster.<br/>
</p>

<p><b>15th December 2004 (V1.0.5.10)</b><br/>
Added "Advanced" button to EditEntry and AddEntry to unhide "Extra" features.<br/>
Added "Extra" features to EditEntry (Change Date, Rons' Coding).<br/>
</p>

<p><b>10th December 2004 (V1.0.5.10)</b><br/>
Fixed spelling mistake of "Occurrence".<br/>
Fixed year roll over bug which meant new years wouldnt be counted in "Archive".<br/>
</p>

<p><b>07th December 2004 (V1.0.5.10)</b><br/>
Updated PictureViewer.asp to allow thumbnails (something about my 22mb photo collection made me added this).<br/>
</p>

<p><b>04th December 2004 (V1.0.5.10)</b><br/>
Updated Mail.asp to not center the message text.<br/>
Updated Mail.asp to convert character returns to new lines.<br/>
Updated Mail.asp's grammer "your mail was sent" to "your message was sent".<br/>
Updated Unsubscribe.asp's InStr checking for my site.<br/>
Updated Comments.asp so it states the entry ID in the e-mail subject.<br/>
Updated mailing list so entries aren't centered.<br/>
</p>

<p><b>02nd December 2004 (V1.0.5.10)</b><br/>
Updated ViewCat.asp to ignore non-numerical page numbers after a failed exploitation (No security vulnerability, just errors).<br/>
</p>

<p><b>28th November 2004 (V1.0.5.10)</b><br/>
Added user submitted hacks to the ZIP.<br/>
Fixed PingOMatic support having my old site URL hardcoded into it.<br/>
Updated EditEntry.asp to replace category "%20"'s with eye friendly spaces.<br/>
</p>

<p><b>17th November 2004 (V1.0.5.09*)</b><br/>
Fixed "EditEntry.asp" to parse out the "&lt;/textarea>"'s.<br/>
Updated "EditEntry.asp" to not convert &amp;nbsp;'s to spaces.<br/>
</p>

<p><b>13th November 2004 (V1.0.5.09*)</b><br/>
Updated "SwimmingPool" theme so the comments are readable.<br/>
Updated RSS so the times are in 24 hour times (Pass Feed Validation).<br/>
</p>

<p><b>11th November 2004 (V1.0.5.09*)</b><br/>
Programmed "PictureViewer" (Beta).<br/>
</p>

<p><b>08th November 2004 (V1.0.5.09*)</b><br/>
Added a RSSReader import facility for AdminToolkit.<br/>
Updated WinblogX to use new site address.<br/>
Updated WinblogX to delete files before downloading new ones (Data corruption).<br/>
Updated site to work with the old "Check For Update" on WinBlogX (At Least I hope).<br/>
Updated WinBlogX to accept an empty blog folder (Seen as my site now doesnt use one).<br/>
</p>

<p><b>06th October 2004 (V1.0.5.09*)</b><br/>
Added a little "help" entry to the database to get people started.<br/>
Added the rest of the Themes at the expense of a larger ZIP file.<br/>
</p>

<p><b>21st October 2004 (V1.0.5.09*)</b><br/>
Fixed Mail support for Mdaemon.<br/>
Updated Mailing code to DIM error message.<br/>
Updated site to not issue e-mails if e-mail is disabled.<br/>
</p>

<p><b>11th October 2004 (V1.0.5.09)</b><br/>
Updated Readme.<br/>
</p>

<p><b>10th October 2004 (V1.0.5.09)</b><br/>
Added ThemeIT Editor
Added BlogXThemer (ThemeIT) Editor support in Header.asp.<br/>
Fixed HTML code in Entry titles messing up E-mail link. (RP4 discovered that)<br/>
</p>

<p><b>07th October 2004 (V1.0.5.09)</b><br/>
Added "On Error Resume Next" to OtherLinks &amp; Links (So WebHosts not supporting FSO work).<br/>
Fixed hacking protection script error (No Vulnerabilities)<br/>
Fixed RTF selected text new window linking glitch<br/>
Fixed Search where an entry category contains NULL.<br/>
Fixed Search to not display categories when categories are turned off.<br/>
Fixed trying to navigate the calendar when on the comments page.<br/>
Updated Mailing List page to actually say what that box was for. (Thanks Ben for that)<br/>
Updated Mailing List page to be more user friendly.<br/>
Updated all site link checking to ignore Matthew1471.co.uk.<br/>
Updated Results page to use the calander CSS.<br/>
</p>

<p><b>26th September 2004 (V1.0.5.08)</b><br/>
Fixed pingback having my website's old address hardcoded into it.<br/>
</p>

<p><b>29th August 2004 (V1.0.5.07)</b><br/>
Fixed rare bug where poll would fail to register a vote.<br/>
Fixed time display if times are set to 24hours on some pages.<br/>
Fixed CheckForUpdate's link.<br/>
Updated CheckForUpdate to verify link exists in future.<br/>
</p>

<p><b>28th August 2004 (V1.0.5.06)</b><br/>
Fixed links to new site layout.<br/>
</p>

<p><b>20th August 2004 (V1.0.5.06)</b><br/>
Fixed problem on IIS3/4 where Response.Redirect failed on Comments.asp.<br/>
</p>

<p><b>12th August 2004 (V1.0.5.05)</b><br/>
Fixed Edit Poll.<br/>
</p>


<p><b>09th August 2004 (V1.0.5.05)</b><br/>
Added List Plugin (Dan's idea).<br/>
Fixed mailing list not showing text on main BlogX domains.<br/>
Updated Image Upload Error to still display smileys even if the upload component is not installed.<br/>
Updated main.asp to display WeekDayName.<br/>
Updated plugin documentation to explain running multiple plugins.<br/>
Updated search to change the font on highlighted words (unreadable yellow highlight on some font colors).<br/>
Updated WhoUses.asp to explain things better.<br/>
Updated ZIP.<br/>
</p>

<p><b>28th July 2004 (V1.0.5.04)</b><br/>
Added several new stylesheets (but not to the ZIP, you'll have to manually steal them).<br/>
Updated RSS feeds to use new domain.<br/>
Updated Comments RSS feed to not confuse feed readers into remarking all comments as new.<br/> 
</p>

<p><b>27th July 2004 (V1.0.5.04)</b><br/>
Added several new stylesheets (but not to the ZIP, you'll have to manually steal them).<br/>
Fixed the link to the Admin EditDisclaimer page (Noticed it this morning).<br/>
</p>

<p><b>19th July 2004 (V1.0.5.04)</b><br/>
Fixed spell's ignore joining words.<br/>
Improved compatibility with FireFox (Poll results).<br/>
Validated &amp; fixed all stylesheets.<br/>
</p>

<p><b>18th July 2004 (V1.0.5.04)</b><br/>
Added ability to remove dictionary (without causing an ASP 500).<br/>
Fixed Config.asp nullifying records.<br/>
Updated Share.asp to use new domain and goto FreeWebs instead of PSC.<br/>
</p>

<p><b>17th July 2004 (V1.0.5.04)</b><br/>
Added "Note.gif" emoticon.<br/>
Fixed a non random PUK code on Comments. (Meaning people could unsubscribe others if they know their e-mail address).<br/>
Fixed link to a password protected entry from the comments page.<br/>
Hopefully fixed a very minor error when the same page is loaded twice at the EXACT same time (and logging is on).<br/>
</p>

<p><b>15th July 2004 (V1.0.5.03)</b><br/>
Fixed "EditEntry.asp" where Title containted quotes.<br/>
Fixed Spell Check were original word contained an appostrophe.<br/>
Fixed absolute path in PingBack.asp (C:\Inetpub\wwwroot\Blog\).<br/>
</p>

<p><b>13th July 2004 (V1.0.5.03)</b><br/>
Added ServerError.asp with Intelligent bug reporting. (You'll need to edit the variables in it)<br/>
Fixed Comments.asp when no Entry specified.<br/>
Fixed ViewCat.asp nullifying records (Causing an error for Calendar) GoogleBot reported that!<br/>
Fixed EditEntry.asp not closing the recordset (Not sure when I created that problem as it worked before).<br/>
Fixed Spell Check error when correction contains appostrophe.<br/>
Fixed Random Quote Plugin (Down to a 1 in 679 chance it would error) I cannot believe im hitting them all today!.<br/>
</p>

<p><b>12th July 2004 (V1.0.5.02)</b><br/>
Fixed RSS Feed validation (Thanks Joe for reporting that).<br/>
Fixed Spell Check failing to accept user corrections with words with symbols e.g. "BlogX,".<br/>
Updated a possible problem with Database collision (Possible I say).<br/>
Updated security to accept the cookie as a direct entry to Admin pages.<br/>
</p>

<p><b>11th July 2004 (V1.0.5.01)</b><br/>
Added "Winamp NowPlaying" plugin.<br/>
Added "Random Quotes" plugin.<br/>
Added my first user submitted template "Black" (Thanks to Kiz for that).<br/>
Fixed link to Poll Results if in an Admin page.<br/>
Fixed link on Default.asp that goes to "EditMainPage.asp".<br/>
</p>

<p><b>10th July 2004 (V1.0.5.01)</b><br/>
Changed Count.asp (Nothing important).<br/>
Fixed Comments.asp problem once and for all.<br/>
Fixed "Results.asp" not checking if user has voted.<br/>
Fixed typo in "CheckForUpdate.asp".<br/>
Updated Footer.asp to nullify "Records".<br/>
Updated all pages to not nullify and recreate "Records" (it "<i>might</i>" cause a problem says a Microsoft Article).<br/>
</p>

<p><b>09th July 2004 (V1.0.5.00)</b><br/>
Added words to ZIP's "UserDictionary".<br/>
Fixed replace.asp adding a comma to a link.<br/>
Fixed "UploadPicture.asp" causing a client side Javascript error when there's an appostrophe in the URL.<br/>
Updated WinBlogX installer to install dependencies (and provide a link in the readme to the Visual Basic ones).<br/>
</p>

<p><b>08th July 2004 (V1.0.5.00)</b><br/>
Fixed demo plugin that shows "Last 5 Entry Titles" showing 6 (Well done to those who can count ;)).<br/>
Modified "Credits.txt" to credit PoorMan's SpellCheck.<br/>
Removed other dictionaries (so zip file is smaller).<br/>
Updated About.asp to credit PMSC, mention beta testers, list new features etc.</p>

<p><b>07th July 2004 (V1.0.5.00)</b><br/>
Added "AllowEditLinks".<br/>
Added "CheckForUpdate".<br/>
Added Pingback viewer to the comments page.<br/>
Fixed "BlogIt" when SiteURL contains an appostrophe.<br/>
Mass problems with Comments.Asp (Naomi &amp; Sarah reported this)...think it's fixed.<br/>
Notified Mailing List.<br/>
<span style="color: red">Added Spell Check.</span><br/>
Updated Download.asp to not run through a server side component (Since the URL is no longer secret).<br/>
Updated Plugin.asp.<br/>
Updated pingback to check for the same IP address pinging back for the same entry.<br/>
Updated ZIP and Freewebs mirror.<br/>
</p>

<p><b>06th July 2004 (V1.0.5.00)</b><br/>
Worked on Spell Check.<br/>
</p>

<p><b>05th July 2004 (V1.0.5.00)</b><br/>
Added a link to SpellCheck in editor (Not yet finished).<br/>
Added "NoDate" for CMS.<br/>
Deleted PlanetSourceCode zip.<br/>
Fixed Count.asp.<br/>
Fixed selecting a day messing up the current day highlighting.<br/>
Researched WinBlogX RTF/HTML (RTF a waste of time).<br/>
Updated Config.asp to include "Dim.asp" and "Database.asp" (Now easier to update).<br/>
Updated WinBlogX (New "Check For Update" option).<br/>
</p>

<p><b>04th July 2004 (V1.0.5.00)</b><br/>
Added moderation to WhoUses.asp.<br/>
Added legacy option to Themes.asp.<br/>
Fixed Javascript error on AddEntry.asp, triggered when an Entry is saved and the page unloaded.<br/>
Updated LegacyMode functions.</p>

<p><b>03rd July 2004 (V1.0.5.00)</b><br/>
Fixed sloppy programming resulting in returning commenters not being able to comment (Unreleased Beta).<br/>
Fixed CommentNotify.asp having undefined variables.<br/>
Updated comments.asp to remove subscription checkbox for logged in users.</p>

<p><b>01st July 2004 (V1.0.5.00)</b><br/>
Fixed today's day being shown on another month.</p>

<p><b>30th June 2004 (V1.0.5.00)</b><br/>
Added the toolbar bookmarklet.</p>

<p><b>28th June 2004 (V1.0.5.00)</b><br/>
Added the ability to edit any entry.<br/>
Added the ability to delete any entry.<br/>
Fixed XMLTimeZones being converted by client RSS (Thanks Dan).<br/>
Updated Orange stylesheet to underline all links.</p>

<p><b>24rd June 2004 (V1.0.5.00)</b><br/>
Added "UseImagesInEditor" option in "Config.asp"<br/>
Added EditLinks.asp.<br/>
Added Links.txt file.<br/>
Updated text editor to set focus.<br/>
Updated text editor to allow highlighting text to link (Thanks Dan).<br/>
Updated text editor to ask if you want the link in a new window (Thanks Dan).<br/>
Updated calendar to show red (I remember now why I gave up with Bold, it was invisible) for Today (Thanks Dan).<br/>
Updated XMLRPC to new version.</p>

<p><b>23rd June 2004 (V1.0.4.00)</b><br/>
Added a new option, "CalendarCheck", which highlights only days with posts (Thanks Dan).<br/>
Added "TimeOffset" to offset the time, Surprising really (Thanks Dan).<br/>
Added "EditPoll.asp".<br/>
Fixed "Any HTTP header modifications must be made before writing page content" on RSS when using IIS 5.0 or less (Thanks Dan).<br/>
Fixed calendar on Admin pages.<br/>
Fixed Comments.asp dimension.<br/>
Fixed ASPUpload not found leaving database open.<br/>
Updated comments.asp to disallow new comments if they are disabled, In event of a comment spam attack (Thanks Tom).<br/>
Updated Links.asp to open links in a new window (Thanks Dan).<br/>
Looked at possible problems with CommentSpam (Thanks Tom).<br/>
Updated config.asp to better handle quotes.<br/>
Updated config.asp to list manual options only (Thanks Dan).<br/>
Updated "Includes/config.asp" to better explain each option (Thanks Dan).</p>

<p><b>22nd June 2004 (V1.0.4.00)</b><br/>
Added another poll results page.<br/>
Fixed comments RSS where the querystring was empty.<br/>
Fixed dimension of calender.<br/>
Fixed dimension of already dimensioned values in EditMainPage.asp.<br/>
Fixed image upload link to stylesheet and progressbar.<br/>
Fixed LegacyMode.<br/>
Fixed logout link.<br/>
Fixed problems with ReaderPassword.<br/>
Fixed someone else's XMLRPC script. (Now I know i'm a pro ;-) )<br/>
Fixed search when on any Admin pages.<br/>
Renamed "Blue.gif" to "Bar.gif" (for more colours).<br/>
Released code.<br/>
Tested new release.<br/>
Updated default.asp to clear buffer.<br/>
Updated Download.asp to fix script timeout.<br/>
Updated documentation.<br/>
Updated planet source code.</p>

<p><b>21th June 2004 (V1.0.4.00)</b><br/>
Added dimension of some variables.<br/>
Added option explicit to some pages.<br/>
<span style="color: red">Added poll.</span><br/>
Fixed double headers when no main page.<br/>
Fixed mailing list.<br/>
Fixed NotFound.asp's case sensitive domain check.<br/>
Fixed WAP not closing the database.<br/>
Removed case sensitivity in usernames and passwords.<br/>
Updated Comments.asp Spam protection (Just incase someone did some clever form spoofing)<br/>
Updated NotFound.asp's download link.</p>

<p><b>20th June 2004 (V1.0.4.00)</b><br/>
Added Ban lookup to pingback.<br/>
Added Password protected entries and modified all pages accordingly.<br/>
Added subscription to comments.<br/>
Fixed the link to disclaimer from Admin pages.<br/>
Fixed photo upload.<br/>
Tested pingback fully.<br/>
Fixed pingback fully.<br/>
Fixed URL linking again.<br/>
Updated comments.asp so "Email" reads as "E-Mail" (Thanks RadicalPuppy4).</p>

<p><b>19th June 2004 (V1.0.4.00)</b><br/>
Added confirmation to exiting the mail page without submitting the form.<br/>
Tested exploiting the XML HTTP to access local file. (No Security risk)<br/>
Tested URL linking fully.<br/>
Worked on URL linking again.</p>

<p><b>18th June 2004 (V1.0.4.00)</b><br/>
Added confirmation to exiting the mail page without submitting the form.<br/>
Added PingBack Client.<br/>
Fixed "main page" link on About.asp</p>

<p><b>17th June 2004 (V1.0.4.00)</b><br/>
Added Printer friendly page.<br/>
Worked on PingBack client.</p>

<p><b>12th June 2004 (V1.0.4.00)</b><br/>
Fixed URL linking, if link was hidden behind a "(".<br/>
Fixed URL linking, if link had a ")" after it.</p>

<p><b>06th June 2004 (V1.0.4.00)</b><br/>
Added "Archive" to nav.<br/>
Fixed page moving on specified dates.<br/>
Fixed case sensitive domain check.<br/>
Re-Branded BlogX to also use new domain "BlogX.co.uk".</p>

<p><b>05th June 2004 (V1.0.4.00)</b><br/>
Optimised code, Database only opened once.<br/>
Optimised code, closed Records reused.</p>

<p><b>04th June 2004 (V1.0.4.00)</b><br/>
Fixed search engine "don't complete words".<br/>
Fixed search engine occurance count on "any" mode.<br/>
Moved all Admin pages to a seperate folder (Thanks Tom for the suggestion).</p>

<p><b>03nd June 2004 (V1.0.4.00)</b><br/>
Added a link to e-mail on the search results (and that solves the mystery of the missing envelope).<br/>
Fixed problems with databases being left open on a few pages (sorry).</p>

<p><b>02nd June 2004 (V1.0.4.00)</b><br/>
Added confirmation to exiting the AddEntry page without submitting the form.<br/>
Fixed entries having a title longer than 80 characters.</p>

<p><b>30th May 2004 (V1.0.4.00)</b><br/>
Added "PingBack" for "Ping-O-Matic".<br/>
Tidyied up the "Config.asp".<br/>
Tested PingBack client (No Problems with "WordPress").<br/>
Updated RSS to use the "Entry=" querystring.</p>

<p><b>29th May 2004 (V1.0.4.00)</b><br/>
Added "credits.txt".<br/>
Added Error handling to image upload main page (forgot to add it to the file page).<br/>
<span style="color: red">Added "PingBack" table to database.</span><br/>
Added untested PingBack client (works internally).<br/>
Fixed category link on "Comments.asp".<br/>
Tidyied up the search engine (both apperance and code).<br/>
Updated "ViewItem.asp" to use the "Entry=" querystring.</p>

<p><b>28th May 2004 (V1.0.4.00)</b><br/>
Added an occurance count to search.<br/>
Added "complete words" to search.<br/>
Fixed search where an URL contained a term, breaking up the link.</p>

<p><b>27th May 2004 (V1.0.4.00)</b><br/>
Added Comments RSS.<br/>
Fixed Category RSS where an Un-Encoded category is empty.<br/>
Fixed Category RSS when an Un-Encoded field is empty.</p>

<p><b>25th May 2004 (V1.0.4.00)</b><br/>
Added a search (both "any order" and exact match).</p>

<p><b>22nd May 2004 (V1.0.4.00)</b><br/>
Added Category select to "Add Entry".<br/>
Added Comments.asp to auto add a "http://" if not already included.<br/>
Added a "legacy" mode to BlogX, so you can now disable all the features I added and "go classic".<br/>
Updated "Email.gif".<br/>
Updated WinBlogX to erase WinBlogX.ini if password/server/folder is wrong.<br/>
Updated WinBlogX to auto advance login if credentials already found.<br/>
Updated WinBlogX to auto save password by default.</p>

<p><b>24th April 2004 (V1.0.3.06)</b><br/>
Added Smiley mode in postings.<br/>
Fixed a few pages links to go to "Main.asp" instead of "Default.asp".<br/>
Fixed null when "0" is the count of "MailingListMembers" or "Refer".<br/>
Updated "Mail.asp" to align center &amp; use site stylesheet.<br/>
Updated "ViewItem.asp" to link to report EOF's to the webmasters.</p>

<p><b>22nd April 2004 (V1.0.3.05)</b><br/>
Added greying out of the "Prev Page" if it's already on the first.<br/>
Updated comments to hide e-mails from non-admins (What was I thinking before!).<br/>
Updated Cat RSS in Zip (been lazy in providing new version).</p>

<p><b>14th April 2004 (V1.0.3.05)</b><br/>
Added BlogX mirroring (Share.asp).<br/>
Fixed RSS on null categories/titles.<br/>
Updated "Orange" theme template.</p>

<p><b>13th April 2004 (V1.0.3.05)</b><br/>
Added BlogX to PlanetSourceCode again (Any publicity is good publicity).<br/>
Updated MailingList.asp &amp; About.asp to reflect this.</p>

<p><b>08th April 2004</b><br/>
Added pagecount to RSS, RSS now shows last 10 entries.<br/>
Fixed the documentation in the ZIP, it was showing up the "WinBlogX Readme" instead.<br/>
Updated the ZIP file again, replaced all files with my copys just incase I missed a few.<br/>
Updated the ZIP file's database to use the orange theme by default.<br/>
Updated the orange theme so work with the comments table.</p>

<p><b>23nd March 2004</b><br/>
Updated WinBlogX in the ZIP files "Bin" directory.</p>

<p><b>21st March 2004</b><br/>
Added ability to hide the OtherLinks from the Config.asp.<br/>
Added installer to WinBlogX (and finally finished).<br/>
Fixed RSS When SiteDescription had an "&amp;" or any other strange symbol.<br/>
Dropped the theme "Sky Blue" from the Zip, download is now 300kb less.<br/>
Dropped a forgotten debugging line of "Application.asp" which threw off the password checks for WinBlogX.<br/>
Updated WinBlogX to fix strange characters.<br/>
Updated WinBlogX to fully encode characters.<br/>
Updated WinBlogX's error handeling.<br/>
Updated License.txt.<br/>
Updated "Sea" template.</p>

<p><b>12th March 2004</b><br/>
Updated footer to link more than just "BlogX".<br/>
Updated a few themes.<br/>
Worked on an undercover script.</p>

<p><b>09th March 2004</b><br/>
Fixed comment notification linking with a wrong URL.<br/>
Updated comment notification to prevent notification when the admin comments.</p>

<p><b>05th March 2004</b><br/>
Updated WinBlogX to include new RTF functions.<br/>
Updated RSS to include optional picture, new information etc.</p>

<p><b>03rd March 2004</b><br/>
Added "Remember Me" to Comments.<br/>
Fixed RSS when title was empty (Feedreader looped a "NEW ENTRY", Sarah's doing ;-) ).</p>

<p><b>02nd March 2004</b><br/>
<span style="color: red">Added Comments Banning System.</span><br/>
Added Comments "You've already posted" System.<br/>
Added Comments Delete Function.</p>

<p><b>01st March 2004</b><br/>
Added Comments.<br/>
Updated "Referers" &amp; "Count.asp" to flag LAN addresses.</p>

<p><b>25th February 2004</b><br/>
Added WAP site.<br/>
Added E-Mail parser.<br/>
Added "SkyBlue" Theme.</p>

<p><b>24th February 2004</b><br/>
Added "Comments".<br/>
Added "Comments" to "Data" table to count comments on a post.<br/>
<span style="color: red">Dropped "CommentsURL" and replaced it with "EnableComments" in the DB.</span><br/>
Fixed page recordset to use file names (Cat is now passed in its own querystring).<br/>
Fixed "ViewCat.asp" to allow switching through pages.<br/>
Updated "MailingList.asp" to warn users of "testing" e-mail addresses.<br/>
Updated "MailingList.asp" to display helpful info for already subscribed users.<br/>
Updated all pages to show the new Comments link.<br/>
Updated "ViewCat.asp" to not link to "ViewCat.asp".</p>

<p><b>23rd February 2004</b><br/>
Fixed a few pages which defaulted to "Default.asp" on the "Go Back" buttons.<br/>
Fixed "Replace.asp".. and its strange bugs.</p>

<p><b>22nd February 2004</b><br/>
Added MainPage (EditMainPage.asp, Default.asp).<br/>
Added "RTF.js" to "EditDisclaimer.asp".<br/>
Fixed a minor glitch in a missing "&lt;/Span&gt;".<br/>
Fixed a minor glitch in "Edit Last Entry" appearing on further pages.<br/>
Updated content boxes to be bigger.<br/>
<span style="color: red">Updated database to include "EnableMainPage".</span><br/>
Updated upload picture to include "MainPage" handeling.<br/>
Updated "RTF.js" to include an if statement for querystrings.<br/>
Updated "Disclaimer.asp" to be held within a box.<br/>
Updated MailingList to block AOL. (Can't send mail to them)</p>

<p><b>15th February 2004</b><br/>
Added "NotFound.asp".<br/>
Added source for "Count.asp" (To settle any concerns).<br/>
Added "WhoUses.asp".<br/>
Added "WinBlogX.exe" to the source.<br/>
Fixed a minor glitch in image uploading.<br/>
Fixed "Default.asp" &amp; "ViewItem.asp" to not error if the Category was set as null.<br/>
Fixed "OtherLinks.txt" not working on a few pages where I used "count" before.<br/>
Fixed "MailingList" after a messup.<br/>
Fixed Images in RSS.<br/>
<span style="color: red">Updated Database to include "ScriptRefer".</span><br/>
Updated Documentation to include "Disallowed Parent Path" information.<br/>
Updated RSS.<br/>
Updated "MailingList" to convert line breaks into HTML line breaks.<br/>
Updated "Includes/Mail.asp" to send the from name and to name with a few components.</p>

<p><b>13th February 2004</b><br/>
Updated RSS.</p>

<p><b>12th February 2004</b><br/>
Added a download license.<br/>
Added mailing list Webmaster notification.<br/>
Crippled PlanetSource Code's Code to cut down on piracy. (Full version available AFTER accepting License).<br/>
Fixed URL replace to handle VbCrlf's better.<br/>
Updated About.asp so now source code can only be downloaded by subscribed users (Sick of people abusing the license).<br/>
Updated "MailingList" sending.</p>

<p><b>10th February 2004</b><br/>
Added Forum support.<br/>
Added &amp; Updated mailing list.<br/>
Added "Upload Picture" ability.<br/>
Fixed hyperlinking ONLY when there is a space or linebreak before a URL (Image URL problems).<br/>
Fixed RSS to convert "Images/Articles/" to the FULL URL.<br/>
Fixed image uploading for thoose whoose paths were different (Now uses MapPath for image paths).<br/>
Updated formatting tools.<br/>
Updated "EditLastEntry" to include formatting tools.</p>

<p><b>09th February 2004</b><br/>
Added "IncludeHTM.txt" (Gets included just before the footer).<br/>
Added "UploadPicture.asp".<br/>
Updated "AddEntry.asp" to include formatting tools.<br/>
Updated clarification of the license agreement a bit (after a few people have removed my ONE line copyright).</p>

<p><b>07th February 2004</b><br/>
Added/Updated "Themes".<br/>
Added "BlankTemplate.zip" (Information on how to make your own theme).<br/>
Updated "Themes.asp".</p>

<p><b>06th February 2004</b><br/>
Added "Themes".<br/>
Added "Themes.asp".<br/>
Added a theme preview querystring to "Header.asp".<br/>
Created "Themes" &amp; updated a few.<br/>
Fixed RSS as somehow setting nothing wasn't the same as nothing?!?. (IsNUll) is now checked.<br/>
Updated "AddEntry.asp" to use "MaxLength" attribute.<br/>
Updated WinBlogX to use "MaxLength" attribute.</p>

<p><b>05th February 2004</b><br/>
<span style="color: red">Added "ReaderPassword", Now You Can Restrict Who Reads The Blog (*Database, RSS, ViewerPassword*).</span><br/>
Fixed RSS to use Password attribute...Should a "ReaderPassword" be implemented.</p>

<p><b>04th February 2004</b><br/>
Updated WinBlogX as someone got mixed up with the Hexidecimal for "&amp;" and the ASCII code for it (*Simple mistake ;-) *).</p>

<p><b>03rd February 2004</b><br/>
Updated disclaimer, it conflicted with the "license.txt".</p>

<p><b>01st February 2004</b><br/>
Added a "Show/Hide" function to the "Recent Changes" (Hidden by default).<br/>
Added WinBlogX documentation, Source code, Zip...<br/>
Fixed RSS handling of "&amp;" after Elin caused a syntax error.<br/>
Fixed RSS handling of tags, after new RSS downloads had paragraphs displayed as tags due to "&amp;" converting further down the script.<br/>
Updated "About.asp" not to show the link to more information on the more information page.<br/>
Updated "About.asp" to hide/allow the ability to hide a few more options.<br/>
Updated documentation to have seperate files for additional information. (Smaller mainpage).<br/>
Updated WinBlogX to convert a few non Alphanumeric characters to more friendly ones. (Don't know the VB equivallant of ASP's Server.URLEncode() )</p>

<p><b>31th January 2004</b><br/>
Added a "Colour Picker" for background colour.<br/>
Added RSS for all (*Really Simple Syndication*).<br/>
Added RSS by category (*Really Simple Syndication*).<br/>
Added "All" To Categories.<br/>
Updated "About.asp" now we have RSS + more.<br/>
Updated "Logging.asp" to include more of the refering URL<br/>
<span style="color: red">Updated Database to allow a longer ReferURL</span><br/>
Updated Header to include "Auto RSS Discovery"
Updated Footer to no longer link to SimpleGeek.com.<br/>
Updated Documentation to include seperate parts (OtherLinks.txt, Updating, Config Definitions).</p>

<p><b>30th January 2004</b><br/>
Added ability to edit the last entry (Still believe that's as far as you should edit ;-) )<br/>
Added "Contact Me" option, EmailServerSettings and the support to use a range of ASP Mail components<br/>
Included "Application.asp" for WinBlogX.<br/>
Updated security (again) so that it knows after line 2 whether you are logged in or not. (this was not a vulnerability).<br/>
Updated a possible vulnerability in which ".Inc" files can be read as plain text if your webServer isn't setup to parse them (Sorry).<br/>
Updated all the includes to the world defined extension of ".asp" extension in case servers arn't configured to handle ".inc" files (See Above).</p>

<p><b>28th January 2004</b><br/>
Updated About Page.<br/>
Updated Pages To Include a "&lt;P&gt;" and updated the post page not to include them<br/>
(Now "No Data Entry" is easier to track..Database fields are a few characters shorter...etc etc).<br/>
Ready To Release WinBlogX 1.0! (No User Documentation Yet)</p>

<p><b>27th January 2004</b><br/>
Updated a HUGE vulnerability in which session state was being called BEFORE "CookieName" was being defined.<br/>
Updated SQL Exploit checking to allow category names with " ' " in them.<br/>
Updated "Matthew1471 Homepage" NOT to show.</p>

<p><b>26th January 2004</b><br/>
Distributed on the Internet.<br/>
Added A "Register" mode in which your site can get added to future versions of "OtherLinks.txt".</p>

<p><b>25th January 2004</b><br/>
Added documentation &amp; started distributing.<br/>
Deleted "ChangeLog.txt" ("About.asp" will replace it).<br/>
Researched RSS &amp; Blogging (I just might get into that).<br/>
Updated "Calendar.inc" to highlight current day. (Bug Fix)<br/>
Updated Windows client to make postings. (Still in Beta State).<br/>
Updated "About.asp".</p>

<p><b>24th January 2004</b><br/>
Added a "CheckSessionCookies" login check (Not Sure How Useful This Is).<br/>
Updated "Remember Login" to store a cookie (Nav.inc, ?ClearCookie).</p>

<p><b>23rd January 2004</b><br/>
Added "ShowCategories" option (Nav.Inc, AddEntry, ViewItem, Default.asp).<br/>
Fixed a problem where ReferURL contained an " ' " (Possible SQL Security Problem).<br/>
Fixed Broken timings of "011:23" because someone got their "<12" and "<10" muddled on the "add a 0 before the number code".<br/>
Allowed user to take advantage of "About.Asp" and Added an acronym for the version number over the "Powered By" link.<br/>
<span style="color: red">Updated database to know what's "Required" and what's not.<br/>
Updated copyright messages to "2004" and added a "License.Txt".</span><br/>
Updated "ChangePassword.Asp" to not show the existing password (It was confusing), also simplified the words more.</p>

<p><b>22nd January 2004</b><br/>
Updated clarity on what is compulsory in "Config.Asp" and what's not.<br/>
Updated CSS to have a darker blue header and a blue entry header.<br/>
Added option to hide "Add comment" should "CommentsURL" be blank.<br/>
Fixed option to change background color.</p>

<p><b>21st January 2004</b><br/>
Fixed problem where "Category" contained a space, then modified all pages to decode and encode it appropriately.<br/>
Added "OtherLinks.txt" and implemented format checking and size checking.</p>

* = A new version number was not required.<br/>
<span style="color: red">Red Text</span> = Database update, please update your database before updating BlogX.
</div>
</div>

</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
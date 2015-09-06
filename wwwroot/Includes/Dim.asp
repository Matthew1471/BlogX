<%
' --------------------------------------------------------------------------
'¦Introduction : Declaration File.                                          ¦
'¦Purpose      : Variables can only be declared once and server-side        ¦
'¦               files are typically included numerous times, this file is  ¦
'¦               run early and seperate from the include files to prevent   ¦
'¦               them from being ran twice.                                 ¦
'¦Ensures      : All variables are defined.                                 ¦
'¦Used By      : Most pages.                                                ¦
'---------------------------------------------------------------------------

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
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

'-- Used in Database.asp --'
Dim Records, Database

'-- Used as config variables --'
Dim AdminUsername, AdminPassword, CookieName, Copyright
Dim EnableEmail, EmailAddress, EmailServer, EmailComponent
Dim EnableComments, EntriesPerPage, Polls, ShowCategories, ShowMonth
Dim SiteName, SiteDescription, SiteSubTitle, SortByDay, BackgroundColor
Dim TimeFormat, Logging, EnableMainPage, ReaderPassword, Template

'-- Auto-generated config variables --'
Dim SpamAttempts, PageName, UTCTimeZoneOffset

'-- Hardcoded in Config.asp --'
Dim SiteURL, SharedFilesPath, Version
Dim DataFile, DataPassword
Dim AboutPage, AllowEditingLinks, ArgoSoftMailServer, CalendarCheck
Dim CommentNotify, NoEmailAddress
Dim LegacyMode, MailingList, NotifyPingOMatic, NoDate
Dim NoticeText, OtherLinks, Register, RSS, RSSImage
Dim UseImagesInEditor, UseExternalPlugin, SSLSupported, ServerTimeOffset

'-- Commonly used page variables --'
Dim AlertBack, PingbackPage, PhotoMode
Dim PluginTitle, PluginText
Dim File, FSO, TextStream, Count, Name, URL, FSODisabled
Dim SplitText, WordLoopCounter, TemplateURL
Dim PageTitle, PageTitleEntryRequest, CookiesDisabled, NoAdvertIP

'-- Mailer specific variables --'
Dim SendOk, gMDUser, mbDllLoaded, gMDMessageInfo

'-- Language related variables --'
Dim MultiLanguage

'-- Used to over-ride setting modified date --'
Dim DontSetModified, GeneralModifiedDate

'-- Login ban variables & manages SSL compatibility --'
Dim Blacklisted, sURL

'-- Calendar variables --'
Dim szYearMonth
Dim DataYear, DataMonth, DataDay, SpecificRequest
%>
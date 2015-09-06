<%
' --------------------------------------------------------------------------
'¦Introduction : Configuration Page.                                        ¦
'¦Purpose      : Allows blog administrator to configure BlogX features.     ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp,                            ¦
'¦               Includes/NAV.asp, Includes/Footer.asp.                     ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

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
AlertBack = True 
PageTitle = "Change Blog Configuration/Options"

'-- If your host does not support parent paths specify the full path here --'
Dim ServerPathToInstalledDirectory
If TemplateURL = "" Then ServerPathToInstalledDirectory = Server.MapPath("..\") Else ServerPathToInstalledDirectory = Server.MapPath("..\..\")
'ServerPathToInstalledDirectory = "C:\inetpub\wwwroot"
%>
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
 <% If Request.Form("Action") <> "Post" Then %>
 <script type="text/javascript">
  function Colors(url) { popupWin = window.open('SelectColor.asp?' + url,'new_page','width=400,height=450,scrollbars=yes') }
 </script>

 <form method="post" action="Config.asp" onsubmit="return setVar()">
 <p>
  <input name="Action" type="hidden" value="Post"/>
 </p>
 <p class="config">
  SiteName<span style="color:red">*</span> : <input name="SiteName" type="text" style="width:90%;" value="<%If SiteName <> "" Then Response.Write Replace(SiteName,"""","&quot;")%>" onchange="return setVarChange()" maxlength="200"/>
 </p>
 <p class="config">
  CookieName<span style="color:red">*</span> : <input name="CookieName" type="text" style="width:20%;" value="<%If CookieName <> "" Then Response.Write CookieName%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  Copyright<span style="color:red">*</span> : <input name="Copyright" type="text" style="width:30%;" value="<%If Copyright <> "" Then Response.Write Replace(Copyright,"""","&quot;")%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  Description : <input name="SiteDescription" type="text" style="width:30%;" value="<%If SiteDescription <> "" Then Response.Write Replace(SiteDescription,"""","&quot;")%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  Comments<span style="color:red">*</span> : <input name="EnableComments" type="checkbox" value="True" onchange="return setVarChange()" <%If EnableComments = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  EntriesPerPage<span style="color:red">*</span> : <input name="EntriesPerPage" type="text" style="width:5%;" value="<%=EntriesPerPage%>"/>
 </p>
 <p class="config">
  ReaderPassword : <input name="ReaderPassword" type="text" style="width:40%;" value="<%If ReaderPassword <> "" Then Response.Write Replace(ReaderPassword,"""","&quot;")%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  SiteSubTitle : <input name="SiteSubTitle" type="text" style="width:40%;" value="<%If SiteSubTitle <> "" Then Response.Write Replace(SiteSubTitle,"""","&quot;")%>" onchange="return setVarChange()" maxlength="50"/>
 </p>
 <p class="config">
  Polls<span style="color:red">*</span> : <input name="Polls" type="checkbox" value="True" onchange="return setVarChange()" <%If Polls = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  ShowCategories<span style="color:red">*</span> : <input name="ShowCategories" type="checkbox" value="True" onchange="return setVarChange()" <%If ShowCategories = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  SortByDay<span style="color:red">*</span> : <input name="SortByDay" type="checkbox" value="True" onchange="return setVarChange()" <%If SortByDay = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  <a href="../Themes.asp">Theme</a> Template<span style="color:red">*</span> : 
  <%
   On Error Resume Next
   Set FSO = Server.CreateObject("Scripting.FileSystemObject")

   If Err = 0 Then
    Response.Write "<select name=""Template"" onchange=""return setVarChange()"">"

    Dim Folder, Folders

    If TemplateURL = "" Then 
     Set Folder = FSO.GetFolder(ServerPathToInstalledDirectory  & "\Templates\")
    Else
     Set Folder = FSO.GetFolder(ServerPathToInstalledDirectory  & "\Includes\Templates\")
    End If

    Set Folders = Folder.SubFolders

    For Each Folder in Folders 
     Response.Write "<option"
     If Template = Folder.Name Then Response.Write " selected=""selected"""
     Response.Write ">" & Folder.Name & "</option>"
    Next

    Response.Write "</select>"

    Set Folders = Nothing
    Set Folder = Nothing
   Else
    Response.Write "<input name=""Template"" type=""text"" style=""width:40%;"" value=""" & Replace(Template,"""","&quot;") & """ onchange=""return setVarChange()"" maxlength=""50""/></p>"
   End If

   On Error GoTo 0

  Set FSO = Nothing
  %>
 </p>
 <p class="config">
  BackgroundColor<span style="color:red">*</span> : <input name="BackgroundColor" type="text" style="width:20%;" value="<%=BackgroundColor%>" maxlength="20"/> <a href="JavaScript:Colors('Box=BackgroundColor')"><img alt="Color icon" src="../Images/Color.gif" style="border:none"/></a>
 </p>
 <p class="config">
  12Hour Times<span style="color:red">*</span> : <input name="TimeFormat" type="checkbox" value="True" onchange="return setVarChange()" <%If TimeFormat = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p class="config">
  Logging<span style="color:red">*</span> : <input name="Logging" type="checkbox" value="True" <%If Logging = True Then Response.Write "checked=""checked"""%>/>
 </p>
 <p>
  <input type="submit" value="Save"/>
 </p>
 <p class="config" style="text-align:Center">
  <span style="color:red">*</span> - Indicates a required field.
 </p>
 <p class="config" style="text-align:center">
  <b>Note</b>: To change CommentNotification, PingOMatic, ArgosoftMail support, CalendarCheck, MailingList, OtherLinks, Register, RSS, RSSImage, UseImagesInEditor and the TimeOffset...Edit Includes/Config.asp using notepad
 </p>
 </form>
 <% Else

 '-- Needed for checkboxes --'
 If Request.Form("EnableComments") = "" Then EnableComments = False Else EnableComments = True
 If Request.Form("Logging")        = "" Then Logging = False Else Logging = True
 If Request.Form("Polls")          = "" Then Polls = False Else Polls = True
 If Request.Form("ShowCategories") = "" Then ShowCategories = False Else Showcategories = True
 If Request.Form("SortByDay")      = "" Then SortByDay = False Else SortByDay = True
 If Request.Form("TimeFormat")     = "" Then TimeFormat = False Else TimeFormat = True

 If IsNumeric(Request.Form("EntriesPerPage")) Then EntriesPerPage = Request.Form("EntriesPerPage")

 '-- Make Changes To Config --'
 Records.CursorType = 2
 Records.LockType = 3
 Records.Open "SELECT ShortTimeFormat, SiteName, CookieName, Copyright, EntriesPerPage, EnableComments, ReaderPassword, Polls, ShowCategories, SiteDescription, SiteSubTitle, Template, SortByDay, BackgroundColor, Logging FROM Config", Database
  Records("SiteName") = Request.Form("SiteName")
  Records("CookieName") = Request.Form("CookieName")
  Records("Copyright") = Request.Form("Copyright")
  Records("EntriesPerPage") = EntriesPerPage
  Records("EnableComments") = EnableComments
  Records("ReaderPassword") = Request.Form("ReaderPassword")
  Records("Polls") = Polls
  Records("ShowCategories") = ShowCategories
  Records("SiteDescription") = Request.Form("SiteDescription")
  Records("SiteSubTitle") = Request.Form("SiteSubTitle")
  If Request.Form("Template") <> "" Then Records("Template") = Request.Form("Template")
  Records("SortByDay") = SortByDay
  Records("BackgroundColor") = Request.Form("BackgroundColor")
  Records("ShortTimeFormat") = TimeFormat
  Records("Logging") = Logging
  Records.Update
 Records.Close

 Response.Write "<p style=""text-align:center"">Config update successful.</p>"
 Response.Write "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"

 End If %>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
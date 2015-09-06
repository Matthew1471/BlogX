<%
' --------------------------------------------------------------------------
'¦Introduction : Viewer Password Check.                                     ¦
'¦Purpose      : Checks and handles viewer password authentication if set.  ¦
'¦Requires     : ReaderLogin.asp.                                           ¦
'¦Used By      : Comments.asp, Default.asp, Disclaimer.asp, Download.asp,   ¦
'¦               Downloads.asp, MailingList.asp, Main.asp, Newspaper.asp,   ¦
'¦               NotFound.asp
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

If (ReaderPassword <> "") Then

 Dim Page, Last, Length
 Page = Request.Servervariables("Script_Name")

 Last = InStrRev(Page,"/")
 Length = Len(Page)

 Page = Right(Page,Length - Last)

 If Session("Reader") = False or IsNull(Session("Reader")) = True AND (Request.Cookies("Reader") <> "True") Then
  Database.Close
  Set Records  = Nothing
  Set Database = Nothing
  Response.Redirect "ReaderLogin.asp?" & Page
 End If

End If
%>
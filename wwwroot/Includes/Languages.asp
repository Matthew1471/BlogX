<%
' --------------------------------------------------------------------------
'¦Introduction : Language Selector (Not yet used).                          ¦
'¦Purpose      : Loads the relevant language file variables in.             ¦
'¦Requires     : Languages/English.asp.                                     ¦
'¦Used By      : Includes/Header.asp.                                       ¦
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

Dim TheLanguages(2, 5)

'0 - Name
'1 - Flag

TheLanguages(0,0) = "English"
TheLanguages(1,0) = "UK.png"

Dim CurrentLanguage
Select Case CurrentLanguage
  Case 1
    <!-- #INCLUDE FILE="Languages\English.asp" -->
  Case Else
    <!-- #INCLUDE FILE="Languages\English.asp" -->
End Select
%> 
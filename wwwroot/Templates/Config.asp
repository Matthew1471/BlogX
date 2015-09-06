<%
' --------------------------------------------------------------------------
'¦Introduction : Template Background Color Include File.                    ¦
'¦Purpose      : Provides browsers that do not support CSS or cannot find   ¦
'¦               our stylesheets a background.                              ¦
'¦Used By      : Includes/Header.asp, Admin/Pingback.asp,                   ¦
'¦               Admin/AddFile_Save.asp, Admin/UploadPicture_Save.asp,      ¦
'¦               Admin/Toolbar.asp.                                         ¦
'¦Requires     : Nothing.                                                   ¦
'¦Standards    : N\A.                                                       ¦
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

Dim CalendarBackground

Select Case Template
 Case "Black"
  CalendarBackground = "fuchsia"
 Case "Busted"
  CalendarBackground = "#FF8C00"
 Case "Default"
  CalendarBackground = "Silver"
 Case "Clouds"
  CalendarBackground = "DodgerBlue"
 Case "Matthew1471"
  CalendarBackground = "#003399"
 Case "Matrix"
  CalendarBackground = "#003300"
 Case "Leaves"
  CalendarBackground = "#008000"
 Case "LightSea"
  CalendarBackground = "darkblue"
 Case "LighterBlue"
  CalendarBackground = "#003366"
 Case "Lotus"
  CalendarBackground = "#8B0000"
 Case "Pebbles"
  CalendarBackground = "RoyalBlue"
 Case "Orange"
  CalendarBackground = "#ff6600"
 Case "Palm Tree"
  CalendarBackground = "#5CACEE"
 Case "Purple"
  CalendarBackground = "#9933FF"
 Case "Puzzle"
  CalendarBackground = "#5CACEE"
 Case "Red"
  CalendarBackground = "DarkOrange"
 Case "Sandy"
  CalendarBackground = "#FFCC66"
 Case "Snails"
  CalendarBackground = "#008000"
 Case "Stary"
  CalendarBackground = "RoyalBlue"
 Case "Sea"
  CalendarBackground = "DarkBlue"
 Case "Swimming Pool"
  CalendarBackground = "#5CACEE"
 Case "TotallyGreen"
  CalendarBackground = "#363"
 Case "WaterFall"
  CalendarBackground = "#536e55"
 Case Else
  CalendarBackground = ""
End select
%> 
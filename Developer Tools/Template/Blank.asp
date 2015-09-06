<%
' --------------------------------------------------------------------------
'¦Introduction : Blank Page                                                 ¦
'¦Purpose      : Useful as a page template to create new pages              ¦
'¦Used By      : Nothing                                                    ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
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
%>

<!-- #INCLUDE FILE="Includes/Header.asp" -->
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

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
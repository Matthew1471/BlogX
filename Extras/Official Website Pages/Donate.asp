<%
' --------------------------------------------------------------------------
'¦Introduction : Donation Page.                                             ¦
'¦Purpose      : Provides the user with the option to donate to the project.¦
'¦Used By      : Links table.                                               ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp,¦
'¦               Includes/Cache.asp.                                        ¦
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

PageTitle = "Donation Information"

'-- Proxy Handler --'
CacheHandle(CDate("21/01/" & Year(Now())))
%>
<!-- #INCLUDE FILE="../Includes/Cache.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<div id="content">
 <div class="entry">
  <h3 class="entryTitle" style="text-align:center">Donate To BlogX</h3>

  <p style="text-align:center">Pleased with your new BlogX installation?</p>
  <p style="text-align:center">Why not after <%=DateDiff("yyyy",#21/01/2004#,Now())%> years of hobbyist student development donate and encourage further development?</p>

  <form action="https://www.paypal.com/cgi-bin/webscr" method="post">
   <p style="text-align:center">
    <input type="hidden" name="cmd" value="_s-xclick"/>
    <input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but04.gif" style="border:none" name="submit" alt="Make payments with PayPal - it's fast, free and secure!"/>
    <input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----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-----END PKCS7-----"/>
   </p>
  </form>

  <p style="text-align:center">All donations truly are appreciated no matter how small.</p>

 </div>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
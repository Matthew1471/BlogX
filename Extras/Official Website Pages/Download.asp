<% 
' --------------------------------------------------------------------------
'¦Introduction : BlogX Download Page.                                       ¦
'¦Purpose      : Allows visitors to my host to download BlogX               ¦
'¦Used By      : About.asp on my server only.                               ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦               Includes/Cache.asp, Includes/ViewerPass.asp.               ¦
'¦Notes        : This page is for downloading BlogX off BlogX.co.uk, but is ¦
'¦               not relevant for other hosts.                              ¦
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
 CacheHandle(CDate("27/07/08 22:53:00"))

PageTitle = "Download BlogX"

If Request.Form("License") <> "Accept" Then%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<!-- #INCLUDE FILE="Includes/Cache.asp" -->
<div id="content">

 <div class="entry">
  <h3 class="entryTitle">Download BlogX</h3><br/>
  <div class="entryBody" style="text-align:center">
   If you have ever used or like this product,<br/>
   consider donating a few dollars to the BlogX development fund :<br/>
   
   <form action="https://www.paypal.com/cgi-bin/webscr" method="post">
    <p><input type="hidden" name="cmd" value="_s-xclick"/>
    <input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but04.gif" name="submit" alt="Make payments with PayPal - it's fast, free and secure!" style="border-style:none"/>
    <input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----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-----END PKCS7-----
"/></p>
   </form>

<br/>

<form method="post" action="Download.asp">
 <table border="0" cellpadding="0" cellspacing="0" style="align:center;margin: 0 auto; width:90%">
 <tr>
  <td align="center" style="background-color:#C0C0C0"><b>Matthew1471 BlogX License Agreement</b></td>
 </tr>
 <tr>
  <td align="center" style="background-color:#F8F8F8"><br/> 
  <textarea name="message" rows="20" class="TextBox" style="width: 90%;" cols="74">Matthew1471 BlogX License Contract
------------------------------

License Last modified: February 12th 2004.
Copyright 2008 Matthew Roberts All Rights Reserved.

This license grants you the right to install, view and run one or multiple instance of this for private, public, non-commercial use.

You may modify source code (at your own risk), but the software (altered or otherwise) may not be distributed to entities beyond the license holder without the explicit written permission of the copyright holder.

You may use parts of this program including source code and images in your own private work, but you may NOT redistribute, repackage, sublicense, copy, 
or sell the whole or any part of this program even if it is modified or reverse engineered in whole or in part without the explicit written permission of the copyright holder. 

Copyright notices must be transferred to any file that any source code from this program in whole or in part used within.

It is forbidden to repackage in full or any part of this software including source code and images as your own work and redistribute for profit or for free without the explicit written permission of the copyright holder.

You may not pass off the whole or any part of this software as your own work.

Any attempt otherwise to copy, modify, sublicense, reverse engineer or distribute any part of this work is void, and will automatically terminate your rights under this License.

You may not deactivate, hide, change, remove, all or any hyper links or links to Matthew1471 and the powered by logo's, text, or images. All must remain visible when the pages are viewed unless you first obtain explicit written permission from Matthew1471.

Permission can be obtained from Matthew1471 to remove the powered bylogo's, text or, images by making a donation to Matthew1471 to help support the development and update of this and future software.

You will keep all copyright notices intact, including the notices embedded in all parts of the software.

The license holder is the legal owner of the magnetic medium on whichthe software is recorded. However, the software itself is not sold : the copyright holder retains the property rights on the software itself,whichever medium is used to store it upon and at any moment in time : original disks, plus any and all copies of the same which may have been produced subsequently.

Material form Matthew1471 including programs, applications, tutorials,scripts, may not be used for anything that would represent or is associated with an Intellectual Property Violation, including, but not limited to, engagingin any activity that infringes or misappropriates the intellectual property rights of others, including copyrights, trademarks, service marks, trade secrets, software piracy, and patents held by individuals,corporations, or other entities. 

You are not required to accept this License, since you have not signed it. 
However, nothing else grants you permission to download, install, run, use, view, modify, the Program or its derivative works.  These actions are prohibited by law if you do not accept this License. Therefore, by downloading,installing, using, viewing, modifying the Program (or any work based on the Program), you indicate your acceptance of this License to do so, and all its terms and conditions for using, copying, or modifying the Program or works based on it.

Any disagreement which may arise between the parties with regard to theinterpretation and/or execution of the present contract shall by default come under the jurisdiction of the United Kingdom Courts.

NO WARRANTY
-----------
THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY
APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE
COPYRIGHT HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS"
WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF 
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE 
RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU. 
SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL 
NECESSARY SERVICING, REPAIR OR CORRECTION.

IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN
WRITING WILL ANY COPYRIGHT HOLDER, BE LIABLE TO YOU FOR DAMAGES, 
INCLUDING ANY GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL
DAMAGES ARISING OUT OF THE USE OR INABILITY TO USE THE PROGRAM 
(INCLUDING BUT NOT LIMITED TO LOSS OF DATA OR DATA BEING RENDERED 
INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD PARTIES OR A 
FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS),
EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE 
POSSIBILITY OF SUCH DAMAGES.

If any terms are violated, Matthew1471 reserves the right to revoke 
the license at any time. No refunds will be granted for revoked licenses.

Matthew1471 reserves the right to modify these terms at any time 
without prior notice. 

This license is governed by the laws of the United Kingdom.

For correspondence : http://matthew1471.co.uk/Contact.asp
</textarea>
  <br/>
  </td>
 </tr>
 <tr>
  <td align="center" style="background-color:#C0C0C0">By clicking or pressing the &quot;ACCEPT&quot; button,<br/>
  &nbsp;You agree to be bound by the terms and conditions of this license.</td>
 </tr>
 <tr>
  <td align="center" style="background-color:#CCCCCC">
   <input type="button" name="Button" value="Cancel" class="button" onclick="window.open('default.asp', '_top')"/>
   <input type="submit" name="License" value="Accept" class="button"/> 
  </td>
 </tr>
</table>
</form>

</div>
</div>
</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<% 
Else
 Response.Redirect "Download/BlogX.zip"
End If
%>
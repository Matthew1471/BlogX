<%
OPTION EXPLICIT
Server.ScriptTimeout = 6000

If Request.Form("License") <> "Accept" Then%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<DIV id=content>

<DIV class=entry>
<h3 class=entryTitle>Download BlogX</h3><br>
<DIV class=entryBody>

<form method="post" name="frmLicense" action="Download.asp">
<table border="0" cellpadding="0" cellspacing="0" width="657" align="center">
<tr>
<td align="center" bgcolor="#C0C0C0" width="655"><b>Matthew1471 BlogX License Agreement</b></td></tr>
<tr>
<td width="655" bgcolor="#F8F8F8" align="center"><br>
<textarea name="message" rows="20" class="TextBox" style="width: 654; height: 402" cols="74">Matthew1471 BlogX License Contract
------------------------------

License Last modified: February 12st 2004 
Copyright 2004 Matthew Roberts All Rights Reserved.

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
<br>
</td>
</tr>
<tr>
<td align="center" bgcolor="#C0C0C0" width="655">By clicking or pressing the &quot;ACCEPT&quot; button,<br>
&nbsp;You agree to be bound by the terms and conditions of this license.</td>
</tr>
<tr>
<td align="center" bgcolor="#CCCCCC" width="655">
<input type="button" name="Button" value="Cancel" class="button" onClick="window.open('default.asp', '_top')">
<input type="submit" name="License" value="Accept" class="button"> 
</td>
</tr>
</table>
</P>
</Div>
</Div>
</Div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<% 
Else

Response.Redirect "Download/WebBlogX.zip"

End If
%>
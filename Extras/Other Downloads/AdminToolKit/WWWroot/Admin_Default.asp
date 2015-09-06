<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Blog Administration</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
<h2 align="center">Welcome To The BlogX Administrative Toolkit</h2>

<form name="Login" method="post" action="Admin_CheckUser.asp">

<table width="273" border="0" align="center" cellspacing="0" cellpadding="0" bgcolor="#0000FF">
<tr><td><font color="#FFFFFF"><b><img src="Includes/Images/icon_key.gif">Login To Your Account!</b></td></font>
</tr>
</table>

  <table width="273" border="0" align="center" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
    <tr>
      <td align="right" height="47" valign="bottom" width="94">UserName: </td>
      <td valign="bottom" width="172"><input type="text" name="UserName"></td>
    </tr>
    <tr>
      <td align="right" width="94">Password: </td>
      <td width="172"><input type="password" name="Password"></td>
    </tr>
    <tr> 
      <td align="right" width="94">&nbsp;</td>
      <td height="44" width="172"> 
        <input type="submit" name="Submit" value="Enter">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Reset" value="Reset">
      </td>
    </tr>
  </table>
</form>

<p align="Center"><font color="Red">Note :</font> An Attempted Login From An Un-Authorised Party, Could Result In Legal Prosecution.</p>

<p></p>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
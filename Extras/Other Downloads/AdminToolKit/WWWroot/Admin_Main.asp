<!-- #INCLUDE FILE="Includes/Admin.asp" -->
<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Blog Administration</font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">

<h2 align="center">Welcome To The BlogX Administrative Toolkit</h2>

<p align="center">
O <a href="Admin_Create.asp">Create/Enable/Update A Blog</a><br>
O <a href="Admin_Refresh.asp">Update <b>ALL</b> Active Blogs</a><br>
O <a href="Admin_Disable.asp">Disable A Blog</a><br>
O <a href="Ranks.asp">Top 10</a><br>
O <a href="Admin_News.asp">Update News</a><br><br>

O <a href="Admin_LogOut.asp">Log Out</a>
</p>

<p align="center"><font color="#FF0000">Note : </font>
This will write to <font color="Red"><%=Root%></font> & <font color="Red"><%=DatabasePath%></font> make sure IIS can write to that folder (and that it exists)</p>

      </td>
      <!--- End Of Content -->
<% WriteFooter %>
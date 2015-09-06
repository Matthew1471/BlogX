<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Blogs Available To Read On <%=Domain%></font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">
      <font color="#000444">
      <b>What is a Blog?</b><br>
      <br>
        A Blog is an online diary type thing of
        whatever the author chooses, you can read it and comment on it, share it and if
        you really want, you can even winge about someone else's on your own blog ;-).<br>

      <br><br>
      <b>What Blogs Are Stored On This Server?</b>
      <br><br>
  <table border="0" cellpadding="0" cellspacing="5">
   <tr>
<%
        Dim FSO, Folder, Folders, Count

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(AppPath)
        Set Folders = Folder.SubFolders

        Count = 0

        For Each Folder in Folders

        If Folder.Name <> "Includes" Then

        If Count => 8 Then Response.Write "   </tr>" & VbCrlf

        If Count => 8 Then Response.Write "   <tr>" & VbCrlf
        Response.Write "    <td><img src=""Includes/Images/eBlog.gif""><a href=""" & Folder.Name & "/"">" & Folder.Name & "</a><br></td>" & VbCrlf

        If Count => 8 Then Count = 0
        Count = Count + 1

        End If

        Next

        Set FSO = Nothing
	Set Folder = Nothing
	Set Folders = Nothing

        Response.Write "   </tr>"
%>
</table>

      <br><br>
      <b>Does The Owner Have A Blog?</b>
      <br><br>
      <% If Len(MyBlog) <> 0 Then %>
      Naturally, If you wish to see my rantings, visit <a title="Owners Blog" href="<%=MyBlog%>">my blog</a>.
      <% Else %>
      No, I don't think the owner does
      <% End If %>

      <br><br>
      <b>How do I get my own blog on <%=Domain%>?</b>

      <br><br>
      <% If EnableGuestSignups <> 0 Then %>
      Just <a title="Free Blog Signup" href="http://<%=Domain%><%=Root%>Signup.asp">click here</a> to signup for your free blog. It's that simple!
      <% Else %>
      Guest registration for <u><%=Domain%></u> is currently <b>Off</b>.<br>
      You can always <a href="Contact.asp">contact the owner</a> and ask what's up with that!
      <% End If %>

</font>
      </td>
      <!--- End Of Content -->
<% WriteFooter %>
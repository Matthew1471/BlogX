<DIV class=sidebar id=rightBar>
<BR>

<% If (ShowMonth <> False) AND (LegacyMode <> True) Then %>
<!--- Archive --->
<DIV class=section>
<H3>Archive</H3>
<UL>No Last Visits</UL>
<BR></DIV><BR>
<%End If%>

<!--- Links --->
<DIV class=section>
<H3>Links <%If (Session(CookieName) = True) Then Response.Write " <acronym title=""Edit Your Links""><a href=""" & SiteURL & "Admin/EditLinks.asp""><Img Border=""0"" Src=""" & SiteURL & "Images/Edit.gif""></a></acronym>"%></H3>
<UL>
  <LI><A href="<%=BlogURL%>">Blog Home</A></LI>
  <LI><A href="<%=SiteURL%>">Proxy Home</A></LI>
  <!-- #INCLUDE FILE="Links.asp" -->
</UL></DIV><BR>

<%
If (UseExternalPlugin = 1) AND (LegacyMode = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!--- <%=PluginTitle%> --->
<DIV class=section>
<H3><%=PluginTitle%></H3>
<UL><%=PluginText%></UL>
</DIV><BR>

<% End If %>

<!--- Login As A Publisher --->
<% If Session(CookieName) = True Then %>
<DIV class="section">
<H3><img border="0" src="<%=SiteURL%>Images/Key.gif">Admin</H3>
    <ul>
        <li><A href="<%=SiteURL%>Admin/AddEntry.asp">Add Entry</A></li>
        <li><A href="<%=SiteURL%>Admin/Log.asp">Log File</A></li>
        <li><a href="<%=SiteURL & PageName %>?ClearCookie">Logout</a></li>
    </ul>
</DIV><BR>
<%Else
Session(CookieName) = False
If Session("CookieTest") = "AOK" Then %>
<Form Name="Login" Method="Post" Action="Default.asp">
<DIV class="section" id=login>   
<H3><img border="0" src="<%=SiteURL%>Images/Key.gif">Admin Sign In</H3>
<P>Username: <INPUT name="username"></P>
<P>Password: <INPUT name="password" type="password"></P>
<P><INPUT name="Remember" type="checkbox" Value="True">Remember Login</P>
<P><INPUT name="SignIn" type="Submit" value="Sign In"></P></DIV><BR></FORM>
<%Else%>
<DIV class="section" id=login>   
<H3>Login Error</H3>
<p>Please <b>Enable</b> Cookies to login</p>
</Div>
<%
End If
End If %>
</DIV>
<!-- #INCLUDE FILE="Config.asp" -->
<%
Dim Domain
Domain = Request.ServerVariables("HTTP_Host")
Domain = Replace(Domain,"www.","")
Domain = UCase(Left(Domain,1)) & Right(Domain,Len(Domain)-1)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<Html>
<Head>
<Title>Blogs Stored On <%=Domain%></Title>
<meta http-equiv="Content-Language" content="en-gb">
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Blog Directory Listing For Blogs Hosted By <%=Domain%>" name="Description">
</Head>

<Body bgcolor="#6699FF">

  <table border="1" cellspacing="0" width="100%" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000" align="center">
    <tr>
      <td bgcolor="#FFCC00" width="70%" height="24" style="PADDING-LEFT: 5px;">
          <font face="Verdana" size="2">
          <b>
          <a title="Blog Directory Listing" href="<%=Root%>Default.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Home</a> 
          ¦ <a title="Free Blog Documents" href="<%=Root%>Documentation.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Documentation</a>
          ¦ <a title="Free Blog Downloads" href="<%=Root%>Downloads.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Downloads</a>
          ¦ <a title="Read A Random Blog" href="<%=Root%>Random.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Random</a>
          <% If EnableGuestSignUps <> 0 Then Response.Write "¦ <a title=""Free Blog Signup"" href=""" & Root & "Signup.asp"" style=""COLOR: #163aa6; TEXT-DECORATION: none"">Signup</a>"%>
          </b>
          </font>
      </td>
      <td bgcolor="#FFCC00" width="*" height="24" align="right" style="PADDING-RIGHT: 5px;">
      <font face="Verdana" size="2">
      <b><a href="<%=Root%>Contact.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Contact</a> 
      ¦ <a href="<%=Root%>Support.asp" style="COLOR: #163aa6; TEXT-DECORATION: none">Support</a></b>
      </font>
      </td>
    </tr>
    <tr>
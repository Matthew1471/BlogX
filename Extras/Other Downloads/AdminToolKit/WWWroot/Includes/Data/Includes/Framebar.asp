<%@EnableSessionState=False%>
<% Response.Expires = -1 %>
<HTML>
<HEAD>
<TITLE>Uploading files</TITLE>
<style type='text/css'>td {font-family:arial; font-size: 9pt }</style>
</HEAD>

<% If Request("b") = "IE" Then %> <!-- Internet Explorer -->
<!-- #INCLUDE FILE="Config.asp" -->
<BODY bgColor="<%=BackgroundColor%>" text="midnightblue" link="darkblue" aLink=red vLink="red">
<IFRAME src="bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" title="Upload Progress" noresize scrolling=no
frameborder=0 framespacing=10 width=369 height=65></IFRAME>
<TABLE BORDER="0" WIDTH="100%" cellpadding="2" cellspacing="0">
  <TR><TD ALIGN="center">
     To cancel uploading, press your browser's <B>STOP</B> button.
  </TD></TR>
</TABLE>
</BODY>

<%Else%> <!-- Netscape Navigator etc ... -->

<FRAMESET ROWS="65%, 35%" COLS="100%" border="0" framespacing="0" frameborder="NO">
<FRAME SRC="bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" noresize scrolling="NO" frameborder="NO" name="sp_body">
<FRAME SRC="note.htm" noresize scrolling="NO" frameborder="NO" name="sp_note">
</FRAMESET>

<%End If%>

</HTML>

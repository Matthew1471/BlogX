<!-- #INCLUDE FILE="Config.asp" -->
<% Session("CookieTest") = "AOK" %>
<!-- #INCLUDE FILE="Security.asp" -->
<%

'-- Remove Security, We Don't Want It --
'Session(CookieName) = True
'Response.Cookies(CookieName) = "True"
'Response.Cookies(CookieName).Expires = "July 31, 2008"
'------------------------------------------------------

If Request.Querystring = "ClearCookie" Then 
Session(CookieName) = False
If Request.Cookies(CookieName) = "True" Then Response.Cookies(CookieName) = ""
End If
%>
<HTML>
<HEAD>
<TITLE><%=SiteDescription%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
// 
// Usage Of This Software Is Subject To The Terms Of The License
//= - - - - - - -
-->
<SCRIPT LANGUAGE="Javascript">
<!-- begin hiding

<% If Request.Querystring("Editor") <> "True" Then%>
  if (parent.frames.length)
  top.location.href= document.location;
<% End If %>
<% 
If (AlertBack = True) AND (Request.Form("Action") <> "Post") Then %>
  var bolIsSubmitted = true;

  function window_onbeforeunload() {
                 if(bolIsSubmitted) {return true;}
         else {event.returnValue="You've modified a textbox or checkbox but haven't saved your changes!";}
     }

function setVar(){
   bolIsSubmitted = true;
   return true;
}

function setVarChange(){
   alert
   bolIsSubmitted = false;
   return true;
}
<% End If%>

// done hiding -->
</SCRIPT>
<%
If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<!-- #INCLUDE FILE="../Templates/Config.asp" -->
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
</HEAD>
<BODY bgColor="<%=BackgroundColor%>" <%If (AlertBack = True) AND (Request.Form("Action") <> "Post") Then Response.Write "onBeforeUnload=""window_onbeforeunload()"""%>>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<DIV id=header>
<H1 id=title><A style="TEXT-DECORATION: none" href="<%=SiteURL%>"><%=SiteName%></A></H1>
<P id=byline><%=SiteDescription%></P>
<P id=sideTitle><SPAN class=blogTitleSub><%=SiteSubTitle%></SPAN></P>
</DIV>
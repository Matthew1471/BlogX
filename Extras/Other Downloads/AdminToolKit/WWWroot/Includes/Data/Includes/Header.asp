<!-- #INCLUDE FILE="Config.asp" -->
<%
If Instr(Request.Querystring(),"LegacyMode") <> 0 Then LegacyMode = True

If LegacyMode = True Then
ArgoSoftMailServer = 0
BackgroundColor = "#FFFFFF"
CalendarCheck = 1
CommentNotify = 0
NotifyPingOMatic = 0
NoDate = 0
MailingList = 0
Polls = 0
RSSImage = 0
SortByDay = True
Template = "Default"
End If

Session("CookieTest") = "AOK" 
%>
<!-- #INCLUDE FILE="Security.asp" -->
<%
Response.AddHeader "X-Pingback",SiteURL & "RSS/PingBack/Default.asp"
If Logging = True Then %>
<!-- #INCLUDE FILE="Logging.asp"-->
<%End If %>
<!-- #INCLUDE FILE="Calendar_Querystrings.asp" -->
<%
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

function PrintPopup(url) {
  popupWin = window.open(url,'Print','width=450,height=525,scrollbars=yes,toolbar=yes,menubar=yes,resizable=yes')
}

<% If (PingBackPage = True) AND (Request.Form("Content") <> "") Then %>
  var winl = (screen.width-275)/2
  var wint = (screen.height-200)/2
  myWindow = window.open("PingBack.asp", "PingBack",'toolbar=no,statusbar=yes,location=no,scrollbars=no,resizable=yes,width=275,height=200,top=' + wint + ',left=' + winl);
<% End If
If (AlertBack = True) AND (Request.Form("Action") <> "Post") Then %>
  var bolIsSubmitted = true;

//Inda : The onbeforeunload event is new to Mozilla (27/12/04) not everyone will have it.
window.onbeforeunload = window_onbeforeunload;

function window_onbeforeunload()
{         
	if(!bolIsSubmitted)
	{
		//this is a non-standard microsoft feature
		//event.returnValue="You've modified a textbox or checkbox but haven't saved your changes!";
		
		return "You've modified a textbox or checkbox but haven't saved your changes!"
	}
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
<Link rel="Alternate" type="application/rss+xml" title="RSS" href="<%=SiteURL%>Rss/<%If (ReaderPassword <> "") AND Session("Reader") = True Then Response.Write "?" & ReaderPassword%>">
<link rel="pingback" href="<%=SiteURL%>RSS/PingBack/Default.asp">
<style><!--
/* image shadow */
.dropshadow {
  clear: both;
  float:left;
  background: url(<%=SiteURL%>Images/shadowAlpha.png) no-repeat bottom right !important;
  background: url(<%=SiteURL%>Images/shadow.gif) no-repeat bottom right;
  margin: 10px 13px 0 6px !important;
  margin: 20px 7px 0 3px;

}

.dropshadow img
{
	display: block;
	position: relative;
	background-color: #fff;
	border: 1px solid #a9a9a9;
	margin: -6px 6px 6px -6px;
	padding: 4px;
}

/* image shadow */
.dropshadowr {
  clear: both;
  float:right;
  background: url(<%=SiteURL%>Images/global/shadowAlpha.png) no-repeat bottom right !important;
  background: url(<%=SiteURL%>Images/global/shadow.gif) no-repeat bottom right;
  margin: 20px 6px 0 30px !important;
  margin: 30px 3px 0 20px;


}

.dropshadowr  img
{
	display: block;
	position: relative;
	background-color: #fff;
	border: 1px solid #a9a9a9;
	margin: -6px 6px 6px -6px;
	padding: 4px;
}
--></style>
</HEAD>
<BODY bgColor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<DIV id=header>
<H1 id=title><A style="TEXT-DECORATION: none" href="<%=SiteURL%>"><%=SiteName%></A></H1>
<P id=byline><%=SiteDescription%></P>
<P id=sideTitle><SPAN class=blogTitleSub><%=SiteSubTitle%></SPAN>
<SPAN class=blogTitleSubDisclaimer>Please read my <A href="<%=SiteURL%>Disclaimer.asp">disclaimer</A>.</SPAN>
</P></DIV>
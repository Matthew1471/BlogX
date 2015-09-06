<% OPTION EXPLICIT %>
<html>

<head>
<title>Download WebBlogX</title>
<% If Request.Querystring() = "" Then %>
</head>
<frameset rows="*,1">
  <frame name="main" src="Share.asp?Download">
  <frame name="main" scrolling="no" noresize src="http://blogX.co.uk/Download.asp" target="main">
  <noframes>
  <body>
  
  <p><a href="http://blogx.co.uk/Download.asp">Click here</a> to download BlogX from Matthew1471.co.uk</p>
  <p><a href="http://freewebs.com/matthew1471/">Click here</a> to download BlogX from Free Webs</p>

  </body>
<% Else %>
<META HTTP-EQUIV="REFRESH" CONTENT="10; URL=http://freewebs.com/matthew1471/">
</head>
<body bgcolor="Orange">
Redirecting To nearest mirror or to <a href="http://freewebs.com/matthew1471/">Freewebs</a> in 10 Seconds
</body>
<% End If %>
  </noframes>
</frameset>

</html>

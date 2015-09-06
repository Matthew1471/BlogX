<html>
<head>
<title>Debugging Client</title>
<meta name="description" content="Using XML within ASP pages" />
<meta name="keywords" content="XML, XML-RPC, ASP, Active Server Pages, Internet Explorer, NT, Windows" />
<meta name="generator" content="Frontier 6.1.1 Win95" /></head>
<body bgcolor="#FFFFFF" alink="#008000" vlink="#800080" link="#0000FF">


<!-- #INCLUDE FILE="../../Includes/xmlrpc.asp" -->
<%
on error resume next

	ReDim paramList(2)
        Dim i, vbNothing, myresp
	paramList(0)="http://blogx.co.uk/ViewItem.asp?Entry=397"
	paramList(1)="http://blogx.co.uk/Comments.asp?Entry=103"

	Response.write("<pre>" & Replace(functionToXML("pingback.ping", paramList), "<", "&lt;", 1, -1, 1) & "</pre>")
	myresp = xmlRPC ("http://BlogX.co.uk/RSS/Pingback/Default.asp", "pingback.ping", paramList)

	Response.write(myresp & "<p>")
	Response.write("<pre>" & Replace(serverResponseText, "<", "&lt;", 1, -1, 1) & "</pre>")

if err.number <>0 then
	response.write("Error number: " & err.number & "<P>")
	response.write("Error description: " & err.description & "<P>")
else
	'response.write(myresp)
end if
%>

<hr>
<h3>The Code</H3>
Let's say that the code wasn't working.
<br>There are obviously a number of places where you can look, but it's nice to be able to look at what you're passing to and getting from the server.
<br>You can do that, and you should also put error handlers in<P>

</body>
</html>
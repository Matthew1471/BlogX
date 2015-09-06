<% OPTION EXPLICIT 
'-- Rendered Obselete due to time restrictions of ASP engine --'
Server.ScriptTimeout = 6000
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<!-- #INCLUDE FILE="../Admin/Admin.asp" -->
<html>
<head>
<title>Verify BlogX Distributions!</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=windows-1252">
<% If Request.Querystring("Step") < 2 Then Response.Write "<META HTTP-EQUIV=Refresh CONTENT=""0; URL=VerifyBlogX.asp?Step=" & Int(Request.Querystring("Step")) + 1 & """>" %>
<!--
//= - - - - - - - 
// Copyright 2004, Matthew Roberts
// Copyright 2003, Chris Anderson
//= - - - - - - -
-->
<% If Request.Querystring("Theme") <> "" Then Template = Request.Querystring("Theme")%>
<Link href="<%=SiteURL%>Templates/<%=Template%>/Blogx.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="<%=BackgroundColor%>">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<br>
<Center>
<%
Select Case Request.Querystring("Step")

 Case 0
	Response.Write "<b>Gathering Links..</b><br><br>"
 Case 1
	Response.Write "<b>Gathering Content Providers..</b><br><br>"
 Case 2
	Response.Write "<b>Contacting Third Party Servers..</b><br><br>" 
		Dim objXMLHTTP, CurrentLine

		'---- Verify The Site Is PingBack'd! ----
		' If you don't have MSXML3 installed you can revert to the old line:
		Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
		'Set objXMLHTTP=Server.CreateObject("MICROSOFT.XMLHTTP")


		'### Open The Records ###
		Records.Open "SELECT ReferURL, Approved FROM ScriptRefer ORDER BY ReferURL;", Database

		Do Until (Records.EOF)

		CurrentLine = Records("ReferURL") & "Application.asp"
                        
		        On Error Resume Next
			objXMLHTTP.open "GET", CurrentLine, true
  			objXMLHTTP.SetRequestHeader "User-Agent", "Matthew1471 BlogX"
  			objXMLhttp.send()
  
			'Wait for up to 10 seconds if we've not gotten the data yet
  			If objXMLHTTP.readyState <> 4 then
    			objXMLHTTP.waitForResponse 10
 			End If

  			If Err.Number <> 0 Then 
    			 Response.Write "<font color=""red"">XMLhttp Error  : </font>" & Hex(Err.Number) & " " & Err.Description & " (" & CurrentLine & ")<br>" & VbCrlf
			 Records("Approved") = False
			 Records.Update

			ElseIf objXMLhttp.status = 404 Then
			 Response.Write "<font color=""red"">Not Verified : </font>" & CurrentLine & "<br>" & VbCrlf
			 Records("Approved") = False
			 Records.Update

  	       		ElseIf objXMLhttp.status <> 200 Then 
    	       		 Response.Write "<font color=""red"">HTTP Error  : </font>" & CStr(objXMLhttp.status) & " " & objXMLhttp.statusText & " (" & CurrentLine & ")<br>" & VbCrlf
			 Records("Approved") = False
			 Records.Update
  			Else 

			  'Abort the XMLHttp request
			  If (objXMLhttp.readyState <> 4) Or (objXMLhttp.Status <> 200) Then objXMLhttp.Abort

                          ' -- Debugging --'
			  'Response.Write objXMLhttp.ResponseText

	      		   '--- QuickCheck Tm ---'
                           If Instr(1, objXMLhttp.responseText,"User/Password Error", 1) <> 0 Then 
			   Response.Write "<font color=""Lime"">Verified : </font>" & CurrentLine & "<br>" & VbCrlf

			     If Records("Approved") = False Then
                             Records("Approved") = True
			     Records.Update
			     Count = Count + 1
                             End If

                           Else
			   Response.Write "<font color=""red"">Not Verified : </font>" & CurrentLine & "<br>" & VbCrlf
			   Records("Approved") = False
			   Records.Update
                           End If
                         
                        End If
                        
                        On Error GoTo 0
                        
                        Response.Flush
			Records.MoveNext
			Loop

			'#### Close Object ###
			Records.Close

			Set objXMLHTTP = Nothing

			If Int(Count) > 0 Then Response.Write "<br><font color=""skyblue"">New Blogs Found : </font>" & Int(Count) & "<br>" & VbCrlf


 Case 3
	Response.Write "<b>Done...</b><br><br>"
	Response.Write "<Script>JavaScript:self.close();</script>"
End Select

Database.Close
Set Records = Nothing
Set Database = Nothing
%>
<br>
<a href="JavaScript:self.close();">Close</a>
</Center>

</Body>
</html>
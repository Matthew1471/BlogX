<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<p align="center">
<%
  			On Error Resume Next 
			Dim objXMLHTTP, VersionResponse

			' If you don't have MSXML3 installed you can revert to the old line:
			Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
			'Set objXMLHTTP=Server.CreateObject("MICROSOFT.XMLHTTP")

			objXMLHTTP.open "GET", "http://BlogX.co.uk/Download/UpdateBlogX.asp", true

			'-- It should really be set as a proper content type, but there's no point --
			objXMLHTTP.setRequestHeader "Content-Type", "text/xml"
			objXMLHTTP.SetRequestHeader "User-Agent", "Matthew1471 BlogX"

  			objXMLhttp.send()

			'Wait for up to 3 seconds if we've not gotten the data yet
  			If objXMLHTTP.readyState <> 4 then
    			objXMLHTTP.waitForResponse 5
 			End If

			'Abort the XMLHttp request
			If (objXMLhttp.readyState <> 4) Or (objXMLhttp.Status <> 200) Then objXMLhttp.Abort

			VersionResponse = ObjXMLHTTP.ResponseText

			' Write it out
			If (VersionResponse = Version) AND (Len(VersionResponse) > 0) Then 
			Response.Write "You Currently Have The <b>LATEST</b> BlogX Engine (V" & VersionResponse & ")"
			ElseIf (Len(VersionResponse) > 0) AND IsNumeric(Replace(VersionResponse,".","")) = True Then
			Response.write "<a href=""http://freewebs.com/matthew1471/"">BlogX V" & VersionResponse & "</a> Is Now Available!"
			Else
			Response.Write "BlogX Update server is currently down, you may <a href=""http://freewebs.com/matthew1471/"">manually download the latest version</a>."
			End If

  		        Set objXMLHTTP = Nothing
			On Error Goto 0
%>
<br><br><a href="<%=SiteURL & PageName%>">Back</a>
</p>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
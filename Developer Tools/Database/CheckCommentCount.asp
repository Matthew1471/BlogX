<% OPTION EXPLICIT%>
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
'-- Send out the page header --'
Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & VbCrlf
Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">" & VbCrlf
Response.Write "<head>" & VbCrlf
Response.Write " <title>Matthew1471's ASP BlogX - Fix Comment Count Inconsistencies</title>" & VbCrlf
Response.Write " <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" & VbCrlf
Response.Write "</head>" & VbCrlf
Response.Write "<body style=""background-color:skyblue"">" & VbCrlf

 '--- Open set ---'
 Records.Open "SELECT Data.RecordID, Data.Comments As [Reported Number], Count(Comments.EntryID) AS [Actual Count] FROM Data LEFT JOIN Comments ON Data.RecordID = Comments.EntryID GROUP BY Data.RecordID, Comments.EntryID, Data.Comments HAVING Data.Comments <> Count(Comments.EntryID);",Database, 1, 1

 '-- Check that there are actually some inaccurate records --'
 If Records.RecordCount > 0 Then 

  Response.Write "<p>" & VbCrlf

  '-- Loop through the inaccurate records until there are no more --'
  Do Until (Records.EOF)

   '-- Did the user want us to modify the DB or just inform them? --'
   If Request.Querystring("Mode") = "Fix" Then 
    Database.Execute "UPDATE Data SET Comments=" & Records("Actual Count") & " WHERE RecordID=" & Records("RecordID")
    Response.Write "Entry #" & Records("RecordID") & " has had its comment count updated to " & Records("Actual Count") & "<br/>" & VbCrlf
   Else
    Response.Write Records("Reported Number") & " records recorded for #" & Records("RecordID") & " but " & Records("Actual Count") & " calculated<br/>" & VbCrlf
   End If

   Records.MoveNext
  Loop

  Response.Write "</p>" & VbCrlf

  Response.Write "<form action="""" method=""get""><p><input name=""Mode"" type=""submit"" value=""Fix""/></p></form>"
 Else
  Response.Write "<p>No Problems Found! :)</p>"
 End If

Response.Write "</body>" & VbCrlf
Response.Write "</html>"

'--- Close The Records & Database ---'
Records.Close
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
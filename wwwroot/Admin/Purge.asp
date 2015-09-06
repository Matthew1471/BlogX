<%
' --------------------------------------------------------------------------
'¦Introduction : Purge Spam Comments Page.                                  ¦
'¦Purpose      : Provides a list of all flagged as spam comments.           ¦
'¦Used By      : Includes/NAV.asp.                                          ¦
'¦Requires     : Includes/Header.asp, Admin.asp, Includes/NAV.asp,          ¦
'¦               Includes/Footer.asp, Includes/Replace.asp.                 ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-08 Matthew Roberts, Chris Anderson
'**
'** This is free software; you can redistribute it and/or
'** modify it under the terms of the GNU General Public License
'** as published by the Free Software Foundation; either version 2
'** of the License, or any later version.
'**
'** All copyright notices regarding Matthew1471's BlogX
'** must remain intact in the scripts and in the outputted HTML
'** The "Powered By" text/logo with the http://www.blogx.co.uk link
'** in the footer of the pages MUST remain visible.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY; without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**********************************************************************
%>
<!-- #INCLUDE FILE="../Includes/Replace.asp" -->
<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<div id="content">
<%
 '-- If we are not wiping, provide a button to wipe --'
 If Request.Form("Confirm") <> "True" Then

  '-- Check we are not trying to SQL exploit the delete --'
  Dim DelRecNo
  DelRecNo = Request.Querystring("Delete")
  If (IsNumeric(DelRecNo) = False) OR (DelRecNo = "") Then DelRecNo = 0 Else DelRecNo = Int(DelRecNo)

  '-- We are welcome to delete comments --'
  If DelRecNo <> 0 Then Database.Execute "DELETE FROM Comments_Unvalidated WHERE CommentID=" & DelRecNo

  Response.Write "<form action=""Purge.asp"" method=""post""><p>The following are a list of "

  '-- Tell Our Users What Sort Of Limiting We Have Here --'
  If Request.Querystring("ShowAll") <> "True" Then
   Response.Write "the last " & EntriesPerPage & " (where applicable) "
  Else
   Response.Write "all "
  End If

  Response.Write "comments that failed BlogXs' Advanced IP Validation System, please confirm that all these messages are malicious and uneeded and then click <input type=""hidden"" name=""Confirm"" value=""True""/><input name=""Submit"" type=""submit"" value=""Purge These Comments""/></p></form>"

  '-- There should be a button to erase STRAIGHT away --'
  Response.Flush

 Else
  '-- Admin requested a purge --'
  Database.Execute "DELETE FROM Comments_Unvalidated"
 End If

 '-- Open the records ready to write --'
 Records.Open "SELECT CommentID, EntryID, PUK, Name, Email, Homepage, Content, CommentedDate, IP FROM Comments_Unvalidated ORDER BY CommentID DESC",Database, 1, 1

  '-- Split records in to pages --'
  If (NOT Request.Querystring("ShowAll") = "True") Then Records.PageSize = EntriesPerPage

  Do Until (Records.EOF) OR (Records.AbsolutePage <> 1 AND Request.Querystring("ShowAll") <> "True")

    '--- We're British, Let's 12Hour Clock Ourselves ---'
    Dim NewTime, NewDate, CommentedDate, CommentedTime

    CommentedDate = Records("CommentedDate")
    CommentedTime = FormatDateTime(CommentedDate,vbLongTime)

    NewTime = ""

    If TimeFormat <> False Then

     If Hour(CommentedTime) > 12 Then 
      NewTime = Hour(CommentedTime) - 12 & ":"
     Else
      NewTime = Hour(CommentedTime) & ":"
     End If
 
     If Minute(CommentedTime) < 10 Then
      NewTime = NewTime & "0" & Minute(CommentedTime)
     Else
      NewTime = NewTime & Minute(CommentedTime)
     End If

     If (Hour(CommentedTime) < 12) AND (Hour(CommentedTime) <> 12) Then
      NewTime = NewTime & " AM"
     Else
      NewTime = NewTime & " PM"
     End If

    Else
     If Hour(CommentedTime) < 10 Then NewTime = "0"
     NewTime = NewTime & Hour(CommentedTime) & ":"
     If Minute(CommentedTime) < 10 Then NewTime = NewTime & "0"
     NewTime = NewTime & Minute(CommentedTime)
    End If

    NewDate = Day(CommentedDate) & "/" & Month(CommentedDate) & "/" & Year(CommentedDate)

    Response.Write "EntryID : " & Records("EntryID") & "<br/>"
%>
<!--- Start Content For Comment <%=Records("CommentID")%> -->
<div class="comment">
 <h3 class="commentTitle">
  <acronym title="Allow This Comment"><a href="<%=SiteURL & "Comments_Validate.asp?CommentID=" & Records("CommentID") & "&amp;PUK=" & Records("PUK")%>"><img alt="Allow this comment" style="border-style: none" src="<%=SiteURL%>Images/Color.gif"/></a></acronym>
  <acronym title="IP Address <%=Records("IP")%>"><img alt="View IP address" style="border-style: none" onclick="javascript:alert('<%=Records("IP")%> failed to validate')" src="<%=SiteURL%>Images/Emoticons/Profile.gif"/></acronym>
  <acronym title="Delete Comment"><a href="?Delete=<%=Records("CommentID")%>"><img alt="Delete comment" style="border-style: none" src="<%=SiteURL%>Images/Key.gif"/></a></acronym>
  <%=NewDate%>&nbsp;<%=NewTime%>
 </h3>

 <span class="commentBody"><%=LinkURLs(Replace(Records("Content"), vbcrlf, "<br/>" & vbcrlf))%></span>

 <p class="commentFooter"><%
 If Records("HomePage") <> "" Then Response.Write "<a class=""permalink"" rel=""nofollow"" href=""" & HTML2Text(Records("Homepage")) & """>"
 Response.Write HTML2Text(Records("Name"))
 If Records("HomePage") <> "" Then Response.Write "</a>"

 If (Session(CookieName) = True) AND (Records("Email") <> "") Then Response.Write " | <span class=""comments""><a href=""mailto:" & HTML2Text(Records("Email")) & """>" & Records("Email") & "</a></span>"%></p>
</div>
<!-- End Content -->
<%
 Response.Flush
 Records.MoveNext
Loop

'-- Warn that we truncated the record set --'
If (Records.AbsolutePage > 1) AND (Request.Form("Confirm") <> "True") Then 
 Response.Write "<p style=""text-align:Center""><b><a href=""?ShowAll=True"">To protect page loading times, " & (Records.RecordCount - EntriesPerPage) & " unverified comments have not been displayed, you may click here to show them but it is not recommended.</a></b></p>"
'-- Notify that we wiped --'
ElseIf Request.Form("Confirm") = "True" Then
 Response.Write "<p style=""text-align:Center"">Unverified Comments Deleted!</p>" & VbCrlf & "<p style=""text-align:center""><a href=""" & SiteURL & PageName & """>Back</a></p>"
End If

Records.Close
%>
</div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
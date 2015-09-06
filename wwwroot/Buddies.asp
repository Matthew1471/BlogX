<%
' --------------------------------------------------------------------------
'¦Introduction : RSS Reader                                                 ¦
'¦Purpose      : Reads the list of "Buddies" in the database and            ¦
'¦               will display the contents of an RSS feed, stripped of HTML.¦
'¦Used By      : Links table if in Database (default).                      ¦
'¦Requires     : Includes/Header.asp, Includes/NAV.asp, Includes/Footer.asp ¦
'¦Notes        : This page is a very crude RSS reader that trims all tags   ¦
'¦               not all tags are bad but I err'd on the side of caution.   ¦
'¦Standards    : XHTML Strict.                                              ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT

'*********************************************************************
'** Copyright (C) 2003-09 Matthew Roberts, Chris Anderson
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
PageTitle = "Buddies"
%>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<div id="content">
 <div class="entry">
  <h3 class="entryTitle" style="text-align:center">My Buddies &amp; Read List</h3>
<%

'### Open The Records Ready To Write ###
Records.Open "SELECT LinkName, LinkURL, LinkRSS, LinkType FROM Links Where LinkType='RSS' ORDER BY LinkName Asc;", Database

 '### Array It! ###'
 If NOT Records.EOF Then
  Dim RSSSites
  RSSSites = Records.GetRows()

  '#### Close Objects ###
  Records.Close

  Response.Write "  <p style=""text-align: center"">These are a few of the websites I frequent:</p>" & VbCrlf
  Response.Write "  <ul>" & VbCrlf
 
  For Count = 0 To Ubound(RSSSites,2)
   ' 0 - Name
   ' 1 - Homepage
   ' 2 - RSS URL
   Response.Write "   <li><a href=""#Site" & Count & """>" & RSSSites(0, Count) & "</a></li>" & VbCrlf
  Next

  Response.Write "  </ul>" & VbCrlf
  Response.Write " </div>" & VbCrlf & VbCrlf

  '-- Lets Let Our User Know We Are Wroking --'
  Response.Flush

   ' Why am I doing this?... 
   '      Because of security, you wouldn't let a commenter use HTML.. 
   '                           so why a random blogger?
   Function KillHTML(Text)

   Dim MaskHTMLInstead
   MaskHTMLInstead = False

    If (MaskHTMLInstead) Then 

    KillHTML = Replace(Text, "<","&lt;")
    KillHTML = Replace(KillHTML, ">","&gt;")

    KillHTML = Replace(KillHTML, "&lt;b&gt;","")
    KillHTML = Replace(KillHTML, "&lt;/b&gt;","")
    KillHTML = Replace(KillHTML, "&lt;u&gt;","")
    KillHTML = Replace(KillHTML, "&lt;/u&gt;","")
    KillHTML = Replace(KillHTML, "&lt;i&gt;","")
    KillHTML = Replace(KillHTML, "&lt;/i&gt;","")
    
    KillHTML = Replace(KillHTML, "&lt;br&gt;"," ")
    KillHTML = Replace(KillHTML, "&lt;p&gt;"," ")
    KillHTML = Replace(KillHTML, "&lt;/p&gt;"," ")

    '-- Just For Dan & His Difficult Blogger Blog --'
    KillHTML = Replace(KillHTML, "&lt;br /&gt;"," ")
    KillHTML = Replace(KillHTML, "&lt;div xmlns=""http://www.w3.org/1999/xhtml""&gt;","")
    KillHTML = Replace(KillHTML, "&lt;/div&gt;","")

    '-- Kill Dan's Tag List --'
    If Instr(KillHTML,"&lt;div class=""tag_list""&gt;") <> 0 Then KillHTML = Left(KillHTML,Instr(KillHTML,"&lt;div class=""tag_list""&gt;") - 1)
    If Instr(KillHTML,"&lt;p style=""font-size:10px;text-align:right;""&gt;technorati tags") <> 0 Then KillHTML = Left(KillHTML,Instr(KillHTML,"&lt;p style=""font-size:10px;text-align:right;""&gt;technorati tags") - 1)
    If Instr(KillHTML,"&lt;p style=""FONT-SIZE: 10px; TEXT-ALIGN: right""&gt;technorati tags") <> 0 Then KillHTML = Left(KillHTML,Instr(KillHTML,"&lt;p style=""FONT-SIZE: 10px; TEXT-ALIGN: right""&gt;technorati tags") - 1)

    Else

    Dim CharacterCount
    Dim OpenTag
    Dim OpenTagPos

     For CharacterCount = 1 To Len(Text)

      '-- We have not already opened and there's a bracket! --'
      If (OpenTag = False) AND (Mid(Text,CharacterCount,1) = "<") Then
       If (Mid(Text,CharacterCount,4) <> "<br/>") AND (Mid(Text,CharacterCount,4) <> "<p>") AND (Mid(Text,CharacterCount,4) <> "</p>") Then
        OpenTag = True
        OpenTagPos = CharacterCount
       End If
      ElseIf (OpenTag = True) AND (Mid(Text,CharacterCount,1) = ">") Then
       OpenTag = False
       Text = Left(Text,OpenTagPos-1) & " " & Right(Text,Len(Text) - CharacterCount)

       '-- Reset back to the old text --'
       CharacterCount = OpenTagPos
       
      End If

     Next

     KillHTML = Text

    End If


   End Function
   
  For Count = 0 To UBound(RSSSites,2)
   Dim xmlHTTP, RSSItem, RSSItems, RSSItemsCount
   Dim RSStitle, RSSlink, RSSdescription, RSSDate
   Dim j, i, child

   Response.Write "<!-- Start Content For " & RSSSites(0,Count) & "(" & Count & ")-->" & VbCrlf
   Response.Write "<div class=""entry"">" & VbCrlf
   Response.Write " <div class=""entryIcon"">" & VbCrlf
   Response.Write "  <h3 class=""entryTitle""><a id=""Site" & Count & """ href=""" & RSSSites(1,Count) & """ onclick=""target='_new';"">" & RSSSites(0,Count) & "</a></h3><br/>" & VbCrlf

   Response.Write " </div>" & VbCrlf
   Response.Write " <div class=""entryBody"">" & VbCrlf
   Response.Write "  <p>" & VbCrlf

  On Error Resume Next

   Dim objXMLHTTP

   '-- ServerXMLHTTP is included with the Microsoft XML Parser (MSXML) version 3.0 or later --'
   '-- If you do not have MSXML6 installed revert to the old line:
   Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
   If Err <> 0 Then Set objXMLHTTP= Server.CreateObject("MSXML2.ServerXMLHTTP")
   objXMLHTTP.open "GET", CStr(RSSSites(2,Count)), true
    objXMLHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
   objXMLhttp.send()

   '-- Wait for up to 1 second if we have not recieved all the data yet --'
   If (objXMLhttp.readyState <> 4) Then objXMLHTTP.waitForResponse 1

   If Err.Number <> 0 Then 
    Response.Write "<b>XMLhttp Error:</b> " & Hex(Err.Number) & " " & Err.Description & "<br/>" & VbCrlf
    Response.Write "An error occured while trying to process <b>" & RSSSites(0,Count) & "</b><br/>Please contact the webmaster."
   Else

    '-- Abort the XMLHttp request --'
    If (objXMLhttp.readyState <> 4) Then
     Response.Write "<b>Timeout Error:</b> Sorry, but the content did not load quick enough so we stopped waiting for it. (State:" & objXMLhttp.readyState & ")<br/>" & VbCrlf
     objXMLhttp.Abort
    '-- Though the page returned, it did not give a "healthy" error code --'
    Else If (objXMLhttp.status <> 200) Then
     Response.Write "<b>HTTP Error:</b> Unfortunately after requesting the content the remote party stated """ & CStr(objXMLhttp.status) & " " & objXMLhttp.statusText & """.<br/>" & VbCrlf
    Else

     Dim xmlDOM

     '-- If you have MS XML 6 it is preferable to use it.. if not replace with the alternative (which calls MS XML 3). --'
     Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument.6.0")
     If Err <> 0 Then Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument")

     xmlDOM.async = False
     xmlDOM.validateOnParse = False
     xmlDom.resolveExternals = False

     '-- Is this valid XML? --'
     If xmlDOM.LoadXml(objXMLHTTP.ResponseText) Then

      '-- Collect All "items" From Downloaded RSS --'
      Set RSSItems = xmlDOM.getElementsByTagName("item")
      RSSItemsCount = RSSItems.Length-1

      If (RSSItemsCount < 1) AND Len(xmlDOM.parseError.reason) = 0 Then Response.Write "There are no ""item""'s in this feed.<br/>" & VbCrlf
 
      J = -1

      For i = 0 To RSSItemsCount
 
       Set RSSItem = RSSItems.Item(i)

       For Each child In RSSItem.childNodes
        Select Case lcase(child.nodeName)
         Case "pubdate" : RSSDate = child.text
         Case "title" : RSStitle = child.text
         Case "link" : RSSlink = child.text
         Case "description" : RSSdescription = child.text
        End Select
       Next
 
       J = J+1

       If J < EntriesPerPage Then Response.Write "  <a onclick=""target='_new';"" href=" & """" & KillHTML(RSSlink) & """" & ">" & KillHTML(RSSTitle) & "</a><br/><small>&nbsp;(" & KillHTML(RSSDate) & ")</small><br/>" & KillHTML(RSSDescription) & "<br/><br/>" & VbCrlf
      Next
  
     Else
      Response.Write "<b>XMLDom Error:</b> The data returned did not make sense (usually due to a malformed document).<br/>" & VbCrlf
      If Len(xmlDOM.parseError.reason) > 0 Then Response.Write "Sorry, but I could not make sense of this content because """ & xmlDOM.parseError.reason & """<br/>" & VbCrlf
    End If

    Set xmlDOM = Nothing

    End If

    If Err.Number <> 0 Then
     Response.Write "<b>General Error:</b> " & Hex(Err.Number) & " " & Err.Description & "(" & Err.LineNumber & ")<br/>" & VbCrlf
     Err.Clear
    End If

    '-- Write Error --'
    If RSSItemsCount <= 0 Then Response.Write "An error occured while trying to process <b>" & RSSSites(0,Count) & "</b><br/>Please contact the webmaster."

    On Error GoTo 0

   End If
  End If

   Set objXMLHTTP = Nothing 

   Response.Write "  </p>" & VbCrlf
   Response.Write " </div>" & VbCrlf

   Response.Write " <p class=""entryFooter"">" & Time() & VbCrlf
   Response.Write " | <span class=""comments""><a title=""This is an external website, visit it to check/write comments"" href=""" & RSSSites(1,Count) & """ onclick=""target='_new';"">Comments [?]</a></span></p></div>" & VbCrlf
   Response.Write " <!-- End Content -->" & VbCrlf

  Next
Else

 '#### Close Objects ###
 Records.Close
 Response.Write "<p style=""text-align: center"">This blog owner has not added any RSS Sites.</p>" & VbCrlf

End If
%>

This website owner is not responsible for the availability of such external sites or 
resources, and does not endorse and is not responsible or liable for any content 
on or available from such sites or resources.

</div>
<!-- #INCLUDE FILE="Includes/Nav.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
<% OPTION EXPLICIT 
Session.LCID = 2057

'-- Small performance gains can be made by hardcoding this in --'
Dim ServerPathToInstalledDirectory
ServerPathToInstalledDirectory = Server.MapPath(".")
'ServerPathToInstalledDirectory = "C:\inetpub\wwwroot"

'-- This host uses PHP --'
Dim PHPEnabled
PHPEnabled = True
%>
<!-- #INCLUDE FILE="Includes/Replace.asp" -->
<!-- #INCLUDE FILE="Includes/Config.asp" -->
<!-- #INCLUDE FILE="Includes/ViewerPass.asp" -->
<%
'--- Open Recordset ---'
    Records.CursorLocation = 3 ' adUseClient
    Records.Open "SELECT RecordID, Title, Text, Password FROM Data WHERE Password='' ORDER BY RecordID DESC",Database, 1, 3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <meta http-equiv="Content-Language" content="en-gb"/>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 <title><%=SiteDescription%> Newsletter</title>
</head>

<body>

<div style="text-align: center;">
  <table border="1" cellpadding="0" cellspacing="6" style="border-collapse: collapse; text-align: left;" bordercolor="#111111" width="80%" bordercolordark="#000000" bordercolorlight="#000000" height="121">
    <tr>
      <td colspan="3" height="72">
      <p align="center"><img alt="Newspaper Logo" border="0" src="Images/NewsPaperLogo.gif"/><br/>&nbsp;</p></td>
    </tr>
    <tr>
      <td height="19" bordercolor="#FFFFFF" valign="top">
      <a href="<%=SiteURL%>Newspaper.asp"><%=SiteURL%>Newspaper.asp</a></td>
      <td height="19" colspan="2" bordercolor="#FFFFFF">
      <p align="right">The BlogX Chronicles<br/><%=FormatDateTime(Now(),2)%></p></td>
    </tr>
    <tr>
      <td width="612" height="1" colspan="3"></td>
    </tr>
    <tr>
      <% 
	Dim RecordID, Title, Text, Password

	'--- Setup Variables ---'
   	Set RecordID = Records("RecordID")
   	Set Title = Records("Title")
   	Set Text = Records("Text")
   	Set Password =  Records("Password")

      Do Until (Records.EOF or Count = 3)
      %>
      <td valign="top" <% If Count <> 2 Then Response.Write "rowspan=""2"""%>>
      <!-- Section <%=Count%> -->
      <h2><%=Title%></h2>
      <%=Replace(LinkURLs(Replace(Text, VbCrlf, "<br/>" & VbCrlf)), "Images/Articles/", "/Images/Articles/")%>
      </td>
    
          <% If Count = 2 Then %>
          </tr>
          <tr>
            <td valign="top" height="10%" align="center">
            <h2 align="center">Random Photo</h2>
            <p>
            <%
            Set FSO = Server.CreateObject("Scripting.FileSystemObject")

            Dim Folder
            Set Folder = FSO.GetFolder(ServerPathToInstalledDirectory & "\Images\Articles\")

            Dim FileCount
            FileCount = -1
     
            ReDim FileArray(20)

            For Each File In Folder.Files
             FileCount = FileCount + 1

             '-- Redim Array? --'	
             If FileCount > UBound(FileArray) then ReDim Preserve FileArray(FileCount + 20)
             FileArray(FileCount) = Replace(File.Path,ServerPathToInstalledDirectory & "\Images\Articles\","")
	
             '-- Make sure FileArray is the right size --'
             Redim Preserve FileArray(FileCount)

            Next

             '-- Did we find any files? --'
             If FileCount > -1 Then

             Randomize

             Dim RandomNumber
             RandomNumber = Int((UBound(FileArray) - 1 + 1) * Rnd + 1)

             Dim strFileName
             strFileName = FileArray(RandomNumber)

             '-- Thumbnail Handler --'
             Dim Ext
             Ext = UCase(Right(strFileName, 3)) 
             If Ext = "JPG" or Ext = "GIF" or Ext = "PNG" or Ext = "BMP" Then

             Dim Thumbnails
              If FSO.FileExists(ServerPathToInstalledDirectory & "\Images\Articles\Thumbnails\tn" & strFileName) Then 
               Thumbnails = "Thumbnails/tn" & strFileName
              ElseIf PHPEnabled = True Then
               '--Lets Generate A Thumbnail --'
               On Error Resume Next
                FSO.CreateFolder ServerPathToInstalledDirectory & "\Images\Articles\Thumbnails"
               On Error GoTo 0

               Thumbnails = "Thumbnail.php?f=" & Server.URLEncode(strFileName)
              Else
               Thumbnails = Server.URLEncode(strFileName)
              End If
             End If

             '-- Trim to get the record number --'
	       Dim ImageNumber
             If Len(strFileName) > 4 Then ImageNumber = Right(strFileName,Len(strFileName) - InstrRev(strFileName,"Entry") - 4)
             If Instr(ImageNumber,"_") Then ImageNumber = Left(ImageNumber,Instr(ImageNumber,"_") - 1)

             If NOT IsNumeric(ImageNumber) Then
   	       Response.Write "Not an entry image : " & strFileName
	      Else

	       Records.Close

	     Records.Open "SELECT RecordID, Title FROM Data WHERE RecordID=" & ImageNumber,Database, 1, 3
	      If NOT Records.EOF Then
	       Dim ImageTitle
	       ImageTitle = Records("Title")
             End If
	    %>
             <a target="_new" href="/Images/Articles/<%=Server.URLEncode(strFileName)%>"><img alt="Photo from entry <%=ImageNumber%>" src="/Images/Articles/<%=Thumbnails%>" border="0"/></a></p>
             <p>This weeks random photo is from entry <%=ImageNumber%> &quot;<b><%=ImageTitle%></b>&quot;</p>
          <%

	     Set Folder = Nothing
	     Set FSO = Nothing

          End If
             End If %>
           </td>
           <%

	   End If

	Count = Count + 1
	Records.MoveNext
	Loop
	
	'--- Close The Records ---
	Records.Close
       %>
    </tr>
    <tr>
      <td height="38" colspan="3" align="center" bgcolor="black" style="color : white;">Have a suggestion 
      on what should go in the BlogX press? <a href="/Mail.asp?BlogX Press"><font color="white">Contact</font></a> our team of editors</td>
    </tr>
  </table>
</div>
<div style="text-align: center;"><%=Copyright%></div>
<!-- Generated by Matthew1471 BlogX v.<%=Version%> -->
</body>
</html>
<% Set Records = Nothing
Database.Close
Set Database = Nothing
%>
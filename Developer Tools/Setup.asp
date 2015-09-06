<% OPTION EXPLICIT

Dim DataFile, Database, Records

'-- !! CHANGE THIS TO YOUR DATABASE PATH !! --'
'DataFile = "C:\Inetpub\Database\BlogX.mdb"
DataFile = "D:\Inetpub\wwwroot\Temp\BlogX\BlogX(upgraded).mdb"

'-- We default to Step 0 --'
Dim Step
If Request.Querystring("Step") <> "" AND IsNumeric(Request.Querystring("Step")) Then 
 Step = Request.Querystring("Step")
ElseIf Request.Querystring("Direction") <> "" AND (Request.Querystring("StepNoJS") <> "" AND IsNumeric(Request.Querystring("StepNoJS"))) Then
  If Request.Querystring("Direction") = "Next-->" Then
   Step = Request.Querystring("StepNoJS") + 1
  ElseIf Request.Querystring("Direction") = "<--Back" Then
   Step = Request.Querystring("StepNoJS") - 1
  Else
   Step = 0
  End If
Else
 Step = 0
End If

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

<head>
 <title>Server Check</title>
 <style type="text/css">
  .header { text-align: center; }
  .black { background: #000000; color: #FFFFFF }

  .installed { background: Green; color: #FFFFFF; text-align: center }
  .notinstalled { background: #FF0000; color: #FFFFFF; text-align: center }

  .smallblue { color:skyblue; font-size:small; font-weight: bold }
  .smallpurple { color:purple; font-size:small; font-weight: bold }
 </style>
</head>

<body>
 <h1 class="header">BlogX Setup</h1>
<%
Select Case Step

'-- Step 1 --'
Case 0

 Response.Write " <h3 style=""text-align: center"">Step 1 - Introduction</h3>" & VbCrlf & VbCrlf

 Response.Write " <p style=""text-align: center"">Welcome to BlogX, this diagnostic page shows what functions of BlogX are available on this server,<br/> upgrades the database to the latest version and checks the database is valid.</p>" & VbCrlf
 Response.Write " <p style=""text-align: center"">Please click &quot;Next&quot; to continue.</p>" & VbCrlf & VbCrlf

'-- Step 2 --'
Case 1

 Function IsObjInstalled(strClassString)
  On Error Resume Next

   '-- Test Each Component --'
   Dim xTestObj
   Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True Else IsObjInstalled = False
   Set xTestObj = Nothing

  On Error GoTo 0

 End Function

 Response.Write " <h3 style=""text-align: center"">Step 2 - Server Component Check</h3>" & VbCrlf & VbCrlf

 Response.Write " <p style=""text-align: center"">Below shows what functions of BlogX are available on this server.<br/>" & VbCrlf
 Response.Write " If a component is not installed, BlogX should run but without the ""Affects"" features.</p>" & VbCrlf & VbCrlf

 Response.Write " <table border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse; text-align: center; margin: 0 auto; width: 40%"">" & VbCrlf
 Response.Write "  <tr>" & VbCrlf
 Response.Write "   <th class=""black"" style=""width: 25%"">Component</th>" & VbCrlf
 Response.Write "   <th class=""black"">Affects</th>" & VbCrlf
 Response.Write "   <th class=""black"" style=""width: 25%"">Status</th>" & VbCrlf
 Response.Write "  </tr>" & VbCrlf

 '-- FSO Check --'
 Response.Write "  <tr>" & VbCrlf
 Response.Write "   <td>FSO<br/><span style=""font-size: small"">(File System Object)</span></td>" & VbCrlf
 Response.Write "   <td>Links, OtherLinks, Pingback, Spell-Check, Dynamic Insertion of &quot;Templates&quot;</td>" & VbCrlf

 If IsObjInstalled("Scripting.FileSystemObject") Then
  Response.Write "   <td class=""installed"">INSTALLED</td>" & VbCrlf
 Else
  Response.Write "   <td class=""notinstalled"">NOT INSTALLED</td>" & VbCrlf
 End If

 Response.Write "  </tr>" & VbCrlf

 '-- XHTML Check --'
 Response.Write "  <tr>" & VbCrlf
 Response.Write "   <td>XMLHTTP</td>" & VbCrlf
 Response.Write "   <td>PingBack</td>" & VbCrlf

 If IsObjInstalled("MSXML2.ServerXMLHTTP") Then 
  Response.Write "   <td class=""installed"">INSTALLED</td>" & VbCrlf
 Else
  Response.Write "   <td class=""notinstalled"">NOT INSTALLED</td>" & VbCrlf
 End If
 
 Response.Write "  </tr>" & VbCrlf

 '-- Upload Component Check --'
 Response.Write "  <tr>" & VbCrlf
 Response.Write "   <td>Upload Component</td>" & VbCrlf
 Response.Write "   <td>PhotoUploads</td>" & VbCrlf
 
 Dim TheComponent(), InstalledUpload, InstalledEmail
 ReDim TheComponent(2,2)

 '-- The Components --'
 TheComponent(0,0) = "Persits.Upload"
 TheComponent(0,1) = "Persits ASPUpload"

 TheComponent(1,0) = "aspSmartUpload.SmartUpload"
 TheComponent(1,1) = "ASPSmartUpload"
	
 Dim I, J
 
 For I = 0 to UBound(TheComponent)
  If IsObjInstalled(TheComponent(i,0)) Then
   If Len(InstalledUpload) > 0 Then InstalledUpload = InstalledUpload & VbCrlf
   InstalledUpload = InstalledUpload & TheComponent(i,1)
  Else
   J = J + 1
  End If
 Next
	
 If J > UBound(TheComponent) Then
  Response.Write "   <td class=""notinstalled"">NOT INSTALLED</td>"
 Else
  Response.Write "   <td class=""installed"">INSTALLED<br/>" & VbCrlf
  Response.Write "   <span class=""smallblue"">" & InstalledUpload & "</span>" & VbCrlf
  Response.Write "   </td>" & VbCrlf
 End If 
 
 Response.Write "  </tr>" & VbCrlf
 Response.Write "  <tr>" & VbCrlf
 Response.Write "   <td>Email Component</td>" & VbCrlf
 Response.Write "   <td>CommentNotification, MailingLists, MailAuthor</td>" & VbCrlf

 ReDim TheComponent(18,2)

 '-- The E-mail Components --'
 TheComponent(0,0) = "ABMailer.Mailman"
 TheComponent(0,1) = "ABMailer v2.2+"

 TheComponent(1,0) = "Persits.MailSender"
 TheComponent(1,1) = "ASPEMail"

 TheComponent(2,0) = "SMTPsvg.Mailer"
 TheComponent(2,1) = "ASPMail"

 TheComponent(3,0) = "SMTPsvg.Mailer"
 TheComponent(3,1) = "ASPQMail"

 TheComponent(4,0) = "CDONTS.NewMail"
 TheComponent(4,1) = "CDONTS (IIS 3/4/5)"

 TheComponent(5,0) = "CDONTS.NewMail"
 TheComponent(5,1) = "Chili!Mail (Chili!Soft ASP)"

 TheComponent(6,0) = "CDO.Message"
 TheComponent(6,1) = "CDOSYS (IIS 5/5.1/6)"

 TheComponent(7,0) = "dkQmail.Qmail"
 TheComponent(7,1) = "dkQMail"

 TheComponent(8,0) = "Dundas.Mailer"
 TheComponent(8,1) = "Dundas Mail (QuickSend)"

 TheComponent(9,0) = "Dundas.Mailer"
 TheComponent(9,1) = "Dundas Mail (SendMail)"

 TheComponent(10,0) = "Geocel.Mailer"
 TheComponent(10,1) = "GeoCel"

 TheComponent(11,0) = "iismail.iismail.1"
 TheComponent(11,1) = "IISMail"

 TheComponent(12,0) = "Jmail.smtpmail"
 TheComponent(12,1) = "JMail"

 TheComponent(13,0) = "MDUserCom.MDUser"
 TheComponent(13,1) = "MDaemon"

 TheComponent(14,0) = "ASPMail.ASPMailCtrl.1"
 TheComponent(14,1) = "OCXMail"

 TheComponent(15,0) = "ocxQmail.ocxQmailCtrl.1"
 TheComponent(15,1) = "OCXQMail"

 TheComponent(16,0) = "SoftArtisans.SMTPMail"
 TheComponent(16,1) = "SA-Smtp Mail"

 TheComponent(17,0) = "SmtpMail.SmtpMail.1"
 TheComponent(17,1) = "SMTP"

 TheComponent(18,0) = "VSEmail.SMTPSendMail"
 TheComponent(18,1) = "VSEmail"

 J = 0

 For I = 0 to UBound(TheComponent)
  If IsObjInstalled(TheComponent(i,0)) Then
   If Len(InstalledUpload) > 0 Then InstalledUpload = InstalledUpload & VbCrlf
   InstalledEmail = InstalledEmail & TheComponent(i,1)
  Else
   J = J + 1
  End If
 Next
	
 If J > UBound(TheComponent) Then
  Response.Write "   <td class=""notinstalled"">NOT INSTALLED</td>"
 Else
  Response.Write "   <td class=""installed"">INSTALLED<br/>" & VbCrlf
  Response.Write "   <span class=""smallblue"">" & InstalledEmail & "</span>" & VbCrlf
  Response.Write "   </td>" & VbCrlf
 End If 

 Response.Write "  </tr>" & VbCrlf
 Response.Write " </table>" & VbCrlf & VbCrlf

 Response.Write " <p style=""text-align: center; font-size: small"">Note: This page does not check your server's ability to write to or read from files.</p>" & VbCrlf

 Response.Write " <p style=""text-align: center"">Please click &quot;Next&quot; to continue.</p>" & VbCrlf & VbCrlf

'-- Step 3 --'
Case 2

 '-- If the database is not okay then we refuse to let the user advance --'
 Dim PreventStep

 Response.Write " <h3 style=""text-align: center"">Step 3 - Server Database Check</h3>" & VbCrlf & VbCrlf

 Set Database = Server.CreateObject("ADODB.connection")

  On Error Resume Next
   Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Datafile & ";"

   If Err <> 0 Then
    Dim DBError
    DBError = True

    '-- We specified a wrong path to the database --'
    If Err = -2147467259 Then
     '-- Try and guess where the database folder might be --'
     If Instr(1, Server.MapPath("."),"wwwroot", 1) Then

      '-- One last chance --'
      Err.Clear

      Dim GuessedPath
      GuessedPath = Left(Server.MapPath("."),Instr(1,Server.MapPath("."),"wwwroot",1)-1) & "Database\BlogX1.mdb"
      
      Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GuessedPath & ";"

      If Err = 0 Then
       DBError = False
       Response.Write " <p style=""text-align: center; background: skyblue;"">The database file was automatically found at <b>" & GuessedPath & "</b> please edit the Includes\Config.asp in notepad to use this file.</p>" & VbCrlf
      ElseIf Err = -2147467259 Then
       Response.Write " <p style=""text-align: center; background: orange;"">Please edit this file (" & Request.ServerVariables("PATH_TRANSLATED") & ") to use the correct database file.<br/>" & VbCrlf
       Response.Write " We tried opening the specified file &quot;" & DataFile & "&quot; (we even tried looking in " & GuessedPath & ") but it does not exist.</p>" & VbCrlf & VbCrlf
      End If

     Else
      Response.Write " <p style=""text-align: center; background: orange;"">Please edit this file (" & Request.ServerVariables("PATH_TRANSLATED") & ") to use the correct database file.<br/>" & VbCrlf
      Response.Write " We tried opening the specified file &quot;" & DataFile & "&quot; but it does not exist.</p>" & VbCrlf & VbCrlf
     End If

    Else
      Response.Write " <p style=""text-align: center; background: orange;"">We could not open the database at " & DataFile & " because the error &quot;<b>" & Err.Description & "</b>&quot; occured.<br/><br/>Please fix this and refresh this page.</p>" & VbCrlf & VbCrlf
    End If

   End If
  On Error GoTo 0

  If DBError = False Then

  Set Records = Server.CreateObject("ADODB.recordset")

  '-- This is where we return any CREATE TABLE statements --'
  Function FindCreationString(Table)

   '-- Compression is an Access hack (http://www.fmsinc.com/Free/NewTips/Access/accesstip44.asp) --'

   Dim CreationStrings(4,1)

   '-- EntryID Should be "Indexed- Yes (Duplicates OK) --'

   CreationStrings(0,0) = "Comments_Unvalidated"
   CreationStrings(0,1) = "CommentID COUNTER NOT NULL, " & _
                          "EntryID LONG NOT NULL, " & _
                          "Name VARCHAR(50) WITH COMP NOT NULL, " & _
                          "Email TEXT(50) WITH COMP, " & _ 
                          "Homepage TEXT(50) WITH COMP, " & _
                          "Content MEMO WITH COMP NOT NULL, " & _
                          "CommentedDate DATETIME NOT NULL, " & _
                          "UTCTimeZoneOffset TEXT(5) WITH COMP, " & _
                          "IP TEXT(50) WITH COMP , " & _
                          "Subscribe YESNO, " & _
                          "PUK LONG, " & _
                          "Primary Key (CommentID)"

   CreationStrings(1,0) = "BannedLoginIP"
   CreationStrings(1,1) = "BannedIP TEXT(15) WITH COMP NOT NULL, " & _
                          "LoginFailCount LONG NOT NULL, " & _
                          "LastLoginFail DATETIME NOT NULL, " & _
                          "Primary Key (BannedIP)"

   CreationStrings(2,0) = "Links"
   CreationStrings(2,1) = "LinkID COUNTER NOT NULL, " & _
                          "LinkName VARCHAR(80) WITH COMP NOT NULL, " & _
                          "LinkURL VARCHAR(255) WITH COMP NOT NULL, " & _
                          "LinkRSS VARCHAR(255) WITH COMP NOT NULL, " & _
                          "LinkType VARCHAR(11) WITH COMP NOT NULL, " & _
                          "Primary Key (LinkID)"



   CreationStrings(3,0) = "Mail_Unvalidated"
   CreationStrings(3,1) = "RecordID COUNTER NOT NULL, " & _
                          "FromEmail VARCHAR(255) WITH COMP NOT NULL, " & _
                          "Subject VARCHAR(255) WITH COMP NOT NULL, " & _
                          "Body MEMO WITH COMP NOT NULL, " & _
                          "IP VARCHAR(50) WITH COMP NOT NULL, " & _
                          "PUK LONG NOT NULL, " & _
                          "Primary Key (RecordID)"

   CreationStrings(4,0) = "NotFound"
   CreationStrings(4,1) = "ID COUNTER NOT NULL, " & _
                          "URL VARCHAR(255) WITH COMP NOT NULL, " & _
                          "ReferringPage VARCHAR(255) WITH COMP NOT NULL, " & _
                          "ErrorCount LONG NOT NULL, " & _
                          "Primary Key (ID)"
   Dim Found
   Found = -1

   Dim LoopCount
   For LoopCount = 0 To UBound(CreationStrings)

   If CreationStrings(LoopCount,0) = Table Then 
    Found = LoopCount
    Exit For
   End If

   Next

   If Found > -1 Then
    Dim TempString
    TempString = "CREATE TABLE " & CreationStrings(Found,0) & "(" & CreationStrings(Found,1) & ")"
    Response.Write "<!-- " & TempString & "-->" & VbCrlf
    FindCreationString = TempString
   End If

  End Function

  '-- JET has a few extra quirks that can't be modified via SQL (http://www.codeguru.com/cpp/data/mfc_database/ado/article.php/c4343) --'
  Sub PerformJETChanges(TheValue)
   Dim TablesRequiringJETQuirks(0,1)

   '- 0 = Table Name
   '- 1 = Array of Field Properties

   '- 0 = Field Name
   '- 1 = Field Default
   '- 2 = Field Zero-Length?
   '- 3 = Field Requires Compression?

   TablesRequiringJETQuirks(0,0) = "Comments_Unvalidated"
   TablesRequiringJETQuirks(0,1) =  Array("CommentedDate","Now()","","")

   TablesRequiringJETQuirks(0,0) = "BannedIP"
   TablesRequiringJETQuirks(0,1) =  Array("IP","","False","True")

   Dim Count
   For Count = 0 To UBound(TablesRequiringJETQuirks)
    If TheValue = TablesRequiringJETQuirks(Count,0) Then

     Dim Catalog, Table, Column

     Set Catalog = server.CreateObject("ADOX.Catalog")
     Set Catalog.ActiveConnection = Database
     Set Table = Catalog.Tables(TheValue)


     For Each Column In Table.Columns

      Dim FieldProperties
      FieldProperties = TablesRequiringJETQuirks(Count,1)

      If FieldProperties(0) = Column.Name Then
       If Len(FieldProperties(1)) > 0 Then
        Column.Properties("Default") = FieldProperties(1)
       End If

       If Len(FieldProperties(2)) > 0 Then
        If FieldProperties(2) = "True" Then
         Column.Properties("Jet OLEDB:Allow Zero Length").Value = True

        ElseIf FieldProperties(2) = "False" Then
         Column.Properties("Jet OLEDB:Allow Zero Length").Value = False
        End If
 

      End If

       If Len(FieldProperties(3)) > 0 Then
        If FieldProperties(3) = "True" Then
         'Column.Properties("UnicodeCompression").Value = True

        ElseIf FieldProperties(3) = "False" Then
         'Column.Properties("Jet OLEDB:Unicode Compression").Value = False
        End If
 

      End If

      End If
     Next
          
    Set Table = Nothing
    Set Catalog = Nothing

    End If
   Next


  End Sub


  Dim SilentlyPerformedAlter

  '-- This is where we return any ALTER TABLE statements --'
  Function FindAlterString(Field)

   Dim RenamedFields(1,2)
   RenamedFields(0,0) = "BannedIP"
   RenamedFields(0,1) = "Date"
   RenamedFields(0,2) = "BannedDate"

   RenamedFields(1,0) = "Comments"
   RenamedFields(1,1) = "Date"
   RenamedFields(1,2) = "CommentedDate"

   For LoopCount = 0 To UBound(RenamedFields)
    If RenamedFields(LoopCount,0) & "." & RenamedFields(LoopCount,2) = Field Then
     Response.Write "<!-- " & RenamedFields(LoopCount,0) & "." & RenamedFields(LoopCount,1) & " should be " & RenamedFields(LoopCount,2) & "-->" & VbCrlf

     Dim Catalog, Table, Column

     Set Catalog = server.CreateObject("ADOX.Catalog")
     Set Catalog.ActiveConnection = Database

     Set Table = Catalog.Tables(RenamedFields(LoopCount,0))
      On Error Resume Next
       Table.Columns(RenamedFields(LoopCount,1)).Name = RenamedFields(LoopCount,2)
       If Err <> 0 Then 
        Response.Write "<!-- Error with ADOX rename " & Err.Description & " -->" & VbCrlf
       Else
        SilentlyPerformedAlter = True
       End If


      On Error GoTo 0


     Set Table = Nothing

     Set Catalog = Nothing

     Exit Function
    End If
   Next

   Dim AlterStrings(6,2)
   AlterStrings(0,0) = "Disclaimer"
   AlterStrings(0,1) = "LastModified"
   AlterStrings(0,2) = "ADD LastModified DATETIME"

   AlterStrings(1,0) = "Data"
   AlterStrings(1,1) = "LastModified"
   AlterStrings(1,2) = "ADD LastModified DATETIME"

   AlterStrings(2,0) = "Data"
   AlterStrings(2,1) = "UTCTimeZoneOffset"
   AlterStrings(2,2) = "ADD UTCTimeZoneOffset VARCHAR(5)"

   AlterStrings(3,0) = "Data"
   AlterStrings(3,1) = "Enclosure"
   AlterStrings(3,2) = "ADD Enclosure VARCHAR(255)"

   AlterStrings(4,0) = "Data"
   AlterStrings(4,1) = "EntryPUK"
   AlterStrings(4,2) = "ADD EntryPUK LONG"

   AlterStrings(5,0) = "Comments"
   AlterStrings(5,1) = "UTCTimeZoneOffset"
   AlterStrings(5,2) = "ADD UTCTimeZoneOffset VARCHAR(5)"

   AlterStrings(6,0) = "Main"
   AlterStrings(6,1) = "LastModified"
   AlterStrings(6,2) = "ADD LastModified DATETIME"

   Dim Found
   Found = -1

   Dim LoopCount
   For LoopCount = 0 To UBound(AlterStrings)

   If AlterStrings(LoopCount,0) & "." & AlterStrings(LoopCount,1) = Field Then 
    Found = LoopCount
    Exit For
   End If

   Next

   If Found > -1 Then
    Dim TempString
    TempString = "ALTER TABLE " & AlterStrings(Found,0) & " " & AlterStrings(Found,2)
    Response.Write "<!-- " & TempString & "-->" & VbCrlf
    FindAlterString = TempString
   End If

  End Function

  '-- This is our general table checker, if we have allowed an upgrade we also call other methods and perform them --'
  Function IsTableOK(strClassString, strClassString2)

   'On Error Resume Next

    '-- Default Values --'
    LastTableError = ""
    IsTableOK = False
    Err = 0

    '-- Test The SQL --'
    Records.Open "Select " & strClassString2 & " FROM " & strClassString,Database, 1, 3
     LastTableError = Err    
    If Records.State = 1 Then Records.Close

     If 0 = Err Then
      If Request.Querystring("Upgrade") = "Allowed" Then PerformJETChanges(strClassString)
      IsTableOK = True
     Else

      '-- We're trying to upgrade? --'
      If Request.Querystring("Upgrade") = "Allowed" Then

       If LastTableError = "-2147217865" Then
        Dim CreationString

        On Error GoTo 0
         CreationString = FindCreationString(strClassString)
        On Error Resume Next

        If Len(CreationString) > 0 Then
         Err.Clear
         Database.Execute CreationString
        On Error GoTo 0

         Dim objErr
         For Each objErr In Database.Errors
          Response.Write "<p>Description: " & objErr.Description & "<br />"
          Response.Write "Help context: " & objErr.HelpContext & "<br />"
          Response.Write "Help file: " & objErr.HelpFile & "<br />"
          Response.Write "Native error: " & objErr.NativeError & "<br />"
          Response.Write "Error number: " & objErr.Number & "<br />"
          Response.Write "Error source: " & objErr.Source & "<br />"
          Response.Write "SQL state: " & objErr.SQLState & "<br />"
          Response.Write "</p>"
         Next

         If 0 = Err Then 
          IsTableOK = True
	  PerformJETChanges(strClassString)
         End If

        End If
       ElseIf LastTableError = "-2147217904" Then
        
        Dim AlterString

        On Error GoTo 0
         AlterString = FindAlterString(strClassString & "." & strClassString2)
        On Error Resume Next

        If Len(AlterString) > 0 Then
         Err.Clear
         Database.Execute AlterString
        On Error GoTo 0

         For Each objErr In Database.Errors
          Response.Write "<p>Description: " & objErr.Description & "<br />"
          Response.Write "Help context: " & objErr.HelpContext & "<br />"
          Response.Write "Help file: " & objErr.HelpFile & "<br />"
          Response.Write "Native error: " & objErr.NativeError & "<br />"
          Response.Write "Error number: " & objErr.Number & "<br />"
          Response.Write "Error source: " & objErr.Source & "<br />"
          Response.Write "SQL state: " & objErr.SQLState & "<br />"
          Response.Write "</p>"
         Next

         If 0 = Err Then
          IsTableOK = True
         End If
       ElseIf SilentlyPerformedAlter Then
        IsTableOK = True
       End If
      End If
     End If
     PreventStep = True
    End If

   On Error GoTo 0

  End Function

  Response.Write " <p style=""text-align: center"">This page informs you of any problems with tables/fields in the database.<br/>" & VbCrlf
  Response.Write " If a table/field is missing, BlogX will be <b>unstable</b> or refuse to run.</p>" & VbCrlf

  '-- We want to be able to place the upgrade button before the table, so only buffer the table into a variable --'
  Dim TheBufferedOutput

  TheBufferedOutput = " <table border=""1"" cellpadding=""1"" cellspacing=""0"" style=""border-collapse: collapse; align: center; margin: 0 auto; width:60%"">" & VbCrlf
  TheBufferedOutput = TheBufferedOutput & "  <tr>" & VbCrlf
  TheBufferedOutput = TheBufferedOutput & "   <th class=""black"">Table</th>" & VbCrlf
  TheBufferedOutput = TheBufferedOutput & "   <th class=""black"">Fields</th>" & VbCrlf
  TheBufferedOutput = TheBufferedOutput & "   <th class=""black"">Status</th>" & VbCrlf
  TheBufferedOutput = TheBufferedOutput & "  </tr>" & VbCrlf

  Dim TheTables(19,3), LastTableError

  '-- The Tables --'
  TheTables(0,0)  = "BannedIP"
  TheTables(0,1)  = "IP, BannedDate"

  TheTables(1,0)  = "BannedLoginIP"
  TheTables(1,1)  = "BannedIP, LoginFailCount, LastLoginFail"

  TheTables(2,0)  = "Comments"
  TheTables(2,1)  = "CommentID, EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK"

  TheTables(3,0)  = "Comments_Unvalidated"
  TheTables(3,1)  = "CommentID, EntryID, Name, Email, Homepage, Content, CommentedDate, UTCTimeZoneOffset, IP, Subscribe, PUK"
   
  TheTables(4,0)  = "Config"
  TheTables(4,1)  = "ConfigID, AdminUsername, AdminPassword, CookieName, Copyright, EnableComments, EnableEmail, EnableMainPage, EmailAddress, EmailComponent, EmailServer, EntriesPerPage, Polls, ReaderPassword, ShowCategories, SiteName, SiteDescription, SiteSubTitle, SortByDay, Template, BackgroundColor, ShortTimeFormat, Logging"

  TheTables(5,0)  = "Data"
  TheTables(5,1)  = "RecordID, Title, Text, Category, Password, Day, Month, Year, Time, UTCTimeZoneOffset, Comments, StopComments, Enclosure, EntryPUK, LastModified"

  TheTables(6,0)  = "Disclaimer"
  TheTables(6,1)  = "DisclaimerID, DisclaimerText, LastModified"

  TheTables(7,0)  = "Draft"
  TheTables(7,1)  = "Title, Text"

  TheTables(8,0)  = "FileExtensions"
  TheTables(8,1)  = "AllowedExtension"

  TheTables(9,0)  = "Links"
  TheTables(9,1)  = "LinkID, LinkName, LinkURL, LinkRSS, LinkType"

  TheTables(10,0)  = "Mail_Unvalidated"
  TheTables(10,1)  = "RecordID, FromEmail, Subject, Body, IP, PUK"

  TheTables(11,0)  = "MailingList"
  TheTables(11,1)  = "EmailID, SubscriberAddress, SubscriberIP, PUK, Active"

  TheTables(12,0)  = "Main"
  TheTables(12,1)  = "MainID, MainText, LastModified"

  TheTables(13,0)  = "NotFound"
  TheTables(13,1)  = "ID, URL, ReferringPage, ErrorCount"

  TheTables(14,0)  = "PingBack"
  TheTables(14,1)  = "PingBackID, EntryID, SourceURI, Error, IP"

  TheTables(15,0) = "Poll"
  TheTables(15,1) = "PollID, Content, Des1, Op1, Des2, Op2, Des3, Op3, Des4, Op4, Total"

  TheTables(16,0) = "Refer"
  TheTables(16,1) = "ReferID, ReferURL, ReferHits"

  TheTables(17,0) = "ScriptRefer"
  TheTables(17,1) = "ReferID, ReferURL, Approved, ReferHits, IP"

  TheTables(18,0) = "UserDictionary"
  TheTables(18,1) = "WordID, Word"

  TheTables(19,0) = "Votes"
  TheTables(19,1) = "VoteID, PollID, IP, Option"

  For I = 0 to UBound(TheTables)

   TheBufferedOutput = TheBufferedOutput & " <tr>" & VbCrlf

   '-- Table Name --'
   TheBufferedOutput = TheBufferedOutput & "  <td>" & TheTables(i,0) & "</td>" & VbCrlf

   '-- As the fields are of variable length we generate a new array --'
   Dim TheFields
   TheFields = Split(TheTables(I,1),", ")

   TheBufferedOutput = TheBufferedOutput & "  <td>"

   Dim Count
   For Count = 0 To UBound(TheFields)
    If NOT IsTableOK(TheTables(i,0), TheFields(Count)) Then 
     TheBufferedOutput = TheBufferedOutput & "<span style=""color: purple; font-weight: bold"">" & Trim(TheFields(Count)) & "</span>"
    Else
     TheBufferedOutput = TheBufferedOutput & Trim(TheFields(Count))
    End If

    If Count < UBound(TheFields) Then TheBufferedOutput = TheBufferedOutput & ", "
   Next

   TheBufferedOutput = TheBufferedOutput & "</td>" & VbCrlf

   Dim NeedsUpgrade

    If IsTableOK(TheTables(i,0), TheTables(i,1)) Then
     TheBufferedOutput = TheBufferedOutput & "  <td class=""installed"">PASSED" 

     '-- Just check we don't have any additional fields --'
     Records.Open "Select * FROM " & TheTables(i,0),Database, 1, 3
     If (Records.Fields.Count-1) <> UBound(TheFields) Then TheBufferedOutput = TheBufferedOutput & "<br/><span style=""font-size: xx-small"">(" & ((Records.Fields.Count-1) - UBound(TheFields)) &" extra field(s))</span>"
     Records.Close
  
     TheBufferedOutput = TheBufferedOutput & "</td>" & VbCrlf

    Else
     TheBufferedOutput = TheBufferedOutput & "  <td class=""notinstalled"">FAILED<br/>" & VbCrlf
     TheBufferedOutput = TheBufferedOutput & "   <span class=""smallpurple"">"

     If LastTableError = "-2147217865" Then 
      TheBufferedOutput = TheBufferedOutput & "(Table does not exist)"
      NeedsUpgrade = True
     ElseIf LastTableError = "-2147217904" Then
      TheBufferedOutput = TheBufferedOutput & "(A field does not exist)"
      NeedsUpgrade = True
     Else
      TheBufferedOutput = TheBufferedOutput & LastTableError
     End If

     TheBufferedOutput = TheBufferedOutput & "   </span>" & VbCrlf

     TheBufferedOutput = TheBufferedOutput & "  </td>" & VbCrlf

   End If

   TheBufferedOutput = TheBufferedOutput & "</tr>" & VbCrlf

  Next

  TheBufferedOutput = TheBufferedOutput & "  </table>" & VbCrlf

  If Request.Querystring("Upgrade") = "Allowed" Then
   Response.Write " <p><div style=""text-align: center; background: purple; color: white; width:30%; align:center; margin: 0 auto;"">Setup attempted to upgrade the database.<br/>Please check below for the results.<br/>" & VbCrlf
   Response.Write " </div><br/></p>" & VbCrlf & VbCrlf
  ElseIf NeedsUpgrade Then
   '-- Beta doesn't have a workable upgrade script!! --'
   'Response.Write " <p><div style=""text-align: center; background: orange; width:30%; align:center; margin: 0 auto;"">Setup detected a few missing fields/tables. If this is an old database you can upgrade it?<br/>" & VbCrlf
   'Response.Write " <input type=""button"" value=""Upgrade Database"" onclick=""if (confirm('Are you *sure* you want to attempt to upgrade this database?\n\nPlease ensure you have kept a backup copy of the old database in case this fails.')) javascript:window.location='?Step=" & Step & "&amp;Upgrade=Allowed';return false;""/></div><br/></p>" & VbCrlf & VbCrlf
  End If

  '-- The buffered table --'
  Response.Write TheBufferedOutput

  Response.Write " <p style=""text-align: center; font-size: small"">Note: This page does not check the field sizes or their types.</p>" & VbCrlf

 Set Records = Nothing

 ElseIf LastTableError <> "3709" Then
  Response.Write LastTableError
 End If

 If Database.State = 1 Then Database.Close
 Set Database = Nothing

'-- Step 4 --'
Case 3

 Response.Write " <h3 style=""text-align: center"">Step 4 - Finished</h3>" & VbCrlf & VbCrlf

 Response.Write " <p style=""text-align: center"">There are no problems with the database or your setup.<br/>Make sure to edit Includes/Config.asp and specify the database file (if you have not already) as listed in the documentation.</p>" & VbCrlf
 Response.Write " <p style=""text-align: center"">I hope you enjoy using BlogX.</p>" & VbCrlf & VbCrlf


End Select

Response.Write "<form style=""display: inline"" method=""get"" action=""Setup.asp"">" & VbCrlf
 Response.Write " <div style=""text-align: center"">"
 Response.Write " <input type=""hidden"" name=""StepNoJS"" value=""" & Step & """/>" & VbCrlf
 If Step > 0 Then Response.Write " <input name=""Direction"" type=""submit"" value=""<--Back"" onclick=""javascript:window.location='?Step=" & Step - 1 & "';return false;""/>"

 If (Step < 3) Then 
  Response.Write " <input name=""Direction"" "
  If PreventStep = True Then Response.Write "disabled=""disabled"""
  Response.Write "type=""submit"" value=""Next-->"" onclick=""javascript:window.location='?Step=" & Step + 1 & "';return false;""/>"
 End If

Response.Write " </div>" & VbCrlf
Response.Write "</form>" & VbCrlf

%>
</body>

</html>
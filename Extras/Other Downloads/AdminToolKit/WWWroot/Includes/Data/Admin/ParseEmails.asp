<!-- #INCLUDE FILE="../Includes/Header.asp" -->
<!-- #INCLUDE FILE="Admin.asp" -->
<DIV id=content>
<%
' Create a filesystem object
Dim MailMessage, FilesExist

Set MailMessage = Server.CreateObject("Scripting.FileSystemObject")

Path = "C:\Program Files\Argo Software Design\Mail Server\_users\_nodomain\Blog\inbox\"

On Error Resume Next

Set Folder = MailMessage.GetFolder(Path)  
Set Files = Folder.files

If err <> "" Then 
Set MailMessage = Nothing
Response.Redirect "../Default.asp"
ENd If

For Each File in Files 

If MailMessage.GetExtensionName(Path & File.Name) = "eml" Then

FilesExist = True
FilePath = Path & File.Name

'---Map The Physical System Path---'

	' Get a handle to the file
	Dim Email	
	Set Email = MailMessage.GetFile(FilePath)

	' Open the file
	Dim EmailStream

        ' Read the file line by line
	Set EmailStream = Email.OpenAsTextStream(1, -2)

        Dim ReadText
        Dim SentDate, SentTime   
        
        Response.Write "<Li> FileName : " & File.Name & "</Li><br>"  
        
	Do While Not EmailStream.AtEndOfStream

                ReadText = EmailStream.Readline

                If InStr(ReadText,"Subject: [BlogX] ") <> 0 Then
                Authorised = True
                Subject = Replace(ReadText,"Subject: [BlogX] ","")
                Response.Write "<Li> Subject : " & Subject & "</Li><br>"
                End If

                If Instr(ReadText,"Date:") <> 0 Then 
                SentDate = Replace(ReadText,"Date: ","")
                Length = Len(SentDate)

                '-- Take Off Day Name --'
                SentDate = Right(SentDate,Length-5)


                '-- Take Off GMT Markup --'
                SentDate = Left(SentDate,Length-11)

                SpacePos = InStrRev(SentDate," ")

                Length = Len(SentDate)
                SentTime = Right(SentDate,Length-SpacePos)
                SentDate = Left(SentDate,SpacePos)

                Response.Write "<Li> Date : " & SentDate & "</Li><br>"
                Response.Write "<Li> Time : " & SentTime & "</Li><br>"
                End If
                              
                If InStr(ReadText,"Category: ") <> 0 Then
                EntryCat = Replace(ReadText,"Category: ","")
                Response.Write "<Li> Category : " & EntryCat & "</Li><br>"
                End If

                If (InStr(ReadText,":") = 0) AND (InStr(ReadText,"=") = 0) AND (Subject <> "") Then
                Body = Body & VbCrlf & ReadText
                Response.Write VbCrlf & "<br>" & ReadText
                End If

                If Instr(ReadText,"This is a multi-part message in MIME format.") <> 0 Then Authorised = False

	Loop   

        EmailStream.Close
	Set EmailStream = nothing
	
'### Create a connection odject ###

If Authorised = True Then

'### Filter & Clean ###
EntryCat = Replace(EntryCat,"'","&#39;")
EntryCat = Replace(EntryCat," ","%20")

'### Open The Records Ready To Write ###
Records.CursorType = 2
Records.LockType = 3
Records.Open "SELECT * FROM Data", Database
Records.AddNew
Records("Title") = Subject
Records("Text") = Body
Records("Category") = EntryCat

Records("Day") = Day(SentDate)
Records("Month") = Month(SentDate)
Records("Year") = Year(SentDate)
Records("Time") = SentTime
Records.Update

'#### Close Objects ###
Records.Close

Response.Write "<p align=""Center"">Entry Submission Successful</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

Else

Response.Write "<p align=""Center"">Invalid Auth Details (Or <b>NOT</b> in ""Plain Text"" Format)</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"

End If

If Authorised = True Then MailMessage.DeleteFile(FilePath)

End If
Next  

Set MailMessage = nothing
Set Folder = Nothing
Set Files = Nothing

If FilesExist <> True Then
Response.Write "<p align=""Center"">No New E-mails</p>"
Response.Write "<p align=""Center""><a href=""" & SiteURL & PageName & """>Back</font></a></p>"
End If
%>
</Div>
<!-- #INCLUDE FILE="../Includes/Nav.asp" -->
<!-- #INCLUDE FILE="../Includes/Footer.asp" -->
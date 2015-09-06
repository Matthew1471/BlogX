<% OPTION EXPLICIT %>
<!-- #INCLUDE FILE="Includes/Header.asp" -->
<DIV id=content>

<!--- Start Installed Themes Information --->
<DIV class=entry>
<H3 class=entryTitle>About The Themes</H3>
<DIV class=entryBody>
<p>The following theme templates are installed :</P>
<p>
<%
	Dim Folder, Folders, ThemePath, Status, JPEG, GIF

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(Server.MapPath("Templates/"))
        Set Folders = Folder.SubFolders

        For Each Folder in Folders 
        Response.Write Folder.Name
        If Template = Folder.Name Then Response.Write " (<b>Current</b>)"
        Response.Write "<br>"
        Next
%>
</p>
</Div></Div>

<!--- Start Make Your Own Information --->
<DIV class=entry>
<H3 class=entryTitle>About Making Themes / Legacy Mode</H3>
<DIV class=entryBody>
<p>For instructions (& Samples) on how to make your own theme, click <a href="Templates/BlankTemplate.zip">here</a></P>
<p>To see what a typical SimpleGeek.com BlogX page looked like, <a href="?LegacyMode">tag ?LegacyMode on to the URL of a page</a>.</P>
</Div></Div>


<!--- Start Individual Breakdown --->
<%
        For Each Folder in Folders

'---Map The Physical System Path---'
ThemePath = Server.MapPath("Templates") & "\" & Folder.Name 
Status = ""

If FSO.FileExists(Server.MapPath("Templates") & "\" & Folder.Name & "\Info.txt") Then 


	' Get a handle to the file
	Dim ThemeFile	
	set ThemeFile = FSO.GetFile(Server.MapPath("Templates") & "\" & Folder.Name & "\Info.txt")

	' Open the file
	Dim ThemeTextStream

        ' Read the file line by line
	Set ThemeTextStream = ThemeFile.OpenAsTextStream(1, -2)
        Dim ThemeName
        Dim ThemeDesc
        Dim AuthorName, AuthorURL

	Do While Not ThemeTextStream.AtEndOfStream
		If ThemeTextStream.Readline = "*" Then
                ThemeName = ThemeTextStream.Readline
                ThemeDesc = ThemeTextStream.Readline
                
 		If ThemeTextStream.Readline = "*" Then
                AuthorName = ThemeTextStream.Readline
                AuthorURL  = ThemeTextStream.Readline
                Else
                AuthorName = ""
                AuthorURL = ""
                Status = "No Author Information"
                End If

                Else
                ThemeName = Folder.Name
                ThemeDesc = ""
                AuthorURL = ""
                AuthorName = ""
                Status = "No Theme Information (File Broken)"
                End If
	Loop
        
        ThemeTextStream.Close
	Set ThemeTextStream = nothing
	Set ThemeFile = Nothing
	
Else
        ThemeName = Folder.Name
        ThemeDesc = ""
        AuthorURL = ""
        AuthorName = ""
	Status = "No Theme Information (File Not Found)"
End If
%>
<!--- Start Information On <%=ThemeName%> --->
<DIV class=entry>
<H3 class=entryTitle>About <A Href="?Theme=<%=Folder.Name%>"><%=ThemeName%></A><%If Template = Folder.Name Then Response.Write " (<b>Current</b>)"%></H3>
<DIV class=entryBody>
<p><b>Name</b> : <%=ThemeName%></p>
<% 
If ThemeDesc <> "" Then Response.Write "<p><b>Description</b> : " & ThemeDesc & "</p>"

If FSO.FileExists(ThemePath & "\Background.jpg") Then JPEG = True Else JPEG = False
If FSO.FileExists(ThemePath & "\Background.gif") Then GIF = True Else GIF = False

If (JPEG <> False) Or (GIF <> False) Then Response.Write "<p><b>Background</b> :<br>"

If JPEG = True Then Response.Write "<A Href=""Templates\" & Folder.Name & "\Background.jpg""><Img Width=""100"" Height=""100"" Src=""Templates\" & Folder.Name & "\Background.jpg""></a></p>"

If GIF = True Then Response.Write "<A Href=""Templates\" & Folder.Name & "\Background.Gif""><Img Width=""100"" Height=""100"" Src=""Templates\" & Folder.Name  & "\Background.gif""></a></p>"

If AuthorName <> "" Then Response.Write "<p><b>Author Information</b> : "
If AuthorURL <> "" Then Response.Write "<a href=""" & AuthorURL & """>" 
If AuthorName <> "" Then Response.Write AuthorName 
If AuthorURL <> "" Then Response.Write "</a><br></p>"

If Status <> "" Then Response.Write "<p> <b>Error</b> : " & Status & "</p>"
%>
<p>
</Div></Div>
<%
Next
Set FSO = Nothing
Set Folder = Nothing
Set Folders = Nothing
%>
</DIV>
<!-- #INCLUDE FILE="Includes/NAV.asp" -->
<!-- #INCLUDE FILE="Includes/Footer.asp" -->
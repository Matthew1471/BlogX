<!-- #INCLUDE FILE="Includes/Config.asp" -->
<%
dim x(1000), FSO, Folder, idx, File, WhichNo, Folders, Page

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Folder = FSO.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") & Root)

Set Folders = Folder.SubFolders

'Step through the files list, keeping track of
'the number of files....
idx=0
For Each Folder in Folders
  idx=idx+1
  x(idx)=Folder.name
Next

'Choose a random picture
Randomize Timer

Do Until (Page <> "Includes/") AND (Page <> "") AND (Count < 15)
Page = x(int(rnd()*idx)+1) & "/"
Count = Count + 1
Loop

'### Kill objects ###
set FSO = Nothing
set Folder = Nothing
set Folders = Nothing

'Send To The New Page
Response.Redirect(Page)
%>
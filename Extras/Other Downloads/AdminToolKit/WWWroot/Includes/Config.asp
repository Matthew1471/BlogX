<%
'***********************************************'
' AdminToolkit Settings
'***********************************************'
Dim AppPath, DatabasePath
Dim AdminUsername, AdminPassword 
Dim Register, Company, Version
Dim DropDownBoxMode, MyBlog

AppPath = "C:\Inetpub\wwwroot\Blogs\"
DatabasePath = "C:\Inetpub\Database\Blogs\"

AdminUsername   = "admin"
AdminPassword   = "letmein"

Company = "Matthew1471" '-- Change This To Reflect Your Name Or Company --'
Register = True
Version = "1.04"

DropDownBoxMode = False ' A Quick way for me to update Blogs (but not create)
MyBlog = "http://BlogX.co.uk"

'***********************************************'
' Viewer Settings
'***********************************************'
Dim DataFile, EnableGuestSignups
Dim EmailAddress, EmailComponent, EmailServer
Dim StatsDatabase, StatsRecords, Blog, Hits, Root, TimeFormat

DataFile = "C:\Inetpub\database\AdminToolKit.mdb"

EmailAddress = "webmaster@matthew1471.co.uk"
EmailComponent = "cdosys"
EmailServer = "SERVER"

Root = "/"
TimeFormat = True

EnableGuestSignups = 1 '-- Lets Guests Signup for blogs! --'
'***********************************************'

Sub WriteFooter() %>
<!--#INCLUDE FILE="Footer.asp"-->
<% End Sub %>
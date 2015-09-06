<%
'--- Open Database ---'
Set Database = Server.CreateObject("ADODB.connection")
'Database.Open  "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=" & DataPassword & "; DBQ=" & DataFile
'Database.Open  "DSN=BlogX;"
Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Datafile & ";"

'--- Open Recordset ---'
set Records = Server.CreateObject("ADODB.recordset")
    Records.Open "SELECT * FROM Config",Database, 1, 3

    If NOT Records.EOF Then
	'Read in the configuration details from the recordset
	AdminUsername   = Records("AdminUsername")
	AdminPassword   = Records("AdminPassword")
	CookieName      = Records("CookieName")
	Copyright       = Records("Copyright")

        On Error Resume Next
	EnableEmail     = Records("EnableEmail")
	EmailAddress    = Records("EmailAddress")
        EmailServer     = Records("EmailServer")
        EmailComponent  = Records("EmailComponent")
        On Error Goto 0

        EnableComments  = Records("EnableComments")
	EntriesPerPage  = Records("EntriesPerPage")
	Polls           = Records("Polls")
	ShowCat         = Records("ShowCategories")
        ShowMonth       = True
	SiteName        = Records("SiteName")
	SiteDescription = Records("SiteDescription")
        SiteSubTitle    = Records("SiteSubTitle")
        SortByDay       = Records("SortByDay")
        BackgroundColor = Records("BackgroundColor")
        TimeFormat      = Records("12HourTimeFormat")
        Logging         = Records("Logging")

        On Error Resume Next
        EnableMainPage  = Records("EnableMainPage")
        ReaderPassword  = Records("ReaderPassword")
        Template        = Records("Template")
        On Error Goto 0
    End If

    Records.Close

If EnableMainPage <> True Then PageName = "Default.asp" Else PageName = "Main.asp"
%>
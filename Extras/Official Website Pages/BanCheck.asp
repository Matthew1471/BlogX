<%
' --------------------------------------------------------------------------
'¦Introduction : Ban Check Service.                                         ¦
'¦Purpose      : Allows external websites to query our ban database.        ¦
'¦Used By      : Matthew1471.co.uk.                                         ¦
'¦Requires     : Includes/Config.asp.                                       ¦
'---------------------------------------------------------------------------

OPTION EXPLICIT
%>
<!-- #INCLUDE FILE="../Includes/Config.asp" -->
<%           
Dim CheckIP       '-- The IP address of the user we want to check the ban of --'
Dim CheckProxy    '-- The IP address of the proxy we want to check the ban of --'

'-- Filter and clean --'
CheckIP = Request.Querystring("IP")
CheckIP = Replace(CheckIP,"'","")

CheckProxy = Request.Querystring("Proxy")
CheckProxy = Replace(CheckProxy,"'","")

'-- Check if they are banned (also check the proxy) --'
Records.Open "SELECT IP FROM BannedIP WHERE IP='" & CheckIP & "' OR IP='" & CheckProxy & "';",Database, 0, 1
 Dim Banned
 Banned = NOT Records.EOF
Records.Close

Response.Write Banned

'--- Close The Database ---'
Database.Close
Set Records = Nothing
Set Database = Nothing
%>
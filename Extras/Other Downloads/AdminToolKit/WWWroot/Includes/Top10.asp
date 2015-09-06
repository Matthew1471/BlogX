<%
'### Create a connection odject ###
Set StatsDatabase = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
StatsDatabase.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set StatsRecords = Server.CreateObject("ADODB.Recordset")
    StatsRecords.CursorLocation = 3

'### Open The Records Ready To Write ###
StatsRecords.Open "SELECT * FROM Top10 ORDER BY Hits DESC;", StatsDatabase
StatsRecords.PageSize = 10

Do Until (StatsRecords.EOF OR StatsRecords.AbsolutePage <> 1)

Set Blog = StatsRecords("Blog")
Set Hits = StatsRecords("Hits")
%>
                    <!--msimagelist--><tr>
                      <!--msimagelist--><td valign="baseline" width="42"><img src="<%=Root%>Includes/Images/eBlog.gif"></td>
                      <td valign="top" width="100%"><A href="<%=Root%><%=Blog%>/" Title="<%=Hits%> Hits"><%=Blog%></A>
                      <!--msimagelist--></td>
                    </tr>
<%
StatsRecords.MoveNext
Loop

StatsRecords.Close
StatsDatabase.Close

'#### Close Objects ###	
Set StatsDatabase = Nothing
Set StatsRecords = Nothing
%>
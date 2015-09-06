<!-- #INCLUDE FILE="Includes/Header.asp" -->
      <td width="934" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">Blog Ranking On <%=Domain%></font></b></td>
      <td width="241" bgcolor="#1843C4" height="28"><b><font color="#FFFFFF" face="Verdana" size="2">News</font></b></td>
    </tr>
    <tr>
      <!--- Content --->
      <td width="934" bgcolor="#FFFFFF" height="317" rowspan="3" valign="top" style="PADDING-LEFT: 5px; PADDING-TOP: 10px;">

<!--msimagelist--><table border="0" cellpadding="0" cellspacing="0" width="25%" align="center">
<%
'### Create a connection odject ###
Set StatsDatabase = Server.CreateObject("ADODB.Connection")

'### Set an active connection to the Connection object ###
StatsDatabase.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=; DBQ=" & DataFile

'### Create a recordset object ###
Set StatsRecords = Server.CreateObject("ADODB.Recordset")

'### Open The Records Ready To Write ###
StatsRecords.Open "SELECT * FROM Top10 ORDER BY Hits DESC;", StatsDatabase

Do Until (StatsRecords.EOF)

Set Blog = StatsRecords("Blog")
Set Hits = StatsRecords("Hits")
%>
                    <!--msimagelist--><tr>
                      <!--msimagelist--><td valign="baseline" width="42"><img src="Includes/Images/eBlog.gif"></td>
                      <td valign="top" width="100%"> <A href="<%=Blog%>/" Title="<%=Hits%> Hits"><%=Blog%></A>
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
<!--msimagelist--></table><br>

<center><small>Note : Blogs which have recieved less than 1 Hit will not be shown</small></center>

<% If Request.ServerVariables("HTTP_REFERER") <> "" Then %>
<br><center><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><< Back</a></center>
<% End If%>

      </td>
      <!--- End Of Content -->
<% WriteFooter %>
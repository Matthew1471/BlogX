<% OPTION EXPLICIT

Dim DataFile, Database

'-- !! CHANGE THIS TO YOUR DATABASE PATH !! --'
'DataFile = "C:\Inetpub\Database\BlogX.mdb"
DataFile = "D:\Inetpub\wwwroot\Temp\BlogX\BlogX(new).mdb"

'-- We default to Step 0 --'
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

<head>
 <title>Show Database Schema</title>
 <style type="text/css">
  .header { text-align: center; }
 </style>
</head>

<body>
 <h1 class="header">JET Database</h1>
<%
 Response.Write " <h3 style=""text-align: center"">Database Schema</h3>" & VbCrlf & VbCrlf

 Set Database = Server.CreateObject("ADODB.connection")

  On Error Resume Next
   Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Datafile & ";"

   If Err <> 0 Then

    '-- We specified a wrong path to the database --'
    If Err = -2147467259 Then
     '-- Try and guess where the database folder might be --'
     If Instr(1, Server.MapPath("."),"wwwroot", 1) Then

      '-- One last chance --'
      Err.Clear

      Dim GuessedPath
      GuessedPath = Left(Server.MapPath("."),Instr(1,Server.MapPath("."),"wwwroot",1)-1) & "Database\BlogX.mdb"
      
      Database.Open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GuessedPath & ";"

      If Err = 0 Then
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

  Function ColumnType(TheValue)

  'http://msdn.microsoft.com/en-us/library/ms806221.aspx
  Dim TheDataTypes(39,2)
  TheDataTypes(0,0) = "AdArray"
  TheDataTypes(0,1) = "0x2000"
  TheDataTypes(0,2) = "(Does not apply to ADOX.) A flag value, always combined with another data type constant, that indicates an array of that other data type."

  TheDataTypes(1,0) = "adBigInt"
  TheDataTypes(1,1) = 20
  TheDataTypes(1,2) = "Indicates an eight-byte signed integer (DBTYPE_I8)."

  TheDataTypes(2,0) = "adBinary"
  TheDataTypes(2,1) = 128
  TheDataTypes(2,2) = "Indicates a binary value (DBTYPE_BYTES)."

  TheDataTypes(3,0) = "adBoolean"
  TheDataTypes(3,1) = 11
  TheDataTypes(3,2) = "Indicates a boolean value (DBTYPE_BOOL)."

  TheDataTypes(4,0) = "adBSTR"
  TheDataTypes(4,1) = 8
  TheDataTypes(4,2) = "Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR)."

  TheDataTypes(5,0) = "adChapter"
  TheDataTypes(5,1) = 136
  TheDataTypes(5,2) = "Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER)."

  TheDataTypes(6,0) = "adChar"
  TheDataTypes(6,1) = 129
  TheDataTypes(6,2) = "Indicates a string value (DBTYPE_STR)."

  TheDataTypes(7,0) = "adCurrency"
  TheDataTypes(7,1) = 6
  TheDataTypes(7,2) = "Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000."

  TheDataTypes(8,0) = "adDate"

  TheDataTypes(8,1) = 7
  TheDataTypes(8,2) = "Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day."

  TheDataTypes(9,0) = "adDBDate"
  TheDataTypes(9,1) = 133
  TheDataTypes(9,2) = "Indicates a date value (yyyymmdd) (DBTYPE_DBDATE)."

  TheDataTypes(10,0) = "adDBTime"
  TheDataTypes(10,1) = 134
  TheDataTypes(10,2) = "Indicates a time value (hhmmss) (DBTYPE_DBTIME)."

  TheDataTypes(11,0) = "adDBTimeStamp"
  TheDataTypes(11,1) = 135
  TheDataTypes(11,2) = "Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP)."

  TheDataTypes(12,0) = "adDecimal"
  TheDataTypes(12,1) = 14
  TheDataTypes(12,2) = "Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL)."

  TheDataTypes(13,0) = "adDouble"
  TheDataTypes(13,1) = 5
  TheDataTypes(13,2) = "Indicates a double-precision floating-point value (DBTYPE_R8)."

  TheDataTypes(14,0) = "adEmpty"
  TheDataTypes(14,1) = 0
  TheDataTypes(14,2) = "Specifies no value (DBTYPE_EMPTY)."

  TheDataTypes(15,0) = "adError"
  TheDataTypes(15,1) = 10
  TheDataTypes(15,2) = "Indicates a 32-bit error code (DBTYPE_ERROR)."

  TheDataTypes(16,0) = "adFileTime"
  TheDataTypes(16,1) = 64
  TheDataTypes(16,2) = "Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME)."

  TheDataTypes(17,0) = "adGUID"
  TheDataTypes(17,1) = 72
  TheDataTypes(17,2) = "Indicates a globally unique identifier (GUID) (DBTYPE_GUID)."

  TheDataTypes(18,0) = "adIDispatch"
  TheDataTypes(18,1) = 9
  TheDataTypes(18,2) = "Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results."

  TheDataTypes(19,0) = "adInteger"
  TheDataTypes(19,1) = 3
  TheDataTypes(19,2) = "Indicates a four-byte signed integer (DBTYPE_I4)."

  TheDataTypes(20,0) = "adIUnknown"
  TheDataTypes(20,1) = 13
  TheDataTypes(20,2) = "Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results."

  TheDataTypes(21,0) = "adLongVarBinary"
  TheDataTypes(21,1) = 205
  TheDataTypes(21,2) = "Indicates a long binary value."

  TheDataTypes(22,0) = "adLongVarChar"
  TheDataTypes(22,1) = 201
  TheDataTypes(22,2) = "Indicates a long string value."

  TheDataTypes(23,0) = "adLongVarWChar"
  TheDataTypes(23,1) = 203
  TheDataTypes(23,2) = "Indicates a long null-terminated Unicode string value."

  TheDataTypes(24,0) = "adNumeric"
  TheDataTypes(24,1) = 131
  TheDataTypes(24,2) = "Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC)."

  TheDataTypes(25,0) = "adPropVariant"
  TheDataTypes(25,1) = 138
  TheDataTypes(25,2) = "Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT)."

  TheDataTypes(26,0) = "adSingle"
  TheDataTypes(26,1) = 4
  TheDataTypes(26,2) = "Indicates a single-precision floating-point value (DBTYPE_R4)."

  TheDataTypes(27,0) = "adSmallInt"
  TheDataTypes(27,1) = 2
  TheDataTypes(27,2) = "Indicates a two-byte signed integer (DBTYPE_I2)."

  TheDataTypes(28,0) = "adTinyInt"
  TheDataTypes(28,1) = 16
  TheDataTypes(28,2) = "Indicates a one-byte signed integer (DBTYPE_I1)."

  TheDataTypes(29,0) = "adUnsignedBigInt"
  TheDataTypes(29,1) = 21
  TheDataTypes(29,2) = "Indicates an eight-byte unsigned integer (DBTYPE_UI8)."

  TheDataTypes(30,0) = "adUnsignedInt"
  TheDataTypes(30,1) = 19
  TheDataTypes(30,2) = "Indicates a four-byte unsigned integer (DBTYPE_UI4)."

  TheDataTypes(31,0) = "adUnsignedSmallInt"
  TheDataTypes(31,1) = 18
  TheDataTypes(31,2) = "Indicates a two-byte unsigned integer (DBTYPE_UI2)."

  TheDataTypes(32,0) = "adUnsignedTinyInt"
  TheDataTypes(32,1) = 17
  TheDataTypes(32,2) = "Indicates a one-byte unsigned integer (DBTYPE_UI1)."

  TheDataTypes(33,0) = "adUserDefined"
  TheDataTypes(33,1) = 132
  TheDataTypes(33,2) = "Indicates a user-defined variable (DBTYPE_UDT)."

  TheDataTypes(34,0) = "adVarBinary"
  TheDataTypes(34,1) = 204
  TheDataTypes(34,2) = "Indicates a binary value."

  TheDataTypes(35,0) = "adVarChar"
  TheDataTypes(35,1) = 200
  TheDataTypes(35,2) = "Indicates a string value."

  TheDataTypes(36,0) = "adVariant"
  TheDataTypes(36,1) = 12
  TheDataTypes(36,2) = "Indicates an Automation Variant (DBTYPE_VARIANT). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results."

  TheDataTypes(37,0) = "adVarNumeric"
  TheDataTypes(37,1) = 139
  TheDataTypes(37,2) = "Indicates a numeric value."

  TheDataTypes(38,0) = "adVarWChar"
  TheDataTypes(38,1) = 202
  TheDataTypes(38,2) = "Indicates a null-terminated Unicode character string."

  TheDataTypes(39,0) = "adWChar"
  TheDataTypes(39,1) = 130
  TheDataTypes(39,2) = "Indicates a null-terminated Unicode character string (DBTYPE_WSTR)."

  Dim Count
  For Count = 0 To UBound(TheDataTypes)
   If TheValue = TheDataTypes(Count,1) Then
    Dim Found
    Found = True
    Exit For
   End If
  Next

  If Found Then 
   ColumnType = "<a title=""" & TheDataTypes(Count,1) & " - " & TheDataTypes(Count,2) & """>" & TheDataTypes(Count,0) & "</a>"
  Else
   ColumnType = "Unknown (" & TheValue & ")"
  End If

  End Function

  '-- JET has a few extra quirks that can't be modified via SQL (http://www.codeguru.com/cpp/data/mfc_database/ado/article.php/c4343) --'
  Dim Catalog, Table, Column

  Set Catalog = server.CreateObject("ADOX.Catalog")
  Set Catalog.ActiveConnection = Database

  For Each Table In Catalog.Tables
   Response.Write "<h1>" & Table.Name & "</h1>" & VbCrlf

   For Each Column In Table.Columns
    Response.Write "<h3 style=""color: red"">" & Column.Name & " (Type: " & ColumnType(Column.Type) & ")</h3>" & VbCrlf
    Response.Write "<p>" & VbCrlf

    Dim Item
    For Each Item In Column.Properties
     Response.Write "<b>" & Item.Name & "</b>=" & Item & "<br/>" & VbCrlf
    Next

    Response.Write "-------------</p>" & VbCrlf

   Next
  Next

  Set Table = Nothing
  Set Catalog = Nothing

  Response.Write " <p style=""text-align: center"">This page informs you of all provider specific settings and tables/fields in the database.<br/>" & VbCrlf
  Response.Write " Not all of these are accessible via SQL.</p>" & VbCrlf

 If Database.State = 1 Then Database.Close
 Set Database = Nothing
%>
</body>

</html>
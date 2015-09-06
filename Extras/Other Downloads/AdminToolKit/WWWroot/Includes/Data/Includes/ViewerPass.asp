<%
If (ReaderPassword <> "") Then

Dim Page, Last, Length
Page = Request.Servervariables("Script_Name")

Last = InStrRev(Page,"/")
Length = Len(Page)

Page = Right(Page,Length - Last)

If Session("Reader") = False or IsNull(Session("Reader")) = True AND (Request.Cookies("Reader") <> "True") Then Response.Redirect "ReaderLogin.asp?" & Page
End If
%>
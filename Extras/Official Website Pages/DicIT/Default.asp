<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #INCLUDE FILE="Includes/Spell.asp" -->
<%
Sub SingleSorter( byRef arrArray )
    Dim row, j
    Dim StartingKeyValue, NewKeyValue, swap_pos

    For row = 0 To UBound( arrArray ) - 1
    'Take a snapshot of the first element
    'in the array because if there is a 
    'smaller value elsewhere in the array 
    'we'll need to do a swap.
        StartingKeyValue = arrArray ( row )
        NewKeyValue = arrArray ( row )
        swap_pos = row
	    	
        For j = row + 1 to UBound( arrArray )
        'Start inner loop.
            If arrArray ( j ) < NewKeyValue Then
            'This is now the lowest number - 
            'remember it's position.
                swap_pos = j
                NewKeyValue = arrArray ( j )
            End If
        Next
	    
        If swap_pos <> row Then
        'If we get here then we are about to do a swap
        'within the array.		
            arrArray ( swap_pos ) = StartingKeyValue
            arrArray ( row ) = NewKeyValue
        End If	
    Next
End Sub
%>
<html><body>
<h1>Bubble Sort for a One-Dimensional Array</h1>

<form method=post id=form1 name=form1>
  Enter a number of strings, each separated by a space (for example: <i>bob scott sue larry elvis</i>):<br>
  <textarea name="txtSearch" cols=50 rows=5><%=Request("txtSearch")%></textarea>
  <p><input type=submit value="Sort!" id=submit1 name=submit1>
</form>

<p><hr><P>

<% If Len(Request("txtSearch")) > 0 Then

      Dim aDigits, intDictSize, Count, x
      aDigits = split(Request("txtSearch"), " ")     
      x = LoadDictArray
      
      ReDim Preserve aDigits(Ubound(aDigits) + x)

      'For Count = 0 To (x - 1)
      'aDigits(Count) = strDictArray(Count)
      'Response.Write aDigits(Count) & "<br>" & VbCrlf
      'Response.Flush
      'Next

      'Display the unsorted array
      Response.Write "<b>Unsorted Array</b>: " & join(aDigits, ", ")

      SingleSorter aDigits
      
      'Display the sorted array
      Response.Write "<p><b>Sorted Array:</b> " & join(aDigits, ", ")
      
      Response.Write "<p><hr><p>"
   End If
%>
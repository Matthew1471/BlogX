<%
' Concatenate a variable number of text files into a single result file
'
' Example:
ConcatenateFiles (Server.MapPath("Includes\UserDictionary.txt")), (Server.MapPath("Includes\dict-large.txt")), (Server.MapPath("Includes\user-archive.txt"))

Sub ConcatenateFiles(ResultFile, SourceFile, SourceFile2)
    Dim FSO, fsSourceStream, fsResStream, i
    
    On Error Resume Next
    
    ' create a new file
    Set fsResStream = FSO.OpenTextFile(ResultFile, ForWriting, True)
    
    ' for each source file in the input array
    For i = 0 To UBound(SourceFiles)

        ' open the file in read mode
        Set fsSourceStream = FSO.OpenTextFile(SourceFiles(i), ForReading)
        ' add its content + a blank line to the result file
        fsResStream.Write fsSourceStream.ReadAll & vbCrLf
        ' close this source file
        fsSourceStream.Close

    Next
    
    fsResStream.Close
    Set fsResStream = Nothing
End Sub
%>
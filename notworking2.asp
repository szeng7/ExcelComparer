<html>
<Body>


<style>

table {
  overflow: hidden;
}

body,td {
    font-family: tahoma,verdana,arial,helvetica;
}
td { 
    color: black; 
    text-decoration: none; 
    background-color: #eeeeee;
}

tr:hover {
  text-shadow: 0.5px 0.5px grey;
}
</style>

<%
const numRowsToCheck = 4
Sub Pr(S)
    Response.Write S
    End Sub

Function excelCols(colNum)
    Dim iAlpha, iRemainder
    iAlpha = Int(colNum / 27)
    iRemainder = colNum - (iAlpha * 26)
    If iAlpha > 0 Then
        excelCols = Chr(iAlpha + 64)
    End If
    If iRemainder > 0 Then
        excelCols = excelCols & Chr(iRemainder + 64)
    End If
End Function

Dim I, numFields, file1rows, file2rows
'Number of rows to look ahead to decide if an insertion
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='1' style='background-color: black' ><tr><td style='background-color: #FFFFCC' ><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td>"&"<td style='background-color: #74A4BC'><b>"&Request.Form("file1")&" Contents</b></td><td style='background-color: #74A4BC'><b>"&Request.Form("file2")&" Contents</b></td>"&"<td style='background-color: #FFFFCC'><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td></tr>"
file1rows = Split(Request.Form("full1"), "\\\")
file2rows = Split(Request.Form("full2"), "\\\")
numFields = Request.Form("fields")
Dim J, insertion1, insertion2, rowNum1, rowNum2, offset1, offset2
rowNum1 = 1
rowNum2 = 1
offset1 = 0
offset2 = 0
Do While I + offset1 < Ubound(file1rows) And I + offset2 < Ubound(file2rows)
    rowNum1 = rowNum1 + 1
    rowNum2 = rowNum2 + 1
    insertion1 = 0
    insertion2 = 0
    If StrComp(file1rows(I+offset1), file2rows(I+offset2)) <> 0 Then
        For J = 1 to numRowsToCheck
            If I+offset1+J < Ubound(file1rows) Then 
                If StrComp(file1rows(I+offset1+J),file2Rows(I+offset2)) = 0 Then
                    insertion1 = 1
                End If
            End If
        Next
        If insertion1 = 0 Then 
            For J = 1 to numRowsToCheck
                If I+offset2+J<Ubound(file2rows) Then
                    If StrComp(file1rows(I+offset1),file2Rows(I+offset2+J)) = 0 Then
                        insertion2 = 1
                    End If
                End If
            Next
        End If
    'Evidence of an insertion if insertion = 1
    End If
    If insertion1 = 1 Then
        Pr "<tr><td style='width:50px; background-color: #FFFFCC' ></td><td  style='width:50px; background-color: #CCFFFF'>"&rowNum1&"</td><td style='width:250px;'>"&file1rows(I+offset1)&"</td>"
        Pr "<td style='width:250px;'>"&file2rows(I+offset2)&"</td><td style='width:50px; background-color: #FFFFCC'></td><td style='width:50px; background-color: #CCFFFF'>"&rowNum2&"</td></tr>"
        offset1 = offset1 + 1
        rowNum1 = rowNum1 + 1
    ElseIf insertion2 = 1 Then
        Pr "<tr><td  style='width:50px; background-color: #FFFFCC' ></td><td style='width:50px; background-color: #CCFFFF'>"&rowNum1&"</td><td style='width:250px;'>"&file1rows(I+offset1)&"</td>"
        Pr "<td style='width:250px;'>"&file2rows(I+offset2)&"</td><td style='width:50px; background-color: #FFFFCC'></td><td style='width:50px; background-color: #CCFFFF'>"&rowNum2&"</td></tr>"
        offset2 = offset2 + 1
        rowNum2 = rowNum2 + 1
    Else
        cells1 = Split(file1rows(I+offset1),"***")
        cells2 = Split(file2rows(I+offset2),"***")
        For J = 0 to Ubound(cells1)
            If StrComp(cells1(J), cells2(J)) <> 0 Then
                Pr "<tr><td style='width:50px; background-color: #FFFFCC' >"&excelCols(J+1)&"</td><td style='width:50px; background-color: #CCFFFF'>"&rowNum1&"</td>"
                Pr "<td style='width:250px;'>"&cells1(J)&"</td><td style='width:250px;'>"&cells2(J)&"</td>" 
                Pr "<td style='width:50px; background-color: #FFFFCC'>"&excelCols(J+1)&"</td><td style='width:50px; background-color: #CCFFFF'>"&rowNum2&"</td></tr>"
            End If
        Next
    End If
    I = I + 1
Loop
%>
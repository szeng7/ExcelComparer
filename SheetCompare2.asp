
<!DOCTYPE html>

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

const adopenforwardonly = 0
const adopenstatic = 3
const adlockreadonly = 1
const adlockpessimistic = 2
const adcmdtext = &H0001
const adcmdtable = &H0002

function rowsDifferent(values1, values2, rowNum1, rowNum2)
    Dim same, I
    same = False
    For I = 0 to Ubound(values1, 2)
        If StrComp(values1(rowNum1, I), values2(rowNum2, I)) <> 0 Then
            same = True
            End If
    Next
    rowsDifferent = same
    End Function

function IsBlank(Value)
'returns True if Empty or NULL or Zero
If IsEmpty(Value) or IsNull(Value) Then
 IsBlank = True
ElseIf VarType(Value) = vbString Then
 If Value = "" Then
  IsBlank = True
 End If
ElseIf IsObject(Value) Then
 If Value Is Nothing Then
  IsBlank = True
 End If
ElseIf IsNumeric(Value) Then
 If Value = 0 Then
  wscript.echo " Zero value found"
  IsBlank = True
 End If
Else
 IsBlank = False
End If
End Function

function getValues(sheetName, file, maxRows, maxFields)
    Dim values()
    Redim values(maxRows - 1, maxFields - 1)
    Dim value, I, J
    Dim CS1, RS1, SQ, columns, rows
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    rows = 0
    Do While Not RS1.EOF
        columns = 0
        For Each F in RS1.Fields
            If IsBlank(F) = True Then
                value = "{Empty}"
            Else
                value = RS1(F.Name)
            End If
            values(rows, columns) = value
            columns = columns + 1
            Next
        Do While columns < Ubound(values, 2) + 1
            values(rows, columns) = "{Empty}"
            columns = columns + 1
            Loop
        RS1.MoveNext
        rows = rows + 1
        Loop
    Do While rows < Ubound(values)
        For I = 0 to maxFields - 1
            values(rows, I) = "{Empty}"
        Next
        rows = rows + 1
    Loop
    RS1.Close
    getValues = values
    End Function


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

Dim I, numFields, numRows, file1name, file2name, values1, values2
'Number of rows to look ahead to decide if an insertion
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='1' style='background-color: black' ><tr><td style='background-color: #FFFFCC' ><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td>"&"<td style='background-color: #74A4BC'><b>"&Request.Form("file1")&" Contents</b></td><td style='background-color: #74A4BC'><b>"&Request.Form("file2")&" Contents</b></td>"&"<td style='background-color: #FFFFCC'><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td></tr>"
file1name = Request.Form("file1")
file2name = Request.Form("file2")
numFields = Request.Form("fields")
numRows = Request.Form("rows")
values1 = getValues(Request.Form("sheet"), file1name, numRows, numFields)
values2 = getValues(Request.Form("sheet"), file2name, numRows, numFields)
Dim J, insertion1, insertion2, rowNum1, rowNum2, offset1, offset2
I = 0
rowNum1 = 1
rowNum2 = 1
offset1 = 0
offset2 = 0
'I - counter, J - check if insertion'
'insertion1 = in first file, insertion2 = in second file'
Do While I + offset1 < numRows - 1 And I + offset2 < numRows - 1
    rowNum1 = rowNum1 + 1
    rowNum2 = rowNum2 + 1
    insertion1 = 0
    insertion2 = 0
    If rowsDifferent(values1, values2, I+offset1, I+offset2) Then
        For J = 1 to numRowsToCheck
            If I+offset1+J < numRows - 1 Then 
                If Not rowsDifferent(values1,values2, I + offset1 + J, I + offset2) Then
                    insertion1 = J
                End If
            End If
        Next
        If insertion1 = 0 Then 
            For J = 1 to numRowsToCheck
                If I + offset2 + J < numRows - 1 Then
                    If Not rowsDifferent(values1,values2, I + offset1, I + offset2 + J) Then
                        insertion2 = J
                    End If
                End If
            Next
        End If
    'Evidence of an insertion if insertion = 1
    End If
    If insertion1 > 0 Then
        For J = 0 to insertion1 - 1
            For A = 1 to Ubound(values1, 2)
                If A = 1 Then
                    Pr "<tr><td style='background-color: #909090'></td><td style='background-color: #909090'>"&rowNum1+J&"</td><td style='background-color: #909090'><i>Inserted Row</i></td><td style='background-color: #909090'></td><td style='background-color: #909090'></td><td style='background-color: #909090'></td></tr>"
                End If
            Next
        Next
        offset1 = offset1 + insertion1
        rowNum1 = rowNum1 + insertion1
    ElseIf insertion2 > 0 Then
        For J = 0 to insertion2 - 1
            For A = 1 to Ubound(values1, 2)
                If (A = 1) Then
                    Pr "<tr><td style='background-color: #909090'></td><td style='background-color: #909090'></td><td style='background-color: #909090'></td><td style='background-color: #909090'><i>Inserted Row</i></td><td style='background-color: #909090'></td><td style='background-color: #909090'>"&rowNum2+J&"</td></tr>"
                End If
            Next
        Next
        offset2 = offset2 + insertion2
        rowNum2 = rowNum2 + insertion2
    Else
        For J = 0 to Ubound(values1, 2)
            If StrComp(values1(I+offset1, J), values2(I+offset2, J)) <> 0 Then
                Pr "<tr><td style='background-color: #FFFFCC' >"&excelCols(J+1)&"</td><td style='background-color: #CCFFFF'>"&rowNum1&"</td>"
                Pr "<td>"&values1(I+offset1, J)&"</td><td>"&values2(I+offset2, J)&"</td>" 
                Pr "<td style='background-color: #FFFFCC'>"&excelCols(J+1)&"</td><td style='background-color: #CCFFFF'>"&rowNum2&"</td></tr>"
            End If
        Next
    End If
    I = I + 1
Loop
%>

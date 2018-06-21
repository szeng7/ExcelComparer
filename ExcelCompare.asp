<html>
<head>
</head>
<b>Excel File Comparison</b>
<Body>
<br>
<Form action='' method='post'>
Excel File 1: 
<input name='File1'>
Excel File 2: 
<input name='File2'>
<input type='submit' value='Compare'>
</Form>
<%
const adopenforwardonly = 0
const adopenstatic = 3
const adlockreadonly = 1
const adlockpessimistic = 2
const adcmdtext = &H0001
const adcmdtable = &H0002

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

'get values of all cells in a file as a string delimited by *'
function getValues(sheetName, file)
    Dim differences, value
    Dim CS1, RS1, SQ
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    Do While Not RS1.EOF 
        For Each F in RS1.Fields 
            If IsBlank(F) = True Then
                value = "{Empty}"
            Else
                value = RS1(F.Name)
            End If
            differences = differences & value & "*" 
            Next
        RS1.MoveNext
        Loop
        
    getValues = differences
    End Function

'get indexes of all cells in a file as a string delimited by *'
function getIndex(sheetName, file)
    Dim differences
    Dim CS1, RS1, SQ
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    dim lineNum
    lineNum=1
    Do While Not RS1.EOF 
        Dim fieldNum
        fieldNum=0 'column number'
        lineNum=lineNum+1
        For Each F in RS1.Fields 
            fieldNum=fieldNum+1
            differences = differences & lineNum & "*" & fieldNum & "*" 
            Next
        RS1.MoveNext
        Loop

    getIndex = differences
    End Function

Sub Pr(S)
    Response.Write S
    End Sub


Dim File1Sheets, File2Sheets
File1Sheets = ""
File2Sheets = ""
If Request.Form <> "" Then 

    Dim oConn1,sConn1,oConn2,sConn2
    Set oConn1 = Server.CreateObject("ADODB.Connection")
    Set oConn2 = Server.CreateObject("ADODB.Connection")
    sConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.Form("File1")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
    sConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.Form("File2")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
    oConn1.Open sConn1
    oConn2.Open sConn2
    Dim oRS1, oRS2
    Set oRS1 = oConn1.OpenSchema(20)
    Set oRS2 = oConn2.OpenSchema(20)
    Do While Not oRS1.EOF
        sSheetName = oRS1.Fields("table_name").Value
        File1Sheets = File1Sheets & sSheetName & ":"
        oRS1.MoveNext()
    Loop
    File1Sheets = StrReverse(File1Sheets)
    File1Sheets = StrReverse(Replace(File1Sheets,":","",1,1))
    Do While Not oRS2.EOF
        sSheetName = oRS2.Fields("table_name").Value
        File2Sheets = File2Sheets & sSheetName & ":"
        oRS2.MoveNext()
    Loop
    File2Sheets = StrReverse(File2Sheets)
    File2Sheets = StrReverse(Replace(File2Sheets,":","",1,1))
    Pr "<center><table border='1' cellspacing='0'>"
    Pr "<tr><td><b>" & Request.Form("File1") & " Sheets</b></td>"
    Pr "<td><b>" & Request.Form("File2") & " Sheets</b></td>"
    Pr "<td><b>Differences</b></td></tr>"

    For Each sheet in Split(File1Sheets,":")
        If Instr(File2Sheets, sheet) Then
            Dim values1, values2, index1, index2, valuesplit, valuesplit2, indexsplit, indexsplit2
            Dim filename1, filename2
            values1 = getValues(sheet, Request.Form("File1"))
            index1 = getIndex(sheet, Request.Form("File1"))
            values2 = getValues(sheet, Request.Form("File2"))
            index2 = getIndex(sheet, Request.Form("File2"))

            If Len(index1) > Len(index2) Then
                valuesplit = Split(values1, "*")
                indexsplit = Split(index1, "*") 'row, column, row, column'
                valuesplit2 = Split(values2, "*")
                indexsplit2 = Split(index2, "*") 'row, column, row, column'
                filename1 = Request.Form("file1")
                filename2 = Request.Form("file2")
            Else
                valuesplit = Split(values2, "*")
                indexsplit = Split(index2, "*") 'row, column, row, column'
                valuesplit2 = Split(values1, "*")
                indexsplit2 = Split(index1, "*") 'row, column, row, column'
                filename2 = Request.Form("file1")
                filename1 = Request.Form("file2")
            End If

            Dim I, J, K, L, X, Y, finaldiff, cellValue, cellValue2
            I = 0 'value1 counter'
            J = 0 'index1 counter'
            K = 0 'value2 counter'
            L = 0 'index2 counter'
            Do While I < Ubound(valuesplit)
                X=""
                Y=""
                cellValue=""
                cellValue2=""
                If (L+1 > Ubound(indexsplit2)) Then 'more info in first, empty in second'
                    cellValue = valuesplit(I)
                    cellValue2 = "{Empty}"
                    If StrComp(cellValue, "{Empty}") <> 0 Then 'if cell in first isn't empty'
                        X = indexsplit(J) 'row num'
                        Y = indexsplit(J+1) 'col num'
                        finaldiff = finaldiff & "(Row " & X & ", Column " & Y & "): " & cellValue & " vs " & cellValue2 & "\"
                        End If
                    I = I + 1
                    J = J + 2
                Else
                    'info in first and info in same corresponding spot in second'
                    If indexsplit(J) = indexsplit2(L) And indexsplit(J+1) = indexsplit2(L+1) Then
                        cellValue = valuesplit(I)
                        cellValue2 = valuesplit2(K)
                        If StrComp(cellValue, cellValue2) <> 0 Then 'if the values are diff'
                            X = indexsplit(J) 'row num'
                            Y = indexsplit(J+1) 'col num'
                            finaldiff = finaldiff & "(Row " & X & ", Column " & Y & "): " & cellValue & " vs " & cellValue2 & "\"
                            End If
                        I = I + 1
                        J = J + 2
                        K = K + 1
                        L = L + 2
                    Else
                        'info in first, no info in second/blank spot'
                        cellValue = valuesplit(I)
                        cellValue2 = "{Empty}"
                        If StrComp(cellValue, "{Empty}") <> 0 Then 'if cell in first isn't empty
                            X = indexsplit(J) 'row num'
                            Y = indexsplit(J+1) 'col num'
                            finaldiff = finaldiff & "(Row " & X & ", Column " & Y & "): " & cellValue & " vs " & cellValue2 & "\"
                        End If
                        I = I + 1
                        J = J + 2
                    End If
                End If
                Loop

            Pr "<tr>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td>" & sheet & "</td>"
            If Len(finaldiff) > 0 Then
                Pr "<td><Form action='sheetCompare.asp' method='post'>"
                Pr "<input type='hidden' name='finaldiff' value='"&finaldiff&"'>"
                Pr "<input type='hidden' name='sheet' value='"&sheet&"'>"
                Pr "<input type='hidden' name='file1' value='"&filename1&"'>"
                Pr "<input type='hidden' name='file2' value='"&filename2&"'>"
                Pr "<input type='submit' value='View Differences'>"
                Pr "</Form></td></tr>"
            Else
                Pr "<td></td></tr>"
                End If
            finaldiff=""
        Else 
            Pr "<tr>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td></td>"
            Pr "<td></td>"
            Pr "</tr>"
            End If
        Next
    For Each sheet in Split(File2Sheets,":")
        If Instr(File1Sheets, sheet) Then
            dim p
        Else
            Pr "<tr>"
            Pr "<td></td>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td></td>"
            Pr "</tr>"
            End If
        Next
    Pr "</center>"
    End If
%>

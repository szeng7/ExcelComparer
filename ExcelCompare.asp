<html>
<head>
</head>
<b>Excel File Comparison</b>
<Body>
<br>
<Form action='' method='get'>
Excel File 1: 
<input type='file' name='File1'>
Excel File 2: 
<input type='file' name='File2'>
<input type='submit' value='Compare'>
</Form>
<%
const adopenforwardonly = 0
const adopenstatic = 3
const adlockreadonly = 1
const adlockpessimistic = 2
const adcmdtext = &H0001
const adcmdtable = &H0002


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

function getColumnHeaders(sheetName, file, maxFields)
    Dim CS1, RS1, SQ, numFields
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    numFields = 0
    For Each F in RS1.Fields
        getColumnHeaders = getColumnHeaders & F.Name & "\"
        numFields = numFields + 1
    Next

    Do While StrComp(numFields, maxFields) <> 0
        getColumnHeaders = getColumnHeaders & "{No Field Name}\" 
        numFields = numFields + 1
    Loop

    End Function

'get number of columns/fields in file'
function numFields(sheetName, file)
    Dim CS1, RS1, SQ
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    numFields = RS1.Fields.Count
    End Function

'get number of rows in file'
function numRows(sheetName, file)
    Dim CS1, RS1, SQ, rows
    rows = 0
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext 
    Do While Not RS1.EOF
        rows = rows + 1
        RS1.MoveNext
        Loop
    numRows = rows
    End Function

'get values of all cells in a file as a string delimited by *'
function getValues(sheetName, file, maxRows, maxFields)
    Dim differences, value
    Dim CS1, RS1, SQ, columns, rows
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    rows = 0
    Do While Not RS1.EOF
        rows = rows + 1
        columns = 0
        For Each F in RS1.Fields
            columns = columns + 1
            If IsBlank(F) = True Then
                value = "{Empty}"
            Else
                value = RS1(F.Name)
            End If
            differences = differences & value & "*" 
            Next
        Do While columns < maxFields 
            differences = differences & "{Empty}*"
            columns = columns + 1
            Loop
        RS1.MoveNext
        Loop
    Do While rows < maxRows
        For I = 0 to maxFields
            differences = differences & "{Empty}*"
        Next
        rows = rows + 1
    Loop
    getValues = differences
    End Function

Sub Pr(S)
    Response.Write S
    End Sub


Dim File1Sheets, File2Sheets
File1Sheets = ""
File2Sheets = ""
If Request.QueryString.Count <> 0 Then 

    Dim oConn1,sConn1,oConn2,sConn2, weboutput, sheetDifferences
    Set oConn1 = Server.CreateObject("ADODB.Connection")
    Set oConn2 = Server.CreateObject("ADODB.Connection")
    sConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.QueryString("File1")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
    sConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.QueryString("File2")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
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
    weboutput = "<center><table border='1' cellspacing='0'>"
    weboutput = weboutput & "<tr><td><b>" & Request.QueryString("File1") & " Sheets</b></td>"
    weboutput = weboutput & "<td><b>" & Request.QueryString("File2") & " Sheets</b></td>"
    weboutput = weboutput & "<td></td></tr>"
    sheetDifferences = 0
    Dim colName
    For Each sheet in Split(File1Sheets,":")
        If Instr(File2Sheets, sheet) Then
            Dim values1, values2, valuesplit, valuesplit2, min
            Dim filename1, filename2, maxRows, maxFields, maxAttempt
            maxRows = numRows(sheet,Request.QueryString("File1")) 'number of rows in file1'
            maxAttempt = numRows(sheet, Request.QueryString("File2")) 'number of rows in file2'
            If maxAttempt > maxRows Then
                maxRows = maxAttempt
                End If
            maxFields = numFields(sheet, Request.QueryString("File1"))
            maxAttempt = numFields(sheet, Request.QueryString("File2"))
            If maxAttempt > maxFields Then
                min = maxFields
                maxFields = maxAttempt
                colName = getColumnHeaders(sheet,Request.QueryString("File1"), maxFields)
            Else
                min = maxAttempt
                colName = getColumnHeaders(sheet,Request.QueryString("File2"), maxFields)
            End If
            colNameSplit = Split(colName, "\")
            values1 = getValues(sheet, Request.QueryString("File1"), maxRows, maxFields)
            values2 = getValues(sheet, Request.QueryString("File2"), maxRows, maxFields)
            valuesplit = Split(values1, "*")
            valuesplit2 = Split(values2, "*")
            filename1 = Request.QueryString("file1")
            filename2 = Request.QueryString("file2")
            Dim I, J, finaldiff, cellValue, cellValue2, column, colRes
            finaldiff = ""
            For I=0 to maxRows-1
                For J=0 to maxFields - 1
                    cellValue = valuesplit(maxFields*I + J)
                    cellValue2 = valuesplit2(maxFields*I + J)
                    If StrComp(cellValue, cellValue2) <> 0 Then
                        finaldiff = finaldiff & " Column " & excelCols(J+1) & " - Row " & I+2 & "\" & cellValue & "\" & cellValue2 & "\"
                    End If
                    Next
                Next

            weboutput = weboutput & "<tr>"
            weboutput = weboutput & "<td>" & sheet & "</td>"
            weboutput = weboutput & "<td>" & sheet & "</td>"
            If Len(finaldiff) > 0 Then
                weboutput = weboutput & "<td><Form action='sheetCompare.asp' method='post'>"
                weboutput = weboutput & "<input type='hidden' name='finaldiff' value='"&finaldiff&"'>"
                weboutput = weboutput & "<input type='hidden' name='sheet' value='"&sheet&"'>"
                weboutput = weboutput & "<input type='hidden' name='file1' value='"&filename1&"'>"
                weboutput = weboutput & "<input type='hidden' name='file2' value='"&filename2&"'>"
                weboutput = weboutput & "<input type='submit' value='View Differences'>"
                weboutput = weboutput & "</Form></td></tr>"
                sheetDifferences = sheetDifferences + 1
            Else
                weboutput = weboutput & "<td>(No Differences)</td></tr>"
                End If
            finaldiff=""
        Else 
            weboutput = weboutput & "<tr>"
            weboutput = weboutput & "<td>" & sheet & "</td>"
            weboutput = weboutput & "<td></td>"
            weboutput = weboutput & "<td></td>"
            weboutput = weboutput & "</tr>"
            sheetDifferences = sheetDifferences + 1
            End If
        Next
    For Each sheet in Split(File2Sheets,":")
        If Instr(File1Sheets, sheet) Then
            dim p
        Else
            weboutput = weboutput & "<tr>"
            weboutput = weboutput & "<td></td>"
            weboutput = weboutput & "<td>" & sheet & "</td>"
            weboutput = weboutput & "<td></td>"
            weboutput = weboutput & "</tr>"
            sheetDifferences = sheetDifferences + 1
            End If
        Next
    weboutput = weboutput & "</center>"
    If sheetDifferences = 0 Then
        Pr "<center>FILES ARE IDENTICAL. NO DIFFERENCES TO SHOW.</center>"
    Else
        Pr weboutput
        End If
    End If
%>
